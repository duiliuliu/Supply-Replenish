# 计算追踪模块 v1.0
# 记录分配过程中的每一步计算，提供计算细节查询接口，生成流程可视化数据
from collections import defaultdict
import copy


# 阶段名称映射
STAGE_NAME_MAP = {
    'broken_size_fix': '断码修复',
    'sales_match': '销量匹配',
    'sell_through_priority': '销尽率优先',
    'remaining_allocation': '剩余分配'
}

# 阶段规则说明
STAGE_RULES = {
    'broken_size_fix': 'SA/A级卖场：核心尺码(160/165)至少2件，非核心尺码至少1件；其他等级卖场：核心尺码至少1件',
    'sales_match': '目标库存 = 平均日需求 × 覆盖周期 + 安全库存，其中安全库存 = 平均日需求 × 安全系数 × 覆盖周期',
    'sell_through_priority': '综合得分 = 销尽率 × 等级权重，按得分降序分配，分配上限 = max(30天销量 × 等级权重, 2)',
    'remaining_allocation': '按等级顺序分配：SA → A → B → C → D → OL，单卖场上限10件'
}

# 阶段公式模板
STAGE_FORMULAS = {
    'broken_size_fix': [
        {'name': '目标库存', 'expression': 'max(0, 目标 - 当前库存)', 'explanation': '根据卖场等级和尺码类型确定最低库存目标'},
        {'name': '分配数量', 'expression': 'min(需分配, 剩余可用)', 'explanation': '确保不超过剩余可分配数量'}
    ],
    'sales_match': [
        {'name': '平均日需求', 'expression': '30天销量 ÷ 30', 'explanation': '计算日均销售需求'},
        {'name': '安全库存', 'expression': '平均日需求 × 安全系数 × 覆盖周期', 'explanation': '为防止缺货准备的缓冲库存'},
        {'name': '目标库存', 'expression': 'max(平均日需求 × 覆盖周期 + 安全库存, 最小目标库存)', 'explanation': '科学计算的理想库存水平'},
        {'name': '分配数量', 'expression': 'min(目标库存 - 当前库存, 剩余可用)', 'explanation': '补足至目标库存'}
    ],
    'sell_through_priority': [
        {'name': '销尽率', 'expression': '30天销量 ÷ (30天销量 + 当前库存)', 'explanation': '表示库存周转快慢'},
        {'name': '综合得分', 'expression': '销尽率 × 等级权重', 'explanation': '综合考虑销尽率和卖场等级'},
        {'name': '分配上限', 'expression': 'max(30天销量 × 等级权重, 2)', 'explanation': '单个卖场在此阶段的最大分配量'},
        {'name': '分配数量', 'expression': 'min(分配上限 - 当前库存, 剩余可用)', 'explanation': '不超过分配上限'}
    ],
    'remaining_allocation': [
        {'name': '分配上限', 'expression': '单卖场上限 = 10件', 'explanation': '剩余分配阶段的固定上限'},
        {'name': '分配数量', 'expression': 'min(10 - 当前库存, 剩余可用)', 'explanation': '按等级顺序依次分配'}
    ]
}


class CalculationStep:
    """单个阶段的计算步骤记录"""

    def __init__(self, stage_name, stage_order):
        self.stage_name = stage_name          # 阶段ID
        self.stage_display_name = STAGE_NAME_MAP.get(stage_name, stage_name)  # 阶段显示名称
        self.stage_order = stage_order        # 阶段顺序
        self.rules = STAGE_RULES.get(stage_name, '')   # 阶段规则
        self.formulas = STAGE_FORMULAS.get(stage_name, [])  # 阶段公式
        self.allocation_results = []          # 分配结果列表
        self.initial_remaining = 0            # 阶段开始时的剩余数量
        self.remaining = 0                    # 阶段结束时的剩余数量

    def add_allocation(self, store_code, level, sku, quantity, reason):
        """添加一条分配记录"""
        self.allocation_results.append({
            "store_code": store_code,
            "level": level,
            "sku": sku,
            "quantity": quantity,
            "reason": reason
        })

    @property
    def total_allocated(self):
        """本阶段总分配量"""
        return sum(r['quantity'] for r in self.allocation_results)

    @property
    def store_count(self):
        """本阶段参与分配的卖场数"""
        return len(self.allocation_results)


class StoreStageDetail:
    """单个卖场在某阶段的计算详情"""

    def __init__(self, stage_name, store_code, sku):
        self.stage_name = stage_name
        self.stage_display_name = STAGE_NAME_MAP.get(stage_name, stage_name)
        self.store_code = store_code
        self.sku = sku
        self.details = []        # 计算详情项列表
        self.allocated = 0       # 本阶段分配数量
        self.skipped = False     # 是否跳过
        self.skip_reason = ""    # 跳过原因

    def add_detail(self, name, expression, value, explanation=""):
        """添加一条计算详情"""
        self.details.append({
            "name": name,
            "expression": expression,
            "value": value,
            "explanation": explanation
        })

    def set_allocated(self, quantity):
        """设置分配数量"""
        self.allocated = quantity

    def set_skipped(self, reason):
        """设置跳过状态"""
        self.skipped = True
        self.skip_reason = reason


class StoreCalculation:
    """单个卖场SKU的完整计算记录"""

    def __init__(self, store_code, level, sku):
        self.store_code = store_code
        self.level = level
        self.sku = sku
        self.stages = []             # 各阶段详情
        self.total_allocation = 0    # 总分配量
        self.initial_inventory = 0   # 初始库存
        self.final_inventory = 0     # 最终库存
        self.sales_30d = 0          # 30天销量
        self.sell_through = 0       # 销尽率

    def add_stage_detail(self, stage_detail):
        """添加阶段详情"""
        self.stages.append(stage_detail)

    @property
    def stage_count(self):
        return len(self.stages)


class SKUTracking:
    """单个SKU的完整分配追踪"""

    def __init__(self, sku, required_qty):
        self.sku = sku
        self.required_qty = required_qty       # 需分配数量
        self.total_allocated = 0               # 总已分配数量
        self.stages = []                       # 各阶段计算步骤
        self.store_calculations = {}           # 卖场计算详情 {store_code: StoreCalculation}

    def add_stage(self, stage):
        """添加阶段计算步骤"""
        self.stages.append(stage)

    def get_or_create_store_calc(self, store_code, level, sku, initial_inv=0, sales_30d=0, sell_through=0):
        """获取或创建卖场计算记录"""
        if store_code not in self.store_calculations:
            calc = StoreCalculation(store_code, level, sku)
            calc.initial_inventory = initial_inv
            calc.sales_30d = sales_30d
            calc.sell_through = sell_through
            self.store_calculations[store_code] = calc
        return self.store_calculations[store_code]

    def finalize(self):
        """完成追踪，计算最终值"""
        self.total_allocated = 0
        for stage in self.stages:
            self.total_allocated += stage.total_allocated
            for store_code, calc in self.store_calculations.items():
                total = sum(s.allocated for s in calc.stages)
                calc.total_allocation = total
                calc.final_inventory = calc.initial_inventory + total


class CalculationTracker:
    """
    计算追踪器，记录整个分配过程

    用法：
        tracker = CalculationTracker()
        tracker.start_sku(sku, required_qty)
        tracker.start_stage('broken_size_fix', remaining_qty)
        # ... 在分配过程中记录详情 ...
        tracker.end_stage(remaining_qty)
        tracker.end_sku()
    """

    def __init__(self):
        self.sku_trackings = {}           # {sku: SKUTracking}
        self.current_sku = None            # 当前追踪的SKU
        self.current_stage = None          # 当前阶段
        self.current_stage_order = 0       # 当前阶段顺序
        self.config_snapshot = {}          # 配置快照

    def set_config_snapshot(self, config):
        """保存配置快照，用于参数影响展示"""
        self.config_snapshot = copy.deepcopy(config)

    def start_sku(self, sku, required_qty):
        """开始追踪一个SKU"""
        self.current_sku = sku
        tracking = SKUTracking(sku, required_qty)
        self.sku_trackings[sku] = tracking
        self.current_stage_order = 0

    def end_sku(self):
        """结束当前SKU追踪"""
        if self.current_sku in self.sku_trackings:
            self.sku_trackings[self.current_sku].finalize()
        self.current_sku = None
        self.current_stage = None

    def start_stage(self, stage_name, remaining_qty):
        """开始一个新阶段"""
        self.current_stage_order += 1
        stage = CalculationStep(stage_name, self.current_stage_order)
        stage.initial_remaining = remaining_qty
        self.current_stage = stage

        if self.current_sku and self.current_sku in self.sku_trackings:
            self.sku_trackings[self.current_sku].add_stage(stage)

    def end_stage(self, remaining_qty):
        """结束当前阶段"""
        if self.current_stage:
            self.current_stage.remaining = remaining_qty
        self.current_stage = None

    def record_allocation(self, store_code, level, sku, quantity, reason):
        """记录一条分配结果"""
        if self.current_stage:
            self.current_stage.add_allocation(store_code, level, sku, quantity, reason)

    def record_store_detail(self, store_code, level, sku, stage_name,
                            initial_inv, sales_30d, sell_through,
                            detail_items, allocated, skipped=False, skip_reason=""):
        """
        记录卖场在某阶段的计算详情

        Args:
            detail_items: list of dict - 详情项 [{'name':..., 'expression':..., 'value':..., 'explanation':...}]
        """
        if self.current_sku and self.current_sku in self.sku_trackings:
            tracking = self.sku_trackings[self.current_sku]
            store_calc = tracking.get_or_create_store_calc(
                store_code, level, sku, initial_inv, sales_30d, sell_through
            )

            stage_detail = StoreStageDetail(stage_name, store_code, sku)
            for item in detail_items:
                stage_detail.add_detail(
                    item.get('name', ''),
                    item.get('expression', ''),
                    item.get('value', ''),
                    item.get('explanation', '')
                )

            if skipped:
                stage_detail.set_skipped(skip_reason)
            else:
                stage_detail.set_allocated(allocated)

            store_calc.add_stage_detail(stage_detail)

    def get_sku_tracking(self, sku):
        """获取SKU的追踪记录"""
        return self.sku_trackings.get(sku)

    def get_store_detail(self, store_code, sku):
        """获取单个卖场SKU的计算详情"""
        tracking = self.sku_trackings.get(sku)
        if tracking:
            return tracking.store_calculations.get(store_code)
        return None

    def get_stage_summary(self, sku, stage_name):
        """
        获取指定SKU在某阶段的统计摘要

        Returns:
            dict with keys: stage_name, rules, total_allocated, store_count, remaining, allocations
        """
        tracking = self.sku_trackings.get(sku)
        if not tracking:
            return None

        for stage in tracking.stages:
            if stage.stage_name == stage_name:
                return {
                    'stage_name': stage.stage_display_name,
                    'rules': stage.rules,
                    'formulas': stage.formulas,
                    'total_allocated': stage.total_allocated,
                    'store_count': stage.store_count,
                    'initial_remaining': stage.initial_remaining,
                    'remaining': stage.remaining,
                    'allocations': stage.allocation_results
                }
        return None

    def generate_flow_overview(self, sku):
        """
        生成SKU的流程概览数据

        Returns:
            dict with keys: sku, required_qty, total_allocated, stages
        """
        tracking = self.sku_trackings.get(sku)
        if not tracking:
            return None

        stages_overview = []
        for stage in tracking.stages:
            allocated = stage.total_allocated
            percentage = 0
            if tracking.required_qty > 0:
                percentage = round(allocated / tracking.required_qty * 100, 1)
            stages_overview.append({
                'stage_name': stage.stage_display_name,
                'stage_id': stage.stage_name,
                'order': stage.stage_order,
                'allocated': allocated,
                'percentage': percentage,
                'store_count': stage.store_count,
                'remaining': stage.remaining,
                'rules': stage.rules
            })

        return {
            'sku': sku,
            'required_qty': tracking.required_qty,
            'total_allocated': tracking.total_allocated,
            'stages': stages_overview
        }

    def get_all_sku_skus(self):
        """获取所有已追踪的SKU列表"""
        return list(self.sku_trackings.keys())

    def get_config_snapshot(self):
        """获取配置快照"""
        return self.config_snapshot


def format_store_detail_text(store_calc, config=None):
    """
    格式化卖场计算详情为可显示的文本

    Args:
        store_calc: StoreCalculation - 卖场计算记录
        config: dict - 配置快照（可选）

    Returns:
        str - 格式化的文本
    """
    if not store_calc:
        return "无计算详情"

    lines = []
    lines.append(f"卖场：{store_calc.store_code}  |  等级：{store_calc.level}  |  SKU：{store_calc.sku}")
    lines.append(f"初始库存：{store_calc.initial_inventory}件  |  30天销量：{store_calc.sales_30d}件  |  销尽率：{store_calc.sell_through:.1%}")
    lines.append("")

    for stage in store_calc.stages:
        lines.append(f"── {stage.stage_display_name} ──")

        if stage.skipped:
            lines.append(f"  跳过原因：{stage.skip_reason}")
        else:
            for detail in stage.details:
                lines.append(f"  {detail['name']} = {detail['expression']} = {detail['value']}")
                if detail.get('explanation'):
                    lines.append(f"    ({detail['explanation']})")
            lines.append(f"  → 已分配：{stage.allocated}件")

        lines.append("")

    lines.append(f"合计分配：{store_calc.total_allocation}件")
    lines.append(f"最终库存：{store_calc.final_inventory}件")

    return '\n'.join(lines)


def format_flow_overview_text(overview):
    """
    格式化流程概览为可显示的文本

    Args:
        overview: dict - generate_flow_overview的返回值

    Returns:
        str - 格式化的文本
    """
    if not overview:
        return "无流程概览数据"

    lines = []
    lines.append(f"SKU: {overview['sku']}  |  需分配: {overview['required_qty']}件  |  已分配: {overview['total_allocated']}件")

    if overview['required_qty'] > 0:
        completion = round(overview['total_allocated'] / overview['required_qty'] * 100, 1)
        lines.append(f"完成度: {completion}%")
    lines.append("")

    for stage in overview['stages']:
        status = "✓" if stage['allocated'] > 0 else "○"
        lines.append(f"{status} 阶段{stage['order']}：{stage['stage_name']}")
        lines.append(f"   分配：{stage['allocated']}件 ({stage['percentage']}%) | 卖场数：{stage['store_count']} | 剩余：{stage['remaining']}件")
        if stage['rules']:
            lines.append(f"   规则：{stage['rules']}")
        lines.append("")

    return '\n'.join(lines)
