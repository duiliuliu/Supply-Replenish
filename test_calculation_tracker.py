# calculation_tracker.py 模块测试
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from calculation_tracker import (
    CalculationTracker,
    CalculationStep,
    StoreStageDetail,
    StoreCalculation,
    SKUTracking,
    STAGE_NAME_MAP,
    STAGE_RULES,
    STAGE_FORMULAS,
    format_store_detail_text,
    format_flow_overview_text
)


def test_calculation_step():
    """测试 CalculationStep 类"""
    print("\n=== 测试 CalculationStep 类 ===")

    step = CalculationStep('broken_size_fix', 1)
    assert step.stage_name == 'broken_size_fix'
    assert step.stage_display_name == '断码修复'
    assert step.stage_order == 1
    assert step.rules != ''
    assert len(step.formulas) > 0

    # 添加分配记录
    step.add_allocation('S001', 'SA', 'SKU001', 2, '断码修复(2)')
    step.add_allocation('S002', 'A', 'SKU001', 1, '断码修复(1)')

    assert step.total_allocated == 3
    assert step.store_count == 2

    # 测试剩余数量
    step.initial_remaining = 50
    step.remaining = 47
    assert step.remaining == 47

    print("  ✓ CalculationStep 类测试通过")


def test_store_stage_detail():
    """测试 StoreStageDetail 类"""
    print("\n=== 测试 StoreStageDetail 类 ===")

    detail = StoreStageDetail('sales_match', 'S001', 'SKU001')

    # 添加计算详情
    detail.add_detail('平均日需求', '6 ÷ 30', '0.2件/天', '计算日均销售需求')
    detail.add_detail('目标库存', '0.2 × 30 + 3', '9件', '科学计算的理想库存水平')

    assert len(detail.details) == 2
    assert detail.details[0]['name'] == '平均日需求'
    assert detail.details[0]['value'] == '0.2件/天'

    # 设置分配数量
    detail.set_allocated(5)
    assert detail.allocated == 5
    assert not detail.skipped

    # 测试跳过状态
    detail2 = StoreStageDetail('sell_through_priority', 'S001', 'SKU001')
    detail2.set_skipped('剩余数量为0')
    assert detail2.skipped
    assert detail2.skip_reason == '剩余数量为0'

    print("  ✓ StoreStageDetail 类测试通过")


def test_store_calculation():
    """测试 StoreCalculation 类"""
    print("\n=== 测试 StoreCalculation 类 ===")

    calc = StoreCalculation('CBBH', 'SA', 'SKU001')
    calc.initial_inventory = 0
    calc.sales_30d = 6
    calc.sell_through = 0.545

    # 添加阶段详情
    stage1 = StoreStageDetail('broken_size_fix', 'CBBH', 'SKU001')
    stage1.set_allocated(1)
    calc.add_stage_detail(stage1)

    stage2 = StoreStageDetail('sales_match', 'CBBH', 'SKU001')
    stage2.add_detail('目标库存', '0.2 × 30 + 3', '9件')
    stage2.set_allocated(4)
    calc.add_stage_detail(stage2)

    stage3 = StoreStageDetail('sell_through_priority', 'CBBH', 'SKU001')
    stage3.set_skipped('剩余数量为0')
    calc.add_stage_detail(stage3)

    assert calc.stage_count == 3
    assert calc.stages[0].allocated == 1
    assert calc.stages[1].allocated == 4
    assert calc.stages[2].skipped

    print("  ✓ StoreCalculation 类测试通过")


def test_calculation_tracker():
    """测试 CalculationTracker 类"""
    print("\n=== 测试 CalculationTracker 类 ===")

    tracker = CalculationTracker()

    # 设置配置快照
    config = {
        'coverage_days': {'SA': 30, 'A': 30, 'B': 14},
        'safety_factors': {'SA': 0.5, 'A': 0.4, 'B': 0.3}
    }
    tracker.set_config_snapshot(config)
    assert tracker.get_config_snapshot() == config

    # 开始追踪SKU
    tracker.start_sku('SKU001', 50)

    # 阶段1：断码修复
    tracker.start_stage('broken_size_fix', 50)

    # 记录分配
    tracker.record_allocation('S001', 'SA', 'SKU001', 2, '断码修复(2)')
    tracker.record_allocation('S002', 'A', 'SKU001', 1, '断码修复(1)')

    # 记录卖场详情
    tracker.record_store_detail(
        'S001', 'SA', 'SKU001', 'broken_size_fix',
        initial_inv=0, sales_30d=6, sell_through=0.5,
        detail_items=[
            {'name': '尺码类型', 'expression': '160', 'value': '核心尺码', 'explanation': '160/165为核心尺码'},
            {'name': '目标库存', 'expression': 'max(0, 2 - 0)', 'value': '2件', 'explanation': 'SA级核心尺码目标2件'},
            {'name': '分配数量', 'expression': 'min(2, 50)', 'value': '2件', 'explanation': '不超过剩余可用'}
        ],
        allocated=2
    )

    tracker.end_stage(47)  # 50 - 3 = 47

    # 阶段2：销量匹配
    tracker.start_stage('sales_match', 47)

    tracker.record_allocation('S001', 'SA', 'SKU001', 5, '销量匹配(5)')

    tracker.record_store_detail(
        'S001', 'SA', 'SKU001', 'sales_match',
        initial_inv=0, sales_30d=6, sell_through=0.5,
        detail_items=[
            {'name': '平均日需求', 'expression': '6 ÷ 30', 'value': '0.2件/天'},
            {'name': '覆盖周期', 'expression': 'SA级', 'value': '30天'},
            {'name': '安全系数', 'expression': 'SA级', 'value': '0.5'},
            {'name': '安全库存', 'expression': '0.2 × 0.5 × 30', 'value': '3件'},
            {'name': '目标库存', 'expression': '0.2 × 30 + 3', 'value': '9件'},
            {'name': '分配数量', 'expression': 'min(9 - 2, 47)', 'value': '5件'}
        ],
        allocated=5
    )

    tracker.end_stage(42)  # 47 - 5 = 42

    # 阶段3：销尽率优先
    tracker.start_stage('sell_through_priority', 42)

    tracker.record_store_detail(
        'S001', 'SA', 'SKU001', 'sell_through_priority',
        initial_inv=0, sales_30d=6, sell_through=0.5,
        detail_items=[
            {'name': '销尽率', 'expression': '6 ÷ (6 + 7)', 'value': '46.2%'},
            {'name': '等级权重', 'expression': 'SA级', 'value': '1.5'},
            {'name': '综合得分', 'expression': '46.2% × 1.5', 'value': '0.693'},
            {'name': '分配上限', 'expression': 'max(6 × 1.5, 2)', 'value': '9件'}
        ],
        allocated=0,
        skipped=True,
        skip_reason='剩余数量充足但当前库存已达上限'
    )

    tracker.end_stage(42)

    # 阶段4：剩余分配
    tracker.start_stage('remaining_allocation', 42)

    tracker.record_allocation('S003', 'B', 'SKU001', 10, '剩余分配(10)')
    tracker.end_stage(32)

    tracker.end_sku()

    # 验证追踪结果
    tracking = tracker.get_sku_tracking('SKU001')
    assert tracking is not None
    assert tracking.sku == 'SKU001'
    assert tracking.required_qty == 50
    assert len(tracking.stages) == 4

    # 验证各阶段
    assert tracking.stages[0].stage_name == 'broken_size_fix'
    assert tracking.stages[0].total_allocated == 3  # 2 + 1
    assert tracking.stages[0].initial_remaining == 50
    assert tracking.stages[0].remaining == 47

    assert tracking.stages[1].stage_name == 'sales_match'
    assert tracking.stages[1].total_allocated == 5

    assert tracking.stages[3].stage_name == 'remaining_allocation'
    assert tracking.stages[3].total_allocated == 10

    # 验证总分配量
    assert tracking.total_allocated == 18  # 3 + 5 + 0 + 10

    print("  ✓ CalculationTracker 类测试通过")


def test_get_store_detail():
    """测试获取卖场计算详情"""
    print("\n=== 测试获取卖场计算详情 ===")

    tracker = CalculationTracker()
    tracker.start_sku('SKU001', 50)

    tracker.start_stage('broken_size_fix', 50)
    tracker.record_store_detail(
        'S001', 'SA', 'SKU001', 'broken_size_fix',
        initial_inv=0, sales_30d=6, sell_through=0.5,
        detail_items=[
            {'name': '目标库存', 'expression': '2', 'value': '2件'}
        ],
        allocated=2
    )
    tracker.end_stage(48)
    tracker.end_sku()

    # 获取卖场详情
    store_calc = tracker.get_store_detail('S001', 'SKU001')
    assert store_calc is not None
    assert store_calc.store_code == 'S001'
    assert store_calc.level == 'SA'
    assert store_calc.initial_inventory == 0
    assert store_calc.sales_30d == 6
    assert store_calc.stage_count == 1
    assert store_calc.stages[0].allocated == 2
    assert store_calc.total_allocation == 2

    # 测试不存在的卖场
    store_calc = tracker.get_store_detail('S999', 'SKU001')
    assert store_calc is None

    print("  ✓ 获取卖场计算详情测试通过")


def test_get_stage_summary():
    """测试获取阶段摘要"""
    print("\n=== 测试获取阶段摘要 ===")

    tracker = CalculationTracker()
    tracker.start_sku('SKU001', 50)

    tracker.start_stage('broken_size_fix', 50)
    tracker.record_allocation('S001', 'SA', 'SKU001', 2, '断码修复(2)')
    tracker.record_allocation('S002', 'A', 'SKU001', 1, '断码修复(1)')
    tracker.end_stage(47)
    tracker.end_sku()

    summary = tracker.get_stage_summary('SKU001', 'broken_size_fix')
    assert summary is not None
    assert summary['stage_name'] == '断码修复'
    assert summary['total_allocated'] == 3
    assert summary['store_count'] == 2
    assert summary['initial_remaining'] == 50
    assert summary['remaining'] == 47
    assert len(summary['allocations']) == 2
    assert summary['rules'] != ''

    # 测试不存在的阶段
    summary = tracker.get_stage_summary('SKU001', 'nonexistent')
    assert summary is None

    print("  ✓ 获取阶段摘要测试通过")


def test_generate_flow_overview():
    """测试生成流程概览"""
    print("\n=== 测试生成流程概览 ===")

    tracker = CalculationTracker()
    tracker.start_sku('SKU001', 50)

    # 阶段1
    tracker.start_stage('broken_size_fix', 50)
    tracker.record_allocation('S001', 'SA', 'SKU001', 3, '断码修复(3)')
    tracker.end_stage(47)

    # 阶段2
    tracker.start_stage('sales_match', 47)
    tracker.record_allocation('S002', 'A', 'SKU001', 5, '销量匹配(5)')
    tracker.end_stage(42)

    # 阶段3
    tracker.start_stage('sell_through_priority', 42)
    tracker.end_stage(42)

    # 阶段4
    tracker.start_stage('remaining_allocation', 42)
    tracker.record_allocation('S003', 'B', 'SKU001', 10, '剩余分配(10)')
    tracker.end_stage(32)

    tracker.end_sku()

    overview = tracker.generate_flow_overview('SKU001')
    assert overview is not None
    assert overview['sku'] == 'SKU001'
    assert overview['required_qty'] == 50
    assert overview['total_allocated'] == 18  # 3 + 5 + 0 + 10
    assert len(overview['stages']) == 4

    # 验证阶段统计
    assert overview['stages'][0]['stage_name'] == '断码修复'
    assert overview['stages'][0]['allocated'] == 3
    assert overview['stages'][0]['percentage'] == 6.0  # 3/50 = 6%

    assert overview['stages'][1]['allocated'] == 5
    assert overview['stages'][1]['percentage'] == 10.0  # 5/50 = 10%

    assert overview['stages'][2]['allocated'] == 0
    assert overview['stages'][2]['percentage'] == 0.0

    assert overview['stages'][3]['allocated'] == 10
    assert overview['stages'][3]['percentage'] == 20.0  # 10/50 = 20%

    print("  ✓ 生成流程概览测试通过")


def test_format_store_detail_text():
    """测试格式化卖场计算详情文本"""
    print("\n=== 测试格式化卖场计算详情文本 ===")

    calc = StoreCalculation('CBBH', 'SA', 'SKU001')
    calc.initial_inventory = 0
    calc.sales_30d = 6
    calc.sell_through = 0.545

    stage1 = StoreStageDetail('broken_size_fix', 'CBBH', 'SKU001')
    stage1.add_detail('核心尺码', '否', '155', '非核心尺码')
    stage1.add_detail('目标', 'SA/A级非核心尺码 ≥ 1件', '1件')
    stage1.set_allocated(1)
    calc.add_stage_detail(stage1)

    stage2 = StoreStageDetail('sales_match', 'CBBH', 'SKU001')
    stage2.add_detail('平均日需求', '6 ÷ 30', '0.2件/天')
    stage2.add_detail('目标库存', '0.2 × 30 + 3', '9件')
    stage2.set_allocated(4)
    calc.add_stage_detail(stage2)

    stage3 = StoreStageDetail('sell_through_priority', 'CBBH', 'SKU001')
    stage3.set_skipped('剩余数量为0')
    calc.add_stage_detail(stage3)

    calc.total_allocation = 5
    calc.final_inventory = 5

    text = format_store_detail_text(calc)
    assert 'CBBH' in text
    assert '断码修复' in text
    assert '销量匹配' in text
    assert '销尽率优先' in text
    assert '跳过' in text or '剩余数量为0' in text
    assert '合计分配：5件' in text

    # 测试空记录
    text = format_store_detail_text(None)
    assert text == '无计算详情'

    print("  ✓ 格式化卖场计算详情文本测试通过")


def test_format_flow_overview_text():
    """测试格式化流程概览文本"""
    print("\n=== 测试格式化流程概览文本 ===")

    overview = {
        'sku': 'SKU001',
        'required_qty': 50,
        'total_allocated': 18,
        'stages': [
            {
                'stage_name': '断码修复',
                'stage_id': 'broken_size_fix',
                'order': 1,
                'allocated': 3,
                'percentage': 6.0,
                'store_count': 2,
                'remaining': 47,
                'rules': 'SA/A级卖场核心尺码≥2件'
            },
            {
                'stage_name': '销量匹配',
                'stage_id': 'sales_match',
                'order': 2,
                'allocated': 5,
                'percentage': 10.0,
                'store_count': 1,
                'remaining': 42,
                'rules': '目标库存 = 平均日需求 × 覆盖周期 + 安全库存'
            }
        ]
    }

    text = format_flow_overview_text(overview)
    assert 'SKU001' in text
    assert '断码修复' in text
    assert '销量匹配' in text
    assert '6.0%' in text or '6%' in text

    # 测试空数据
    text = format_flow_overview_text(None)
    assert text == '无流程概览数据'

    print("  ✓ 格式化流程概览文本测试通过")


def test_multiple_sku_tracking():
    """测试多SKU追踪"""
    print("\n=== 测试多SKU追踪 ===")

    tracker = CalculationTracker()

    # 追踪第一个SKU
    tracker.start_sku('SKU001', 50)
    tracker.start_stage('broken_size_fix', 50)
    tracker.record_allocation('S001', 'SA', 'SKU001', 2, '断码修复(2)')
    tracker.end_stage(48)
    tracker.end_sku()

    # 追踪第二个SKU
    tracker.start_sku('SKU002', 30)
    tracker.start_stage('broken_size_fix', 30)
    tracker.record_allocation('S002', 'A', 'SKU002', 1, '断码修复(1)')
    tracker.end_stage(29)
    tracker.end_sku()

    # 验证两个SKU都有追踪记录
    skus = tracker.get_all_sku_skus()
    assert len(skus) == 2
    assert 'SKU001' in skus
    assert 'SKU002' in skus

    tracking1 = tracker.get_sku_tracking('SKU001')
    tracking2 = tracker.get_sku_tracking('SKU002')
    assert tracking1.required_qty == 50
    assert tracking2.required_qty == 30

    print("  ✓ 多SKU追踪测试通过")


def run_all_tests():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("  calculation_tracker.py 模块测试")
    print("=" * 60)

    tests = [
        test_calculation_step,
        test_store_stage_detail,
        test_store_calculation,
        test_calculation_tracker,
        test_get_store_detail,
        test_get_stage_summary,
        test_generate_flow_overview,
        test_format_store_detail_text,
        test_format_flow_overview_text,
        test_multiple_sku_tracking
    ]

    passed = 0
    failed = 0

    for test in tests:
        try:
            test()
            passed += 1
        except Exception as e:
            failed += 1
            print(f"  ✗ {test.__name__} 失败: {e}")
            import traceback
            traceback.print_exc()

    print("\n" + "=" * 60)
    print(f"  测试结果: {passed} 通过, {failed} 失败, 共 {passed + failed} 项")
    print("=" * 60)

    return failed == 0


if __name__ == '__main__':
    success = run_all_tests()
    sys.exit(0 if success else 1)
