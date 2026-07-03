# 数据验证模块 v1.0
# 提供Excel模板生成、数据格式验证、智能纠错建议、验证报告生成
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime


# 必需的工作表名称
REQUIRED_SHEETS = {
    'inventory': '库存',
    'sales': '销售',
    'store_level': '卖场等级',
    'add_order': '加单数量'
}

# 各工作表必需的列名
REQUIRED_COLUMNS = {
    '库存': ['卖场代码', '条码', '库存数量'],
    '销售': ['条码.条码', '店仓.卖场代码', '数量'],
    '卖场等级': ['代码', '卖场等级'],
    '加单数量': ['SKU', 'SKC', '需分配数量']
}

# 列名模糊匹配映射（常见错误列名 -> 标准列名）
COLUMN_FUZZY_MAP = {
    '库存': {
        '卖场代码': ['卖场', '店铺代码', '店仓代码', '门店代码', 'store', 'store_code', '卖场编码'],
        '条码': ['SKU', '商品条码', '条形码', 'barcode', '商品代码', '货号'],
        '库存数量': ['库存', '数量', '库存数', 'stock', '现有库存', '当前库存']
    },
    '销售': {
        '条码.条码': ['条码', 'SKU', '商品条码', 'barcode', '商品代码', '货号'],
        '店仓.卖场代码': ['卖场代码', '卖场', '店铺代码', '店仓代码', '门店代码', 'store', '卖场编码'],
        '数量': ['销售数量', '销量', '销售数', 'sales', '销售件数']
    },
    '卖场等级': {
        '代码': ['卖场代码', '卖场', '店铺代码', '店仓代码', '门店代码', 'store', '卖场编码'],
        '卖场等级': ['等级', '级别', 'level', '店铺等级', '卖场级别']
    },
    '加单数量': {
        'SKU': ['条码', '商品条码', '条形码', 'barcode', '商品代码', '货号'],
        'SKC': ['SKC代码', '分类代码', '颜色代码', 'color_code'],
        '需分配数量': ['分配数量', '加单数量', '需求数量', '数量', 'qty', '需分配']
    }
}

# 有效的卖场等级
VALID_LEVELS = ['SA', 'A', 'B', 'C', 'D', 'OL']

# 等级修复映射（常见错误值 -> 标准值）
LEVEL_FIX_MAP = {
    'A级': 'A', 'a级': 'A', 'a': 'A', 'SA级': 'SA', 'sa': 'SA',
    'b级': 'B', 'b': 'B', 'B级': 'B',
    'c级': 'C', 'c': 'C', 'C级': 'C',
    'd级': 'D', 'd': 'D', 'D级': 'D',
    'ol级': 'OL', 'ol': 'OL', 'OL级': 'OL', '线上': 'OL'
}


class ValidationResult:
    """验证结果类，存储验证过程中的通过项、警告和错误"""

    def __init__(self):
        self.passed = []      # 通过的验证项
        self.warnings = []    # 警告项
        self.errors = []      # 错误项
        self.fixed = []       # 已自动修复的项

    def add_passed(self, sheet_name, message):
        """添加通过的验证项"""
        self.passed.append({"sheet": sheet_name, "message": message})

    def add_warning(self, sheet_name, row, column, message, fixable=False, fix_action=None):
        """添加警告项"""
        self.warnings.append({
            "sheet": sheet_name,
            "row": row,
            "column": column,
            "message": message,
            "fixable": fixable,
            "fix_action": fix_action
        })

    def add_error(self, sheet_name, row, column, message, fixable=False, fix_action=None):
        """添加错误项"""
        self.errors.append({
            "sheet": sheet_name,
            "row": row,
            "column": column,
            "message": message,
            "fixable": fixable,
            "fix_action": fix_action
        })

    def add_fixed(self, sheet_name, row, column, original, fixed, message):
        """添加已修复的项"""
        self.fixed.append({
            "sheet": sheet_name,
            "row": row,
            "column": column,
            "original": original,
            "fixed": fixed,
            "message": message
        })

    @property
    def has_errors(self):
        """是否有错误"""
        return len(self.errors) > 0

    @property
    def has_warnings(self):
        """是否有警告"""
        return len(self.warnings) > 0

    @property
    def is_passed(self):
        """是否通过验证（无错误）"""
        return not self.has_errors

    def get_summary(self):
        """获取验证摘要"""
        return {
            "passed": len(self.passed),
            "warnings": len(self.warnings),
            "errors": len(self.errors),
            "fixed": len(self.fixed)
        }

    def get_report_by_sheet(self):
        """按工作表分组获取验证报告"""
        report = {}
        for sheet in REQUIRED_SHEETS.values():
            report[sheet] = {
                "passed": [p for p in self.passed if p["sheet"] == sheet],
                "warnings": [w for w in self.warnings if w["sheet"] == sheet],
                "errors": [e for e in self.errors if e["sheet"] == sheet],
                "fixed": [f for f in self.fixed if f["sheet"] == sheet]
            }
        return report


def generate_template(output_path):
    """
    生成标准格式的空白Excel模板

    Args:
        output_path: str - 输出文件路径

    Returns:
        bool - 是否生成成功
    """
    try:
        wb = Workbook()

        # 定义样式
        header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
        required_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
        example_fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        example_font = Font(color="9CA3AF", italic=True, size=10)
        thin_border = Border(
            left=Side(style='thin', color='E5E7EB'),
            right=Side(style='thin', color='E5E7EB'),
            top=Side(style='thin', color='E5E7EB'),
            bottom=Side(style='thin', color='E5E7EB')
        )

        # 1. 库存表
        ws_inv = wb.active
        ws_inv.title = '库存'
        inv_headers = ['卖场代码', '条码', '库存数量']
        for col, header in enumerate(inv_headers, 1):
            cell = ws_inv.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            # 必填列标记浅黄色背景
            for row in range(2, 5):
                ws_inv.cell(row=row, column=col).fill = required_fill
        # 示例数据
        inv_examples = [['S001', 'PRKWG2303M99160', 10], ['S002', 'PRKWG2303M99160', 5]]
        for row_idx, row_data in enumerate(inv_examples, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_inv.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = example_fill
                cell.font = example_font
                cell.border = thin_border
        # 数值验证
        dv_inv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1=0, allow_blank=True)
        ws_inv.add_data_validation(dv_inv)
        dv_inv.add(f'C2:C1000')
        # 列宽
        for col in range(1, 4):
            ws_inv.column_dimensions[get_column_letter(col)].width = 18

        # 2. 销售表
        ws_sales = wb.create_sheet('销售')
        sales_headers = ['条码.条码', '店仓.卖场代码', '数量']
        for col, header in enumerate(sales_headers, 1):
            cell = ws_sales.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            for row in range(2, 5):
                ws_sales.cell(row=row, column=col).fill = required_fill
        sales_examples = [['PRKWG2303M99160', 'S001', 30], ['PRKWG2303M99160', 'S002', 15]]
        for row_idx, row_data in enumerate(sales_examples, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_sales.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = example_fill
                cell.font = example_font
                cell.border = thin_border
        dv_sales = DataValidation(type="whole", operator="greaterThanOrEqual", formula1=0, allow_blank=True)
        ws_sales.add_data_validation(dv_sales)
        dv_sales.add(f'C2:C1000')
        for col in range(1, 4):
            ws_sales.column_dimensions[get_column_letter(col)].width = 20

        # 3. 卖场等级表
        ws_level = wb.create_sheet('卖场等级')
        level_headers = ['代码', '卖场等级']
        for col, header in enumerate(level_headers, 1):
            cell = ws_level.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            for row in range(2, 5):
                ws_level.cell(row=row, column=col).fill = required_fill
        level_examples = [['S001', 'SA'], ['S002', 'A']]
        for row_idx, row_data in enumerate(level_examples, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_level.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = example_fill
                cell.font = example_font
                cell.border = thin_border
        # 等级下拉选择
        dv_level = DataValidation(type="list", formula1='"SA,A,B,C,D,OL"', allow_blank=True)
        ws_level.add_data_validation(dv_level)
        dv_level.add(f'B2:B1000')
        for col in range(1, 3):
            ws_level.column_dimensions[get_column_letter(col)].width = 15

        # 4. 加单数量表
        ws_order = wb.create_sheet('加单数量')
        order_headers = ['SKU', 'SKC', '需分配数量']
        for col, header in enumerate(order_headers, 1):
            cell = ws_order.cell(row=1, column=col, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = thin_border
            for row in range(2, 5):
                ws_order.cell(row=row, column=col).fill = required_fill
        order_examples = [['PRKWG2303M99160', 'PRKWG2303M991', 50], ['PRKWG2303M99165', 'PRKWG2303M991', 30]]
        for row_idx, row_data in enumerate(order_examples, 2):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_order.cell(row=row_idx, column=col_idx, value=value)
                cell.fill = example_fill
                cell.font = example_font
                cell.border = thin_border
        dv_order = DataValidation(type="whole", operator="greaterThanOrEqual", formula1=1, allow_blank=True)
        ws_order.add_data_validation(dv_order)
        dv_order.add(f'C2:C1000')
        for col in range(1, 4):
            ws_order.column_dimensions[get_column_letter(col)].width = 22

        # 5. 使用说明Sheet
        ws_guide = wb.create_sheet('使用说明')
        guide_content = [
            ['加单商品分配系统 - Excel模板使用说明', ''],
            ['', ''],
            ['生成日期:', datetime.now().strftime('%Y-%m-%d')],
            ['', ''],
            ['工作表说明:', ''],
            ['库存', '包含卖场代码、条码、库存数量三列'],
            ['销售', '包含条码.条码、店仓.卖场代码、数量三列，数量为近30天销售总量'],
            ['卖场等级', '包含代码、卖场等级两列，等级取值：SA/A/B/C/D/OL'],
            ['加单数量', '包含SKU、SKC、需分配数量三列'],
            ['', ''],
            ['注意事项:', ''],
            ['1', '请勿修改工作表名称和列名'],
            ['2', '黄色背景为必填列，灰色字体为示例数据（可删除）'],
            ['3', '数值列请输入非负整数'],
            ['4', '卖场等级请使用下拉选择（SA/A/B/C/D/OL）'],
            ['5', '销售数量应为近30天总销量，而非平均日销量'],
            ['6', '加单的需分配数量应≥1'],
        ]
        for row_idx, row_data in enumerate(guide_content, 1):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws_guide.cell(row=row_idx, column=col_idx, value=value)
                if row_idx == 1:
                    cell.font = Font(bold=True, size=14, color="2563EB")
                elif col_idx == 1 and value:
                    cell.font = Font(bold=True, size=11)
                else:
                    cell.font = Font(size=10)
        ws_guide.column_dimensions['A'].width = 20
        ws_guide.column_dimensions['B'].width = 60

        wb.save(output_path)
        print(f"模板已生成: {output_path}")
        return True
    except Exception as e:
        print(f"生成模板失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def _fuzzy_match_column(actual_columns, required_column, sheet_name):
    """
    模糊匹配列名

    Args:
        actual_columns: list - 实际的列名列表
        required_column: str - 需要匹配的标准列名
        sheet_name: str - 工作表名称

    Returns:
        str or None - 匹配到的实际列名，未匹配返回None
    """
    if required_column in actual_columns:
        return required_column

    fuzzy_map = COLUMN_FUZZY_MAP.get(sheet_name, {}).get(required_column, [])
    for candidate in fuzzy_map:
        for actual in actual_columns:
            if actual == candidate or actual.lower() == candidate.lower():
                return actual
            # 包含关系匹配
            if candidate.lower() in actual.lower() or actual.lower() in candidate.lower():
                return actual
    return None


def validate_sheet(df, sheet_type):
    """
    验证单个工作表

    Args:
        df: DataFrame - 工作表数据
        sheet_type: str - 工作表类型（库存/销售/卖场等级/加单数量）

    Returns:
        ValidationResult - 验证结果
    """
    result = ValidationResult()

    if df is None or df.empty:
        result.add_error(sheet_type, 0, "", f"工作表'{sheet_type}'为空或不存在", fixable=False)
        return result

    actual_columns = list(df.columns)
    required_columns = REQUIRED_COLUMNS.get(sheet_type, [])

    # 第一级：结构验证（列名）
    column_mapping = {}
    for req_col in required_columns:
        matched = _fuzzy_match_column(actual_columns, req_col, sheet_type)
        if matched:
            if matched != req_col:
                result.add_warning(
                    sheet_type, 0, req_col,
                    f"列名'{matched}'与标准列名'{req_col}'不完全匹配，已自动映射",
                    fixable=True,
                    fix_action=f"rename:{matched}->{req_col}"
                )
            column_mapping[req_col] = matched
        else:
            result.add_error(
                sheet_type, 0, req_col,
                f"缺少必需列: '{req_col}'",
                fixable=False
            )

    # 如果有缺失列，无法继续验证
    if len(column_mapping) < len(required_columns):
        return result

    if len(column_mapping) == len(required_columns):
        result.add_passed(sheet_type, "列名匹配通过")

    # 第二级：数据类型验证
    for index, row in df.iterrows():
        row_num = index + 2  # Excel行号（从2开始，1是表头）

        if sheet_type == '库存':
            _validate_inventory_row(result, row, column_mapping, row_num)
        elif sheet_type == '销售':
            _validate_sales_row(result, row, column_mapping, row_num)
        elif sheet_type == '卖场等级':
            _validate_store_level_row(result, row, column_mapping, row_num)
        elif sheet_type == '加单数量':
            _validate_add_order_row(result, row, column_mapping, row_num)

    return result


def _validate_inventory_row(result, row, column_mapping, row_num):
    """验证库存表数据行"""
    sheet = '库存'
    store_col = column_mapping['卖场代码']
    barcode_col = column_mapping['条码']
    qty_col = column_mapping['库存数量']

    # 卖场代码不能为空
    store_val = row[store_col]
    if pd.isna(store_val) or str(store_val).strip() == '':
        result.add_error(sheet, row_num, '卖场代码', "卖场代码为空", fixable=False)

    # 条码不能为空
    barcode_val = row[barcode_col]
    if pd.isna(barcode_val) or str(barcode_val).strip() == '':
        result.add_error(sheet, row_num, '条码', "条码为空", fixable=False)

    # 库存数量验证
    qty_val = row[qty_col]
    if pd.isna(qty_val):
        result.add_error(sheet, row_num, '库存数量', "库存数量为空", fixable=False)
    else:
        try:
            qty_num = float(qty_val)
            if qty_num < 0:
                result.add_error(
                    sheet, row_num, '库存数量',
                    f"库存数量为负数 ({qty_num})",
                    fixable=True,
                    fix_action=f"set_zero:{qty_col}"
                )
            elif qty_num != int(qty_num):
                result.add_warning(
                    sheet, row_num, '库存数量',
                    f"库存数量非整数 ({qty_num})",
                    fixable=True,
                    fix_action=f"to_int:{qty_col}"
                )
            elif qty_num > 9999:
                result.add_warning(
                    sheet, row_num, '库存数量',
                    f"库存数量异常大 ({int(qty_num)})",
                    fixable=False
                )
        except (ValueError, TypeError):
            result.add_error(
                sheet, row_num, '库存数量',
                f"库存数量不是有效数字: '{qty_val}'",
                fixable=False
            )


def _validate_sales_row(result, row, column_mapping, row_num):
    """验证销售表数据行"""
    sheet = '销售'
    barcode_col = column_mapping['条码.条码']
    store_col = column_mapping['店仓.卖场代码']
    qty_col = column_mapping['数量']

    barcode_val = row[barcode_col]
    if pd.isna(barcode_val) or str(barcode_val).strip() == '':
        result.add_error(sheet, row_num, '条码.条码', "条码为空", fixable=False)

    store_val = row[store_col]
    if pd.isna(store_val) or str(store_val).strip() == '':
        result.add_error(sheet, row_num, '店仓.卖场代码', "卖场代码为空", fixable=False)

    qty_val = row[qty_col]
    if pd.isna(qty_val):
        result.add_error(sheet, row_num, '数量', "销售数量为空", fixable=False)
    else:
        try:
            qty_num = float(qty_val)
            if qty_num < 0:
                result.add_error(
                    sheet, row_num, '数量',
                    f"销售数量为负数 ({qty_num})",
                    fixable=True,
                    fix_action=f"set_zero:{qty_col}"
                )
            elif qty_num != int(qty_num):
                result.add_warning(
                    sheet, row_num, '数量',
                    f"销售数量非整数 ({qty_num})",
                    fixable=True,
                    fix_action=f"to_int:{qty_col}"
                )
        except (ValueError, TypeError):
            result.add_error(
                sheet, row_num, '数量',
                f"销售数量不是有效数字: '{qty_val}'",
                fixable=False
            )


def _validate_store_level_row(result, row, column_mapping, row_num):
    """验证卖场等级表数据行"""
    sheet = '卖场等级'
    code_col = column_mapping['代码']
    level_col = column_mapping['卖场等级']

    code_val = row[code_col]
    if pd.isna(code_val) or str(code_val).strip() == '':
        result.add_error(sheet, row_num, '代码', "卖场代码为空", fixable=False)

    level_val = row[level_col]
    if pd.isna(level_val) or str(level_val).strip() == '':
        result.add_error(sheet, row_num, '卖场等级', "卖场等级为空", fixable=False)
    else:
        level_str = str(level_val).strip()
        if level_str not in VALID_LEVELS:
            # 检查是否可以通过映射修复
            if level_str in LEVEL_FIX_MAP:
                fixed_level = LEVEL_FIX_MAP[level_str]
                result.add_error(
                    sheet, row_num, '卖场等级',
                    f"无效等级 '{level_str}' (应为 {fixed_level})",
                    fixable=True,
                    fix_action=f"fix_level:{level_col}->{fixed_level}"
                )
            else:
                result.add_error(
                    sheet, row_num, '卖场等级',
                    f"无效等级 '{level_str}'，有效值: SA/A/B/C/D/OL",
                    fixable=False
                )


def _validate_add_order_row(result, row, column_mapping, row_num):
    """验证加单数量表数据行"""
    sheet = '加单数量'
    sku_col = column_mapping['SKU']
    skc_col = column_mapping['SKC']
    qty_col = column_mapping['需分配数量']

    sku_val = row[sku_col]
    if pd.isna(sku_val) or str(sku_val).strip() == '':
        result.add_error(sheet, row_num, 'SKU', "SKU为空", fixable=False)

    skc_val = row[skc_col]
    if pd.isna(skc_val) or str(skc_val).strip() == '':
        result.add_warning(sheet, row_num, 'SKC', "SKC为空", fixable=False)

    qty_val = row[qty_col]
    if pd.isna(qty_val):
        result.add_error(sheet, row_num, '需分配数量', "需分配数量为空", fixable=False)
    else:
        try:
            qty_num = float(qty_val)
            if qty_num < 1:
                result.add_error(
                    sheet, row_num, '需分配数量',
                    f"需分配数量小于1 ({qty_num})",
                    fixable=False
                )
            elif qty_num != int(qty_num):
                result.add_warning(
                    sheet, row_num, '需分配数量',
                    f"需分配数量非整数 ({qty_num})",
                    fixable=True,
                    fix_action=f"to_int:{qty_col}"
                )
        except (ValueError, TypeError):
            result.add_error(
                sheet, row_num, '需分配数量',
                f"需分配数量不是有效数字: '{qty_val}'",
                fixable=False
            )


def detect_sales_type(df_sales):
    """
    检测销售数据类型（30天销量 vs 平均日销量）

    Args:
        df_sales: DataFrame - 销售表数据

    Returns:
        dict - {'type': '30day' or 'avg', 'confidence': 0-1, 'message': str}
    """
    try:
        if df_sales is None or df_sales.empty:
            return {'type': 'unknown', 'confidence': 0, 'message': '销售数据为空'}

        qty_col = '数量' if '数量' in df_sales.columns else None
        if qty_col is None:
            return {'type': 'unknown', 'confidence': 0, 'message': '未找到数量列'}

        quantities = pd.to_numeric(df_sales[qty_col], errors='coerce').dropna()
        if len(quantities) == 0:
            return {'type': 'unknown', 'confidence': 0, 'message': '无有效数量数据'}

        # 统计数值分布
        total_count = len(quantities)
        small_values = (quantities < 5).sum()
        small_ratio = small_values / total_count
        median_val = quantities.median()
        mean_val = quantities.mean()

        # 启发式规则
        # 如果大多数值<5（占比>60%），可能是平均日销量
        if small_ratio > 0.6 and median_val < 5:
            confidence = min(0.5 + small_ratio * 0.4, 0.95)
            return {
                'type': 'avg',
                'confidence': round(confidence, 2),
                'message': f'检测到数据可能是平均日销量（{small_ratio*100:.0f}%的值<5，中位数={median_val}）'
            }

        # 如果中位数>=10，更可能是30天销量
        if median_val >= 10:
            confidence = min(0.6 + (median_val / 30) * 0.3, 0.95)
            return {
                'type': '30day',
                'confidence': round(confidence, 2),
                'message': f'数据符合30天销量特征（中位数={median_val}）'
            }

        # 不确定的情况
        return {
            'type': 'uncertain',
            'confidence': 0.5,
            'message': f'无法确定数据类型（中位数={median_val}，小值占比={small_ratio*100:.0f}%）'
        }
    except Exception as e:
        return {'type': 'unknown', 'confidence': 0, 'message': f'检测失败: {e}'}


def validate_file(file_path, data_frames=None):
    """
    验证Excel文件的三级验证

    Args:
        file_path: str - Excel文件路径
        data_frames: dict or None - 预加载的DataFrame字典，键为sheet名。
                                   如果为None，则从文件读取。

    Returns:
        tuple - (ValidationResult, dict of DataFrames)
    """
    result = ValidationResult()
    dfs = {}

    # 读取Excel文件
    if data_frames is None:
        try:
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names

            # 第一级：结构验证（Sheet存在性）
            for sheet_key, sheet_name in REQUIRED_SHEETS.items():
                if sheet_name in sheet_names:
                    dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                else:
                    result.add_error(
                        sheet_name, 0, "",
                        f"缺少必需工作表: '{sheet_name}'",
                        fixable=False
                    )

            xls.close()
        except Exception as e:
            result.add_error("文件", 0, "", f"读取Excel文件失败: {e}", fixable=False)
            return result, dfs
    else:
        dfs = data_frames.copy()
        # 检查预加载数据是否缺少必需的工作表
        for sheet_key, sheet_name in REQUIRED_SHEETS.items():
            if sheet_name not in dfs:
                result.add_error(
                    sheet_name, 0, "",
                    f"缺少必需工作表: '{sheet_name}'",
                    fixable=False
                )

    # 验证各工作表
    for sheet_name in REQUIRED_SHEETS.values():
        if sheet_name in dfs:
            sheet_result = validate_sheet(dfs[sheet_name], sheet_name)
            # 合并验证结果
            result.passed.extend(sheet_result.passed)
            result.warnings.extend(sheet_result.warnings)
            result.errors.extend(sheet_result.errors)
            result.fixed.extend(sheet_result.fixed)

    # 第三级：业务逻辑验证
    _validate_business_logic(result, dfs)

    return result, dfs


def _validate_business_logic(result, dfs):
    """
    业务逻辑验证

    Args:
        result: ValidationResult - 验证结果对象
        dfs: dict - DataFrame字典
    """
    # 30天销量检测
    if '销售' in dfs and not dfs['销售'].empty:
        sales_detection = detect_sales_type(dfs['销售'])
        if sales_detection['type'] == 'avg':
            result.add_warning(
                '销售', 0, '数量',
                sales_detection['message'],
                fixable=True,
                fix_action=f"multiply_30:数量"
            )

    # 数据完整性检查
    if '加单数量' in dfs and '库存' in dfs and not dfs['加单数量'].empty and not dfs['库存'].empty:
        _check_sku_integrity(result, dfs)

    if '库存' in dfs and '卖场等级' in dfs and not dfs['库存'].empty and not dfs['卖场等级'].empty:
        _check_store_integrity(result, dfs)

    # 重复数据检查
    if '库存' in dfs and not dfs['库存'].empty:
        _check_duplicates(result, dfs['库存'], '库存', ['卖场代码', '条码'])

    if '销售' in dfs and not dfs['销售'].empty:
        _check_duplicates(result, dfs['销售'], '销售', ['条码.条码', '店仓.卖场代码'])


def _check_sku_integrity(result, dfs):
    """检查SKU完整性：加单SKU是否存在于库存表中"""
    try:
        inv_barcodes = set()
        if '条码' in dfs['库存'].columns:
            inv_barcodes = set(dfs['库存']['条码'].dropna().astype(str).str.strip())

        order_skus = set()
        if 'SKU' in dfs['加单数量'].columns:
            order_skus = set(dfs['加单数量']['SKU'].dropna().astype(str).str.strip())

        # 模糊匹配列名
        if not inv_barcodes and '库存' in dfs:
            for col in dfs['库存'].columns:
                if '条码' in str(col).lower() or 'sku' in str(col).lower():
                    inv_barcodes = set(dfs['库存'][col].dropna().astype(str).str.strip())
                    break

        if not order_skus and '加单数量' in dfs:
            for col in dfs['加单数量'].columns:
                if 'sku' in str(col).lower() or '条码' in str(col).lower():
                    order_skus = set(dfs['加单数量'][col].dropna().astype(str).str.strip())
                    break

        missing_skus = order_skus - inv_barcodes
        if missing_skus:
            result.add_warning(
                '加单数量', 0, 'SKU',
                f"有{len(missing_skus)}个加单SKU不在库存表中: {', '.join(list(missing_skus)[:5])}{'...' if len(missing_skus) > 5 else ''}",
                fixable=False
            )
        else:
            result.add_passed('加单数量', "加单SKU在库存表中均存在")
    except Exception as e:
        print(f"SKU完整性检查失败: {e}")


def _check_store_integrity(result, dfs):
    """检查卖场代码完整性：库存表中的卖场是否都在卖场等级表中"""
    try:
        inv_stores = set()
        if '卖场代码' in dfs['库存'].columns:
            inv_stores = set(dfs['库存']['卖场代码'].dropna().astype(str).str.strip())

        level_stores = set()
        if '代码' in dfs['卖场等级'].columns:
            level_stores = set(dfs['卖场等级']['代码'].dropna().astype(str).str.strip())

        # 模糊匹配
        if not inv_stores:
            for col in dfs['库存'].columns:
                if '卖场' in str(col) or '店铺' in str(col) or '店仓' in str(col):
                    inv_stores = set(dfs['库存'][col].dropna().astype(str).str.strip())
                    break

        if not level_stores:
            for col in dfs['卖场等级'].columns:
                if '代码' in str(col) or '卖场' in str(col):
                    level_stores = set(dfs['卖场等级'][col].dropna().astype(str).str.strip())
                    break

        missing_stores = inv_stores - level_stores
        if missing_stores:
            result.add_warning(
                '库存', 0, '卖场代码',
                f"有{len(missing_stores)}个库存表中的卖场不在卖场等级表中: {', '.join(list(missing_stores)[:5])}{'...' if len(missing_stores) > 5 else ''}",
                fixable=False
            )
        else:
            result.add_passed('库存', "卖场代码完整性检查通过")
    except Exception as e:
        print(f"卖场代码完整性检查失败: {e}")


def _check_duplicates(result, df, sheet_name, key_columns):
    """检查重复数据"""
    try:
        # 模糊匹配列名
        actual_cols = list(df.columns)
        matched_cols = []
        for key_col in key_columns:
            matched = _fuzzy_match_column(actual_cols, key_col, sheet_name)
            if matched:
                matched_cols.append(matched)

        if len(matched_cols) == len(key_columns):
            duplicates = df.duplicated(subset=matched_cols, keep=False)
            dup_count = duplicates.sum()
            if dup_count > 0:
                result.add_warning(
                    sheet_name, 0, ', '.join(matched_cols),
                    f"有{dup_count}行重复数据（{', '.join(matched_cols)}相同）",
                    fixable=False
                )
    except Exception as e:
        print(f"重复数据检查失败: {e}")


def fix_errors(df, errors, sheet_name):
    """
    应用修复操作到DataFrame

    Args:
        df: DataFrame - 要修复的数据
        errors: list - 错误/警告列表
        sheet_name: str - 工作表名称

    Returns:
        tuple - (修复后的DataFrame, 修复记录列表)
    """
    if df is None:
        return df, []

    df = df.copy()
    fix_records = []

    for error in errors:
        if not error.get('fixable', False):
            continue

        fix_action = error.get('fix_action', '')
        row_num = error.get('row', 0)
        column = error.get('column', '')

        try:
            if fix_action.startswith('rename:'):
                # 列名重命名
                parts = fix_action.split(':', 1)[1].split('->')
                if len(parts) == 2:
                    old_name, new_name = parts[0], parts[1]
                    if old_name in df.columns:
                        df.rename(columns={old_name: new_name}, inplace=True)
                        fix_records.append({
                            'sheet': sheet_name,
                            'row': 0,
                            'column': new_name,
                            'original': old_name,
                            'fixed': new_name,
                            'message': f"列名'{old_name}'已重命名为'{new_name}'"
                        })

            elif fix_action.startswith('set_zero:'):
                # 设置为0
                col = fix_action.split(':', 1)[1]
                if col in df.columns and row_num > 1:
                    idx = row_num - 2  # 转换为DataFrame索引
                    if idx < len(df):
                        original = df.iloc[idx][col]
                        df.iloc[idx, df.columns.get_loc(col)] = 0
                        fix_records.append({
                            'sheet': sheet_name,
                            'row': row_num,
                            'column': col,
                            'original': original,
                            'fixed': 0,
                            'message': f"负值已修复为0"
                        })

            elif fix_action.startswith('to_int:'):
                # 转为整数
                col = fix_action.split(':', 1)[1]
                if col in df.columns and row_num > 1:
                    idx = row_num - 2
                    if idx < len(df):
                        original = df.iloc[idx][col]
                        try:
                            df.iloc[idx, df.columns.get_loc(col)] = int(float(original))
                            fix_records.append({
                                'sheet': sheet_name,
                                'row': row_num,
                                'column': col,
                                'original': original,
                                'fixed': int(float(original)),
                                'message': f"已转为整数"
                            })
                        except (ValueError, TypeError):
                            pass

            elif fix_action.startswith('fix_level:'):
                # 修复等级
                parts = fix_action.split(':', 1)[1].split('->')
                if len(parts) == 2 and row_num > 1:
                    col = parts[0]
                    new_level = parts[1]
                    if col in df.columns:
                        idx = row_num - 2
                        if idx < len(df):
                            original = df.iloc[idx][col]
                            df.iloc[idx, df.columns.get_loc(col)] = new_level
                            fix_records.append({
                                'sheet': sheet_name,
                                'row': row_num,
                                'column': col,
                                'original': original,
                                'fixed': new_level,
                                'message': f"等级'{original}'已修复为'{new_level}'"
                            })

            elif fix_action.startswith('multiply_30:'):
                # 销量×30（平均日销量转30天销量）
                col = fix_action.split(':', 1)[1]
                if col in df.columns:
                    original_values = df[col].copy()
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) * 30
                    df[col] = df[col].astype(int)
                    fix_records.append({
                        'sheet': sheet_name,
                        'row': 0,
                        'column': col,
                        'original': '平均日销量',
                        'fixed': '30天销量(×30)',
                        'message': f"销售数据已从平均日销量转换为30天销量（×30）"
                    })

        except Exception as e:
            print(f"修复操作失败: {e}, fix_action: {fix_action}")

    return df, fix_records


def apply_all_fixes(dfs, validation_result):
    """
    应用所有可修复的错误到DataFrame字典

    Args:
        dfs: dict - DataFrame字典
        validation_result: ValidationResult - 验证结果

    Returns:
        tuple - (修复后的DataFrame字典, 所有修复记录列表)
    """
    fixed_dfs = {}
    all_fix_records = []

    # 合并错误和警告中可修复的项
    all_issues = validation_result.errors + validation_result.warnings

    for sheet_name, df in dfs.items():
        sheet_errors = [e for e in all_issues if e.get('sheet') == sheet_name]
        if sheet_errors:
            fixed_df, fix_records = fix_errors(df, sheet_errors, sheet_name)
            fixed_dfs[sheet_name] = fixed_df
            all_fix_records.extend(fix_records)
        else:
            fixed_dfs[sheet_name] = df

    return fixed_dfs, all_fix_records
