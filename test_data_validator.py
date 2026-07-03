# data_validator.py 模块测试
import sys
import os
import tempfile
import pandas as pd
import numpy as np

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_validator import (
    ValidationResult,
    generate_template,
    validate_sheet,
    validate_file,
    detect_sales_type,
    fix_errors,
    apply_all_fixes,
    REQUIRED_SHEETS,
    REQUIRED_COLUMNS,
    VALID_LEVELS,
    LEVEL_FIX_MAP
)


def test_validation_result_class():
    """测试 ValidationResult 类"""
    print("\n=== 测试 ValidationResult 类 ===")
    result = ValidationResult()

    # 测试添加通过项
    result.add_passed('库存', '列名匹配通过')
    assert len(result.passed) == 1
    assert result.passed[0]['sheet'] == '库存'

    # 测试添加警告
    result.add_warning('销售', 5, '数量', '数值异常', fixable=True, fix_action='multiply_30:数量')
    assert len(result.warnings) == 1
    assert result.warnings[0]['fixable'] == True

    # 测试添加错误
    result.add_error('库存', 15, '库存数量', '负数', fixable=True, fix_action='set_zero:库存数量')
    assert len(result.errors) == 1

    # 测试属性
    assert result.has_errors == True
    assert result.has_warnings == True
    assert result.is_passed == False

    # 测试摘要
    summary = result.get_summary()
    assert summary['passed'] == 1
    assert summary['warnings'] == 1
    assert summary['errors'] == 1

    # 测试按工作表分组
    report = result.get_report_by_sheet()
    assert '库存' in report
    assert len(report['库存']['passed']) == 1
    assert len(report['库存']['errors']) == 1

    print("  ✓ ValidationResult 类测试通过")


def test_generate_template():
    """测试模板生成"""
    print("\n=== 测试模板生成 ===")

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        output_path = tmp.name

    try:
        # 生成模板
        success = generate_template(output_path)
        assert success == True, "模板生成应返回True"
        assert os.path.exists(output_path), "模板文件应存在"

        # 读取模板验证内容
        xls = pd.ExcelFile(output_path)
        sheet_names = xls.sheet_names

        # 验证4个必需工作表 + 使用说明
        assert '库存' in sheet_names, "应包含库存表"
        assert '销售' in sheet_names, "应包含销售表"
        assert '卖场等级' in sheet_names, "应包含卖场等级表"
        assert '加单数量' in sheet_names, "应包含加单数量表"
        assert '使用说明' in sheet_names, "应包含使用说明"

        # 验证列名
        df_inv = pd.read_excel(xls, sheet_name='库存')
        assert list(df_inv.columns) == ['卖场代码', '条码', '库存数量']

        df_sales = pd.read_excel(xls, sheet_name='销售')
        assert list(df_sales.columns) == ['条码.条码', '店仓.卖场代码', '数量']

        df_level = pd.read_excel(xls, sheet_name='卖场等级')
        assert list(df_level.columns) == ['代码', '卖场等级']

        df_order = pd.read_excel(xls, sheet_name='加单数量')
        assert list(df_order.columns) == ['SKU', 'SKC', '需分配数量']

        xls.close()
        print("  ✓ 模板生成测试通过")
    finally:
        if os.path.exists(output_path):
            os.remove(output_path)


def test_validate_sheet_valid_data():
    """测试有效数据验证"""
    print("\n=== 测试有效数据验证 ===")

    # 创建有效的库存表数据
    df_inv = pd.DataFrame({
        '卖场代码': ['S001', 'S002', 'S003'],
        '条码': ['SKU001', 'SKU001', 'SKU001'],
        '库存数量': [10, 5, 8]
    })

    result = validate_sheet(df_inv, '库存')
    assert not result.has_errors, f"有效数据不应有错误: {[e['message'] for e in result.errors]}"
    print("  ✓ 有效库存数据验证通过")

    # 创建有效的销售表数据
    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001', 'SKU001', 'SKU001'],
        '店仓.卖场代码': ['S001', 'S002', 'S003'],
        '数量': [30, 15, 20]
    })

    result = validate_sheet(df_sales, '销售')
    assert not result.has_errors, f"有效销售数据不应有错误: {[e['message'] for e in result.errors]}"
    print("  ✓ 有效销售数据验证通过")

    # 创建有效的卖场等级表
    df_level = pd.DataFrame({
        '代码': ['S001', 'S002', 'S003'],
        '卖场等级': ['SA', 'A', 'B']
    })

    result = validate_sheet(df_level, '卖场等级')
    assert not result.has_errors, f"有效等级数据不应有错误: {[e['message'] for e in result.errors]}"
    print("  ✓ 有效卖场等级数据验证通过")

    # 创建有效的加单数量表
    df_order = pd.DataFrame({
        'SKU': ['SKU001'],
        'SKC': ['SKC001'],
        '需分配数量': [50]
    })

    result = validate_sheet(df_order, '加单数量')
    assert not result.has_errors, f"有效加单数据不应有错误: {[e['message'] for e in result.errors]}"
    print("  ✓ 有效加单数量数据验证通过")


def test_validate_sheet_column_mismatch():
    """测试列名不匹配验证"""
    print("\n=== 测试列名不匹配验证 ===")

    # 列名模糊匹配测试 - 库存表
    df = pd.DataFrame({
        '卖场': ['S001'],  # 应匹配到"卖场代码"
        'SKU': ['SKU001'],  # 应匹配到"条码"
        '库存': [10]  # 应匹配到"库存数量"
    })

    result = validate_sheet(df, '库存')
    # 应该有警告但不应有错误（因为可以模糊匹配）
    assert len(result.errors) == 0, f"模糊匹配不应报错: {[e['message'] for e in result.errors]}"
    assert len(result.warnings) > 0, "应有列名不匹配的警告"
    print("  ✓ 库存表列名模糊匹配测试通过")

    # 缺少必需列测试
    df_missing = pd.DataFrame({
        '卖场代码': ['S001'],
        '条码': ['SKU001']
        # 缺少库存数量列
    })

    result = validate_sheet(df_missing, '库存')
    assert result.has_errors, "缺少必需列应报错"
    assert any('库存数量' in e['message'] for e in result.errors), "应提示缺少库存数量列"
    print("  ✓ 缺少必需列验证通过")


def test_validate_sheet_data_errors():
    """测试数据类型错误验证"""
    print("\n=== 测试数据类型错误验证 ===")

    # 负数库存
    df = pd.DataFrame({
        '卖场代码': ['S001', 'S002'],
        '条码': ['SKU001', 'SKU001'],
        '库存数量': [10, -3]  # 第二行负数
    })

    result = validate_sheet(df, '库存')
    assert result.has_errors, "负数库存应报错"
    assert any('负数' in e['message'] for e in result.errors), "应提示负数错误"
    assert any(e.get('fixable') for e in result.errors), "负数错误应可修复"
    print("  ✓ 负数库存验证通过")

    # 非数字库存
    df = pd.DataFrame({
        '卖场代码': ['S001'],
        '条码': ['SKU001'],
        '库存数量': ['abc']  # 非数字
    })

    result = validate_sheet(df, '库存')
    assert result.has_errors, "非数字库存应报错"
    assert any('不是有效数字' in e['message'] for e in result.errors)
    print("  ✓ 非数字库存验证通过")

    # 无效卖场等级
    df = pd.DataFrame({
        '代码': ['S001', 'S002'],
        '卖场等级': ['SA', 'A级']  # 第二行无效
    })

    result = validate_sheet(df, '卖场等级')
    assert result.has_errors, "无效等级应报错"
    assert any('A级' in e['message'] for e in result.errors), "应提示A级错误"
    assert any(e.get('fixable') for e in result.errors), "A级应可修复为A"
    print("  ✓ 无效卖场等级验证通过")


def test_detect_sales_type():
    """测试销售数据类型检测"""
    print("\n=== 测试销售数据类型检测 ===")

    # 30天销量数据（中位数>=10）
    df_30day = pd.DataFrame({'数量': [30, 25, 40, 15, 20, 35, 10]})
    result = detect_sales_type(df_30day)
    assert result['type'] == '30day', f"应检测为30天销量: {result}"
    print(f"  ✓ 30天销量检测通过: {result['message']}")

    # 平均日销量数据（多数值<5）
    df_avg = pd.DataFrame({'数量': [1, 2, 3, 1, 2, 4, 1, 0, 3, 2]})
    result = detect_sales_type(df_avg)
    assert result['type'] == 'avg', f"应检测为平均日销量: {result}"
    print(f"  ✓ 平均日销量检测通过: {result['message']}")

    # 空数据
    df_empty = pd.DataFrame({'数量': []})
    result = detect_sales_type(df_empty)
    assert result['type'] == 'unknown'
    print("  ✓ 空数据检测通过")


def test_fix_errors():
    """测试错误修复"""
    print("\n=== 测试错误修复 ===")

    # 测试负数修复
    df = pd.DataFrame({
        '卖场代码': ['S001', 'S002'],
        '条码': ['SKU001', 'SKU001'],
        '库存数量': [10, -3]
    })

    result = validate_sheet(df, '库存')
    fixed_df, fix_records = fix_errors(df, result.errors, '库存')

    assert len(fix_records) == 1, f"应有1条修复记录: {fix_records}"
    assert fixed_df.iloc[1]['库存数量'] == 0, "负数应修复为0"
    print("  ✓ 负数修复测试通过")

    # 测试等级修复
    df_level = pd.DataFrame({
        '代码': ['S001', 'S002'],
        '卖场等级': ['SA', 'A级']
    })

    result = validate_sheet(df_level, '卖场等级')
    fixed_df, fix_records = fix_errors(df_level, result.errors, '卖场等级')

    assert len(fix_records) == 1, f"应有1条修复记录: {fix_records}"
    assert fixed_df.iloc[1]['卖场等级'] == 'A', "A级应修复为A"
    print("  ✓ 等级修复测试通过")

    # 测试销量转换
    df_sales = pd.DataFrame({'数量': [1, 2, 3, 4]})
    errors = [{
        'sheet': '销售',
        'row': 0,
        'column': '数量',
        'message': '可能是平均日销量',
        'fixable': True,
        'fix_action': 'multiply_30:数量'
    }]

    fixed_df, fix_records = fix_errors(df_sales, errors, '销售')
    assert len(fix_records) == 1, "应有1条修复记录"
    assert fixed_df.iloc[0]['数量'] == 30, f"1×30应等于30: {fixed_df.iloc[0]['数量']}"
    print("  ✓ 销量转换测试通过")


def test_validate_file_integration():
    """测试文件级验证集成"""
    print("\n=== 测试文件级验证集成 ===")

    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        file_path = tmp.name

    try:
        # 先生成模板
        generate_template(file_path)

        # 验证模板文件
        result, dfs = validate_file(file_path)

        assert '库存' in dfs, "应读取到库存表"
        assert '销售' in dfs, "应读取到销售表"
        assert '卖场等级' in dfs, "应读取到卖场等级表"
        assert '加单数量' in dfs, "应读取到加单数量表"

        # 模板有示例数据，可能有警告但不应有错误
        assert not result.has_errors, f"模板不应有错误: {[e['message'] for e in result.errors]}"
        print("  ✓ 模板文件验证通过")

        # 测试缺少工作表的情况
        df_only_inv = {'库存': pd.DataFrame({'卖场代码': ['S001'], '条码': ['SKU001'], '库存数量': [10]})}
        result, _ = validate_file("dummy.xlsx", data_frames=df_only_inv)
        assert result.has_errors, "缺少工作表应报错"
        assert any('销售' in e['message'] for e in result.errors), "应提示缺少销售表"
        print("  ✓ 缺少工作表验证通过")

    finally:
        if os.path.exists(file_path):
            os.remove(file_path)


def test_apply_all_fixes():
    """测试批量修复"""
    print("\n=== 测试批量修复 ===")

    dfs = {
        '库存': pd.DataFrame({
            '卖场代码': ['S001', 'S002'],
            '条码': ['SKU001', 'SKU001'],
            '库存数量': [10, -5]
        }),
        '卖场等级': pd.DataFrame({
            '代码': ['S001', 'S002'],
            '卖场等级': ['SA', 'A级']
        })
    }

    # 验证
    result, _ = validate_file("dummy.xlsx", data_frames=dfs)

    # 批量修复
    fixed_dfs, all_fixes = apply_all_fixes(dfs, result)

    assert len(all_fixes) >= 2, f"应有至少2条修复记录: {all_fixes}"
    assert fixed_dfs['库存'].iloc[1]['库存数量'] == 0, "负数应修复为0"
    assert fixed_dfs['卖场等级'].iloc[1]['卖场等级'] == 'A', "A级应修复为A"
    print("  ✓ 批量修复测试通过")


def run_all_tests():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("  data_validator.py 模块测试")
    print("=" * 60)

    tests = [
        test_validation_result_class,
        test_generate_template,
        test_validate_sheet_valid_data,
        test_validate_sheet_column_mismatch,
        test_validate_sheet_data_errors,
        test_detect_sales_type,
        test_fix_errors,
        test_validate_file_integration,
        test_apply_all_fixes
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
