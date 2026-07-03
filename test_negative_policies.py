# 负数处理策略测试
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from allocation_core import allocate_add_order, get_30day_sales, get_inventory, DEFAULT_CONFIG
from data_validator import validate_file, validate_sheet


def test_negative_sales_policy_zero():
    """测试销量负数按0处理策略"""
    print("=== 测试销量负数按0处理策略 ===")

    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001', 'SKU001', 'SKU001'],
        '店仓.卖场代码': ['S001', 'S001', 'S001'],
        '数量': [10, -5, 3]
    })

    # 默认策略（zero），负数按0处理，总和应为13
    total = get_30day_sales(df_sales, 'SKU001', 'S001', policy="zero")
    assert total == 13, f"按0处理策略，总和应为13，实际为{total}"
    print(f"  ✓ 按0处理策略: 销量[-5] → 0, 总和=13")

    # original策略，负数按原值处理，总和应为8
    total = get_30day_sales(df_sales, 'SKU001', 'S001', policy="original")
    assert total == 8, f"按原值处理策略，总和应为8，实际为{total}"
    print(f"  ✓ 按原值处理策略: 销量[-5] → -5, 总和=8")


def test_negative_inventory_policy_zero():
    """测试库存负数按0处理策略"""
    print("\n=== 测试库存负数按0处理策略 ===")

    df_inv = pd.DataFrame({
        '卖场代码': ['S001'],
        '条码': ['SKU001'],
        '库存数量': [-2]
    })

    # 默认策略（zero），负数按0处理
    inv = get_inventory(df_inv, 'S001', 'SKU001', policy="zero")
    assert inv == 0, f"按0处理策略，库存应为0，实际为{inv}"
    print(f"  ✓ 按0处理策略: 库存[-2] → 0")

    # original策略，负数按原值处理
    inv = get_inventory(df_inv, 'S001', 'SKU001', policy="original")
    assert inv == -2, f"按原值处理策略，库存应为-2，实际为{inv}"
    print(f"  ✓ 按原值处理策略: 库存[-2] → -2")


def test_validation_negative_warnings():
    """测试数据验证对负数的警告提示"""
    print("\n=== 测试数据验证对负数的警告提示 ===")

    df_inv = pd.DataFrame({
        '卖场代码': ['S001', 'S002'],
        '条码': ['SKU001', 'SKU001'],
        '库存数量': [10, -5]
    })

    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001'],
        '店仓.卖场代码': ['S001'],
        '数量': [-3]
    })

    # zero策略：负数应提示警告，可修复
    policies = {'sales': 'zero', 'inventory': 'zero'}
    result_inv = validate_sheet(df_inv, '库存', policies=policies)
    result_sales = validate_sheet(df_sales, '销售', policies=policies)

    assert len(result_inv.warnings) == 1, f"库存表应1个警告，实际{len(result_inv.warnings)}"
    assert "负数" in result_inv.warnings[0]['message']
    assert "按0处理" in result_inv.warnings[0]['message']
    assert result_inv.warnings[0]['fixable'] == True
    print(f"  ✓ 库存负数(zero策略): 警告提示'将按0处理'")

    assert len(result_sales.warnings) == 1, f"销售表应1个警告，实际{len(result_sales.warnings)}"
    assert "负数" in result_sales.warnings[0]['message']
    assert "按0处理" in result_sales.warnings[0]['message']
    assert result_sales.warnings[0]['fixable'] == True
    print(f"  ✓ 销量负数(zero策略): 警告提示'将按0处理'")

    # original策略：负数应提示警告，不可修复
    policies = {'sales': 'original', 'inventory': 'original'}
    result_inv = validate_sheet(df_inv, '库存', policies=policies)
    result_sales = validate_sheet(df_sales, '销售', policies=policies)

    assert len(result_inv.warnings) == 1, f"库存表应1个警告，实际{len(result_inv.warnings)}"
    assert "按原值处理" in result_inv.warnings[0]['message']
    assert result_inv.warnings[0]['fixable'] == False
    print(f"  ✓ 库存负数(original策略): 警告提示'将按原值处理'")

    assert len(result_sales.warnings) == 1, f"销售表应1个警告，实际{len(result_sales.warnings)}"
    assert "按原值处理" in result_sales.warnings[0]['message']
    assert result_sales.warnings[0]['fixable'] == False
    print(f"  ✓ 销量负数(original策略): 警告提示'将按原值处理'")


def test_integration_negative_policies():
    """测试分配过程中负数处理策略的集成"""
    print("\n=== 测试分配过程中负数处理策略的集成 ===")

    df_inventory = pd.DataFrame({
        '卖场代码': ['S001', 'S002'],
        '条码': ['SKU001', 'SKU001'],
        '库存数量': [-5, 10]  # S001库存为负
    })

    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001', 'SKU001'],
        '店仓.卖场代码': ['S001', 'S002'],
        '数量': [-2, 15]  # S001销量为负（退货）
    })

    df_store_level = pd.DataFrame({
        '代码': ['S001', 'S002'],
        '卖场等级': ['SA', 'A']
    })

    df_add_order = pd.DataFrame({
        'SKU': ['SKU001'],
        'SKC': ['SKC001'],
        '需分配数量': [50]
    })

    # zero策略：负数按0处理
    config_zero = DEFAULT_CONFIG.copy()
    config_zero['allocation_config']['negative_sales_policy'] = 'zero'
    config_zero['allocation_config']['negative_inventory_policy'] = 'zero'

    alloc_zero, _, stores, skus, _ = allocate_add_order(
        df_inventory, df_sales, df_store_level, df_add_order, config=config_zero
    )

    total_zero = sum(alloc_zero[store]['SKU001'] for store in stores)
    print(f"  zero策略: 总分配量={total_zero}")

    # original策略：负数按原值处理
    config_orig = DEFAULT_CONFIG.copy()
    config_orig['allocation_config']['negative_sales_policy'] = 'original'
    config_orig['allocation_config']['negative_inventory_policy'] = 'original'

    alloc_orig, _, stores, skus, _ = allocate_add_order(
        df_inventory, df_sales, df_store_level, df_add_order, config=config_orig
    )

    total_orig = sum(alloc_orig[store]['SKU001'] for store in stores)
    print(f"  original策略: 总分配量={total_orig}")

    # 验证两种策略结果可能不同（因为原始数据包含负数）
    print(f"  ✓ 两种策略分配结果对比完成")


if __name__ == '__main__':
    test_negative_sales_policy_zero()
    test_negative_inventory_policy_zero()
    test_validation_negative_warnings()
    test_integration_negative_policies()
    print("\n=== 所有负数处理策略测试通过 ===")
