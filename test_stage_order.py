#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""测试阶段顺序调换后分配是否正常"""

import pandas as pd
import numpy as np
from allocation_core import allocate_add_order, generate_result_dataframe, DEFAULT_CONFIG

def create_test_data():
    """创建测试数据"""
    stores = ['S001', 'S002', 'S003', 'S004', 'S005']
    levels = ['SA', 'A', 'B', 'C', 'D']
    
    df_store_level = pd.DataFrame({
        '代码': stores,
        '卖场等级': levels
    })
    
    skus = ['SKU001160', 'SKU001165']
    df_inventory = pd.DataFrame({
        '卖场代码': stores * 2,
        '条码': [skus[0]] * 5 + [skus[1]] * 5,
        '库存数量': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    })
    
    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001160'] * 5 + ['SKU001165'] * 5,
        '店仓.卖场代码': stores * 2,
        '数量': [100, 80, 60, 40, 20, 100, 80, 60, 40, 20]
    })
    
    df_add_order = pd.DataFrame({
        'SKU': skus,
        'SKC': ['SKC001', 'SKC001'],
        '需分配数量': [50, 50]
    })
    
    return df_inventory, df_sales, df_store_level, df_add_order

def count_allocation_reasons(reason_df, reason_type):
    """统计某种分配原因的数量"""
    count = 0
    for col in reason_df.columns:
        if col in ['卖场', '卖场等级']:
            continue
        for val in reason_df[col]:
            if reason_type in str(val):
                count += 1
    return count

def test_stage_order(order_str):
    """测试指定阶段顺序"""
    order_map = {'1': 'broken_size_fix', '2': 'sales_match', '3': 'sell_through_priority'}
    order = [int(c) - 1 for c in order_str]
    stage_names = ['broken_size_fix', 'sales_match', 'sell_through_priority']
    stage_priority = [stage_names[i] for i in order]
    
    config = {
        "version": "2.6",
        "allocation_config": {
            "coverage_days": {"SA": 30, "A": 30, "B": 14, "C": 14, "D": 14, "OL": 14},
            "level_weights": {"SA": 1.5, "A": 1.3, "B": 1.2, "C": 1.1, "D": 1.1, "OL": 1.0},
            "safety_factors": {"SA": 0.5, "A": 0.4, "B": 0.3, "C": 0.25, "D": 0.2, "OL": 0.2},
            "min_target_inventory": {"SA": 0, "A": 0, "B": 0, "C": 0, "D": 0, "OL": 0},
            "stage_priority": stage_priority,
            "max_remaining_per_store": 10
        }
    }
    
    df_inventory, df_sales, df_store_level, df_add_order = create_test_data()
    
    allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
        df_inventory, df_sales, df_store_level, df_add_order, config
    )
    
    result_df, reason_df = generate_result_dataframe(
        allocation_result, allocation_reasons, stores_sorted, skus, store_level_map
    )
    
    total_allocated = 0
    for col in result_df.columns:
        if col != '卖场':
            total_allocated += result_df[col].sum()
    
    broken_fix_count = count_allocation_reasons(reason_df, '断码修复')
    sales_match_count = count_allocation_reasons(reason_df, '销量匹配')
    sell_through_count = count_allocation_reasons(reason_df, '销尽率优先')
    remaining_count = count_allocation_reasons(reason_df, '剩余分配')
    
    return {
        'order': order_str,
        'stage_priority': stage_priority,
        'total_allocated': total_allocated,
        'broken_fix_count': broken_fix_count,
        'sales_match_count': sales_match_count,
        'sell_through_count': sell_through_count,
        'remaining_count': remaining_count,
        'reason_df': reason_df
    }

def main():
    print("=" * 60)
    print("测试阶段顺序调换后分配是否正常")
    print("=" * 60)
    
    orders = ['123', '213', '321', '132', '231', '312']
    
    for order in orders:
        result = test_stage_order(order)
        print(f"\n顺序 {order} ({' -> '.join(result['stage_priority'])}):")
        print(f"  总分配数量: {result['total_allocated']}")
        print(f"  断码修复次数: {result['broken_fix_count']}")
        print(f"  销量匹配次数: {result['sales_match_count']}")
        print(f"  销尽率优先次数: {result['sell_through_count']}")
        print(f"  剩余分配次数: {result['remaining_count']}")
    
    print("\n" + "=" * 60)
    print("验证: 不同顺序下总分配数量应该相同 (100)")
    print("=" * 60)
    
    all_same = True
    base_total = None
    for order in orders:
        result = test_stage_order(order)
        if base_total is None:
            base_total = result['total_allocated']
        elif result['total_allocated'] != base_total:
            all_same = False
            print(f"错误: 顺序 {order} 的总分配数量 {result['total_allocated']} 与基准 {base_total} 不同")
    
    if all_same:
        print("✓ 所有顺序的总分配数量一致")
    else:
        print("✗ 存在分配数量不一致的情况")
    
    print("\n" + "=" * 60)
    print("验证: 阶段执行顺序影响分配原因")
    print("=" * 60)
    
    result_123 = test_stage_order('123')
    result_321 = test_stage_order('321')
    
    print(f"顺序123: 断码修复={result_123['broken_fix_count']}, 销量匹配={result_123['sales_match_count']}, 销尽率优先={result_123['sell_through_count']}")
    print(f"顺序321: 断码修复={result_321['broken_fix_count']}, 销量匹配={result_321['sales_match_count']}, 销尽率优先={result_321['sell_through_count']}")
    
    if result_123['broken_fix_count'] != result_321['broken_fix_count'] or \
       result_123['sales_match_count'] != result_321['sales_match_count'] or \
       result_123['sell_through_count'] != result_321['sell_through_count']:
        print("✓ 不同顺序导致分配原因不同，说明阶段顺序生效")
    else:
        print("✗ 不同顺序的分配原因相同，可能阶段顺序未生效")
    
    print("\n" + "=" * 60)
    print("测试完成")
    print("=" * 60)

if __name__ == "__main__":
    main()
