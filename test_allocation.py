
# 测试加单分配逻辑
import pandas as pd
import os
from allocation_core import allocate_add_order, generate_result_dataframe, load_config, DEFAULT_CONFIG

# 读取Excel文件
excel_path = "/workspace/💡加单商品分配_0424.xlsx"
if not os.path.exists(excel_path):
    print(f"Error: File {excel_path} not found")
    exit(1)

print("=" * 60)
print("读取Excel文件...")
print("=" * 60)

try:
    df_inventory = pd.read_excel(excel_path, sheet_name='库存')
    df_sales = pd.read_excel(excel_path, sheet_name='销售')
    df_store_level = pd.read_excel(excel_path, sheet_name='卖场等级')
    df_add_order = pd.read_excel(excel_path, sheet_name='加单数量')
    
    print("✓ 所有工作表读取成功")
    print()
    
    print("=" * 60)
    print("工作表数据预览")
    print("=" * 60)
    
    print("\n【库存表】:")
    print(df_inventory.head(10).to_string())
    print(f"总记录数: {len(df_inventory)}")
    
    print("\n【销售表】:")
    print(df_sales.head(10).to_string())
    print(f"总记录数: {len(df_sales)}")
    
    print("\n【卖场等级表】:")
    print(df_store_level.head(10).to_string())
    print(f"等级分布:")
    print(df_store_level['卖场等级'].value_counts().to_string())
    
    print("\n【加单数量表】:")
    print(df_add_order.head(10).to_string())
    print(f"需分配的SKU数: {len(df_add_order)}")
    print(f"总需分配数量: {df_add_order['需分配数量'].sum()}")
    
    print("\n" + "=" * 60)
    print("执行分配逻辑...")
    print("=" * 60)
    
    allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
        df_inventory, df_sales, df_store_level, df_add_order, load_config()
    )
    
    print("✓ 分配逻辑执行完成")
    print()
    
    # 生成结果
    result_df, reason_df = generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus, store_level_map)
    
    print("=" * 60)
    print("分配结果分析")
    print("=" * 60)
    
    # 统计分配原因
    print("\n【分配原因统计】:")
    for col in reason_df.columns[2:]:
        reason_series = reason_df[col]
        non_empty = reason_series[reason_series != '']
        if len(non_empty) > 0:
            print(f"\n--- {col} ---")
            reasons = []
            for r in non_empty:
                if r:
                    reasons.extend(r.split(','))
            # 统计各阶段分配次数和数量
            stage_counts = {}
            stage_total = {}
            for r in reasons:
                if r:
                    # 解析 "销量匹配(3)" -> stage="销量匹配", qty=3
                    import re
                    match = re.match(r'(.*)\((\d+)\)', r)
                    if match:
                        stage = match.group(1)
                        qty = int(match.group(2))
                        stage_counts[stage] = stage_counts.get(stage, 0) + 1
                        stage_total[stage] = stage_total.get(stage, 0) + qty
            
            for stage in sorted(stage_counts.keys()):
                print(f"  {stage}: {stage_counts[stage]}次，总计{stage_total[stage]}件")
    
    # 打印分配结果前20行
    print("\n" + "=" * 60)
    print("分配结果预览（前20行）")
    print("=" * 60)
    print(result_df.head(20).to_string())
    
    print("\n" + "=" * 60)
    print("分配原因预览（前20行）")
    print("=" * 60)
    print(reason_df.head(20).to_string())
    
    # 保存测试结果
    test_output_path = "/workspace/💡加单商品分配_测试结果.xlsx"
    with pd.ExcelWriter(test_output_path, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='分配数量', index=False)
        reason_df.to_excel(writer, sheet_name='分配原因', index=False)
    print(f"\n✓ 测试结果已保存到: {test_output_path}")
    
except Exception as e:
    print(f"Error: {str(e)}")
    import traceback
    print(traceback.format_exc())
