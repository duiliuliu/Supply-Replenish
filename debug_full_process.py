
# 完整调试分配流程
import pandas as pd
import os
from collections import defaultdict
from allocation_core import get_30day_sales, get_store_level, get_inventory, extract_size, is_core_size, load_config, DEFAULT_CONFIG

# 读取Excel文件
excel_path = "/workspace/💡加单商品分配_0424.xlsx"

df_inventory = pd.read_excel(excel_path, sheet_name='库存')
df_sales = pd.read_excel(excel_path, sheet_name='销售')
df_store_level = pd.read_excel(excel_path, sheet_name='卖场等级')
df_add_order = pd.read_excel(excel_path, sheet_name='加单数量')

config = load_config()
alloc_config = config.get('allocation_config', DEFAULT_CONFIG['allocation_config'])
coverage_days = alloc_config.get('coverage_days', DEFAULT_CONFIG['allocation_config']['coverage_days'])
level_weights = alloc_config.get('level_weights', DEFAULT_CONFIG['allocation_config']['level_weights'])
safety_factors = alloc_config.get('safety_factors', DEFAULT_CONFIG['allocation_config']['safety_factors'])
max_remaining_per_store = alloc_config.get('max_remaining_per_store', 10)

level_order = ['SA', 'A', 'B', 'C', 'D', 'OL']
stores_sorted = []
store_level_map = {}
for level in level_order:
    level_stores = df_store_level[df_store_level['卖场等级'] == level]['代码'].tolist()
    stores_sorted.extend(level_stores)
    for store in level_stores:
        store_level_map[store] = level

skus = []
for idx, row in df_add_order.iterrows():
    skus.append({
        'sku': row['SKU'],
        'skc': row['SKC'],
        'required_qty': int(row['需分配数量'])
    })

print("=" * 100)
print("完整分配流程调试")
print("=" * 100)

for sku_info in skus[:1]:  # 只分析第一个SKU
    sku = sku_info['sku']
    remaining_qty = sku_info['required_qty']
    print(f"\n【SKU: {sku}】")
    print(f"总需分配: {remaining_qty} 件")
    
    # 初始化结果
    allocation_result = defaultdict(lambda: defaultdict(int))
    allocation_reasons = defaultdict(lambda: defaultdict(str))
    
    # 收集store_data
    store_data = {}
    for store in stores_sorted:
        inv = get_inventory(df_inventory, store, sku)
        sales_30d = get_30day_sales(df_sales, sku, store)
        level = get_store_level(df_store_level, store)
        store_data[store] = {
            'inventory': inv,
            'sales_30d': sales_30d,
            'level': level
        }
    
    # 阶段1: 断码修复
    print(f"\n--- 阶段1: 断码修复 ---")
    stage1_allocated = 0
    for store in stores_sorted[:15]:
        current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
        size = extract_size(sku)
        level = store_data[store]['level']
        
        target = 0
        if level in ['SA', 'A']:
            if is_core_size(size):
                target = 2
            else:
                target = 1
        else:
            if is_core_size(size):
                target = 1
        
        if current_inv < target:
            to_allocate = min(target - current_inv, remaining_qty)
            if to_allocate > 0:
                allocation_result[store][sku] += to_allocate
                remaining_qty -= to_allocate
                stage1_allocated += to_allocate
                if allocation_reasons[store][sku]:
                    allocation_reasons[store][sku] += f',断码修复({to_allocate})'
                else:
                    allocation_reasons[store][sku] = f'断码修复({to_allocate})'
                print(f"  {store}: 当前{current_inv} → 目标{target} → 分配{to_allocate}")
    
    print(f"阶段1总计: {stage1_allocated}件，剩余: {remaining_qty}件")
    
    # 阶段2: 销量匹配
    print(f"\n--- 阶段2: 销量匹配 ---")
    stage2_allocated = 0
    print(f"{'卖场':<10} {'等级':<5} {'库存':<8} {'已分配':<8} {'合计':<8} {'30天销量':<10} {'目标库存':<10} {'可分配':<10}")
    print("-" * 90)
    
    for store in stores_sorted[:15]:
        current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
        level = store_data[store]['level']
        sales_30d = store_data[store]['sales_30d']
        
        # 计算目标库存
        daily_demand = sales_30d / 30.0
        coverage = coverage_days.get(level, 14)
        safety_factor = safety_factors.get(level, 0.3)
        safety_stock = daily_demand * safety_factor * coverage
        target_inv = int(daily_demand * coverage + safety_stock)
        
        can_allocate = max(0, target_inv - current_inv)
        
        if current_inv < target_inv and remaining_qty > 0:
            to_allocate = min(target_inv - current_inv, remaining_qty)
            if to_allocate > 0:
                allocation_result[store][sku] += to_allocate
                remaining_qty -= to_allocate
                stage2_allocated += to_allocate
                if allocation_reasons[store][sku]:
                    allocation_reasons[store][sku] += f',销量匹配({to_allocate})'
                else:
                    allocation_reasons[store][sku] = f'销量匹配({to_allocate})'
                print(f"{store:<10} {level:<5} {store_data[store]['inventory']:<8} {allocation_result[store][sku]-to_allocate:<8} {current_inv:<8} {sales_30d:<10} {target_inv:<10} +{to_allocate:<9}")
            else:
                print(f"{store:<10} {level:<5} {store_data[store]['inventory']:<8} {allocation_result[store][sku]:<8} {current_inv:<8} {sales_30d:<10} {target_inv:<10} 0         [跳过：库存已达目标]")
        else:
            reason = ""
            if current_inv >= target_inv:
                reason = "[跳过：库存已达目标]"
            elif remaining_qty <= 0:
                reason = "[跳过：已无剩余数量]"
            print(f"{store:<10} {level:<5} {store_data[store]['inventory']:<8} {allocation_result[store][sku]:<8} {current_inv:<8} {sales_30d:<10} {target_inv:<10} {can_allocate:<10} {reason}")
    
    print(f"阶段2总计: {stage2_allocated}件，剩余: {remaining_qty}件")
    
    print(f"\n【小结】")
    print(f"断码修复分配: {stage1_allocated}件")
    print(f"销量匹配分配: {stage2_allocated}件")
    print(f"剩余: {remaining_qty}件")
    if stage2_allocated == 0:
        print("\n⚠️ 销量匹配阶段没有分配任何商品！")
        print("原因分析：")
        print("1. 大部分卖场30天销量很低或为0，导致目标库存计算为0")
        print("2. 当前库存（包括断码修复分配的）已经达到或超过目标库存")
