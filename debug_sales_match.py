
# 调试销量匹配阶段
import pandas as pd
import os
from allocation_core import allocate_add_order, load_config, DEFAULT_CONFIG, get_store_level, get_inventory, get_30day_sales

# 读取Excel文件
excel_path = "/workspace/💡加单商品分配_0424.xlsx"

df_inventory = pd.read_excel(excel_path, sheet_name='库存')
df_sales = pd.read_excel(excel_path, sheet_name='销售')
df_store_level = pd.read_excel(excel_path, sheet_name='卖场等级')
df_add_order = pd.read_excel(excel_path, sheet_name='加单数量')

config = load_config()
alloc_config = config.get('allocation_config', DEFAULT_CONFIG['allocation_config'])
coverage_days = alloc_config.get('coverage_days', DEFAULT_CONFIG['allocation_config']['coverage_days'])
safety_factors = alloc_config.get('safety_factors', DEFAULT_CONFIG['allocation_config']['safety_factors'])

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

print("=" * 80)
print("调试销量匹配阶段")
print("=" * 80)

for sku_info in skus[:1]:  # 只分析第一个SKU
    sku = sku_info['sku']
    print(f"\n【分析SKU: {sku}】")
    print(f"需分配数量: {sku_info['required_qty']}")
    
    # 收集每个卖场的数据
    store_data = {}
    print(f"\n{'卖场':<10} {'等级':<5} {'库存':<8} {'30天销量':<10} {'平均日需求':<12} {'覆盖周期':<8} {'安全系数':<8} {'安全库存':<10} {'目标库存':<10} {'可分配':<10}")
    print("-" * 100)
    
    for store in stores_sorted[:15]:  # 查看前15个卖场
        level = get_store_level(df_store_level, store)
        inv = get_inventory(df_inventory, store, sku)
        sales_30d = get_30day_sales(df_sales, sku, store)
        
        # 计算目标库存
        daily_demand = sales_30d / 30.0
        coverage = coverage_days.get(level, 14)
        safety_factor = safety_factors.get(level, 0.3)
        safety_stock = daily_demand * safety_factor * coverage
        target_inv = int(daily_demand * coverage + safety_stock)
        
        # 计算可分配数量
        current_inv = inv
        can_allocate = max(0, target_inv - current_inv)
        
        store_data[store] = {
            'inventory': inv,
            'sales_30d': sales_30d,
            'level': level,
            'daily_demand': daily_demand,
            'coverage': coverage,
            'safety_factor': safety_factor,
            'safety_stock': safety_stock,
            'target_inv': target_inv,
            'can_allocate': can_allocate
        }
        
        print(f"{store:<10} {level:<5} {inv:<8} {sales_30d:<10.2f} {daily_demand:<12.4f} {coverage:<8} {safety_factor:<8} {safety_stock:<10.2f} {target_inv:<10} {can_allocate:<10}")

print("\n" + "=" * 80)
print("问题分析：")
print("=" * 80)
print("1. 从上面的数据可以看出，大部分卖场的库存数量已经超过了目标库存")
print("2. 目标库存 = 平均日需求 × 覆盖周期 + 安全库存")
print("3. 如果当前库存已经达到或超过目标库存，销量匹配阶段就不会分配任何商品")
print("\n可能的原因：")
print("a) 覆盖周期设置较长（SA/A级30天）")
print("b) 30天销量较低，导致计算出的目标库存较低")
print("c) 当前库存已经比较充足")
print("\n建议：")
print("可以根据实际业务需求调整以下参数：")
print("- 覆盖周期天数（当前SA/A: 30, B/C/D/OL: 14）")
print("- 安全系数（当前SA: 0.5, A: 0.4, B: 0.3, C: 0.25, D/OL: 0.2）")
