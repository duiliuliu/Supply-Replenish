# 加单分配核心逻辑
import pandas as pd
from collections import defaultdict

def get_30day_sales(df_sales, sku, store_code):
    df_filtered = df_sales[(df_sales['条码.条码'] == sku) & (df_sales['店仓.卖场代码'] == store_code)].copy()
    df_filtered['数量'] = df_filtered['数量'].apply(lambda x: max(0, x))
    return df_filtered['数量'].sum()

def extract_size(sku):
    try:
        return int(str(sku)[-3:])
    except:
        return 0

def is_core_size(size):
    return size in [160, 165]

def get_store_level(df_store_level, store_code):
    filtered = df_store_level[df_store_level['代码'] == store_code]
    if len(filtered) > 0:
        return filtered.iloc[0]['卖场等级']
    return 'C'

def get_inventory(df_inventory, store_code, sku):
    filtered = df_inventory[(df_inventory['卖场代码'] == store_code) & (df_inventory['条码'] == sku)]
    if len(filtered) > 0:
        return max(0, int(filtered.iloc[0]['库存数量']))
    return 0

def allocate_add_order(df_inventory, df_sales, df_store_level, df_add_order):
    # 获取按等级排序的卖场列表
    level_order = ['SA', 'A', 'B', 'C', 'D', 'OL']
    stores_sorted = []
    for level in level_order:
        level_stores = df_store_level[df_store_level['卖场等级'] == level]['代码'].tolist()
        stores_sorted.extend(level_stores)
    
    # 获取SKU列表
    skus = []
    for idx, row in df_add_order.iterrows():
        skus.append({
            'sku': row['SKU'],
            'skc': row['SKC'],
            'required_qty': int(row['需分配数量'])
        })
    
    # 初始化分配结果
    allocation_result = defaultdict(lambda: defaultdict(int))
    allocation_reasons = defaultdict(lambda: defaultdict(str))
    
    # 对每个SKU执行分配
    for sku_info in skus:
        sku = sku_info['sku']
        remaining_qty = sku_info['required_qty']
        core_sizes = [160, 165]
        max_single_size = 15
        
        # 收集数据
        store_data = {}
        for store in stores_sorted:
            inv = get_inventory(df_inventory, store, sku)
            sales_30d = get_30day_sales(df_sales, store, sku)
            level = get_store_level(df_store_level, store)
            
            total = sales_30d + inv
            if total > 0:
                sell_through = sales_30d / total
            else:
                sell_through = 0
            
            store_data[store] = {
                'inventory': inv,
                'sales_30d': sales_30d,
                'level': level,
                'sell_through': sell_through
            }
        
        # 1. 断码修复
        for store in stores_sorted:
            if remaining_qty <= 0:
                break
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
                to_allocate = target - current_inv
                to_allocate = min(to_allocate, remaining_qty)
                to_allocate = min(to_allocate, max_single_size - current_inv)
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',断码修复({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'断码修复({to_allocate})'
        
        # 2. 销量匹配
        for store in stores_sorted:
            if remaining_qty <= 0:
                break
            current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
            target_inv = int(store_data[store]['sales_30d'])
            
            if current_inv < target_inv:
                to_allocate = target_inv - current_inv
                to_allocate = min(to_allocate, remaining_qty)
                to_allocate = min(to_allocate, max_single_size - current_inv)
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',销量匹配({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'销量匹配({to_allocate})'
        
        # 3. B/C/D/OL级按销尽率降序
        bc_stores_sorted = []
        for level in ['B', 'C', 'D', 'OL']:
            level_stores = [(store, store_data[store]['sell_through']) 
                           for store in stores_sorted 
                           if store_data[store]['level'] == level]
            level_stores.sort(key=lambda x: x[1], reverse=True)
            bc_stores_sorted.extend([s[0] for s in level_stores])
        
        for store in bc_stores_sorted:
            if remaining_qty <= 0:
                break
            current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
            
            if current_inv < max_single_size:
                to_allocate = max_single_size - current_inv
                to_allocate = min(to_allocate, remaining_qty)
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',销尽率优先({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'销尽率优先({to_allocate})'
        
        # 4. 剩余分配
        for store in stores_sorted:
            if remaining_qty <= 0:
                break
            current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
            
            if current_inv < max_single_size:
                to_allocate = max_single_size - current_inv
                to_allocate = min(to_allocate, remaining_qty)
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',剩余分配({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'剩余分配({to_allocate})'
    
    return allocation_result, allocation_reasons, stores_sorted, skus

def generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus):
    # 生成分配数量的DataFrame
    data = []
    for store in stores_sorted:
        row = {'卖场': store}
        for sku_info in skus:
            row[sku_info['sku']] = allocation_result[store][sku_info['sku']]
        data.append(row)
    df_quantity = pd.DataFrame(data)
    
    # 生成分配原因的DataFrame
    reason_data = []
    for store in stores_sorted:
        row = {'卖场': store}
        for sku_info in skus:
            row[sku_info['sku']] = allocation_reasons[store][sku_info['sku']]
        reason_data.append(row)
    df_reason = pd.DataFrame(reason_data)
    
    return df_quantity, df_reason
