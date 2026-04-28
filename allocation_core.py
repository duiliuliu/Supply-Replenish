# 加单分配核心逻辑 v2.0
import pandas as pd
import json
import os
from collections import defaultdict

DEFAULT_CONFIG = {
    "version": "2.0",
    "allocation_config": {
        "coverage_days": {
            "SA": 30, "A": 30, "B": 14, "C": 14, "D": 14, "OL": 14
        },
        "level_weights": {
            "SA": 1.5, "A": 1.3, "B": 1.2, "C": 1.1, "D": 1.1, "OL": 1.0
        },
        "safety_factors": {
            "SA": 0.5, "A": 0.4, "B": 0.3, "C": 0.25, "D": 0.2, "OL": 0.2
        },
        "max_remaining_per_store": 10
    }
}

def load_config():
    config_path = os.path.join(os.path.dirname(__file__), 'allocation_config.json')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return DEFAULT_CONFIG

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

def allocate_add_order(df_inventory, df_sales, df_store_level, df_add_order, config=None):
    if config is None:
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
    
    allocation_result = defaultdict(lambda: defaultdict(int))
    allocation_reasons = defaultdict(lambda: defaultdict(str))
    
    for sku_info in skus:
        sku = sku_info['sku']
        remaining_qty = sku_info['required_qty']
        core_sizes = [160, 165]
        
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
        
        # 阶段1: 断码修复
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
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',断码修复({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'断码修复({to_allocate})'
        
        # 阶段2: 销量匹配（基于供应链公式）
        # 设置最小目标库存，确保即使销量为0也有基础库存
        min_target_by_level = {'SA': 3, 'A': 2, 'B': 2, 'C': 2, 'D': 1, 'OL': 1}
        
        for store in stores_sorted:
            if remaining_qty <= 0:
                break
            
            level = store_data[store]['level']
            sales_30d = store_data[store]['sales_30d']
            daily_demand = sales_30d / 30
            
            coverage = coverage_days.get(level, 14)
            safety_factor = safety_factors.get(level, 0.3)
            safety_stock = daily_demand * safety_factor * coverage
            target_inv = int(daily_demand * coverage + safety_stock)
            
            # 应用最小目标库存
            min_target = min_target_by_level.get(level, 2)
            target_inv = max(target_inv, min_target)
            
            current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
            
            if current_inv < target_inv:
                to_allocate = target_inv - current_inv
                to_allocate = min(to_allocate, remaining_qty)
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',销量匹配({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'销量匹配({to_allocate})'
        
        # 阶段3: 销尽率优先分配（所有等级参与，按权重排序）
        all_stores_with_score = []
        for store in stores_sorted:
            level = store_data[store]['level']
            sell_through = store_data[store]['sell_through']
            weight = level_weights.get(level, 1.0)
            weighted_score = sell_through * weight
            all_stores_with_score.append((store, weighted_score, level))
        
        all_stores_with_score.sort(key=lambda x: x[1], reverse=True)
        
        for store, score, level in all_stores_with_score:
            if remaining_qty <= 0:
                break
            
            sales_30d = store_data[store]['sales_30d']
            current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
            weight = level_weights.get(level, 1.0)
            
            max_for_store = max(int(sales_30d * weight), 2)
            
            if current_inv < max_for_store:
                to_allocate = max_for_store - current_inv
                to_allocate = min(to_allocate, remaining_qty)
                
                if to_allocate > 0:
                    allocation_result[store][sku] += to_allocate
                    remaining_qty -= to_allocate
                    if allocation_reasons[store][sku]:
                        allocation_reasons[store][sku] += f',销尽率优先({to_allocate})'
                    else:
                        allocation_reasons[store][sku] = f'销尽率优先({to_allocate})'
        
        # 阶段4: 剩余分配（按等级优先级）
        for level in level_order:
            if remaining_qty <= 0:
                break
            
            level_stores = [s for s in stores_sorted if store_data[s]['level'] == level]
            
            if level_stores:
                for store in level_stores:
                    if remaining_qty <= 0:
                        break
                    
                    current_inv = store_data[store]['inventory'] + allocation_result[store][sku]
                    
                    if current_inv < max_remaining_per_store:
                        to_allocate = max_remaining_per_store - current_inv
                        to_allocate = min(to_allocate, remaining_qty)
                        
                        if to_allocate > 0:
                            allocation_result[store][sku] += to_allocate
                            remaining_qty -= to_allocate
                            if allocation_reasons[store][sku]:
                                allocation_reasons[store][sku] += f',剩余分配({to_allocate})'
                            else:
                                allocation_reasons[store][sku] = f'剩余分配({to_allocate})'
    
    return allocation_result, allocation_reasons, stores_sorted, skus, store_level_map

def generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus, store_level_map=None):
    data = []
    for store in stores_sorted:
        row = {'卖场': store}
        for sku_info in skus:
            row[sku_info['sku']] = allocation_result[store][sku_info['sku']]
        data.append(row)
    df_quantity = pd.DataFrame(data)
    
    reason_data = []
    for store in stores_sorted:
        level = store_level_map.get(store, '未知') if store_level_map else '未知'
        row = {'卖场': store, '卖场等级': level}
        for sku_info in skus:
            row[sku_info['sku']] = allocation_reasons[store][sku_info['sku']]
        reason_data.append(row)
    df_reason = pd.DataFrame(reason_data)
    
    return df_quantity, df_reason
