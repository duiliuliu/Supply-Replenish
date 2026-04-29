# 加单分配核心逻辑 v2.6.0
import pandas as pd
import json
import os
import sys
from collections import defaultdict

try:
    import tomllib
except ImportError:
    try:
        import tomli as tomllib
    except ImportError:
        tomllib = None

def get_version():
    """从 pyproject.toml 读取版本号"""
    try:
        pyproject_paths = []
        
        if getattr(sys, 'frozen', False):
            pyproject_paths.append(os.path.join(sys._MEIPASS, 'pyproject.toml'))
            if sys.platform == 'darwin':
                content_path = os.path.dirname(sys.executable)
                pyproject_paths.append(os.path.join(content_path, 'pyproject.toml'))
            else:
                pyproject_paths.append(os.path.join(os.path.dirname(sys.executable), 'pyproject.toml'))
        
        pyproject_paths.append('pyproject.toml')
        
        try:
            script_path = os.path.dirname(os.path.abspath(__file__))
            pyproject_paths.append(os.path.join(script_path, 'pyproject.toml'))
        except:
            pass
        
        for pyproject_path in pyproject_paths:
            try:
                if os.path.exists(pyproject_path):
                    if tomllib is not None:
                        with open(pyproject_path, 'rb') as f:
                            data = tomllib.load(f)
                            version = data.get('project', {}).get('version', '2.6.0')
                            print(f'Loaded version from {pyproject_path}: {version}')
                            return version
                    else:
                        import re
                        with open(pyproject_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            match = re.search(r'version\s*=\s*["\']([^"\']+)["\']', content)
                            if match:
                                version = match.group(1)
                                print(f'Loaded version from {pyproject_path}: {version}')
                                return version
            except Exception as e:
                continue
        
        return '2.6.0'
    except Exception as e:
        print(f'Error reading version: {e}')
        return '2.6.0'

VERSION = get_version()

DEFAULT_CONFIG = {
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
        "min_target_inventory": {
            "SA": 0, "A": 0, "B": 0, "C": 0, "D": 0, "OL": 0
        },
        "stage_priority": [
            "broken_size_fix",
            "sales_match",
            "sell_through_priority"
        ],
        "max_remaining_per_store": 10
    }
}

def load_config():
    """加载配置文件，包含多个路径尝试"""
    config_paths = []
    
    # 先尝试PyInstaller打包时的路径
    if getattr(sys, 'frozen', False):
        config_paths.append(os.path.join(sys._MEIPASS, 'allocation_config.json'))
        if sys.platform == 'darwin':
            # Mac .app的Contents目录
            content_path = os.path.dirname(sys.executable)
            config_paths.append(os.path.join(content_path, 'allocation_config.json'))
        else:
            config_paths.append(os.path.join(os.path.dirname(sys.executable), 'allocation_config.json'))
    
    # 当前工作目录
    config_paths.append('allocation_config.json')
    
    # 源代码目录
    try:
        script_path = os.path.dirname(os.path.abspath(__file__))
        config_paths.append(os.path.join(script_path, 'allocation_config.json'))
    except:
        pass
    
    print(f'Trying to load config from: {config_paths}')
    
    for config_path in config_paths:
        try:
            if os.path.exists(config_path):
                print(f'Loading config from {config_path}')
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    print('Config loaded successfully')
                    return config
        except Exception as e:
            print(f'Warning: Error reading config from {config_path}: {e}')
            continue
    
    # 如果找不到配置文件，使用默认配置
    print('Using default config')
    return DEFAULT_CONFIG

def get_30day_sales(df_sales, sku, store_code):
    try:
        df_filtered = df_sales[(df_sales['条码.条码'] == sku) & (df_sales['店仓.卖场代码'] == store_code)].copy()
        df_filtered['数量'] = df_filtered['数量'].apply(lambda x: max(0, x))
        return df_filtered['数量'].sum()
    except Exception as e:
        return 0

def extract_size(sku):
    try:
        return int(str(sku)[-3:])
    except:
        return 0

def is_core_size(size):
    return size in [160, 165]

def get_store_level(df_store_level, store_code):
    try:
        filtered = df_store_level[df_store_level['代码'] == store_code]
        if len(filtered) > 0:
            return filtered.iloc[0]['卖场等级']
    except Exception as e:
        print(f'Warning: Error getting store level for {store_code}: {e}')
    return 'C'

def get_inventory(df_inventory, store_code, sku):
    try:
        filtered = df_inventory[(df_inventory['卖场代码'] == store_code) & (df_inventory['条码'] == sku)]
        if len(filtered) > 0:
            return max(0, int(filtered.iloc[0]['库存数量']))
    except Exception as e:
        print(f'Warning: Error getting inventory for {store_code} - {sku}: {e}')
    return 0

def stage_broken_size_fix(stores_sorted, store_data, sku, allocation_result, allocation_reasons, remaining_qty):
    """阶段1: 断码修复"""
    try:
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
        return remaining_qty
    except Exception as e:
        print(f'Warning: Error in broken size fix stage: {e}')
        return remaining_qty

def stage_sales_match(stores_sorted, store_data, sku, allocation_result, allocation_reasons, remaining_qty, 
                     coverage_days, safety_factors, min_target_inventory):
    """阶段2: 销量匹配（基于供应链公式）"""
    try:
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
            
            min_target = min_target_inventory.get(level, 0)
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
        return remaining_qty
    except Exception as e:
        print(f'Warning: Error in sales match stage: {e}')
        return remaining_qty

def stage_sell_through_priority(stores_sorted, store_data, sku, allocation_result, allocation_reasons, remaining_qty, level_weights):
    """阶段3: 销尽率优先分配（所有等级参与，按权重排序）"""
    try:
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
        return remaining_qty
    except Exception as e:
        print(f'Warning: Error in sell through priority stage: {e}')
        return remaining_qty

def stage_remaining_allocation(stores_sorted, store_data, sku, allocation_result, allocation_reasons, remaining_qty, 
                               level_order, max_remaining_per_store, store_level_map):
    """阶段4: 剩余分配（按等级优先级）"""
    try:
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
        return remaining_qty
    except Exception as e:
        print(f'Warning: Error in remaining allocation stage: {e}')
        return remaining_qty

def allocate_add_order(df_inventory, df_sales, df_store_level, df_add_order, config=None):
    try:
        if config is None:
            config = load_config()
        
        alloc_config = config.get('allocation_config', DEFAULT_CONFIG['allocation_config'])
        coverage_days = alloc_config.get('coverage_days', DEFAULT_CONFIG['allocation_config']['coverage_days'])
        level_weights = alloc_config.get('level_weights', DEFAULT_CONFIG['allocation_config']['level_weights'])
        safety_factors = alloc_config.get('safety_factors', DEFAULT_CONFIG['allocation_config']['safety_factors'])
        min_target_inventory = alloc_config.get('min_target_inventory', DEFAULT_CONFIG['allocation_config']['min_target_inventory'])
        stage_priority = alloc_config.get('stage_priority', DEFAULT_CONFIG['allocation_config']['stage_priority'])
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
        
        stage_map = {
            'broken_size_fix': lambda *args: stage_broken_size_fix(*args[:6]),
            'sales_match': lambda *args: stage_sales_match(*args[:6], coverage_days, safety_factors, min_target_inventory),
            'sell_through_priority': lambda *args: stage_sell_through_priority(*args[:6], level_weights)
        }
        
        for sku_info in skus:
            sku = sku_info['sku']
            remaining_qty = sku_info['required_qty']
            core_sizes = [160, 165]
            
            store_data = {}
            for store in stores_sorted:
                inv = get_inventory(df_inventory, store, sku)
                sales_30d = get_30day_sales(df_sales, sku, store)
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
            
            for stage_name in stage_priority:
                if remaining_qty <= 0:
                    break
                if stage_name in stage_map:
                    remaining_qty = stage_map[stage_name](stores_sorted, store_data, sku, allocation_result, 
                                                           allocation_reasons, remaining_qty)
            
            if remaining_qty > 0:
                remaining_qty = stage_remaining_allocation(stores_sorted, store_data, sku, allocation_result, 
                                                          allocation_reasons, remaining_qty, level_order, 
                                                          max_remaining_per_store, store_level_map)
        
        return allocation_result, allocation_reasons, stores_sorted, skus, store_level_map
    except Exception as e:
        print(f'Error in allocate_add_order: {e}')
        import traceback
        traceback.print_exc()
        return defaultdict(lambda: defaultdict(int)), defaultdict(lambda: defaultdict(str)), [], [], {}

def generate_result_dataframe(allocation_result, allocation_reasons, stores_sorted, skus, store_level_map=None):
    try:
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
    except Exception as e:
        print(f'Error in generate_result_dataframe: {e}')
        import traceback
        traceback.print_exc()
        return pd.DataFrame(), pd.DataFrame()
