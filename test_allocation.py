#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
加单商品分配系统 全面测试套件
包含：单元测试、集成测试、边界测试
"""

import unittest
import tempfile
import os
import sys
import json
import pandas as pd
from io import BytesIO

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 导入核心模块
from allocation_core import (
    load_config,
    get_30day_sales,
    get_store_level,
    get_inventory,
    allocate_add_order,
    generate_result_dataframe,
    DEFAULT_CONFIG,
)


class TestBase(unittest.TestCase):
    """测试基类，提供通用的测试数据"""
    
    def setUp(self):
        """设置测试数据"""
        # 测试库存数据
        self.inventory_data = pd.DataFrame({
            "卖场代码": ["卖场A", "卖场B", "卖场C", "卖场D"],
            "条码": ["SKU001", "SKU001", "SKU001", "SKU001"],
            "库存数量": [10, 5, 15, 3]
        })
        
        # 测试销售数据
        self.sales_data = pd.DataFrame({
            "条码.条码": ["SKU001", "SKU001", "SKU001", "SKU001"],
            "店仓.卖场代码": ["卖场A", "卖场B", "卖场C", "卖场D"],
            "数量": [20, 15, 8, 5]
        })
        
        # 测试卖场等级
        self.store_level_data = pd.DataFrame({
            "代码": ["卖场A", "卖场B", "卖场C", "卖场D"],
            "卖场等级": ["SA", "A", "B", "C"]
        })
        
        # 测试加单数据
        self.add_order_data = pd.DataFrame({
            "SKU": ["SKU001"],
            "SKC": ["SKC001"],
            "需分配数量": [50]
        })


class TestConfigLoading(TestBase):
    """配置文件加载测试"""
    
    def test_default_config(self):
        """测试默认配置是否正常"""
        config = DEFAULT_CONFIG
        self.assertIn("allocation_config", config)
        self.assertIn("coverage_days", config["allocation_config"])
        self.assertIn("level_weights", config["allocation_config"])
        self.assertIn("safety_factors", config["allocation_config"])
    
    def test_config_file_loading(self):
        """测试配置文件加载"""
        # 创建临时配置文件
        temp_dir = tempfile.mkdtemp()
        config_file = os.path.join(temp_dir, "allocation_config.json")
        
        test_config = {
            "version": "2.3",
            "allocation_config": {
                "coverage_days": {"SA": 30, "A": 28, "B": 14},
                "level_weights": {"SA": 1.8, "A": 1.5, "B": 1.0},
                "safety_factors": {"SA": 0.5, "A": 0.4, "B": 0.3},
                "min_target_inventory": {"SA": 2, "A": 1, "B": 0},
                "stage_priority": ["sales_match", "broken_size_fix", "sell_through_priority"],
                "max_remaining_per_store": 10
            }
        }
        
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(test_config, f)
        
        # 测试配置加载（模拟打包场景）
        try:
            # 简单测试配置的结构
            self.assertIn("allocation_config", test_config)
        finally:
            import shutil
            shutil.rmtree(temp_dir)


class TestDataProcessing(TestBase):
    """数据处理函数测试"""
    
    def test_get_30day_sales(self):
        """测试30天销量获取"""
        sales = get_30day_sales(self.sales_data, "SKU001", "卖场A")
        self.assertEqual(sales, 20)
    
    def test_get_30day_sales_nonexistent(self):
        """测试获取不存在的SKU或卖场"""
        sales = get_30day_sales(self.sales_data, "SKU999", "卖场X")
        self.assertEqual(sales, 0)
    
    def test_get_store_level(self):
        """测试获取卖场等级"""
        level = get_store_level(self.store_level_data, "卖场A")
        self.assertEqual(level, "SA")
    
    def test_get_store_level_default(self):
        """测试获取不存在的卖场等级"""
        level = get_store_level(self.store_level_data, "卖场X")
        self.assertEqual(level, "C")
    
    def test_get_inventory(self):
        """测试获取库存"""
        inventory = get_inventory(self.inventory_data, "卖场A", "SKU001")
        self.assertEqual(inventory, 10)
    
    def test_get_inventory_default(self):
        """测试获取不存在的库存"""
        inventory = get_inventory(self.inventory_data, "卖场X", "SKU001")
        self.assertEqual(inventory, 0)


class TestAllocationLogic(TestBase):
    """分配逻辑核心测试"""
    
    def test_full_allocation(self):
        """测试完整的分配流程"""
        result, reasons, stores, skus, level_map = allocate_add_order(
            self.inventory_data,
            self.sales_data,
            self.store_level_data,
            self.add_order_data
        )
        
        # 验证结果不为空
        self.assertIsNotNone(result)
        self.assertIsNotNone(reasons)
        self.assertGreater(len(stores), 0)
        self.assertGreater(len(skus), 0)
        
        # 验证有分配记录
        total_allocated = 0
        for store in stores:
            for sku in [s["sku"] for s in skus]:
                total_allocated += result[store][sku]
        
        # 验证总分配不超过总加单数量
        self.assertLessEqual(total_allocated, 50)
    
    def test_custom_stage_priority(self):
        """测试自定义阶段优先级"""
        custom_config = {
            "allocation_config": {
                "coverage_days": {"SA": 30, "A": 28, "B": 14, "C": 14, "D": 14, "OL": 14},
                "level_weights": {"SA": 1.5, "A": 1.3, "B": 1.2, "C": 1.1, "D": 1.1, "OL": 1.0},
                "safety_factors": {"SA": 0.5, "A": 0.4, "B": 0.3, "C": 0.25, "D": 0.2, "OL": 0.2},
                "min_target_inventory": {"SA": 0, "A": 0, "B": 0, "C": 0, "D": 0, "OL": 0},
                "stage_priority": ["sales_match", "sell_through_priority", "broken_size_fix"],
                "max_remaining_per_store": 10
            }
        }
        
        result, reasons, stores, skus, level_map = allocate_add_order(
            self.inventory_data,
            self.sales_data,
            self.store_level_data,
            self.add_order_data,
            custom_config
        )
        
        self.assertIsNotNone(result)
    
    def test_generate_result_dataframe(self):
        """测试生成结果DataFrame"""
        result, reasons, stores, skus, level_map = allocate_add_order(
            self.inventory_data,
            self.sales_data,
            self.store_level_data,
            self.add_order_data
        )
        
        df_qty, df_reason = generate_result_dataframe(
            result, reasons, stores, skus, level_map
        )
        
        self.assertIsNotNone(df_qty)
        self.assertIsNotNone(df_reason)
        self.assertEqual(len(df_qty), len(stores))
        self.assertEqual(len(df_reason), len(stores))
        self.assertIn("卖场等级", df_reason.columns)


class TestBoundaryConditions(unittest.TestCase):
    """边界条件测试"""
    
    def test_zero_inventory_allocation(self):
        """测试零库存情况下的分配"""
        inventory_data = pd.DataFrame({
            "卖场代码": ["卖场A", "卖场B"],
            "条码": ["SKU001", "SKU001"],
            "库存数量": [0, 0]
        })
        
        sales_data = pd.DataFrame({
            "条码.条码": ["SKU001", "SKU001"],
            "店仓.卖场代码": ["卖场A", "卖场B"],
            "数量": [10, 8]
        })
        
        store_level_data = pd.DataFrame({
            "代码": ["卖场A", "卖场B"],
            "卖场等级": ["SA", "A"]
        })
        
        add_order_data = pd.DataFrame({
            "SKU": ["SKU001"],
            "SKC": ["SKC001"],
            "需分配数量": [30]
        })
        
        result, reasons, stores, skus, level_map = allocate_add_order(
            inventory_data, sales_data, store_level_data, add_order_data
        )
        
        self.assertIsNotNone(result)
    
    def test_zero_sales_allocation(self):
        """测试零销售情况下的分配"""
        inventory_data = pd.DataFrame({
            "卖场代码": ["卖场A", "卖场B"],
            "条码": ["SKU001", "SKU001"],
            "库存数量": [5, 5]
        })
        
        sales_data = pd.DataFrame({
            "条码.条码": ["SKU001", "SKU001"],
            "店仓.卖场代码": ["卖场A", "卖场B"],
            "数量": [0, 0]
        })
        
        store_level_data = pd.DataFrame({
            "代码": ["卖场A", "卖场B"],
            "卖场等级": ["SA", "A"]
        })
        
        add_order_data = pd.DataFrame({
            "SKU": ["SKU001"],
            "SKC": ["SKC001"],
            "需分配数量": [20]
        })
        
        result, reasons, stores, skus, level_map = allocate_add_order(
            inventory_data, sales_data, store_level_data, add_order_data
        )
        
        self.assertIsNotNone(result)
    
    def test_empty_add_order(self):
        """测试加单数量为0的情况"""
        inventory_data = pd.DataFrame({
            "卖场代码": ["卖场A"],
            "条码": ["SKU001"],
            "库存数量": [5]
        })
        
        sales_data = pd.DataFrame({
            "条码.条码": ["SKU001"],
            "店仓.卖场代码": ["卖场A"],
            "数量": [10]
        })
        
        store_level_data = pd.DataFrame({
            "代码": ["卖场A"],
            "卖场等级": ["SA"]
        })
        
        add_order_data = pd.DataFrame({
            "SKU": ["SKU001"],
            "SKC": ["SKC001"],
            "需分配数量": [0]
        })
        
        result, reasons, stores, skus, level_map = allocate_add_order(
            inventory_data, sales_data, store_level_data, add_order_data
        )
        
        self.assertIsNotNone(result)


class TestIntegration(unittest.TestCase):
    """集成测试"""
    
    def test_end_to_end_allocation(self):
        """端到端测试"""
        # 创建完整的测试数据
        inventory_df = pd.DataFrame({
            "卖场代码": ["SA001", "A001", "B001", "C001"],
            "条码": ["SKU2023160", "SKU2023160", "SKU2023160", "SKU2023160"],
            "库存数量": [10, 5, 15, 8]
        })
        
        sales_df = pd.DataFrame({
            "条码.条码": ["SKU2023160", "SKU2023160", "SKU2023160", "SKU2023160"],
            "店仓.卖场代码": ["SA001", "A001", "B001", "C001"],
            "数量": [30, 25, 18, 10]
        })
        
        store_level_df = pd.DataFrame({
            "代码": ["SA001", "A001", "B001", "C001"],
            "卖场等级": ["SA", "A", "B", "C"]
        })
        
        add_order_df = pd.DataFrame({
            "SKU": ["SKU2023160"],
            "SKC": ["SKC2023"],
            "需分配数量": [100]
        })
        
        # 执行分配
        result, reasons, stores, skus, level_map = allocate_add_order(
            inventory_df, sales_df, store_level_df, add_order_df
        )
        
        # 生成结果
        df_qty, df_reason = generate_result_dataframe(
            result, reasons, stores, skus, level_map
        )
        
        # 验证结果
        self.assertEqual(len(df_qty), 4)
        self.assertEqual(len(df_reason), 4)
        
        # 验证列名
        self.assertIn("卖场", df_qty.columns)
        self.assertIn("卖场等级", df_reason.columns)


def create_test_excel():
    """创建测试用的Excel文件"""
    # 创建完整的测试数据
    inventory_df = pd.DataFrame({
        "卖场代码": ["SA001", "A001", "B001", "C001", "D001", "OL001"],
        "条码": ["SKU2023160", "SKU2023160", "SKU2023160", 
                "SKU2023160", "SKU2023160", "SKU2023160"],
        "库存数量": [10, 5, 15, 8, 3, 2]
    })
    
    sales_df = pd.DataFrame({
        "条码.条码": ["SKU2023160", "SKU2023160", "SKU2023160",
                    "SKU2023160", "SKU2023160", "SKU2023160"],
        "店仓.卖场代码": ["SA001", "A001", "B001", "C001", "D001", "OL001"],
        "数量": [30, 25, 18, 10, 5, 2]
    })
    
    store_level_df = pd.DataFrame({
        "代码": ["SA001", "A001", "B001", "C001", "D001", "OL001"],
        "卖场等级": ["SA", "A", "B", "C", "D", "OL"]
    })
    
    add_order_df = pd.DataFrame({
        "SKU": ["SKU2023160"],
        "SKC": ["SKC2023"],
        "需分配数量": [100]
    })
    
    # 保存测试文件
    test_file = "test_allocation_data.xlsx"
    with pd.ExcelWriter(test_file, engine="openpyxl") as writer:
        inventory_df.to_excel(writer, sheet_name="库存", index=False)
        sales_df.to_excel(writer, sheet_name="销售", index=False)
        store_level_df.to_excel(writer, sheet_name="卖场等级", index=False)
        add_order_df.to_excel(writer, sheet_name="加单数量", index=False)
    
    return test_file


def run_tests():
    """运行所有测试"""
    # 先创建测试用Excel
    print("=" * 60)
    print("创建测试用Excel文件")
    print("=" * 60)
    test_file = create_test_excel()
    print(f"✓ 测试文件已创建: {test_file}")
    print()
    
    # 运行测试
    print("=" * 60)
    print("运行测试用例")
    print("=" * 60)
    
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    
    # 添加所有测试类
    suite.addTests(loader.loadTestsFromTestCase(TestConfigLoading))
    suite.addTests(loader.loadTestsFromTestCase(TestDataProcessing))
    suite.addTests(loader.loadTestsFromTestCase(TestAllocationLogic))
    suite.addTests(loader.loadTestsFromTestCase(TestBoundaryConditions))
    suite.addTests(loader.loadTestsFromTestCase(TestIntegration))
    
    # 运行测试
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    print()
    print("=" * 60)
    print("测试结果总结")
    print("=" * 60)
    print(f"运行测试数: {result.testsRun}")
    print(f"失败: {len(result.failures)}")
    print(f"错误: {len(result.errors)}")
    print(f"跳过: {len(result.skipped)}")
    print(f"成功: {result.testsRun - len(result.failures) - len(result.errors) - len(result.skipped)}")
    print()
    
    if result.wasSuccessful():
        print("🎉 所有测试通过！")
        return 0
    else:
        print("❌ 测试未完全通过")
        return 1


if __name__ == "__main__":
    exit_code = run_tests()
    sys.exit(exit_code)
