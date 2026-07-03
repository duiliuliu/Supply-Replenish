# allocation_core.py 集成追踪功能测试
import sys
import os
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from allocation_core import allocate_add_order, DEFAULT_CONFIG
from calculation_tracker import CalculationTracker, format_store_detail_text, format_flow_overview_text


def create_test_data():
    """创建测试数据"""
    df_inventory = pd.DataFrame({
        '卖场代码': ['S001', 'S002', 'S003', 'S004', 'S005'],
        '条码': ['SKU001', 'SKU001', 'SKU001', 'SKU001', 'SKU001'],
        '库存数量': [0, 2, 5, 0, 1]
    })

    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001', 'SKU001', 'SKU001', 'SKU001', 'SKU001'],
        '店仓.卖场代码': ['S001', 'S002', 'S003', 'S004', 'S005'],
        '数量': [30, 15, 20, 0, 10]
    })

    df_store_level = pd.DataFrame({
        '代码': ['S001', 'S002', 'S003', 'S004', 'S005'],
        '卖场等级': ['SA', 'A', 'B', 'C', 'D']
    })

    df_add_order = pd.DataFrame({
        'SKU': ['SKU001'],
        'SKC': ['SKC001'],
        '需分配数量': [200]
    })

    return df_inventory, df_sales, df_store_level, df_add_order


def test_backward_compatibility():
    """测试向后兼容性（不传tracker）"""
    print("\n=== 测试向后兼容性 ===")
    df_inv, df_sales, df_level, df_order = create_test_data()

    # 不传tracker，应该正常工作
    result = allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG)
    assert len(result) == 5, "应返回5个值"

    allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = result
    assert len(stores_sorted) == 5
    assert len(skus) == 1

    # 验证总分配量（各阶段有上限，可能不等于需分配量）
    total = sum(allocation_result[store]['SKU001'] for store in stores_sorted)
    assert total > 0, f"总分配量应大于0，实际为{total}"
    assert total <= 200, f"总分配量不应超过需分配量200，实际为{total}"

    print("  ✓ 向后兼容性测试通过")


def test_tracking_integration():
    """测试追踪功能集成"""
    print("\n=== 测试追踪功能集成 ===")
    df_inv, df_sales, df_level, df_order = create_test_data()

    # 创建tracker并传入
    tracker = CalculationTracker()
    result = allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker)

    allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = result

    # 验证tracker记录
    assert len(tracker.sku_trackings) == 1, "应追踪1个SKU"
    assert 'SKU001' in tracker.sku_trackings

    tracking = tracker.get_sku_tracking('SKU001')
    assert tracking is not None
    assert tracking.required_qty == 200

    # 验证有4个阶段
    assert len(tracking.stages) == 4, f"应有4个阶段，实际{len(tracking.stages)}"

    # 验证配置快照
    config_snap = tracker.get_config_snapshot()
    assert 'coverage_days' in config_snap
    assert 'safety_factors' in config_snap

    print("  ✓ 追踪功能集成测试通过")


def test_flow_overview_generation():
    """测试流程概览生成"""
    print("\n=== 测试流程概览生成 ===")
    df_inv, df_sales, df_level, df_order = create_test_data()

    tracker = CalculationTracker()
    allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker)

    overview = tracker.generate_flow_overview('SKU001')
    assert overview is not None
    assert overview['sku'] == 'SKU001'
    assert overview['required_qty'] == 200
    assert overview['total_allocated'] > 0, f"总分配应大于0，实际{overview['total_allocated']}"
    assert overview['total_allocated'] <= 200, f"总分配不应超过200，实际{overview['total_allocated']}"
    assert len(overview['stages']) == 4

    # 验证各阶段有分配记录
    for stage in overview['stages']:
        print(f"  阶段{stage['order']} - {stage['stage_name']}: 分配{stage['allocated']}件 ({stage['percentage']}%), 卖场数{stage['store_count']}")

    print("  ✓ 流程概览生成测试通过")


def test_store_detail_lookup():
    """测试卖场计算详情查询"""
    print("\n=== 测试卖场计算详情查询 ===")
    df_inv, df_sales, df_level, df_order = create_test_data()

    tracker = CalculationTracker()
    allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker)

    # 查询有分配的卖场详情
    for store in ['S001', 'S002', 'S003', 'S004', 'S005']:
        store_calc = tracker.get_store_detail(store, 'SKU001')
        if store_calc:
            print(f"\n  --- 卖场 {store} (等级: {store_calc.level}) ---")
            print(f"  初始库存: {store_calc.initial_inventory}, 30天销量: {store_calc.sales_30d}")
            print(f"  总分配: {store_calc.total_allocation}件, 最终库存: {store_calc.final_inventory}件")
            print(f"  阶段数: {store_calc.stage_count}")

            for stage in store_calc.stages:
                if stage.skipped:
                    print(f"    {stage.stage_display_name}: 跳过 ({stage.skip_reason})")
                else:
                    print(f"    {stage.stage_display_name}: 分配{stage.allocated}件")
                    for detail in stage.details:
                        print(f"      {detail['name']}: {detail['value']}")

    # 验证S001（SA级，库存0，销量30）应有分配
    store_calc = tracker.get_store_detail('S001', 'SKU001')
    assert store_calc is not None
    assert store_calc.level == 'SA'
    assert store_calc.total_allocation > 0, "S001应有分配"

    print("\n  ✓ 卖场计算详情查询测试通过")


def test_formatted_text_output():
    """测试格式化文本输出"""
    print("\n=== 测试格式化文本输出 ===")
    df_inv, df_sales, df_level, df_order = create_test_data()

    tracker = CalculationTracker()
    allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker)

    # 测试流程概览文本
    overview = tracker.generate_flow_overview('SKU001')
    overview_text = format_flow_overview_text(overview)
    assert 'SKU001' in overview_text
    assert '断码修复' in overview_text
    assert '销量匹配' in overview_text

    print("\n  --- 流程概览文本 ---")
    print(overview_text)

    # 测试卖场详情文本
    store_calc = tracker.get_store_detail('S001', 'SKU001')
    if store_calc:
        detail_text = format_store_detail_text(store_calc)
        assert 'S001' in detail_text
        assert 'SA' in detail_text

        print("\n  --- 卖场S001计算详情 ---")
        print(detail_text)

    print("\n  ✓ 格式化文本输出测试通过")


def test_allocation_consistency():
    """测试追踪结果与分配结果的一致性"""
    print("\n=== 测试追踪结果与分配结果一致性 ===")
    df_inv, df_sales, df_level, df_order = create_test_data()

    # 不带tracker的分配
    result1 = allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG)
    alloc1, _, stores1, _, _ = result1

    # 带tracker的分配
    tracker = CalculationTracker()
    result2 = allocate_add_order(df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker)
    alloc2, _, stores2, _, _ = result2

    # 验证分配结果一致
    for store in stores1:
        for sku in ['SKU001']:
            val1 = alloc1[store][sku]
            val2 = alloc2[store][sku]
            assert val1 == val2, f"卖场{store}的{sku}分配不一致: {val1} vs {val2}"

    # 验证追踪的总分配量与实际分配量一致
    tracking = tracker.get_sku_tracking('SKU001')
    actual_total = sum(alloc2[store]['SKU001'] for store in stores2)
    assert tracking.total_allocated == actual_total, f"追踪总量{tracking.total_allocated}与实际{actual_total}不一致"

    print("  ✓ 追踪结果与分配结果一致性测试通过")


def run_all_tests():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("  allocation_core.py 集成追踪功能测试")
    print("=" * 60)

    tests = [
        test_backward_compatibility,
        test_tracking_integration,
        test_flow_overview_generation,
        test_store_detail_lookup,
        test_formatted_text_output,
        test_allocation_consistency
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
