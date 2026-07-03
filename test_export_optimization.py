# 验证导出分配原因含完整计算详情的功能
import sys
import os
import tempfile
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from allocation_core import allocate_add_order, generate_result_dataframe, DEFAULT_CONFIG
from calculation_tracker import CalculationTracker, format_store_detail_text


def create_test_data():
    df_inventory = pd.DataFrame({
        '卖场代码': ['S001', 'S002', 'S003'],
        '条码': ['SKU001', 'SKU001', 'SKU001'],
        '库存数量': [0, 2, 5]
    })
    df_sales = pd.DataFrame({
        '条码.条码': ['SKU001', 'SKU001', 'SKU001'],
        '店仓.卖场代码': ['S001', 'S002', 'S003'],
        '数量': [30, 15, 20]
    })
    df_store_level = pd.DataFrame({
        '代码': ['S001', 'S002', 'S003'],
        '卖场等级': ['SA', 'A', 'B']
    })
    df_add_order = pd.DataFrame({
        'SKU': ['SKU001'],
        'SKC': ['SKC001'],
        '需分配数量': [200]
    })
    return df_inventory, df_sales, df_store_level, df_add_order


def test_detail_reason_export():
    """测试导出的分配原因包含完整计算详情"""
    print("=== 测试导出分配原因含完整计算详情 ===")

    df_inv, df_sales, df_level, df_order = create_test_data()
    tracker = CalculationTracker()
    allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
        df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker
    )

    result_df, reason_df, stage_order_header = generate_result_dataframe(
        allocation_result, allocation_reasons, stores_sorted, skus, store_level_map
    )

    # 模拟 _build_detail_reason_df 的逻辑
    detail_df = reason_df.copy()
    sku_columns = [c for c in detail_df.columns if c not in ('卖场', '卖场等级')]

    for idx in detail_df.index:
        store_code = str(detail_df.at[idx, '卖场'])
        for sku in sku_columns:
            store_calc = tracker.get_store_detail(store_code, sku)
            if store_calc is not None:
                full_text = format_store_detail_text(store_calc, tracker.get_config_snapshot())
                detail_df.at[idx, sku] = full_text

    # 验证详情文本包含关键字段
    s001_detail = detail_df.loc[detail_df['卖场'] == 'S001', 'SKU001'].iloc[0]
    assert 'S001' in s001_detail, "应包含卖场代码"
    assert '断码修复' in s001_detail or '销量匹配' in s001_detail, "应包含阶段名称"
    assert '目标库存' in s001_detail or '分配数量' in s001_detail, "应包含计算公式"
    assert '已分配' in s001_detail, "应包含分配结果"

    print("\n--- S001 的完整计算详情 ---")
    print(s001_detail)

    # 验证写入Excel正常
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        out_path = tmp.name
    try:
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            result_df.to_excel(writer, sheet_name="分配数量", index=False)
            detail_df.to_excel(writer, sheet_name="分配原因", index=False)
        # 读取验证
        df_read = pd.read_excel(out_path, sheet_name="分配原因")
        assert len(df_read) == 3, "应有3行数据"
        assert 'SKU001' in df_read.columns, "应有SKU001列"
        s001_read = df_read.loc[df_read['卖场'] == 'S001', 'SKU001'].iloc[0]
        assert '断码修复' in s001_read or '销量匹配' in s001_read, "Excel中应包含阶段名称"
        print("\n✓ Excel导出验证通过")
    finally:
        if os.path.exists(out_path):
            os.remove(out_path)

    print("✓ 测试通过")


def test_search_filter_logic():
    """测试搜索过滤逻辑"""
    print("\n=== 测试搜索过滤逻辑 ===")

    df_inv, df_sales, df_level, df_order = create_test_data()
    tracker = CalculationTracker()
    allocation_result, allocation_reasons, stores_sorted, skus, store_level_map = allocate_add_order(
        df_inv, df_sales, df_level, df_order, config=DEFAULT_CONFIG, tracker=tracker
    )
    result_df, _, _ = generate_result_dataframe(
        allocation_result, allocation_reasons, stores_sorted, skus, store_level_map
    )

    # 模拟搜索 "S001"
    keyword = "S001"
    mask = result_df['卖场'].astype(str).str.contains(keyword, case=False, na=False)
    filtered = result_df[mask]
    assert len(filtered) == 1, f"搜索S001应返回1条，实际{len(filtered)}"
    assert filtered.iloc[0]['卖场'] == 'S001'

    # 模拟搜索 "S" (匹配所有)
    keyword = "S"
    mask = result_df['卖场'].astype(str).str.contains(keyword, case=False, na=False)
    filtered = result_df[mask]
    assert len(filtered) == 3, f"搜索S应返回3条，实际{len(filtered)}"

    # 模拟空搜索
    keyword = ""
    mask = result_df['卖场'].astype(str).str.contains(keyword, case=False, na=False)
    filtered = result_df[mask]
    assert len(filtered) == 3, f"空搜索应返回全部3条，实际{len(filtered)}"

    # 模拟无匹配搜索
    keyword = "XYZ"
    mask = result_df['卖场'].astype(str).str.contains(keyword, case=False, na=False)
    filtered = result_df[mask]
    assert len(filtered) == 0, f"无匹配搜索应返回0条，实际{len(filtered)}"

    print("✓ 搜索过滤逻辑测试通过")


if __name__ == '__main__':
    test_detail_reason_export()
    test_search_filter_logic()
    print("\n=== 所有测试通过 ===")
