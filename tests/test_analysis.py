"""Tests for analysis tools: remove_duplicates, sort_data, filter_data, create_pivot_summary."""

from tools.analysis import remove_duplicates, sort_data, filter_data, create_pivot_summary


def test_remove_duplicates_guard_and_execution(sample_manager):
    # Intentional duplicates exist in sample_data (ORD-1005, ORD-1013)
    res_guard = remove_duplicates(sample_manager, sheet_name="Sales", subset_columns=["Order_ID"], confirmed=False)
    assert res_guard.get("requires_confirmation") is True
    assert res_guard["duplicate_count"] >= 2

    # Execute with confirmed=True
    res_exec = remove_duplicates(sample_manager, sheet_name="Sales", subset_columns=["Order_ID"], confirmed=True)
    assert res_exec["success"] is True
    assert res_exec["removed_count"] >= 2


def test_sort_data(sample_manager):
    res = sort_data(sample_manager, sheet_name="Sales", by_columns="Units", ascending=False)
    assert res["success"] is True


def test_filter_data_output_sheet(sample_manager):
    res = filter_data(sample_manager, sheet_name="Sales", column="Region", condition="equals", value="North", output_sheet="North_Sales")
    assert res["success"] is True
    assert res["matched_rows"] > 0
    assert "North_Sales" in sample_manager.get_sheet_names()


def test_create_pivot_summary(sample_manager):
    res = create_pivot_summary(
        sample_manager,
        sheet_name="Sales",
        index_columns=["Region"],
        values_column="Units",
        aggfunc="sum",
        output_sheet="Region_Units_Summary"
    )
    assert res["success"] is True
    assert "Region_Units_Summary" in sample_manager.get_sheet_names()
    assert res["rows_generated"] > 0
