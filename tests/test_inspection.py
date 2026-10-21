"""Tests for inspection tools: list_sheets, read_range, get_sheet_summary."""

from tools.inspection import list_sheets, read_range, get_sheet_summary


def test_list_sheets(sample_manager):
    res = list_sheets(sample_manager)
    assert "sheets" in res
    assert res["total_sheets"] >= 2
    names = [s["name"] for s in res["sheets"]]
    assert "Sales" in names
    assert "Targets" in names


def test_read_range_with_formula(sample_manager):
    # Read row 2 which has formula =F2*G2
    res = read_range(sample_manager, sheet_name="Sales", cell_range="H2", show_formulas=True)
    assert res["row_count"] == 1
    assert "=F2*G2" in str(res["rows"][0][0])


def test_read_range_all_and_limit(sample_manager):
    res = read_range(sample_manager, sheet_name="Sales", max_rows=5)
    assert res["row_count"] == 5
    assert len(res["rows"]) == 5
    assert res["truncated"] is True


def test_get_sheet_summary(sample_manager):
    res = get_sheet_summary(sample_manager, sheet_name="Sales")
    assert res["total_rows"] >= 20
    assert res["total_columns"] >= 9
    col_names = [c["name"] for c in res["columns"]]
    assert "Order_ID" in col_names
    assert "Total_Revenue" in col_names
    assert len(res["sample_rows"]) > 0


def test_read_range_invalid_sheet(sample_manager):
    res = read_range(sample_manager, sheet_name="NonExistentSheet")
    assert "error" in res
