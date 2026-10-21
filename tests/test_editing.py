"""Tests for editing tools: write_cells, find_and_replace, clean_data."""

from tools.editing import write_cells, find_and_replace, clean_data
from tools.inspection import read_range


def test_write_cells_and_verify(sample_manager):
    # Write a formula and a value
    updates = [
        {"cell": "H25", "value": "=SUM(H2:H23)"},
        {"cell": "A25", "value": "GRAND TOTAL"}
    ]
    res = write_cells(sample_manager, sheet_name="Sales", updates=updates)
    assert res["success"] is True
    assert res["updated_cells_count"] == 2
    # Verify post-write verification was performed
    assert "post_write_verification" in res
    assert res["post_write_verification"]["H25"] == "=SUM(H2:H23)"
    assert res["post_write_verification"]["A25"] == "GRAND TOTAL"


def test_find_and_replace(sample_manager):
    res = find_and_replace(sample_manager, sheet_name="Sales", find_text="Laptop", replace_text="Notebook")
    assert res["success"] is True
    assert res["replacements_count"] > 0

    # Verify that Laptop was replaced
    read_res = read_range(sample_manager, sheet_name="Sales", max_rows=30)
    flattened = [str(cell) for row in read_res["rows"] for cell in row]
    assert "Notebook" in flattened


def test_clean_data_trim_and_casing(sample_manager):
    # In sample_data, Fatima Noor has trailing space "Fatima Noor " and bilal ahmed is lowercase
    # Also verify that column name 'Sales_Rep' works as well as column letter 'D'
    res = clean_data(sample_manager, sheet_name="Sales", columns=["Sales_Rep"], trim_spaces=True, fix_casing="title")
    assert res["success"] is True
    assert res["cleaned_cells_count"] > 0

    read_res = read_range(sample_manager, sheet_name="Sales", cell_range="Sales!D1:D10")
    reps = [r[0] for r in read_res["rows"]]
    # Reps should now be properly titled with no trailing space
    for rep in reps[1:]:
        assert rep == rep.strip()
        assert rep == rep.title()


def test_write_cells_with_sheet_prefix(sample_manager):
    res = write_cells(sample_manager, cell="Sales!H26", value="=100*2")
    assert res["success"] is True
    assert res["post_write_verification"]["H26"] == "=100*2"
