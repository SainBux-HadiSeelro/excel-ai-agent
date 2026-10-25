"""Tests for merge_workbooks and undo_last_change."""

from pathlib import Path
from tools.workbook_ops import merge_workbooks, undo_last_change
from tools.structure import create_sheet
from tools.editing import write_cells


def test_merge_workbooks(sample_manager):
    extra_file = Path("sample_data/regional_extra.xlsx")
    assert extra_file.exists()

    res = merge_workbooks(sample_manager, source_files=[str(extra_file)], mode="add_sheets")
    assert res["success"] is True
    assert len(res["imported_sheets"]) > 0
    assert any("Q1_Summary" in s for s in sample_manager.get_sheet_names())


def test_undo_last_change(sample_manager):
    initial_sheets = list(sample_manager.get_sheet_names())

    # Perform a modifying action
    create_sheet(sample_manager, title="UndoTestSheet")
    assert "UndoTestSheet" in sample_manager.get_sheet_names()

    # Perform undo
    undo_res = undo_last_change(sample_manager)
    assert undo_res["success"] is True
    assert "UndoTestSheet" not in sample_manager.get_sheet_names()
    assert sample_manager.get_sheet_names() == initial_sheets


def test_undo_multiple_mutations(sample_manager):
    wb = sample_manager.open_workbook()
    original_val = wb.active["A1"].value
    assert sample_manager.get_history_count() == 0

    # Mutation 1: write A1=1
    write_cells(sample_manager, updates=[{"cell": "A1", "value": 1}])
    wb = sample_manager.open_workbook()
    assert wb.active["A1"].value == 1
    assert sample_manager.get_history_count() == 1

    # Mutation 2: write A1=2
    write_cells(sample_manager, updates=[{"cell": "A1", "value": 2}])
    wb = sample_manager.open_workbook()
    assert wb.active["A1"].value == 2
    assert sample_manager.get_history_count() == 2

    # Undo once: assert A1 == 1
    res1 = undo_last_change(sample_manager)
    assert res1["success"] is True
    wb = sample_manager.open_workbook()
    assert wb.active["A1"].value == 1
    assert sample_manager.get_history_count() == 1

    # Undo again: assert original value
    res2 = undo_last_change(sample_manager)
    assert res2["success"] is True
    wb = sample_manager.open_workbook()
    assert wb.active["A1"].value == original_val
    assert sample_manager.get_history_count() == 0

    # Undo when count is 0: fails gracefully
    res3 = undo_last_change(sample_manager)
    assert res3["success"] is False
    assert sample_manager.get_history_count() == 0

