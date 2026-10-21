"""Tests for sheet structure tools: create, delete, rename, insert/delete rows/cols."""

from tools.structure import (
    create_sheet,
    delete_sheet,
    rename_sheet,
    insert_rows,
    delete_rows,
    insert_columns,
    delete_columns
)
from tools.inspection import list_sheets


def test_create_and_rename_sheet(sample_manager):
    res_create = create_sheet(sample_manager, title="NewReport")
    assert res_create["success"] is True
    assert "NewReport" in sample_manager.get_sheet_names()

    res_rename = rename_sheet(sample_manager, old_title="NewReport", new_title="FinalReport")
    assert res_rename["success"] is True
    assert "FinalReport" in sample_manager.get_sheet_names()
    assert "NewReport" not in sample_manager.get_sheet_names()


def test_delete_sheet_guard_and_confirmed(sample_manager):
    create_sheet(sample_manager, title="TempSheet")
    assert "TempSheet" in sample_manager.get_sheet_names()

    # Call delete without confirmation -> must require confirmation
    res_guard = delete_sheet(sample_manager, title="TempSheet", confirmed=False)
    assert res_guard.get("requires_confirmation") is True
    assert "TempSheet" in sample_manager.get_sheet_names()

    # Call with confirmed=True -> should delete
    res_del = delete_sheet(sample_manager, title="TempSheet", confirmed=True)
    assert res_del["success"] is True
    assert "TempSheet" not in sample_manager.get_sheet_names()


def test_insert_and_delete_rows(sample_manager):
    # Insert 2 rows at row 2
    res_insert = insert_rows(sample_manager, sheet_name="Sales", row_idx=2, count=2)
    assert res_insert["success"] is True

    # Attempt delete rows without confirmation -> guard
    res_guard = delete_rows(sample_manager, sheet_name="Sales", row_idx=2, count=2, confirmed=False)
    assert res_guard.get("requires_confirmation") is True

    # Delete with confirmed=True
    res_del = delete_rows(sample_manager, sheet_name="Sales", row_idx=2, count=2, confirmed=True)
    assert res_del["success"] is True


def test_insert_and_delete_columns(sample_manager):
    res_ins = insert_columns(sample_manager, sheet_name="Sales", col_idx="B", count=1)
    assert res_ins["success"] is True

    # Guard check
    res_guard = delete_columns(sample_manager, sheet_name="Sales", col_idx="B", count=1, confirmed=False)
    assert res_guard.get("requires_confirmation") is True

    # Confirmed delete
    res_del = delete_columns(sample_manager, sheet_name="Sales", col_idx="B", count=1, confirmed=True)
    assert res_del["success"] is True
