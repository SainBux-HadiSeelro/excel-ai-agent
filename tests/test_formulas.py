"""Tests for advanced formula evaluation (SUMIF, AVERAGE, IF, and nested formulas)."""

import openpyxl
from workspace_manager import evaluate_formula_cell


def test_evaluate_formulas_advanced():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CalcSheet"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["A3"] = 30

    ws["B1"] = "=AVERAGE(A1:A3)"
    ws["B2"] = '=IF(A1>5, "Pass", "Fail")'
    ws["B3"] = '=SUMIF(A1:A3, ">15")'
    ws["B4"] = '=IF(SUM(A1:A3)>50, AVERAGE(A1:A3)*2, 0)'

    # AVERAGE test
    val_avg = evaluate_formula_cell(ws["B1"].value, None, ws, cell_coord="B1")
    assert round(float(val_avg), 2) == 20.0

    # IF test
    val_if = evaluate_formula_cell(ws["B2"].value, None, ws, cell_coord="B2")
    assert str(val_if).strip() == "Pass"

    # SUMIF test
    val_sumif = evaluate_formula_cell(ws["B3"].value, None, ws, cell_coord="B3")
    assert round(float(val_sumif), 2) == 50.0

    # Nested formula test: IF + SUM + AVERAGE
    val_nested = evaluate_formula_cell(ws["B4"].value, None, ws, cell_coord="B4")
    assert round(float(val_nested), 2) == 40.0


def test_evaluate_formula_read_range_integration(sample_manager):
    from tools.editing import write_cells
    from tools.inspection import read_range

    # Write test formulas to Sales sheet
    write_cells(sample_manager, updates=[
        {"cell": "K2", "value": "=AVERAGE(F2:F5)"},
        {"cell": "K3", "value": '=IF(F2>10, "HighQty", "LowQty")'},
        {"cell": "K4", "value": '=SUMIF(F2:F5, ">5")'},
        {"cell": "K5", "value": '=IF(SUM(F2:F5)>0, AVERAGE(F2:F5), -1)'}
    ], sheet_name="Sales")

    # Read without show_formulas (should evaluate)
    res = read_range(sample_manager, sheet_name="Sales", cell_range="K2:K5", show_formulas=False)
    assert res["row_count"] == 4
    # Values should be evaluated, not raw formulas
    assert not str(res["rows"][0][0]).startswith("=")
    assert str(res["rows"][1][0]) in ("HighQty", "LowQty")
    assert float(res["rows"][2][0]) > 0
    assert float(res["rows"][3][0]) > 0


def test_formula_evaluation_failure_never_crashes():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "=COMPLETELY_INVALID_UNKNOWN_FUNC(999, 'bad')"
    # Should safely return raw formula without raising
    val = evaluate_formula_cell(ws["A1"].value, None, ws, cell_coord="A1")
    assert val == "=COMPLETELY_INVALID_UNKNOWN_FUNC(999, 'bad')"


def test_formula_cache_hit_does_not_call_evaluator_twice(sample_manager, monkeypatch):
    import workspace_manager
    workspace_manager.invalidate_formula_cache()

    call_count = 0
    orig_run = workspace_manager._run_pure_python_formulas

    def mock_evaluator(path):
        nonlocal call_count
        call_count += 1
        return orig_run(path)

    monkeypatch.setattr(workspace_manager, "_run_pure_python_formulas", mock_evaluator)

    # First call: should evaluate
    df1, s1 = sample_manager.get_sheet_dataframe("Sales")
    assert call_count == 1

    # Second call on unchanged file: should hit cache and NOT call evaluator again
    df2, s2 = sample_manager.get_sheet_dataframe("Sales")
    assert call_count == 1
    assert len(df1) == len(df2)


def test_formula_cache_invalidates_after_save_workbook(sample_manager, monkeypatch):
    import workspace_manager
    workspace_manager.invalidate_formula_cache()

    call_count = 0
    orig_run = workspace_manager._run_pure_python_formulas

    def mock_evaluator(path):
        nonlocal call_count
        call_count += 1
        return orig_run(path)

    monkeypatch.setattr(workspace_manager, "_run_pure_python_formulas", mock_evaluator)

    # First call: evaluator called once
    sample_manager.get_sheet_dataframe("Sales")
    assert call_count == 1

    # Mutate and save workbook
    wb = sample_manager.open_workbook()
    wb["Sales"]["A2"] = "ModifiedCustomer"
    sample_manager.save_workbook(wb)

    # Next call: cache must be invalidated, so evaluator is called again
    sample_manager.get_sheet_dataframe("Sales")
    assert call_count == 2


def test_oversize_workbook_falls_back_without_crashing(sample_manager, monkeypatch):
    import workspace_manager
    workspace_manager.invalidate_formula_cache()

    # Set threshold low so sample workbook is considered oversize
    monkeypatch.setattr(workspace_manager, "FORMULA_CELL_LIMIT", 5)

    call_count = 0
    def mock_evaluator(path):
        nonlocal call_count
        call_count += 1
        return {}

    monkeypatch.setattr(workspace_manager, "_run_pure_python_formulas", mock_evaluator)

    # Load dataframe for oversize file
    df, s, is_oversize = workspace_manager.load_sheet_dataframe(sample_manager.get_current_file_path())
    assert is_oversize is True
    # Evaluator was skipped!
    assert call_count == 0
    # DataFrame loaded successfully without crash
    assert not df.empty


def test_error_values_never_crash_preview_or_json_serialization(sample_manager):
    import json
    from tools.editing import write_cells
    from tools.inspection import read_range
    from workspace_manager import load_sheet_dataframe

    # Write error-producing formulas
    write_cells(sample_manager, updates=[
        {"cell": "K2", "value": "=1/0"},
        {"cell": "K3", "value": "=NA()"}
    ], sheet_name="Sales")

    # Read via read_range
    res = read_range(sample_manager, sheet_name="Sales", cell_range="K2:K3", show_formulas=False)
    assert res["row_count"] == 2
    # Verify error values like #DIV/0! and #N/A are present
    vals = [r[0] for r in res["rows"]]
    assert any("#DIV/0!" in str(v) for v in vals)
    assert any("#N/A" in str(v) for v in vals)

    # JSON serialization of tool result must not crash
    serialized_res = json.dumps(res)
    assert "#DIV/0!" in serialized_res
    assert "#N/A" in serialized_res

    # Preview dataframe load and JSON conversion must not crash
    df, _, _ = load_sheet_dataframe(sample_manager.get_current_file_path(), sheet_name="Sales")
    df = df.convert_dtypes()
    df_json = df.to_json()
    assert ("#DIV/0!" in df_json or "#DIV\\/0!" in df_json)
    dict_json = json.dumps(df.to_dict(), default=str)
    assert "#DIV/0!" in dict_json
    assert "#N/A" in dict_json


def test_libreoffice_temp_dir_cleanup_on_error(tmp_path, monkeypatch):
    import subprocess
    import tempfile
    import workspace_manager

    created_dirs = []
    orig_mkdtemp = tempfile.mkdtemp

    def spy_mkdtemp(*args, **kwargs):
        d = orig_mkdtemp(*args, **kwargs)
        created_dirs.append(workspace_manager.Path(d))
        return d

    monkeypatch.setattr(workspace_manager.tempfile, "mkdtemp", spy_mkdtemp)
    monkeypatch.setattr(workspace_manager.shutil, "which", lambda cmd: "/usr/bin/soffice")

    # 1. Non-zero return code
    def mock_subprocess_nonzero(*args, **kwargs):
        return subprocess.CompletedProcess(args=args, returncode=1, stdout=b"", stderr=b"error")

    monkeypatch.setattr(workspace_manager.subprocess, "run", mock_subprocess_nonzero)

    test_file = tmp_path / "test.xlsx"
    test_file.write_text("dummy")

    res = workspace_manager._evaluate_with_libreoffice(test_file)
    assert res is None
    assert len(created_dirs) == 1
    assert not created_dirs[0].exists(), "Temp dir must be deleted on non-zero return code"

    # 2. Exception in subprocess
    def mock_subprocess_exc(*args, **kwargs):
        raise subprocess.SubprocessError("Subprocess failed")

    monkeypatch.setattr(workspace_manager.subprocess, "run", mock_subprocess_exc)
    res2 = workspace_manager._evaluate_with_libreoffice(test_file)
    assert res2 is None
    assert len(created_dirs) == 2
    assert not created_dirs[1].exists(), "Temp dir must be deleted on exception"


def test_formula_cache_lru_eviction(monkeypatch):
    import workspace_manager
    workspace_manager.invalidate_formula_cache()
    monkeypatch.setattr(workspace_manager, "MAX_CACHE_ENTRIES", 3)

    k1 = ("file1.xlsx", 100, 10)
    k2 = ("file2.xlsx", 200, 20)
    k3 = ("file3.xlsx", 300, 30)
    k4 = ("file4.xlsx", 400, 40)

    workspace_manager._cache_put(k1, {"A": 1}, False)
    workspace_manager._cache_put(k2, {"B": 2}, False)
    workspace_manager._cache_put(k3, {"C": 3}, False)
    assert list(workspace_manager._MODULE_FORMULA_CACHE.keys()) == [k1, k2, k3]

    # Inserting 4th item evicts oldest (k1)
    workspace_manager._cache_put(k4, {"D": 4}, False)
    assert k1 not in workspace_manager._MODULE_FORMULA_CACHE
    assert list(workspace_manager._MODULE_FORMULA_CACHE.keys()) == [k2, k3, k4]

    # Accessing k2 moves it to the end
    workspace_manager._MODULE_FORMULA_CACHE.move_to_end(k2)
    assert list(workspace_manager._MODULE_FORMULA_CACHE.keys()) == [k3, k4, k2]

    # Inserting k5 evicts k3 (new oldest)
    k5 = ("file5.xlsx", 500, 50)
    workspace_manager._cache_put(k5, {"E": 5}, False)
    assert k3 not in workspace_manager._MODULE_FORMULA_CACHE
    assert list(workspace_manager._MODULE_FORMULA_CACHE.keys()) == [k4, k2, k5]


def test_cleanup_invalidates_formula_cache_for_deleted_sessions(tmp_path):
    import time
    import os
    import workspace_manager

    ws_root = tmp_path / "workspace"
    ws_root.mkdir()
    sess_dir = ws_root / "sess_stale_cache"
    sess_dir.mkdir()
    curr_file = sess_dir / "current.xlsx"
    curr_file.write_text("data")

    # Put entry in cache
    file_resolved = str(curr_file.resolve())
    cache_key = (file_resolved, 9999, 123)
    workspace_manager._cache_put(cache_key, {"VAL": 42}, False)
    assert cache_key in workspace_manager._MODULE_FORMULA_CACHE

    # Make stale
    old_time = time.time() - (40 * 3600)
    os.utime(sess_dir, (old_time, old_time))
    os.utime(curr_file, (old_time, old_time))

    # Run cleanup
    workspace_manager.cleanup_stale_sessions(workspace_root=ws_root, temp_uploads_root=tmp_path / "uploads", max_age_hours=24.0)

    # Session deleted and cache invalidated
    assert not sess_dir.exists()
    assert cache_key not in workspace_manager._MODULE_FORMULA_CACHE


def test_temp_path_evaluations_leave_cache_unchanged():
    import workspace_manager

    # Clear and snapshot cache keys
    workspace_manager.invalidate_formula_cache()
    assert len(workspace_manager._MODULE_FORMULA_CACHE) == 0

    # In-memory workbook without _file_path
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 10
    ws["A2"] = 20
    ws["B1"] = "=A1*A2"

    val = workspace_manager.evaluate_formula_cell(ws["B1"].value, None, ws, cell_coord="B1")
    assert float(val) == 200.0
    # Cache must still be empty!
    assert len(workspace_manager._MODULE_FORMULA_CACHE) == 0

