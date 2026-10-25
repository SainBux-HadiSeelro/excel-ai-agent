"""Tests for stale session cleanup and upload path security."""

import os
import time
from pathlib import Path
from workspace_manager import cleanup_stale_sessions


def test_cleanup_stale_sessions_older_than_24h(tmp_path):
    ws_root = tmp_path / "workspace"
    up_root = tmp_path / "temp_uploads"
    ws_root.mkdir()
    up_root.mkdir()

    # 1. Stale session (25 hours old)
    stale_session = ws_root / "session_stale"
    stale_session.mkdir()
    stale_file = stale_session / "workbook.xlsx"
    stale_file.write_text("data")
    old_time = time.time() - (25 * 3600)
    os.utime(stale_session, (old_time, old_time))
    os.utime(stale_file, (old_time, old_time))

    # 2. Fresh session (2 hours old)
    fresh_session = ws_root / "session_fresh"
    fresh_session.mkdir()
    fresh_file = fresh_session / "workbook.xlsx"
    fresh_file.write_text("data")
    fresh_time = time.time() - (2 * 3600)
    os.utime(fresh_session, (fresh_time, fresh_time))
    os.utime(fresh_file, (fresh_time, fresh_time))

    # 3. Active session (even if 25 hours old, must NOT be deleted)
    active_session = ws_root / "session_active"
    active_session.mkdir()
    active_file = active_session / "workbook.xlsx"
    active_file.write_text("data")
    os.utime(active_session, (old_time, old_time))
    os.utime(active_file, (old_time, old_time))

    # 4. Stale upload dir (25 hours old)
    stale_up = up_root / "upload_stale"
    stale_up.mkdir()
    stale_up_file = stale_up / "upload.xlsx"
    stale_up_file.write_text("data")
    os.utime(stale_up, (old_time, old_time))
    os.utime(stale_up_file, (old_time, old_time))

    # Run cleanup with active_session preserved
    res = cleanup_stale_sessions(
        workspace_root=ws_root,
        temp_uploads_root=up_root,
        max_age_hours=24.0,
        active_session_ids=["session_active"]
    )

    assert res["removed_workspaces"] == 1
    assert res["removed_uploads"] == 1

    # Stale directory was deleted
    assert not stale_session.exists()
    assert not stale_up.exists()

    # Fresh and active directories are preserved
    assert fresh_session.exists()
    assert active_session.exists()


def test_cleanup_configurable_via_env(tmp_path, monkeypatch):
    monkeypatch.setenv("EXCEL_SESSION_MAX_AGE_HOURS", "2")
    ws_root = tmp_path / "workspace"
    ws_root.mkdir()

    # 3 hours old -> should be cleaned up
    sess = ws_root / "session_3h"
    sess.mkdir()
    os.utime(sess, (time.time() - (3 * 3600), time.time() - (3 * 3600)))

    # 1 hour old -> should be kept
    sess_recent = ws_root / "session_1h"
    sess_recent.mkdir()
    os.utime(sess_recent, (time.time() - (1 * 3600), time.time() - (1 * 3600)))

    res = cleanup_stale_sessions(workspace_root=ws_root, temp_uploads_root=tmp_path / "uploads")
    assert res["removed_workspaces"] == 1
    assert not sess.exists()
    assert sess_recent.exists()


def test_cleanup_preserves_old_folder_with_recently_modified_history_file(tmp_path):
    ws_root = tmp_path / "workspace"
    ws_root.mkdir()

    session_dir = ws_root / "session_old_folder_recent_history"
    session_dir.mkdir()
    history_dir = session_dir / "history"
    history_dir.mkdir()

    old_time = time.time() - (30 * 3600)  # 30 hours ago
    recent_time = time.time() - (1 * 3600)  # 1 hour ago

    # Old session folder, history folder, and current.xlsx
    current_file = session_dir / "current.xlsx"
    current_file.write_text("dummy current")
    os.utime(session_dir, (old_time, old_time))
    os.utime(current_file, (old_time, old_time))
    os.utime(history_dir, (old_time, old_time))

    # Recently modified file inside history/
    history_file = history_dir / "v001_snapshot.xlsx"
    history_file.write_text("snapshot content")
    os.utime(history_file, (recent_time, recent_time))

    # Run cleanup with 24-hour limit
    res = cleanup_stale_sessions(workspace_root=ws_root, temp_uploads_root=tmp_path / "temp_uploads", max_age_hours=24.0)

    # Must NOT be deleted because history file was recently modified
    assert res["removed_workspaces"] == 0
    assert session_dir.exists()
    assert history_file.exists()


def test_upload_path_sanitization():
    from workspace_manager import safe_upload_name

    # 1. Forward slash path traversal
    assert safe_upload_name("../../../etc/passwd.xlsx") == "passwd.xlsx"

    # 2. Backslash path traversal & Windows absolute paths on any OS
    assert safe_upload_name("C:\\Windows\\System32\\file.xlsx") == "file.xlsx"
    assert safe_upload_name("..\\..\\secret.xlsx") == "secret.xlsx"
    assert safe_upload_name("subfolder\\nested/data.xlsx") == "data.xlsx"

    # 3. Leading dots stripped (prevents hidden files and dot-relative names)
    assert safe_upload_name("...hidden_sheet.xlsx") == "hidden_sheet.xlsx"
    assert safe_upload_name(".gitignore.xlsx") == "gitignore.xlsx"

    # 4. Fallback to upload.xlsx if empty, whitespace, or only dots
    assert safe_upload_name("") == "upload.xlsx"
    assert safe_upload_name("   ") == "upload.xlsx"
    assert safe_upload_name("..") == "upload.xlsx"
    assert safe_upload_name("...") == "upload.xlsx"
    assert safe_upload_name(None) == "upload.xlsx"

    # 5. Normal safe filenames preserved
    assert safe_upload_name("sales_q1.xlsx") == "sales_q1.xlsx"
    assert safe_upload_name("monthly_report.csv") == "monthly_report.csv"
