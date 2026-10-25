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
    (stale_session / "workbook.xlsx").write_text("data")
    old_time = time.time() - (25 * 3600)
    os.utime(stale_session, (old_time, old_time))

    # 2. Fresh session (2 hours old)
    fresh_session = ws_root / "session_fresh"
    fresh_session.mkdir()
    (fresh_session / "workbook.xlsx").write_text("data")
    fresh_time = time.time() - (2 * 3600)
    os.utime(fresh_session, (fresh_time, fresh_time))

    # 3. Active session (even if 25 hours old, must NOT be deleted)
    active_session = ws_root / "session_active"
    active_session.mkdir()
    (active_session / "workbook.xlsx").write_text("data")
    os.utime(active_session, (old_time, old_time))

    # 4. Stale upload dir (25 hours old)
    stale_up = up_root / "upload_stale"
    stale_up.mkdir()
    (stale_up / "upload.xlsx").write_text("data")
    os.utime(stale_up, (old_time, old_time))

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


def test_upload_path_sanitization():
    malicious_name = "../../../etc/passwd.xlsx"
    sanitized = Path(malicious_name).name
    assert sanitized == "passwd.xlsx"
    assert ".." not in sanitized
    assert "/" not in sanitized
    assert "\\" not in sanitized
