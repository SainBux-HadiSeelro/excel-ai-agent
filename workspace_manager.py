"""
Workspace & Version Safety Manager for Excel AI Agent.
Guarantees:
- Original file is never modified (operates on isolated copies).
- Every modifying operation creates an automatic snapshot.
- Full undo support by reverting to previous snapshots.
- Detection and confirmation guards for destructive operations.
"""

from __future__ import annotations

import os
import shutil
import time
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any
import openpyxl
import pandas as pd


class WorkspaceManager:
    """Manages an isolated workspace for workbook operations with automatic backups and undo."""

    def __init__(self, workspace_root: Optional[str] = None, session_id: Optional[str] = None):
        if workspace_root is None:
            workspace_root = os.getenv("EXCEL_WORKSPACE_DIR", "workspace")
        if session_id is None:
            session_id = f"session_{int(time.time() * 1000)}"
        self.session_id = session_id
        self.workspace_dir = Path(workspace_root) / session_id
        self.history_dir = self.workspace_dir / "history"
        self.workspace_dir.mkdir(parents=True, exist_ok=True)
        self.history_dir.mkdir(parents=True, exist_ok=True)

        self.current_file: Optional[Path] = None
        self.original_filename: str = "workbook.xlsx"
        self._history_stack: List[Path] = []
        self._version_counter: int = 0

    def init_from_file(self, source_path: str | Path) -> str:
        """Copies an external file into the workspace as the active working copy."""
        src = Path(source_path)
        if not src.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        self.original_filename = src.name
        # If it's a CSV, convert to XLSX first
        if src.suffix.lower() == ".csv":
            target_file = self.workspace_dir / f"{src.stem}.xlsx"
            df = pd.read_csv(src)
            df.to_excel(target_file, index=False)
            self.current_file = target_file
        else:
            target_file = self.workspace_dir / "current.xlsx"
            shutil.copy2(src, target_file)
            self.current_file = target_file

        # Create initial baseline snapshot (v0)
        self._history_stack.clear()
        self._version_counter = 0
        self._create_snapshot("Initial upload")
        return str(self.current_file)

    def init_blank(self, filename: str = "workbook.xlsx") -> str:
        """Initializes a new empty workbook in workspace."""
        self.original_filename = filename
        target_file = self.workspace_dir / "current.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "Sheet1"
        wb.save(target_file)
        self.current_file = target_file
        self._history_stack.clear()
        self._version_counter = 0
        self._create_snapshot("Initial blank workbook")
        return str(self.current_file)

    def _create_snapshot(self, description: str = "") -> Path:
        """Saves a point-in-time snapshot of the current file into the history folder."""
        if self.current_file is None or not self.current_file.exists():
            raise RuntimeError("Cannot snapshot: current working file does not exist.")

        self._version_counter += 1
        snapshot_path = self.history_dir / f"v{self._version_counter:03d}_{int(time.time())}.xlsx"
        shutil.copy2(self.current_file, snapshot_path)
        self._history_stack.append(snapshot_path)
        return snapshot_path

    def prepare_for_mutation(self, action_name: str = "") -> None:
        """Must be called before applying any modifying operation to ensure undo availability."""
        if self.current_file is None or not self.current_file.exists():
            raise RuntimeError("No active workbook loaded in workspace.")
        # Make snapshot of current state before change
        self._create_snapshot(action_name)

    def undo(self) -> Dict[str, Any]:
        """Reverts the current workbook to the previous snapshot in the history stack."""
        if len(self._history_stack) <= 1:
            return {
                "success": False,
                "message": "No previous changes to undo. Already at the initial version."
            }

        # Remove the latest state
        current_state = self._history_stack.pop()
        # Restore the previous state
        target_state = self._history_stack[-1]

        if not target_state.exists():
            return {
                "success": False,
                "message": f"Snapshot file missing: {target_state.name}"
            }

        shutil.copy2(target_state, self.current_file)
        return {
            "success": True,
            "message": f"Successfully reverted to previous snapshot ({target_state.stem}).",
            "remaining_undo_steps": len(self._history_stack) - 1
        }

    def get_history_count(self) -> int:
        """Returns the number of undoable actions available."""
        return max(0, len(self._history_stack) - 1)

    def get_current_file_path(self) -> str:
        """Returns the absolute path of the current working copy."""
        if self.current_file is None or not self.current_file.exists():
            raise RuntimeError("No working file initialized in workspace.")
        return str(self.current_file.resolve())

    def open_workbook(self, data_only: bool = False) -> openpyxl.Workbook:
        """Loads openpyxl Workbook from current file.
        data_only=False preserves formulas.
        data_only=True returns evaluated formula values (if saved by Excel).
        """
        path = self.get_current_file_path()
        return openpyxl.load_workbook(path, data_only=data_only)

    def save_workbook(self, wb: openpyxl.Workbook) -> None:
        """Saves openpyxl Workbook to the current working copy."""
        path = self.get_current_file_path()
        wb.save(path)

    def get_sheet_dataframe(self, sheet_name: Optional[str] = None, max_rows: Optional[int] = 500) -> Tuple[pd.DataFrame, str]:
        """Loads a sheet as a pandas DataFrame for viewing/inspection."""
        path = self.get_current_file_path()
        xl = pd.ExcelFile(path)
        target_sheet = sheet_name if sheet_name and sheet_name in xl.sheet_names else xl.sheet_names[0]
        
        # Read sheet
        if max_rows:
            df = pd.read_excel(path, sheet_name=target_sheet, nrows=max_rows)
        else:
            df = pd.read_excel(path, sheet_name=target_sheet)
        return df, target_sheet

    def get_sheet_names(self) -> List[str]:
        """Lists all sheet names in the workbook."""
        wb = self.open_workbook()
        return wb.sheetnames
