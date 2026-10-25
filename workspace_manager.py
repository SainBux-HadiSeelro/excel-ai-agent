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
import re
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd

try:
    import formulas
    HAS_FORMULAS = True
except ImportError:
    HAS_FORMULAS = False

_MODULE_FORMULA_CACHE: Dict[Tuple[str, int, int], Dict[str, Any]] = {}
_MODULE_OVERSIZE_FLAGS: Dict[Tuple[str, int, int], bool] = {}
FORMULA_CELL_LIMIT = 20000


def get_file_cache_key(file_path: str | Path) -> Optional[Tuple[str, int, int]]:
    """Generates cache key: (resolved_path, mtime_ns, size)."""
    try:
        p = Path(file_path).resolve()
        if not p.exists():
            return None
        st = p.stat()
        return (str(p), st.st_mtime_ns, st.st_size)
    except Exception:
        return None


def invalidate_formula_cache(file_path: Optional[str | Path] = None) -> None:
    """Invalidates formula cache entries for a specific path, or all entries if None."""
    global _MODULE_FORMULA_CACHE, _MODULE_OVERSIZE_FLAGS
    if file_path is None:
        _MODULE_FORMULA_CACHE.clear()
        _MODULE_OVERSIZE_FLAGS.clear()
    else:
        resolved = str(Path(file_path).resolve())
        keys_to_remove = [k for k in _MODULE_FORMULA_CACHE if k[0] == resolved]
        for k in keys_to_remove:
            _MODULE_FORMULA_CACHE.pop(k, None)
            _MODULE_OVERSIZE_FLAGS.pop(k, None)


def count_non_empty_cells(wb: openpyxl.Workbook, limit: Optional[int] = None) -> int:
    """Counts non-empty cells across all worksheets up to limit + 1."""
    if limit is None:
        limit = FORMULA_CELL_LIMIT
    count = 0
    for s in wb.worksheets:
        for row in s.iter_rows(values_only=True):
            for v in row:
                if v is not None and v != "":
                    count += 1
                    if count > limit:
                        return count
    return count


def is_file_oversize(file_path: str | Path, limit: Optional[int] = None) -> bool:
    """Checks if the workbook at file_path has > limit non-empty cells."""
    if limit is None:
        limit = FORMULA_CELL_LIMIT
    key = get_file_cache_key(file_path)
    if key and key in _MODULE_OVERSIZE_FLAGS:
        return _MODULE_OVERSIZE_FLAGS[key]
    try:
        p = Path(file_path).resolve()
        if not p.exists():
            return False
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        count = count_non_empty_cells(wb, limit=limit)
        wb.close()
        is_over = count > limit
        if key:
            _MODULE_OVERSIZE_FLAGS[key] = is_over
        return is_over
    except Exception:
        return False


def _normalize_eval_value(v: Any) -> Any:
    """Normalizes evaluated formula values to standard JSON/DataFrame safe types."""
    if v is None:
        return ""
    if hasattr(v, "item"):
        try:
            v = v.item()
        except Exception:
            pass
    if hasattr(v, "tolist"):
        v = v.tolist()
    while isinstance(v, list) and len(v) == 1:
        v = v[0]

    v_str = str(v)
    if v_str.startswith("#") and any(err in v_str for err in ["#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!"]):
        return v_str

    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        if v != v or abs(v) == float("inf"):
            return "#NUM!"
        return v
    return v_str


def _evaluate_with_libreoffice(file_path: Path) -> Optional[Path]:
    """Headless LibreOffice recalculation fallback if LibreOffice is available on PATH."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp(prefix="libreoffice_recalc_")
        cmd = [soffice, "--headless", "--convert-to", "xlsx", "--outdir", temp_dir, str(file_path)]
        res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=20)
        if res.returncode == 0:
            out_file = Path(temp_dir) / file_path.name
            if out_file.exists():
                return out_file
        # Clean up temp_dir on non-zero return code or missing out_file
        if temp_dir and Path(temp_dir).exists():
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        if temp_dir and Path(temp_dir).exists():
            shutil.rmtree(temp_dir, ignore_errors=True)
    return None


def _run_pure_python_formulas(file_path: Path) -> Dict[str, Any]:
    """Evaluates formulas in an Excel file using the formulas library."""
    cache: Dict[str, Any] = {}
    if not HAS_FORMULAS:
        return cache
    try:
        model = formulas.ExcelModel().loads(str(file_path)).finish()
        calc_res = model.calculate()
        for k, cell_obj in calc_res.items():
            m = re.search(r"\]([^']+)'!([A-Za-z0-9:]+)", str(k))
            if m:
                sheet_name, coord = m.groups()
                v = getattr(cell_obj, "value", cell_obj)
                cache[f"{sheet_name.upper()}!{coord.upper()}"] = _normalize_eval_value(v)
    except Exception:
        pass
    return cache


def _evaluate_workbook_formulas(file_path: Path, is_oversize: bool = False) -> Dict[str, Any]:
    """Evaluates formulas across a workbook using LibreOffice or pure-Python formulas library."""
    cache: Dict[str, Any] = {}
    # 1. LibreOffice headless recalculation if available
    recalc_path = _evaluate_with_libreoffice(file_path)
    if recalc_path and recalc_path.exists():
        try:
            recalc_wb = openpyxl.load_workbook(recalc_path, data_only=True)
            for sheet_name in recalc_wb.sheetnames:
                s_calc = recalc_wb[sheet_name]
                for r in range(1, s_calc.max_row + 1):
                    for c in range(1, s_calc.max_column + 1):
                        v = s_calc.cell(row=r, column=c).value
                        if v is not None:
                            coord = f"{get_column_letter(c)}{r}"
                            cache[f"{sheet_name.upper()}!{coord}"] = _normalize_eval_value(v)
        except Exception:
            pass
        finally:
            try:
                shutil.rmtree(recalc_path.parent, ignore_errors=True)
            except Exception:
                pass

    # 2. Pure-Python formulas library evaluator (skipped if oversize)
    if not cache and HAS_FORMULAS and not is_oversize:
        cache = _run_pure_python_formulas(file_path)

    return cache


def _build_formula_cache(ws_raw: Any, ws_data: Any = None) -> Dict[str, Any]:
    """Pre-calculates formulas in a workbook using module cache, LibreOffice, or formulas library."""
    cache: Dict[str, Any] = {}
    wb = getattr(ws_raw, "parent", None)
    if not wb:
        return cache

    file_path = getattr(wb, "_file_path", None)
    temp_wb_path = None
    if not file_path or not Path(file_path).exists():
        try:
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                temp_wb_path = Path(tmp.name)
            wb.save(temp_wb_path)
            file_path = temp_wb_path
        except Exception:
            file_path = None

    if file_path and Path(file_path).exists():
        cache_key = get_file_cache_key(file_path)
        if cache_key and cache_key in _MODULE_FORMULA_CACHE:
            cache = dict(_MODULE_FORMULA_CACHE[cache_key])
        else:
            is_oversize = is_file_oversize(file_path)
            cache = _evaluate_workbook_formulas(Path(file_path), is_oversize=is_oversize)
            if cache_key:
                _MODULE_FORMULA_CACHE[cache_key] = dict(cache)
                _MODULE_OVERSIZE_FLAGS[cache_key] = is_oversize

    if temp_wb_path and temp_wb_path.exists():
        try:
            os.remove(temp_wb_path)
        except Exception:
            pass

    # Map raw formula strings to evaluated results
    try:
        sheet_name = getattr(ws_raw, "title", "").upper()
        for row in ws_raw.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formula_key = f"{sheet_name}!{cell.value.upper().strip()}"
                    coord_key = f"{sheet_name}!{cell.coordinate.upper()}"
                    if coord_key in cache:
                        cache[formula_key] = cache[coord_key]
    except Exception:
        pass

    return cache


def evaluate_formula_cell(
    raw_val: Any,
    data_val: Any,
    ws_raw: Any,
    ws_data: Any = None,
    cell_coord: Optional[str] = None
) -> Any:
    """Evaluates or extracts formula value if data_only cached value is None."""
    if data_val is not None and data_val != "":
        return data_val
    if raw_val is None:
        return ""
    if not isinstance(raw_val, str) or not raw_val.startswith("="):
        return raw_val

    # 1. Pure-Python formulas or LibreOffice evaluation from cache
    if ws_raw is not None:
        if not hasattr(ws_raw, "_formula_eval_cache"):
            setattr(ws_raw, "_formula_eval_cache", _build_formula_cache(ws_raw, ws_data))
        cache = getattr(ws_raw, "_formula_eval_cache", {})
        sheet_title = getattr(ws_raw, "title", "").upper()

        if cell_coord:
            key_coord = f"{sheet_title}!{cell_coord.upper()}"
            if key_coord in cache:
                return cache[key_coord]

        key_formula = f"{sheet_title}!{raw_val.upper().strip()}"
        if key_formula in cache:
            return cache[key_formula]

    # 2. Legacy fallback regex evaluator for simple operations
    formula = raw_val[1:].strip()

    # Try simple product: =F2*G2
    m_mul = re.match(r"^([A-Za-z]+)(\d+)\s*\*\s*([A-Za-z]+)(\d+)$", formula)
    if m_mul:
        try:
            c1, r1, c2, r2 = m_mul.groups()
            v1 = ws_raw[f"{c1}{r1}"].value or 0
            v2 = ws_raw[f"{c2}{r2}"].value or 0
            return float(v1) * float(v2)
        except Exception:
            pass

    # Try simple addition: =F2+G2
    m_add = re.match(r"^([A-Za-z]+)(\d+)\s*\+\s*([A-Za-z]+)(\d+)$", formula)
    if m_add:
        try:
            c1, r1, c2, r2 = m_add.groups()
            v1 = ws_raw[f"{c1}{r1}"].value or 0
            v2 = ws_raw[f"{c2}{r2}"].value or 0
            return float(v1) + float(v2)
        except Exception:
            pass

    # Try simple subtraction: =F2-G2
    m_sub = re.match(r"^([A-Za-z]+)(\d+)\s*\-\s*([A-Za-z]+)(\d+)$", formula)
    if m_sub:
        try:
            c1, r1, c2, r2 = m_sub.groups()
            v1 = ws_raw[f"{c1}{r1}"].value or 0
            v2 = ws_raw[f"{c2}{r2}"].value or 0
            return float(v1) - float(v2)
        except Exception:
            pass

    # Try simple division: =F2/G2
    m_div = re.match(r"^([A-Za-z]+)(\d+)\s*\/\s*([A-Za-z]+)(\d+)$", formula)
    if m_div:
        try:
            c1, r1, c2, r2 = m_div.groups()
            v1 = ws_raw[f"{c1}{r1}"].value or 0
            v2 = ws_raw[f"{c2}{r2}"].value
            if v2 is None or v2 == 0 or float(v2) == 0:
                return "#DIV/0!"
            return float(v1) / float(v2)
        except ZeroDivisionError:
            return "#DIV/0!"
        except Exception:
            pass

    # Try NA(): =NA()
    if formula.upper() == "NA()":
        return "#N/A"

    # Try SUM: =SUM(H2:H23)
    m_sum = re.match(r"^SUM\(([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)\)$", formula, re.IGNORECASE)
    if m_sum:
        try:
            from openpyxl.utils.cell import column_index_from_string
            c1, r1, c2, r2 = m_sum.groups()
            start_c = column_index_from_string(c1)
            end_c = column_index_from_string(c2)
            total = 0.0
            for r in range(int(r1), int(r2) + 1):
                for c in range(start_c, end_c + 1):
                    raw_sub = ws_raw.cell(row=r, column=c).value
                    data_sub = ws_data.cell(row=r, column=c).value if ws_data else None
                    sub_val = evaluate_formula_cell(raw_sub, data_sub, ws_raw, ws_data)
                    try:
                        total += float(sub_val or 0)
                    except Exception:
                        pass
            return total
        except Exception:
            pass

    # 3. Last fallback: never crash, return raw formula
    return raw_val


def load_sheet_dataframe(
    file_path: str | Path,
    sheet_name: Optional[str] = None,
    max_rows: Optional[int] = 500
) -> Tuple[pd.DataFrame, str, bool]:
    """Loads a sheet as a pandas DataFrame, evaluating formulas, with oversize workbook detection."""
    p = Path(file_path).resolve()
    wb_raw = openpyxl.load_workbook(p, data_only=False)
    wb_raw._file_path = p
    wb_data = openpyxl.load_workbook(p, data_only=True)
    wb_data._file_path = p
    target_sheet = sheet_name if sheet_name and sheet_name in wb_raw.sheetnames else wb_raw.sheetnames[0]

    ws_raw = wb_raw[target_sheet]
    ws_data = wb_data[target_sheet]

    is_oversize = is_file_oversize(p)

    if ws_raw.max_row == 0 or ws_raw.max_column == 0:
        return pd.DataFrame(), target_sheet, is_oversize

    headers = [str(ws_raw.cell(row=1, column=c).value or f"Col_{c}") for c in range(1, ws_raw.max_column + 1)]

    limit_rows = min(ws_raw.max_row, (max_rows + 1) if max_rows else ws_raw.max_row)
    rows_data = []
    for r in range(2, limit_rows + 1):
        row = []
        for c in range(1, ws_raw.max_column + 1):
            raw_v = ws_raw.cell(row=r, column=c).value
            data_v = ws_data.cell(row=r, column=c).value
            cell_coord = f"{get_column_letter(c)}{r}"
            row.append(evaluate_formula_cell(raw_v, data_v, ws_raw, ws_data, cell_coord=cell_coord))
        rows_data.append(row)

    df = pd.DataFrame(rows_data, columns=headers)
    return df, target_sheet, is_oversize


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

        # Initialize clean history
        self._history_stack.clear()
        self._version_counter = 0
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
        if not self._history_stack:
            return {
                "success": False,
                "message": "No previous changes to undo. Already at the initial version."
            }

        # Restore the state immediately prior to the latest mutation
        target_state = self._history_stack.pop()

        if not target_state.exists():
            return {
                "success": False,
                "message": f"Snapshot file missing: {target_state.name}"
            }

        shutil.copy2(target_state, self.current_file)
        return {
            "success": True,
            "message": f"Successfully reverted to previous snapshot ({target_state.stem}).",
            "remaining_undo_steps": len(self._history_stack)
        }

    def get_history_count(self) -> int:
        """Returns the number of undoable actions available."""
        return len(self._history_stack)

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
        wb = openpyxl.load_workbook(path, data_only=data_only)
        wb._file_path = Path(path).resolve()
        return wb

    def save_workbook(self, wb: openpyxl.Workbook) -> None:
        """Saves openpyxl Workbook to the current working copy and invalidates formula cache."""
        path = self.get_current_file_path()
        wb.save(path)
        invalidate_formula_cache(path)

    def is_current_file_oversize(self) -> bool:
        """Checks if current file has > 20,000 non-empty cells."""
        if not self.current_file:
            return False
        return is_file_oversize(self.current_file)

    def get_sheet_dataframe(self, sheet_name: Optional[str] = None, max_rows: Optional[int] = 500) -> Tuple[pd.DataFrame, str]:
        """Loads a sheet as a pandas DataFrame for viewing/inspection, evaluating formulas if unevaluated."""
        path = self.get_current_file_path()
        df, target_sheet, _ = load_sheet_dataframe(path, sheet_name=sheet_name, max_rows=max_rows)
        return df, target_sheet

    def get_sheet_names(self) -> List[str]:
        """Lists all sheet names in the workbook."""
        wb = self.open_workbook()
        return wb.sheetnames


def _get_newest_mtime(folder: Path) -> float:
    """Finds the newest mtime recursively inside folder, including the folder itself."""
    try:
        newest = folder.stat().st_mtime
    except Exception:
        newest = 0.0

    try:
        for p in folder.rglob("*"):
            try:
                t = p.stat().st_mtime
                if t > newest:
                    newest = t
            except Exception:
                pass
    except Exception:
        pass

    return newest


def cleanup_stale_sessions(
    workspace_root: Optional[str | Path] = None,
    temp_uploads_root: Optional[str | Path] = None,
    max_age_hours: Optional[float] = None,
    active_session_ids: Optional[List[str]] = None
) -> Dict[str, int]:
    """Removes session folders and temp upload directories older than max_age_hours (default 24h).
    Decides staleness by the newest mtime found inside the folder (recursively, including history/ and current.xlsx).
    Ensures active_session_ids are never touched. Configurable via EXCEL_SESSION_MAX_AGE_HOURS env var.
    """
    if max_age_hours is None:
        try:
            max_age_hours = float(os.getenv("EXCEL_SESSION_MAX_AGE_HOURS", "24"))
        except (ValueError, TypeError):
            max_age_hours = 24.0

    if workspace_root is None:
        workspace_root = os.getenv("EXCEL_WORKSPACE_DIR", "workspace")
    if temp_uploads_root is None:
        temp_uploads_root = os.getenv("EXCEL_TEMP_UPLOADS_DIR", "temp_uploads")

    max_age_seconds = max_age_hours * 3600
    now = time.time()
    active_set = set(active_session_ids or [])

    removed_workspaces = 0
    removed_uploads = 0

    ws_path = Path(workspace_root)
    if ws_path.exists() and ws_path.is_dir():
        for session_dir in ws_path.iterdir():
            if not session_dir.is_dir():
                continue
            if session_dir.name in active_set:
                continue
            try:
                newest_mtime = _get_newest_mtime(session_dir)
                if (now - newest_mtime) > max_age_seconds:
                    shutil.rmtree(session_dir, ignore_errors=True)
                    removed_workspaces += 1
            except Exception:
                pass

    up_path = Path(temp_uploads_root)
    if up_path.exists() and up_path.is_dir():
        for item in up_path.iterdir():
            if item.name in active_set:
                continue
            try:
                newest_mtime = _get_newest_mtime(item) if item.is_dir() else item.stat().st_mtime
                if (now - newest_mtime) > max_age_seconds:
                    if item.is_dir():
                        shutil.rmtree(item, ignore_errors=True)
                    else:
                        item.unlink(missing_ok=True)
                    removed_uploads += 1
            except Exception:
                pass

    return {
        "removed_workspaces": removed_workspaces,
        "removed_uploads": removed_uploads
    }
