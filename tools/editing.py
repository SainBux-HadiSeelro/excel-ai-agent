"""
Editing tools for Excel AI Agent:
- write_cells (with post-write verification)
- find_and_replace
- clean_data
"""

from __future__ import annotations

import re
from typing import Dict, Any, List, Optional, Union
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple
from workspace_manager import WorkspaceManager


def write_cells(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    updates: Optional[List[Dict[str, Any]]] = None,
    cell: Optional[str] = None,
    value: Optional[Any] = None
) -> Dict[str, Any]:
    """
    Writes values or Excel formulas to specified cells.
    Includes post-write verification by re-reading the written range.
    Updates can be passed as a list of dicts [{'cell': 'A1', 'value': 10}],
    or single cell & value arguments.
    """
    items_to_write = []
    if updates:
        items_to_write.extend(updates)
    if cell is not None and value is not None:
        items_to_write.append({"cell": cell, "value": value})

    if not items_to_write:
        return {"error": "No cell updates specified."}

    wb = manager.open_workbook(data_only=False)

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            return {"error": f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}"}
        ws = wb[sheet_name]
    else:
        ws = wb.active

    validated_updates = []
    for item in items_to_write:
        coord = str(item.get("cell", "")).strip().upper()
        if not coord:
            continue
        if "!" in coord:
            sheet_part, _, cell_part = coord.rpartition("!")
            clean_sheet = sheet_part.strip().strip("'")
            coord = cell_part.strip().strip("'")
            if not sheet_name and clean_sheet in wb.sheetnames:
                ws = wb[clean_sheet]
        try:
            # Validate coordinate syntax (e.g. A1, B25) before creating a snapshot
            coordinate_to_tuple(coord)
        except Exception as e:
            return {"error": f"Invalid cell coordinate '{coord}': {str(e)}"}
        validated_updates.append((coord, item.get("value")))

    if not validated_updates:
        return {"error": "No valid cell coordinates specified."}

    manager.prepare_for_mutation("write_cells")
    written_coordinates = []
    for coord, val in validated_updates:
        # Convert numeric string to int/float so formulas and Excel treat them numerically
        if isinstance(val, str) and not val.startswith("="):
            clean_num = val.strip().replace(",", "")
            try:
                if "." in clean_num or "e" in clean_num.lower():
                    val = float(clean_num)
                else:
                    val = int(clean_num)
            except (ValueError, TypeError):
                pass

        # Write value or formula
        ws[coord] = val
        written_coordinates.append(coord)

    manager.save_workbook(wb)

    # Post-write verification: Re-read the affected cells from saved workbook
    wb_verify = manager.open_workbook(data_only=False)
    ws_verify = wb_verify[ws.title]
    verification_results = {}
    for coord in written_coordinates:
        read_val = ws_verify[coord].value
        verification_results[coord] = read_val

    return {
        "success": True,
        "sheet_name": ws.title,
        "updated_cells_count": len(written_coordinates),
        "post_write_verification": verification_results,
        "message": f"Successfully wrote {len(written_coordinates)} cell(s) and verified contents."
    }


def find_and_replace(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    whole_cell: bool = False,
    columns: Optional[List[str]] = None
) -> Dict[str, Any]:
    """Finds and replaces text across the entire sheet or specified columns."""
    if not find_text:
        return {"error": "find_text parameter cannot be empty."}

    wb = manager.open_workbook(data_only=False)

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            return {"error": f"Sheet '{sheet_name}' not found."}
        ws = wb[sheet_name]
    else:
        ws = wb.active

    manager.prepare_for_mutation(f"find_replace: {find_text} -> {replace_text}")
    matched_cells = []
    flags = 0 if match_case else re.IGNORECASE
    target_pattern = re.escape(str(find_text))
    if whole_cell:
        target_pattern = f"^{target_pattern}$"
    regex = re.compile(target_pattern, flags)

    # Resolve column letters from names or letters
    headers = [str(ws.cell(row=1, column=c).value or f"Col_{c}") for c in range(1, ws.max_column + 1)]
    col_targets = []
    if columns:
        for c in columns:
            c_str = str(c).strip()
            if c_str.upper() in [get_column_letter(i+1) for i in range(len(headers))]:
                col_targets.append(c_str.upper())
            else:
                for i, h in enumerate(headers):
                    if h.lower() == c_str.lower():
                        col_targets.append(get_column_letter(i+1))
                        break

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            # If columns filter is specified
            if col_targets and cell.column_letter not in col_targets:
                continue

            str_val = str(cell.value)
            if regex.search(str_val):
                old_val = str_val
                new_val = regex.sub(replace_text, str_val)
                # If was number originally and remained purely numeric
                if isinstance(cell.value, (int, float)):
                    try:
                        new_val = float(new_val) if "." in new_val else int(new_val)
                    except ValueError:
                        pass
                cell.value = new_val
                matched_cells.append({
                    "coordinate": cell.coordinate,
                    "old_value": old_val,
                    "new_value": new_val
                })

    manager.save_workbook(wb)

    return {
        "success": True,
        "sheet_name": ws.title,
        "replacements_count": len(matched_cells),
        "sample_replacements": matched_cells[:10],
        "message": f"Replaced {len(matched_cells)} occurrence(s) in sheet '{ws.title}'."
    }


def clean_data(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    columns: Optional[List[str]] = None,
    trim_spaces: bool = True,
    fix_casing: Optional[str] = None,  # 'title', 'lower', 'upper', None
    convert_types: bool = True
) -> Dict[str, Any]:
    """
    Cleans messy data:
    - Trims whitespace from string values
    - Standardizes text casing ('title', 'lower', 'upper')
    - Converts string digits to int/float and parseable date strings
    """
    wb = manager.open_workbook(data_only=False)

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            return {"error": f"Sheet '{sheet_name}' not found."}
        ws = wb[sheet_name]
    else:
        ws = wb.active

    manager.prepare_for_mutation("clean_data")
    cleaned_cells_count = 0

    headers = [str(ws.cell(row=1, column=c).value or f"Col_{c}") for c in range(1, ws.max_column + 1)]
    col_targets = []
    if columns:
        for c in columns:
            c_str = str(c).strip()
            if c_str.upper() in [get_column_letter(i+1) for i in range(len(headers))]:
                col_targets.append(c_str.upper())
            else:
                for i, h in enumerate(headers):
                    if h.lower() == c_str.lower():
                        col_targets.append(get_column_letter(i+1))
                        break

    # Date parsing formats
    date_formats = ["%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"]

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        if row_idx == 1:
            # Skip header row unless explicitly requested
            continue

        for cell in row:
            if cell.value is None or str(cell.value).startswith("="):
                # Skip formulas or empty cells
                continue

            if col_targets and cell.column_letter not in col_targets:
                continue

            val = cell.value
            changed = False

            if isinstance(val, str):
                # 1. Trim spaces
                if trim_spaces:
                    trimmed = val.strip()
                    # Also collapse multiple internal spaces
                    trimmed = re.sub(r"\s+", " ", trimmed)
                    if trimmed != val:
                        val = trimmed
                        changed = True

                # 2. Fix casing
                if fix_casing and isinstance(val, str) and not val.isnumeric():
                    mode = fix_casing.lower()
                    if mode == "title":
                        new_casing = val.title()
                    elif mode == "lower":
                        new_casing = val.lower()
                    elif mode == "upper":
                        new_casing = val.upper()
                    else:
                        new_casing = val

                    if new_casing != val:
                        val = new_casing
                        changed = True

                # 3. Type conversion
                if convert_types and isinstance(val, str):
                    # Check integer
                    if re.match(r"^-?\d+$", val):
                        try:
                            val = int(val)
                            changed = True
                        except ValueError:
                            pass
                    # Check float
                    elif re.match(r"^-?\d+\.\d+$", val):
                        try:
                            val = float(val)
                            changed = True
                        except ValueError:
                            pass
                    # Check dates
                    else:
                        for fmt in date_formats:
                            try:
                                parsed_date = datetime.strptime(val, fmt).date()
                                val = parsed_date
                                cell.number_format = "YYYY-MM-DD"
                                changed = True
                                break
                            except ValueError:
                                pass

            if changed:
                cell.value = val
                cleaned_cells_count += 1

    manager.save_workbook(wb)

    return {
        "success": True,
        "sheet_name": ws.title,
        "cleaned_cells_count": cleaned_cells_count,
        "message": f"Successfully cleaned {cleaned_cells_count} cell(s) in sheet '{ws.title}'."
    }
