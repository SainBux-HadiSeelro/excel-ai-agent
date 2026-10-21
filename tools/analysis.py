"""
Analysis tools for Excel AI Agent:
- remove_duplicates (destructive guard)
- sort_data
- filter_data
- create_pivot_summary (pandas group by / pivot)
"""

from __future__ import annotations

from typing import Dict, Any, List, Optional, Union
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from workspace_manager import WorkspaceManager


def remove_duplicates(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    columns: Optional[List[str]] = None,
    subset_columns: Optional[List[str]] = None,
    keep: str = "first",
    confirmed: bool = False
) -> Dict[str, Any]:
    """
    Removes duplicate rows based on subset columns (or all columns).
    Destructive operation requiring confirmation.
    """
    wb = manager.open_workbook(data_only=False)
    ws = wb[sheet_name] if sheet_name else wb.active

    target_columns = columns or subset_columns

    if ws.max_row <= 1:
        return {"message": "Sheet has no data rows to deduplicate.", "removed_count": 0}

    # Load data rows into pandas for fast duplicate detection
    headers = [str(ws.cell(row=1, column=c).value or f"Col_{c}") for c in range(1, ws.max_column + 1)]
    data = []
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        data.append(row_vals)

    df = pd.DataFrame(data, columns=headers)
    total_before = len(df)

    # Determine subset columns
    subset = None
    if target_columns:
        valid_cols = []
        for c in target_columns:
            c_str = str(c).strip()
            if c_str in headers:
                valid_cols.append(c_str)
            elif c_str.upper() in [get_column_letter(i+1) for i in range(len(headers))]:
                idx = [get_column_letter(i+1) for i in range(len(headers))].index(c_str.upper())
                valid_cols.append(headers[idx])
            else:
                for h in headers:
                    if h.lower() == c_str.lower():
                        valid_cols.append(h)
                        break
        if valid_cols:
            subset = valid_cols

    duplicates_mask = df.duplicated(subset=subset, keep=keep)
    duplicate_count = int(duplicates_mask.sum())

    if duplicate_count == 0:
        return {
            "success": True,
            "removed_count": 0,
            "message": "No duplicate rows found matching the criteria."
        }

    if not confirmed:
        # Preview duplicates
        sample_dup_rows = df[duplicates_mask].head(5).to_dict(orient="records")
        return {
            "requires_confirmation": True,
            "destructive_action": "remove_duplicates",
            "sheet_name": ws.title,
            "subset_columns": subset or "all columns",
            "duplicate_count": duplicate_count,
            "total_rows_before": total_before,
            "total_rows_after": total_before - duplicate_count,
            "sample_duplicates": sample_dup_rows,
            "warning": f"Found {duplicate_count} duplicate row(s). Removing them will permanently delete those rows.",
            "message": f"Confirmation required to remove {duplicate_count} duplicate row(s)."
        }

    # Proceed with deduplication
    manager.prepare_for_mutation(f"remove_duplicates: {duplicate_count} rows")
    dedup_df = df[~duplicates_mask]

    # Clear old data rows from worksheet
    while ws.max_row > 1:
        ws.delete_rows(2)

    # Re-append deduplicated rows
    for row in dedup_df.values.tolist():
        ws.append(row)

    manager.save_workbook(wb)

    return {
        "success": True,
        "sheet_name": ws.title,
        "removed_count": duplicate_count,
        "rows_remaining": len(dedup_df),
        "message": f"Successfully removed {duplicate_count} duplicate row(s) from sheet '{ws.title}'."
    }


def sort_data(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    by_columns: Optional[Union[str, List[str]]] = None,
    ascending: Union[bool, List[bool]] = True,
    has_headers: bool = True
) -> Dict[str, Any]:
    """Sorts sheet data by one or more columns."""
    wb = manager.open_workbook(data_only=False)
    if sheet_name and sheet_name not in wb.sheetnames:
        return {"error": f"Sheet '{sheet_name}' does not exist. Available: {wb.sheetnames}"}
    ws = wb[sheet_name] if sheet_name else wb.active

    if ws.max_row <= 1:
        return {"message": "Sheet has insufficient data to sort."}

    start_row = 2 if has_headers else 1
    headers = [str(ws.cell(row=1, column=c).value or f"Col_{c}") for c in range(1, ws.max_column + 1)] if has_headers else [f"Col_{c}" for c in range(1, ws.max_column + 1)]

    data = []
    for r in range(start_row, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        data.append(row_vals)

    df = pd.DataFrame(data, columns=headers)

    # Match sort column
    if isinstance(by_columns, str):
        sort_cols = [by_columns]
    elif isinstance(by_columns, list):
        sort_cols = by_columns
    else:
        sort_cols = [headers[0]]

    # Map column letters to column headers if needed
    matched_cols = []
    for sc in sort_cols:
        sc_str = str(sc).strip()
        if sc_str in headers:
            matched_cols.append(sc_str)
        elif sc_str.upper() in [get_column_letter(i+1) for i in range(len(headers))]:
            idx = [get_column_letter(i+1) for i in range(len(headers))].index(sc_str.upper())
            matched_cols.append(headers[idx])
        else:
            for h in headers:
                if h.lower() == sc_str.lower():
                    matched_cols.append(h)
                    break

    if not matched_cols:
        matched_cols = [headers[0]]

    # Handle sorting in pandas
    # Convert numeric strings safely if mixed
    df_sorted = df.sort_values(by=matched_cols, ascending=ascending)

    manager.prepare_for_mutation("sort_data")

    # Clear and rewrite rows
    while ws.max_row >= start_row:
        ws.delete_rows(start_row)

    for row_vals in df_sorted.values.tolist():
        ws.append(row_vals)

    manager.save_workbook(wb)

    return {
        "success": True,
        "sheet_name": ws.title,
        "sorted_by": matched_cols,
        "ascending": ascending,
        "row_count": len(df_sorted),
        "message": f"Successfully sorted {len(df_sorted)} rows by {matched_cols} in sheet '{ws.title}'."
    }


def filter_data(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    column: str = "",
    condition: str = "equals",  # equals, contains, greater_than, less_than, not_equals, is_null
    value: Any = None,
    output_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """
    Filters rows matching a condition. If output_sheet is specified, saves the filtered
    data into that sheet; otherwise returns filtered rows summary.
    """
    wb = manager.open_workbook(data_only=False)
    ws = wb[sheet_name] if sheet_name else wb.active

    headers = [str(ws.cell(row=1, column=c).value or f"Col_{c}") for c in range(1, ws.max_column + 1)]
    data = []
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        data.append(row_vals)

    df = pd.DataFrame(data, columns=headers)

    # Resolve column
    target_col = None
    col_str = str(column).strip()
    if col_str in headers:
        target_col = col_str
    elif col_str.upper() in [get_column_letter(i+1) for i in range(len(headers))]:
        idx = [get_column_letter(i+1) for i in range(len(headers))].index(col_str.upper())
        target_col = headers[idx]
    else:
        # Case-insensitive header match
        for h in headers:
            if h.lower() == col_str.lower():
                target_col = h
                break

    if not target_col:
        return {"error": f"Column '{column}' not found. Available columns: {headers}"}

    series = df[target_col]
    cond = condition.lower()

    if cond in ("equals", "eq", "=="):
        mask = series.astype(str).str.lower() == str(value).lower()
    elif cond in ("not_equals", "neq", "!="):
        mask = series.astype(str).str.lower() != str(value).lower()
    elif cond in ("contains", "like"):
        mask = series.astype(str).str.contains(str(value), case=False, na=False)
    elif cond in ("greater_than", "gt", ">"):
        numeric_series = pd.to_numeric(series, errors="coerce")
        mask = numeric_series > float(value)
    elif cond in ("less_than", "lt", "<"):
        numeric_series = pd.to_numeric(series, errors="coerce")
        mask = numeric_series < float(value)
    elif cond in ("is_null", "empty"):
        mask = series.isna() | (series == "")
    elif cond in ("not_null", "not_empty"):
        mask = series.notna() & (series != "")
    else:
        return {"error": f"Unsupported condition '{condition}'."}

    filtered_df = df[mask]

    if output_sheet:
        manager.prepare_for_mutation(f"filter_data -> {output_sheet}")
        wb = manager.open_workbook()
        if output_sheet in wb.sheetnames:
            target_ws = wb[output_sheet]
            while target_ws.max_row > 0:
                target_ws.delete_rows(1)
        else:
            target_ws = wb.create_sheet(title=output_sheet)

        target_ws.append(headers)
        for row_vals in filtered_df.values.tolist():
            target_ws.append(row_vals)
        manager.save_workbook(wb)

        return {
            "success": True,
            "matched_rows": len(filtered_df),
            "output_sheet": output_sheet,
            "message": f"Filtered {len(filtered_df)} rows matching {target_col} {condition} '{value}' into sheet '{output_sheet}'."
        }

    return {
        "success": True,
        "matched_rows": len(filtered_df),
        "sample_results": filtered_df.head(10).to_dict(orient="records"),
        "message": f"Found {len(filtered_df)} rows matching {target_col} {condition} '{value}'."
    }


def create_pivot_summary(
    manager: WorkspaceManager,
    sheet_name: Optional[str] = None,
    index_columns: Union[str, List[str]] = "Region",
    agg_dict: Optional[Dict[str, Union[str, List[str]]]] = None,
    values_column: Optional[str] = None,
    aggfunc: str = "sum",
    output_sheet: Optional[str] = None
) -> Dict[str, Any]:
    """
    Creates a grouped summary / pivot table using pandas and writes to output_sheet.
    Supports index_columns (e.g. ['Region', 'Product']) and aggregations (sum, mean, count).
    """
    df, active_sheet = manager.get_sheet_dataframe(sheet_name=sheet_name, max_rows=None)

    if isinstance(index_columns, str):
        idx_cols = [index_columns]
    else:
        idx_cols = index_columns

    # Validate index columns against dataframe
    valid_idx = []
    for ic in idx_cols:
        matched = False
        for c in df.columns:
            if str(c).strip().lower() == str(ic).strip().lower():
                valid_idx.append(c)
                matched = True
                break
        if not matched:
            return {"error": f"Group column '{ic}' not found in sheet. Available columns: {list(df.columns)}"}

    if output_sheet is None:
        output_sheet = f"Summary_{valid_idx[0]}"

    # Determine aggregation
    try:
        if agg_dict:
            # e.g. {"Total_Revenue": "sum", "Units": "sum"}
            pivot = df.groupby(valid_idx).agg(agg_dict).reset_index()
        elif values_column:
            # find matching values column
            val_col = None
            for c in df.columns:
                if str(c).strip().lower() == str(values_column).strip().lower():
                    val_col = c
                    break
            if not val_col:
                return {"error": f"Values column '{values_column}' not found. Available: {list(df.columns)}"}
            
            # Ensure numeric
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce")
            pivot = df.groupby(valid_idx)[val_col].agg(aggfunc).reset_index()
        else:
            # Default to all numeric columns aggregated by sum
            numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
            numeric_cols = [c for c in numeric_cols if c not in valid_idx]
            if not numeric_cols:
                return {"error": "No numeric columns found to aggregate."}
            pivot = df.groupby(valid_idx)[numeric_cols].agg(aggfunc).reset_index()

        # Flatten multi-index columns if any
        if isinstance(pivot.columns, pd.MultiIndex):
            pivot.columns = ['_'.join(filter(None, col)).strip() for col in pivot.columns.values]

        manager.prepare_for_mutation(f"pivot_summary -> {output_sheet}")
        wb = manager.open_workbook()

        if output_sheet in wb.sheetnames:
            ws_target = wb[output_sheet]
            while ws_target.max_row > 0:
                ws_target.delete_rows(1)
        else:
            ws_target = wb.create_sheet(title=output_sheet)

        # Write pivot headers
        ws_target.append(list(pivot.columns))

        # Write rows
        for row in pivot.itertuples(index=False):
            ws_target.append(list(row))

        manager.save_workbook(wb)

        return {
            "success": True,
            "output_sheet": output_sheet,
            "group_by": valid_idx,
            "rows_generated": len(pivot),
            "columns": list(pivot.columns),
            "preview": pivot.head(10).to_dict(orient="records"),
            "message": f"Created pivot summary sheet '{output_sheet}' with {len(pivot)} grouped rows."
        }
    except Exception as e:
        return {"error": f"Failed creating pivot summary: {str(e)}"}
