"""
Tool Registry and Dispatcher for Excel AI Agent.
Provides OpenAI / Anthropic / Gemini compatible JSON schemas and tool execution.
"""

from __future__ import annotations

import json
from typing import Dict, Any, List, Optional
from workspace_manager import WorkspaceManager
from tools.inspection import list_sheets, read_range, get_sheet_summary
from tools.editing import write_cells, find_and_replace, clean_data
from tools.structure import (
    create_sheet,
    delete_sheet,
    rename_sheet,
    insert_rows,
    delete_rows,
    insert_columns,
    delete_columns,
)
from tools.analysis import remove_duplicates, sort_data, filter_data, create_pivot_summary
from tools.formatting import format_range
from tools.visualization import add_chart
from tools.workbook_ops import merge_workbooks, undo_last_change

# Full list of tool schemas for LLM function calling
TOOL_DEFINITIONS = [
    {
        "name": "list_sheets",
        "description": "Lists all sheets in the workbook with row/column dimensions and active sheet name.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": []
        }
    },
    {
        "name": "read_range",
        "description": "Reads data or formulas from a cell range (e.g. 'A1:E20' or entire sheet). Limits row count for safety.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Name of the sheet to read from (optional, defaults to active sheet)."},
                "cell_range": {"type": "string", "description": "Range expression like 'A1:D10' or 'A1' (optional, defaults to entire sheet up to max_rows)."},
                "max_rows": {"type": "integer", "description": "Maximum number of rows to return (default 50).", "default": 50},
                "show_formulas": {"type": "boolean", "description": "True to show underlying Excel formulas (e.g. =SUM(A1:A5)), False to show computed values.", "default": True}
            },
            "required": []
        }
    },
    {
        "name": "get_sheet_summary",
        "description": "Returns sheet overview: total rows/cols, column headers, inferred types, null counts, and sample rows. Handles 50k+ rows safely.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Name of the sheet (optional, defaults to active sheet)."},
                "max_sample_rows": {"type": "integer", "description": "Number of sample rows to return (default 5).", "default": 5}
            },
            "required": []
        }
    },
    {
        "name": "write_cells",
        "description": "Writes values or Excel formulas to specific cells and automatically verifies them by re-reading the written range.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Name of sheet (optional, defaults to active sheet)."},
                "updates": {
                    "type": "array",
                    "description": "List of cell updates, e.g. [{'cell': 'H2', 'value': '=F2*G2'}, {'cell': 'I2', 'value': 'Approved'}]",
                    "items": {
                        "type": "object",
                        "properties": {
                            "cell": {"type": "string", "description": "Cell coordinate, e.g. 'A1', 'H20'."},
                            "value": {"type": "string", "description": "Value to write (number, string, or formula starting with '=')"}
                        },
                        "required": ["cell", "value"]
                    }
                },
                "cell": {"type": "string", "description": "Single cell coordinate (optional shortcut if only writing 1 cell)."},
                "value": {"type": "string", "description": "Value or formula for single cell (optional shortcut)."}
            },
            "required": []
        }
    },
    {
        "name": "find_and_replace",
        "description": "Finds and replaces text or numbers across a sheet or specified columns.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "find_text": {"type": "string", "description": "Text to search for."},
                "replace_text": {"type": "string", "description": "Replacement text."},
                "match_case": {"type": "boolean", "description": "Whether to match exact casing (default False).", "default": False},
                "whole_cell": {"type": "boolean", "description": "Whether the match must be the entire cell value (default False).", "default": False},
                "columns": {"type": "array", "items": {"type": "string"}, "description": "Optional list of column letters or names to restrict replacement to."}
            },
            "required": ["find_text", "replace_text"]
        }
    },
    {
        "name": "clean_data",
        "description": "Cleans data in a sheet: trims leading/trailing whitespace, standardizes text casing, and converts numbers/dates from text.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "columns": {"type": "array", "items": {"type": "string"}, "description": "Optional list of column letters (e.g. ['C', 'D']) to clean."},
                "trim_spaces": {"type": "boolean", "description": "Trim whitespace from start and end of strings (default True).", "default": True},
                "fix_casing": {"type": "string", "description": "'title', 'lower', 'upper', or null.", "enum": ["title", "lower", "upper"]},
                "convert_types": {"type": "boolean", "description": "Convert string digits to int/float and parseable date strings (default True).", "default": True}
            },
            "required": []
        }
    },
    {
        "name": "create_sheet",
        "description": "Creates a new blank sheet with the specified title.",
        "parameters": {
            "type": "object",
            "properties": {
                "title": {"type": "string", "description": "Name for the new sheet."},
                "index": {"type": "integer", "description": "Optional tab index position."}
            },
            "required": ["title"]
        }
    },
    {
        "name": "delete_sheet",
        "description": "Deletes a sheet. Destructive operation: requires confirmation from the user.",
        "parameters": {
            "type": "object",
            "properties": {
                "title": {"type": "string", "description": "Name of sheet to delete."},
                "confirmed": {"type": "boolean", "description": "Set to True once the user has explicitly confirmed deletion.", "default": False}
            },
            "required": ["title"]
        }
    },
    {
        "name": "rename_sheet",
        "description": "Renames an existing sheet.",
        "parameters": {
            "type": "object",
            "properties": {
                "old_title": {"type": "string", "description": "Existing sheet name."},
                "new_title": {"type": "string", "description": "New sheet name."}
            },
            "required": ["old_title", "new_title"]
        }
    },
    {
        "name": "insert_rows",
        "description": "Inserts one or more blank rows before the given row index.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "row_idx": {"type": "integer", "description": "1-based row index before which to insert.", "default": 1},
                "count": {"type": "integer", "description": "Number of rows to insert (default 1).", "default": 1}
            },
            "required": ["row_idx"]
        }
    },
    {
        "name": "delete_rows",
        "description": "Deletes rows starting from row_idx. Destructive operation: requires confirmation.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "row_idx": {"type": "integer", "description": "1-based row index to start deleting from."},
                "count": {"type": "integer", "description": "Number of rows to delete.", "default": 1},
                "confirmed": {"type": "boolean", "description": "Set to True once user confirmed.", "default": False}
            },
            "required": ["row_idx"]
        }
    },
    {
        "name": "insert_columns",
        "description": "Inserts one or more blank columns before the specified column letter or index.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "col_idx": {"type": "string", "description": "Column letter (e.g. 'C') or 1-based integer index."},
                "count": {"type": "integer", "description": "Number of columns to insert.", "default": 1}
            },
            "required": ["col_idx"]
        }
    },
    {
        "name": "delete_columns",
        "description": "Deletes columns starting from col_idx. Destructive operation: requires confirmation.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "col_idx": {"type": "string", "description": "Column letter (e.g. 'C') or 1-based integer index."},
                "count": {"type": "integer", "description": "Number of columns to delete.", "default": 1},
                "confirmed": {"type": "boolean", "description": "Set to True once user confirmed.", "default": False}
            },
            "required": ["col_idx"]
        }
    },
    {
        "name": "remove_duplicates",
        "description": "Identifies and removes duplicate rows based on chosen columns. Destructive: requires user confirmation before final removal.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "columns": {"type": "array", "items": {"type": "string"}, "description": "List of column names or letters to check duplicates for (e.g. ['Order_ID'] or ['A']). Defaults to all columns."},
                "keep": {"type": "string", "description": "'first' or 'last' occurrence to keep.", "enum": ["first", "last"], "default": "first"},
                "confirmed": {"type": "boolean", "description": "Must be set to True after user confirms the removal.", "default": False}
            },
            "required": []
        }
    },
    {
        "name": "sort_data",
        "description": "Sorts rows in a sheet by one or more columns.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "by_columns": {"type": "string", "description": "Single column name/letter or list of columns to sort by (e.g. 'Region' or ['Region', 'Total_Revenue'])."},
                "ascending": {"type": "boolean", "description": "True for ascending, False for descending (default True).", "default": True},
                "has_headers": {"type": "boolean", "description": "Whether row 1 is header (default True).", "default": True}
            },
            "required": ["by_columns"]
        }
    },
    {
        "name": "filter_data",
        "description": "Filters rows by condition. Can preview or write matching rows to an output sheet.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "column": {"type": "string", "description": "Column name or letter to filter by."},
                "condition": {"type": "string", "description": "Condition type.", "enum": ["equals", "not_equals", "contains", "greater_than", "less_than", "is_null", "not_null"]},
                "value": {"type": "string", "description": "Value to compare against."},
                "output_sheet": {"type": "string", "description": "Optional name of a sheet to save filtered rows into."}
            },
            "required": ["column", "condition"]
        }
    },
    {
        "name": "format_range",
        "description": "Formats cells: bold, font color, background color, number formatting (currency/dates), borders, text alignment, and column auto-fit.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "cell_range": {"type": "string", "description": "Range to format, e.g. 'A1:H1' or 'F2:F20'."},
                "bold": {"type": "boolean", "description": "Make font bold."},
                "italic": {"type": "boolean", "description": "Make font italic."},
                "font_color": {"type": "string", "description": "Hex font color e.g. 'FFFFFF' or 'FF0000'."},
                "bg_color": {"type": "string", "description": "Hex background color e.g. '1F4E78' or 'E7E6E6'."},
                "font_size": {"type": "integer", "description": "Font size in points."},
                "number_format": {"type": "string", "description": "Excel number format string, e.g. '$#,##0.00', '#,##0', '0.0%', 'YYYY-MM-DD'."},
                "border": {"type": "string", "description": "Border style.", "enum": ["thin", "medium", "thick", "double"]},
                "align_horizontal": {"type": "string", "description": "Horizontal alignment.", "enum": ["left", "center", "right"]},
                "auto_fit_columns": {"type": "boolean", "description": "Whether to auto-adjust column widths based on contents.", "default": False}
            },
            "required": []
        }
    },
    {
        "name": "add_chart",
        "description": "Embeds a native Excel chart (Bar, Column, Line, or Pie) into the worksheet.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Sheet name (optional)."},
                "chart_type": {"type": "string", "description": "Type of chart.", "enum": ["bar", "column", "line", "pie"]},
                "data_range": {"type": "string", "description": "Cell range of numeric data with headers, e.g. 'H1:H10' or 'B1:C5'."},
                "categories_range": {"type": "string", "description": "Cell range of category labels for X-axis, e.g. 'C2:C10'."},
                "title": {"type": "string", "description": "Title of the chart."},
                "target_cell": {"type": "string", "description": "Top-left cell where the chart is placed (default 'K2').", "default": "K2"},
                "x_title": {"type": "string", "description": "Optional X-axis title."},
                "y_title": {"type": "string", "description": "Optional Y-axis title."}
            },
            "required": ["chart_type", "data_range"]
        }
    },
    {
        "name": "create_pivot_summary",
        "description": "Groups sheet data using pandas and creates an aggregated pivot summary table on a new sheet (sum, mean, count).",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet_name": {"type": "string", "description": "Source sheet name (optional)."},
                "index_columns": {"type": "string", "description": "Column or list of columns to group by (e.g. 'Region' or ['Region', 'Product'])."},
                "values_column": {"type": "string", "description": "Column to aggregate (e.g. 'Total_Revenue')."},
                "aggfunc": {"type": "string", "description": "Aggregation function.", "enum": ["sum", "mean", "count", "min", "max"], "default": "sum"},
                "output_sheet": {"type": "string", "description": "Name for the new summary sheet."}
            },
            "required": ["index_columns"]
        }
    },
    {
        "name": "merge_workbooks",
        "description": "Combines sheets or data rows from other Excel or CSV files into the current workbook.",
        "parameters": {
            "type": "object",
            "properties": {
                "source_files": {"type": "array", "items": {"type": "string"}, "description": "Paths to external files to merge."},
                "mode": {"type": "string", "description": "'add_sheets' to add as new tabs, or 'append_rows' to append into active sheet.", "enum": ["add_sheets", "append_rows"], "default": "add_sheets"},
                "target_sheet_names": {"type": "array", "items": {"type": "string"}, "description": "Optional custom names for imported sheets."}
            },
            "required": ["source_files"]
        }
    },
    {
        "name": "undo_last_change",
        "description": "Reverts the last modifying action and restores the previous snapshot.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": []
        }
    }
]


def execute_tool(manager: WorkspaceManager, tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
    """Executes a tool by name with arguments against the given WorkspaceManager."""
    try:
        if tool_name == "list_sheets":
            return list_sheets(manager)
        elif tool_name == "read_range":
            return read_range(manager, **arguments)
        elif tool_name == "get_sheet_summary":
            return get_sheet_summary(manager, **arguments)
        elif tool_name == "write_cells":
            return write_cells(manager, **arguments)
        elif tool_name == "find_and_replace":
            return find_and_replace(manager, **arguments)
        elif tool_name == "clean_data":
            return clean_data(manager, **arguments)
        elif tool_name == "create_sheet":
            return create_sheet(manager, **arguments)
        elif tool_name == "delete_sheet":
            return delete_sheet(manager, **arguments)
        elif tool_name == "rename_sheet":
            return rename_sheet(manager, **arguments)
        elif tool_name == "insert_rows":
            return insert_rows(manager, **arguments)
        elif tool_name == "delete_rows":
            return delete_rows(manager, **arguments)
        elif tool_name == "insert_columns":
            return insert_columns(manager, **arguments)
        elif tool_name == "delete_columns":
            return delete_columns(manager, **arguments)
        elif tool_name == "remove_duplicates":
            return remove_duplicates(manager, **arguments)
        elif tool_name == "sort_data":
            return sort_data(manager, **arguments)
        elif tool_name == "filter_data":
            return filter_data(manager, **arguments)
        elif tool_name == "format_range":
            return format_range(manager, **arguments)
        elif tool_name == "add_chart":
            return add_chart(manager, **arguments)
        elif tool_name == "create_pivot_summary":
            return create_pivot_summary(manager, **arguments)
        elif tool_name == "merge_workbooks":
            return merge_workbooks(manager, **arguments)
        elif tool_name == "undo_last_change":
            return undo_last_change(manager)
        else:
            return {"error": f"Unknown tool name: '{tool_name}'"}
    except Exception as e:
        return {"error": f"Tool execution failed for '{tool_name}': {str(e)}"}
