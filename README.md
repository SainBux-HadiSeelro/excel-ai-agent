# 📊 Excel AI Agent

A complete, production-quality autonomous AI Agent in Python that interprets plain language instructions in **English** and **Roman Urdu** (e.g., *"total sales nikalo"*, *"duplicates hata do"*, *"region ke hisaab se chart banao"*, *"do files merge karo"*) and executes them safely and reliably on Excel (`.xlsx`) files.

---

## 🌟 Key Features

- **Multilingual Understanding**: Seamlessly understands and responds in both **English** and **Roman Urdu**.
- **15 Built-in Tools**: Full suite of spreadsheet operations including reading, writing formulas, cleaning, deduplication, sorting, filtering, pivot tables, native Excel charts, sheet management, and merging.
- **Safety First (Non-Destructive)**:
  - Never touches or overwrites your original file; works strictly inside an isolated `workspace/` session folder.
  - Automatic version snapshotting before every modifying action.
  - **Full Undo Support**: Revert any change with the click of a button or the `undo` command.
  - **Confirmation Guards**: Destructive operations (deleting rows, deleting sheets, removing duplicates) always prompt for explicit user confirmation before applying.
- **Live Formulas Over Hardcoded Values**: Prefers Excel formulas (`=SUM(...)`, `=AVERAGE(...)`, `=F2*G2`) so spreadsheets stay live and reactive.
- **Post-Write Verification**: Re-reads affected cell ranges after every write to verify correctness.
- **Large File Optimization**: Automatically samples 50k+ row datasets instead of overloading memory.
- **Flexible LLM Provider Support**: Compatible with Google **Gemini**, **OpenAI**, **Anthropic**, and an offline **Mock/Demo mode** for evaluation without API keys.
- **Dual Interfaces**: Interactive **Streamlit Web UI** and terminal **CLI REPL**.

---

## 📁 Project Structure

```text
excel/
├── app.py                     # Streamlit Web UI (Upload, Chat, Traces, Preview, Download, Undo)
├── cli.py                     # Terminal CLI interface (python cli.py file.xlsx)
├── agent.py                   # Autonomous tool-calling loop (capped at 15 steps)
├── llm.py                     # Provider abstraction (Gemini, OpenAI, Anthropic, Mock)
├── workspace_manager.py       # Safe isolated workspace, snapshotting, and undo manager
├── generate_sample_data.py    # Sample dataset generator script
├── verify_end_to_end.py       # End-to-end integration test runner
├── tools/
│   ├── __init__.py
│   ├── inspection.py          # list_sheets, read_range, get_sheet_summary
│   ├── editing.py             # write_cells (formulas + verification), find_and_replace, clean_data
│   ├── structure.py           # create_sheet, delete_sheet, rename_sheet, insert/delete rows & cols
│   ├── analysis.py            # remove_duplicates, sort_data, filter_data, create_pivot_summary
│   ├── formatting.py          # format_range (bold, colors, borders, number formats, auto-fit)
│   ├── visualization.py       # add_chart (Bar, Column, Line, Pie)
│   ├── workbook_ops.py        # merge_workbooks, undo_last_change
│   └── registry.py            # Tool dispatcher & JSON schemas
├── tests/
│   ├── test_inspection.py
│   ├── test_editing.py
│   ├── test_structure.py
│   ├── test_analysis.py
│   ├── test_formatting_and_chart.py
│   ├── test_workbook_ops.py
│   └── test_agent_loop.py
├── sample_data/
│   ├── sales.xlsx             # Realistic dataset with formulas, messy text & duplicates
│   └── regional_extra.xlsx    # Secondary workbook for testing merges
├── requirements.txt
├── .env.example
└── README.md
```

---

## 🛠️ Tech Stack & Requirements

- **Python**: 3.11+ (Tested on Python 3.12)
- **openpyxl**: Editing cell values, formulas, formatting, and charts
- **pandas**: High-performance data manipulation, filtering, and pivot aggregations
- **streamlit**: Web UI with live tool execution traces and interactive sheet preview
- **pytest**: Test suite with 100% pass rate

---

## 🚀 Quickstart & Setup

### 1. Clone & Set Up Virtual Environment

```bash
cd "d:/1-New folder/excel"
python -m venv .venv

# On Windows PowerShell:
.\.venv\Scripts\Activate.ps1

# On Linux/macOS:
source .venv/bin/activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure API Keys

Copy the sample environment file and provide your API key:

```bash
cp .env.example .env
```

Edit `.env`:
```env
LLM_PROVIDER=gemini
GEMINI_API_KEY=your_gemini_api_key_here

# Or for OpenAI / Anthropic:
# OPENAI_API_KEY=your_openai_api_key
# ANTHROPIC_API_KEY=your_anthropic_api_key
```

*(Note: If you don't have an API key right now, you can select **Offline Demo (Mock)** in the Streamlit UI or pass `--mock` in the CLI to test all functionality immediately!)*

### 4. Generate Sample Data

```bash
python generate_sample_data.py
```
This generates `sample_data/sales.xlsx` with realistic sales rows, formulas, messy strings, and duplicate entries.

---

## 🖥️ Running the Applications

### 1. Streamlit Web UI

Launch the web interface:

```bash
streamlit run app.py
```

**UI Highlights**:
- **Sidebar**: Upload `.xlsx` or `.csv`, select provider and model, enter API key, view undo step counter, and click **Download Updated File**.
- **Interactive Chat**: Type your request in English or Roman Urdu. Watch live collapsible tool call traces showing arguments and execution results.
- **Confirmation Prompts**: Automatic modal card warning before destructive operations (e.g. deleting rows, sheets, or duplicates).
- **Sheet Preview Tab**: View active sheet data with pagination and metadata metrics.

### 2. Terminal CLI Mode

Run on any workbook:

```bash
python cli.py sample_data/sales.xlsx
```

Or run in offline demo mode without an API key:

```bash
python cli.py --mock sample_data/sales.xlsx
```

**Interactive CLI Commands**:
- Type any plain language instruction (e.g., `total sales nikalo`)
- `undo` : Revert the last change
- `sheets` : List all sheets in the workbook
- `exit` or `quit` : Exit the REPL

---

## 💬 10 Example Prompts (English & Roman Urdu)

| # | Task | Language | Example Command |
|---|---|---|---|
| 1 | **Total Sales Calculation** | Roman Urdu | `total sales nikalo aur formula H25 mein likh do` |
| 2 | **Deduplication** | Roman Urdu | `Order_ID ke hisaab se duplicates hata do` |
| 3 | **Chart Creation** | Roman Urdu | `Targets sheet mein region ke hisaab se column chart banao` |
| 4 | **Workbook Merging** | Roman Urdu | `sample_data/regional_extra.xlsx file ko is workbook mein merge karo` |
| 5 | **Data Cleaning** | Roman Urdu | `Sales Rep column D ke extra spaces hatao aur text ko title case kardo` |
| 6 | **Total Calculation** | English | `Calculate total units and total revenue using Excel SUM formulas at the bottom` |
| 7 | **Pivot Table** | English | `Create a pivot summary grouped by Region showing the sum of Units` |
| 8 | **Filter Data** | English | `Filter all rows where Region is 'North' and save them to a new sheet called 'North_Sales'` |
| 9 | **Find and Replace** | English | `Find all occurrences of 'Laptop' and replace with 'Notebook'` |
| 10 | **Formatting** | English | `Make header row A1:I1 bold, dark blue background, white text, and auto-fit column widths` |

---

## 🧰 Available Tools Reference

1. **`list_sheets`**: Enumerates all worksheets, indices, row/column dimensions, and active sheet.
2. **`read_range`**: Reads cell ranges with optional row limiting and formula preservation.
3. **`get_sheet_summary`**: Scans sheet headers, inferred data types, null counts, and row counts (handles 50k+ rows via sampling).
4. **`write_cells`**: Writes values and live Excel formulas with automatic post-write verification.
5. **`create_sheet`**: Adds a new worksheet with duplicate name collision guards.
6. **`delete_sheet`**: Deletes a sheet (*Destructive: requires user confirmation*).
7. **`rename_sheet`**: Renames existing worksheets.
8. **`insert_rows` / `delete_rows`**: Inserts or removes rows (*delete_rows is destructive and guarded*).
9. **`insert_columns` / `delete_columns`**: Inserts or removes columns (*delete_columns is destructive and guarded*).
10. **`find_and_replace`**: Replaces text or numeric matches across sheets or specified columns.
11. **`remove_duplicates`**: Detects and purges duplicate rows based on subset columns (*Destructive: requires user confirmation*).
12. **`sort_data`**: Sorts rows by one or more columns ascending or descending.
13. **`filter_data`**: Filters rows by conditions (`equals`, `contains`, `greater_than`, `is_null`, etc.) and optionally writes to a new sheet.
14. **`format_range`**: Sets font weights, colors, cell fills, borders, number formats (`$#,##0.00`, dates), and auto-fits columns.
15. **`add_chart`**: Embeds native Excel Bar, Column, Line, or Pie charts.
16. **`create_pivot_summary`**: Generates grouped pivot summaries using pandas and outputs a formatted table.
17. **`merge_workbooks`**: Imports sheets or appends rows from external `.xlsx` and `.csv` files.
18. **`clean_data`**: Strips whitespace, standardizes casing, and converts numbers/dates stored as text.
19. **`undo_last_change`**: Rolls back the workbook to the prior version snapshot.

---

## 🧪 Running the Test Suite

Run the full pytest suite:

```bash
pytest -v
```

All 23 test suites pass, verifying tool execution, snapshot restoration, formula preservation, confirmation guards, and error self-correction:
```text
tests/test_agent_loop.py::test_agent_one_step_tool_call PASSED
tests/test_agent_loop.py::test_agent_destructive_confirmation_pause PASSED
tests/test_agent_loop.py::test_agent_error_recovery PASSED
tests/test_analysis.py::test_remove_duplicates_guard_and_execution PASSED
tests/test_analysis.py::test_sort_data PASSED
tests/test_analysis.py::test_filter_data_output_sheet PASSED
tests/test_analysis.py::test_create_pivot_summary PASSED
tests/test_editing.py::test_write_cells_and_verify PASSED
tests/test_editing.py::test_find_and_replace PASSED
tests/test_editing.py::test_clean_data_trim_and_casing PASSED
tests/test_formatting_and_chart.py::test_format_range PASSED
tests/test_formatting_and_chart.py::test_add_chart_bar_and_pie PASSED
tests/test_inspection.py::test_list_sheets PASSED
tests/test_inspection.py::test_read_range_with_formula PASSED
tests/test_inspection.py::test_read_range_all_and_limit PASSED
tests/test_inspection.py::test_get_sheet_summary PASSED
tests/test_inspection.py::test_read_range_invalid_sheet PASSED
tests/test_structure.py::test_create_and_rename_sheet PASSED
tests/test_structure.py::test_delete_sheet_guard_and_confirmed PASSED
tests/test_structure.py::test_insert_and_delete_rows PASSED
tests/test_structure.py::test_insert_and_delete_columns PASSED
tests/test_workbook_ops.py::test_merge_workbooks PASSED
tests/test_workbook_ops.py::test_undo_last_change PASSED

============================= 23 passed in 1.91s ==============================
```

Run end-to-end integration test:

```bash
python verify_end_to_end.py
```

---

## 📜 License

MIT License. Free for commercial and private use.
