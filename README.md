<div align="center">

# 📊 Excel AI Agent
### *Autonomous Multilingual Spreadsheet Assistant (English & Roman Urdu)*

[![Live Demo](https://img.shields.io/badge/Live_Demo-Streamlit_App-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://excel-agent-ai.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.11%2B-blue?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.42%2B-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io)
[![Status: Live](https://img.shields.io/badge/Status-Live-00ff66?style=for-the-badge)](https://excel-agent-ai.streamlit.app/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

<br/>

### 🌐 **[👉 Click Here to Open Live Demo (excel-agent-ai.streamlit.app)](https://excel-agent-ai.streamlit.app/)**

[🎮 Live Web App](https://excel-agent-ai.streamlit.app/) • [🌟 Features](#-key-features) • [📸 Screenshots](#-screenshots) • [💬 Example Prompts](#-10-example-prompts-english--roman-urdu) • [🚀 Quick Start](#-quickstart--setup) • [🧰 Tools Reference](#-available-tools-reference)

---

</div>

## 📖 About the Project & Live Demo

**Excel AI Agent** is a production-quality autonomous spreadsheet assistant in Python that interprets natural language instructions in **English** and **Roman Urdu** (e.g., *"total sales nikalo"*, *"duplicates hata do"*, *"region ke hisaab se chart banao"*, *"do files merge karo"*) and executes them safely and reliably on Excel (`.xlsx`) files.

- 🌐 **Live Web App**: [https://excel-agent-ai.streamlit.app/](https://excel-agent-ai.streamlit.app/)
- 🚀 **Instant Evaluation**: No installation or local setup required—open the link, upload your workbook or click "Load Sample Data", and start chatting with the agent immediately.
- 🛡️ **Safe & Non-Destructive**: Isolated workspace with automatic snapshots, one-click undo, and confirmation guards for deletions.

---

## 🌟 Key Features

- **Multilingual Understanding**: Seamlessly understands and responds in both **English** and **Roman Urdu**.
- **21 Built-in Tools**: Full suite of spreadsheet operations including reading, writing formulas, cleaning, deduplication, sorting, filtering, pivot tables, native Excel charts, sheet management, and merging.
- **Safety First (Non-Destructive & Isolated)**:
  - Never touches or overwrites your original file; works strictly inside an isolated `workspace/<session_id>` session folder.
  - Automatic version snapshotting before every modifying action.
  - **Granular Single-Step Undo**: Each Undo reverts exactly one mutation, with an accurate live undo step counter in the UI.
  - **Bound Confirmation Guards**: Destructive actions (deleting sheets, rows, columns, removing duplicates) pause execution for user confirmation and bind strictly to the exact approved tool arguments upon execution without allowing the LLM to change parameters.
  - **Session Isolation & Auto-Cleanup**: User uploads are stored in unique per-session directories with sanitized filenames; stale sessions older than 24 hours (configurable via `EXCEL_SESSION_MAX_AGE_HOURS`) are automatically purged on startup.
- **Tiered Live Formula Evaluation**:
  - Dynamically evaluates formulas in preview and agent reads using a pure-Python engine (`formulas`), supporting functions like `SUMIF`, `AVERAGE`, `IF`, `COUNTIF`, and nested expressions without heavy external processes.
  - Automatically utilizes headless LibreOffice recalculation when available on `PATH`, and falls back to regex arithmetic as a safety net. Never crashes on complex or unknown formulas.
- **Post-Write Verification**: Re-reads affected cell ranges after every write to verify correctness.
- **Large File Optimization & Context Management**: Samples 50k+ row datasets, caps LLM tool output history at ~4,000 characters to conserve tokens, and enforces bounded conversation history without splitting tool calls.
- **Flexible LLM Provider Support**: Compatible with Google **Gemini** (`gemini-3.8-flash`), **OpenAI** (`gpt-6-luna`), **Anthropic** (`claude-haiku-5-5`), and an offline **Mock/Demo mode** for evaluation without API keys.
- **Dual Interfaces**: Interactive **Streamlit Web UI** and terminal **CLI REPL**.

---

## 📸 Screenshots

![Main Interface](docs/screenshots/app_main.png)
*Interactive Streamlit UI with file status bar, quick action pills, and chat panel.*

![Chat and Tool Execution Traces](docs/screenshots/chat_traces.png)
*Live collapsible tool execution traces with parameters, inspection results, and destructive confirmation cards.*

![Sheet Preview and Live Formulas](docs/screenshots/sheet_preview.png)
*Sheet preview table displaying evaluated formulas and data metrics.*

> [!NOTE]
> Community Cloud free Streamlit applications may enter sleep mode when inactive. If the app displays a sleep screen, simply click the **"Yes, get this app back up!"** button to wake it up in a few seconds.

---

## 📁 Project Structure

```text
excel-ai-agent/
├── app.py                     # Streamlit Web UI (Upload, Chat, Traces, Preview, Download, Undo)
├── cli.py                     # Terminal CLI interface (python cli.py file.xlsx)
├── agent.py                   # Autonomous tool-calling loop (history cap, bounded steps)
├── llm.py                     # Provider abstraction (Gemini, OpenAI, Anthropic, Mock)
├── workspace_manager.py       # Safe isolated workspace, snapshotting, undo, formula evaluator
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
│   └── registry.py            # Tool dispatcher & JSON schemas (21 tools)
├── tests/
│   ├── test_inspection.py
│   ├── test_editing.py
│   ├── test_structure.py
│   ├── test_analysis.py
│   ├── test_formatting_and_chart.py
│   ├── test_formulas.py       # Tests for SUMIF, AVERAGE, IF, nested formulas
│   ├── test_ui_escaping.py    # Tests for HTML escaping in Streamlit
│   ├── test_cleanup_and_uploads.py # Tests for upload isolation and session cleanup
│   ├── test_workbook_ops.py
│   └── test_agent_loop.py
├── sample_data/
│   ├── sales.xlsx             # Realistic dataset with formulas, messy text & duplicates
│   └── regional_extra.xlsx    # Secondary workbook for testing merges
├── docs/
│   └── screenshots/           # UI screenshots
├── .github/
│   └── workflows/
│       └── ci.yml             # GitHub Actions CI for Python 3.11 & 3.12
├── requirements.txt           # Runtime dependencies
├── requirements-dev.txt       # Development & test dependencies
├── .env.example
└── README.md
```

---

## 🛠️ Tech Stack & Requirements

- **Python**: 3.11+ (Tested on Python 3.11 & 3.12)
- **openpyxl**: Editing cell values, formulas, formatting, and charts
- **pandas**: High-performance data manipulation, filtering, and pivot aggregations
- **formulas**: Pure-Python Excel formula evaluation engine
- **streamlit**: Web UI with live tool execution traces and interactive sheet preview
- **pytest**: Test suite with 100% pass rate (39 passed tests)

---

## 🚀 Quickstart & Setup

### 1. Clone & Set Up Virtual Environment

```bash
git clone https://github.com/SainBux-HadiSeelro/excel-ai-agent.git
cd excel-ai-agent
python -m venv .venv

# On Windows PowerShell:
.\.venv\Scripts\Activate.ps1

# On Linux/macOS:
source .venv/bin/activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
pip install -r requirements-dev.txt
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

> 🌐 **Live Cloud Demo**: You can run the application directly in your browser without local installation at **[excel-agent-ai.streamlit.app](https://excel-agent-ai.streamlit.app/)**.

To run locally, launch the web interface:

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

The agent includes 21 function-calling tools defined in `tools/registry.py`:

1. **`list_sheets`**: Enumerates all worksheets, indices, row/column dimensions, and active sheet.
2. **`read_range`**: Reads cell ranges with optional row limiting and formula preservation.
3. **`get_sheet_summary`**: Scans sheet headers, inferred data types, null counts, and row counts (handles 50k+ rows via sampling).
4. **`write_cells`**: Writes values and live Excel formulas with automatic post-write verification.
5. **`find_and_replace`**: Replaces text or numeric matches across sheets or specified columns.
6. **`clean_data`**: Strips whitespace, standardizes casing, and converts numbers/dates stored as text.
7. **`create_sheet`**: Adds a new worksheet with duplicate name collision guards.
8. **`delete_sheet`**: Deletes a sheet (*Destructive: requires user confirmation*).
9. **`rename_sheet`**: Renames existing worksheets.
10. **`insert_rows`**: Inserts one or more blank rows at a specific index.
11. **`delete_rows`**: Deletes rows (*Destructive: requires user confirmation*).
12. **`insert_columns`**: Inserts one or more blank columns at a specific index.
13. **`delete_columns`**: Deletes columns (*Destructive: requires user confirmation*).
14. **`remove_duplicates`**: Detects and purges duplicate rows based on subset columns (*Destructive: requires user confirmation*).
15. **`sort_data`**: Sorts rows by one or more columns ascending or descending.
16. **`filter_data`**: Filters rows by conditions (`equals`, `contains`, `greater_than`, `is_null`, etc.) and optionally writes to a new sheet.
17. **`format_range`**: Sets font weights, colors, cell fills, borders, number formats (`$#,##0.00`, dates), and auto-fits columns.
18. **`add_chart`**: Embeds native Excel Bar, Column, Line, or Pie charts.
19. **`create_pivot_summary`**: Generates grouped pivot summaries using pandas and outputs a formatted table.
20. **`merge_workbooks`**: Imports sheets or appends rows from external `.xlsx` and `.csv` files.
21. **`undo_last_change`**: Rolls back the workbook to the prior version snapshot.

---

## 🧪 Running the Test Suite

Run the full pytest suite:

```bash
pytest -v
```

All 39 tests pass with 100% pass rate, verifying undo snapshots, confirmation binding, formula evaluation, HTML escaping, session cleanup, tool execution, and error self-correction:
```text
============================= test session starts =============================
platform win32 -- Python 3.12.2, pytest-9.1.1, pluggy-1.6.0
rootdir: excel-ai-agent
collected 39 items

tests/test_agent_loop.py::test_agent_one_step_tool_call PASSED           [  2%]
tests/test_agent_loop.py::test_agent_destructive_confirmation_pause PASSED [  5%]
tests/test_agent_loop.py::test_agent_error_recovery PASSED               [  7%]
tests/test_agent_loop.py::test_agent_destructive_confirm_execution PASSED [ 10%]
tests/test_agent_loop.py::test_agent_destructive_cancel_makes_no_changes PASSED [ 12%]
tests/test_agent_loop.py::test_agent_truncates_large_tool_output PASSED  [ 15%]
tests/test_agent_loop.py::test_agent_history_cap_boundary PASSED         [ 17%]
tests/test_agent_loop.py::test_agent_max_steps_custom_limit PASSED       [ 20%]
tests/test_analysis.py::test_remove_duplicates_guard_and_execution PASSED [ 23%]
tests/test_analysis.py::test_sort_data PASSED                            [ 25%]
tests/test_analysis.py::test_filter_data_output_sheet PASSED             [ 28%]
tests/test_analysis.py::test_create_pivot_summary PASSED                 [ 30%]
tests/test_cleanup_and_uploads.py::test_cleanup_stale_sessions_older_than_24h PASSED [ 33%]
tests/test_cleanup_and_uploads.py::test_cleanup_configurable_via_env PASSED [ 35%]
tests/test_cleanup_and_uploads.py::test_upload_path_sanitization PASSED  [ 38%]
tests/test_editing.py::test_write_cells_and_verify PASSED                [ 41%]
tests/test_editing.py::test_find_and_replace PASSED                      [ 43%]
tests/test_editing.py::test_clean_data_trim_and_casing PASSED            [ 46%]
tests/test_editing.py::test_write_cells_with_sheet_prefix PASSED         [ 48%]
tests/test_formatting_and_chart.py::test_format_range PASSED             [ 51%]
tests/test_formatting_and_chart.py::test_add_chart_bar_and_pie PASSED    [ 53%]
tests/test_formulas.py::test_evaluate_formulas_advanced PASSED           [ 56%]
tests/test_formulas.py::test_evaluate_formula_read_range_integration PASSED [ 58%]
tests/test_formulas.py::test_formula_evaluation_failure_never_crashes PASSED [ 61%]
tests/test_inspection.py::test_list_sheets PASSED                        [ 64%]
tests/test_inspection.py::test_read_range_with_formula PASSED            [ 66%]
tests/test_inspection.py::test_read_range_all_and_limit PASSED           [ 69%]
tests/test_inspection.py::test_get_sheet_summary PASSED                  [ 71%]
tests/test_inspection.py::test_read_range_invalid_sheet PASSED           [ 74%]
tests/test_structure.py::test_create_and_rename_sheet PASSED             [ 76%]
tests/test_structure.py::test_delete_sheet_guard_and_confirmed PASSED    [ 79%]
tests/test_structure.py::test_insert_and_delete_rows PASSED              [ 82%]
tests/test_structure.py::test_insert_and_delete_columns PASSED           [ 84%]
tests/test_ui_escaping.py::test_sheet_name_html_escaping PASSED          [ 87%]
tests/test_ui_escaping.py::test_filename_html_escaping PASSED            [ 89%]
tests/test_ui_escaping.py::test_destructive_warning_html_escaping PASSED [ 92%]
tests/test_workbook_ops.py::test_merge_workbooks PASSED                  [ 94%]
tests/test_workbook_ops.py::test_undo_last_change PASSED                 [ 97%]
tests/test_workbook_ops.py::test_undo_multiple_mutations PASSED          [100%]

============================= 39 passed in 19.95s =============================
```

Run end-to-end integration test:

```bash
python verify_end_to_end.py
```

---

## 📜 License

MIT License. Free for commercial and private use.
