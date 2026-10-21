"""
Command-Line Interface (CLI) for Excel AI Agent.
Usage:
    python cli.py sample_data/sales.xlsx
"""

from __future__ import annotations

import sys
import os
from pathlib import Path
from dotenv import load_dotenv

# Ensure UTF-8 output on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# Ensure local imports work cleanly
sys.path.insert(0, str(Path(__file__).parent.resolve()))

from workspace_manager import WorkspaceManager
from llm import get_llm_provider
from agent import ExcelAgent

load_dotenv()


def print_banner():
    print("=" * 65)
    print(" 📊 EXCEL AI AGENT — Autonomous Spreadsheet Assistant")
    print(" Works in English & Roman Urdu (e.g. 'total sales nikalo')")
    print(" Commands: 'undo', 'sheets', 'history', 'quit' or 'exit'")
    print("=" * 65)


def cli_on_step(step: int, tool_name: str, args: dict, result: dict):
    print(f"\n  ⚡ [Step {step}] Calling Tool: \033[1;36m{tool_name}\033[0m")
    if args:
        summary_args = {k: v for k, v in args.items() if k != "updates"}
        if "updates" in args:
            summary_args["updates_count"] = len(args["updates"])
        print(f"     Args: {summary_args}")
    if "error" in result:
        print(f"     ❌ \033[1;31mError: {result['error']}\033[0m")
    elif "message" in result:
        print(f"     ✅ \033[1;32m{result['message']}\033[0m")
    elif "warning" in result:
        print(f"     ⚠️ \033[1;33m{result['warning']}\033[0m")


def main():
    use_mock = False
    args = [a for a in sys.argv[1:] if a not in ("--mock", "--demo")]
    if "--mock" in sys.argv or "--demo" in sys.argv:
        use_mock = True

    if not args:
        print("Usage: python cli.py [--mock] <path_to_excel_file.xlsx>")
        file_input = input("Enter path to Excel file (or press enter for sample_data/sales.xlsx): ").strip()
        file_path = file_input if file_input else "sample_data/sales.xlsx"
    else:
        file_path = args[0]

    if not Path(file_path).exists():
        print(f"❌ File not found: {file_path}")
        sys.exit(1)

    print_banner()
    print(f"📂 Loading target workbook: {file_path}")

    manager = WorkspaceManager()
    working_copy = manager.init_from_file(file_path)
    print(f"🛡️  Safety mode active: Working on isolated copy at: {working_copy}")
    print(f"📋 Available sheets: {manager.get_sheet_names()}\n")

    provider = None
    if use_mock:
        from llm import MockProvider
        provider = MockProvider()
        print("🤖 LLM Provider: Offline Demo (MockProvider) — no API key required")
    else:
        try:
            provider = get_llm_provider()
            print(f"🤖 LLM Provider: {provider.__class__.__name__} ({provider.model or 'default'})")
        except Exception as e:
            print(f"⚠️  LLM Configuration Notice: {e}")
            print("Switching to Offline Demo mode so you can interact with the agent right away.\n")
            from llm import MockProvider
            provider = MockProvider()

    agent = ExcelAgent(manager=manager, provider=provider)

    while True:
        try:
            user_input = input("\n🧑 User > ").strip()
            if not user_input:
                continue

            cmd = user_input.lower()
            if cmd in ("exit", "quit", "q"):
                print("👋 Exiting Excel AI Agent. Working file preserved at:")
                print(f"   {manager.get_current_file_path()}")
                break

            if cmd == "undo":
                undo_res = manager.undo()
                if undo_res["success"]:
                    print(f"⏪ \033[1;32m{undo_res['message']}\033[0m")
                else:
                    print(f"⚠️ \033[1;33m{undo_res['message']}\033[0m")
                continue

            if cmd in ("sheets", "list"):
                print(f"📋 Sheets: {manager.get_sheet_names()}")
                continue

            # Run agent loop
            result = agent.run(user_input, on_step=cli_on_step)

            if result.pending_confirmation:
                print(f"\n{result.final_response}")
                confirm_choice = input("Do you confirm this action? [y/N]: ").strip().lower()
                is_confirmed = confirm_choice in ("y", "yes")
                resumed = agent.confirm_and_resume(confirmed=is_confirmed, on_step=cli_on_step)
                print(f"\n🤖 Agent > {resumed.final_response}")
            else:
                print(f"\n🤖 Agent > {result.final_response}")

        except KeyboardInterrupt:
            print("\n👋 Operation cancelled. Type 'exit' to quit.")
        except Exception as ex:
            print(f"\n❌ Error: {str(ex)}")


if __name__ == "__main__":
    main()
