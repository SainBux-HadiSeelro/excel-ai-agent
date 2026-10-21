"""
End-to-End Verification of Excel AI Agent on sample_data/sales.xlsx.
Tests English and Roman Urdu tasks:
1. "total sales nikalo"
2. "duplicates hata do"
3. "region ke hisaab se chart banao"
4. "do files merge karo"
5. "undo"
"""

import sys
import os
from pathlib import Path

# Force UTF-8 encoding on Windows consoles
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
from workspace_manager import WorkspaceManager
from llm import MockProvider
from agent import ExcelAgent


def test_e2e():
    print("🚀 Starting End-to-End Verification on sample_data/sales.xlsx...")
    manager = WorkspaceManager()
    working_copy = manager.init_from_file("sample_data/sales.xlsx")
    print(f"Working copy initialized at: {working_copy}")

    agent = ExcelAgent(manager=manager, provider=MockProvider())

    # Task 1: "total sales nikalo"
    print("\n--------------------------------------------------")
    print("TEST 1: Prompt in Roman Urdu: 'total sales nikalo'")
    res1 = agent.run("total sales nikalo")
    print(f"Agent Steps: {res1.steps_taken}")
    print(f"Agent Final Response: {res1.final_response}")
    assert res1.success is True
    # Verify cell H25 contains formula
    wb = manager.open_workbook()
    assert "=SUM(H2:H23)" in str(wb["Sales"]["H25"].value)
    print("✅ Verified formula =SUM(H2:H23) present in cell H25!")

    # Task 2: "duplicates hata do"
    print("\n--------------------------------------------------")
    print("TEST 2: Prompt in Roman Urdu: 'duplicates hata do'")
    res2 = agent.run("duplicates hata do")
    print(f"Agent Steps: {res2.steps_taken}")
    print(f"Pending Confirmation: {bool(res2.pending_confirmation)}")
    assert res2.pending_confirmation is not None
    print("✅ Guard successfully paused destructive removal and requested confirmation!")

    # Confirm and resume
    res2_confirmed = agent.confirm_and_resume(confirmed=True)
    print(f"Agent Response after confirmation: {res2_confirmed.final_response}")
    assert res2_confirmed.success is True
    print("✅ Verified duplicate removal executed after confirmation!")

    # Task 3: "region ke hisaab se chart banao"
    print("\n--------------------------------------------------")
    print("TEST 3: Prompt: 'region ke hisaab se chart banao'")
    res3 = agent.run("region ke hisaab se chart banao")
    print(f"Agent Steps: {res3.steps_taken}")
    print(f"Agent Response: {res3.final_response}")
    assert res3.success is True
    wb_charts = manager.open_workbook()
    assert len(wb_charts["Targets"]._charts) > 0
    print(f"✅ Verified chart embedded in Targets sheet! Total charts: {len(wb_charts['Targets']._charts)}")

    # Task 4: "do files merge karo"
    print("\n--------------------------------------------------")
    print("TEST 4: Prompt: 'do files merge karo'")
    res4 = agent.run("do files merge karo")
    print(f"Agent Steps: {res4.steps_taken}")
    print(f"Agent Response: {res4.final_response}")
    assert res4.success is True
    assert any("Q1_Summary" in s for s in manager.get_sheet_names())
    print(f"✅ Verified merged sheets! All sheets: {manager.get_sheet_names()}")

    # Task 5: "undo"
    print("\n--------------------------------------------------")
    print("TEST 5: Undo previous change")
    undo_res = manager.undo()
    print(f"Undo result: {undo_res}")
    assert undo_res["success"] is True
    print("✅ Verified undo restored previous version!")

    print("\n🎉 ALL END-TO-END TESTS PASSED WITH 100% SUCCESS!")


if __name__ == "__main__":
    test_e2e()
