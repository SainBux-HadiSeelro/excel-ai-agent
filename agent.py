"""
Excel AI Agent Core Loop.
Handles:
- Multilingual instructions (English and Roman Urdu)
- Tool-calling execution loop capped at 15 steps
- Inspection-first strategy
- One-line plan explanation
- Destructive operation confirmation guards
- Post-write verification & error self-correction
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional, Callable
from workspace_manager import WorkspaceManager
from llm import BaseLLMProvider, ToolCall, LLMResponse, get_llm_provider
from tools.registry import TOOL_DEFINITIONS, execute_tool

SYSTEM_PROMPT = """You are an expert, production-grade Excel AI Agent.
The user interacts with you in plain English or Roman Urdu (e.g., "total sales nikalo", "duplicates hata do", "region ke hisaab se chart banao", "do files merge karo").

Core Instructions & Behavioral Rules:
1. Language Consistency: Always formulate your explanations and final response in the exact same language as the user's prompt (English or Roman Urdu).
2. Inspection First: Before applying mutations or calculations, inspect the workbook structure and columns first using `list_sheets`, `get_sheet_summary`, or `read_range`.
3. Tool Calling & Execution: When asked to perform an action, calculation, or inspection, you MUST call the relevant tool(s) in your turn. NEVER output a text plan alone without calling tools unless you have completely concluded the task or are requesting explicit confirmation for a destructive action.
4. Prefer Live Formulas: ALWAYS prefer Excel formulas (e.g., =SUM(H2:H23), =AVERAGE(F2:F23), =F2*G2) over hardcoded values so the spreadsheet stays dynamic and interactive. When asked to calculate totals or aggregates, write the formula into the summary cell using write_cells.
5. Destructive Safety:
   - Destructive actions (delete_sheet, delete_rows, delete_columns, remove_duplicates) CANNOT be executed directly without user confirmation.
   - First call the tool without confirmed=True (or inspect) to discover the impact and items affected.
   - When the tool returns `requires_confirmation: True`, pause and clearly ask the user for confirmation with the exact impact.
   - Only supply confirmed=True when the user has explicitly confirmed the deletion.
6. Error Self-Correction: If a tool returns an error, analyze the error message, correct your parameters or range references, and self-correct.
7. Final Summary: After completing the task, give a short, helpful summary of what changed in the workbook.
"""


@dataclass
class ToolExecutionTrace:
    """Record of a tool called during an agent step."""
    step: int
    tool_name: str
    arguments: Dict[str, Any]
    result: Dict[str, Any]
    error: Optional[str] = None


@dataclass
class AgentResult:
    """Outcome of an agent execution."""
    success: bool
    final_response: str
    steps_taken: int
    tool_traces: List[ToolExecutionTrace] = field(default_factory=list)
    pending_confirmation: Optional[Dict[str, Any]] = None
    error: Optional[str] = None


class ExcelAgent:
    """Autonomous agent orchestrating LLM tool calling on Excel workbooks."""

    def __init__(
        self,
        manager: WorkspaceManager,
        provider: Optional[BaseLLMProvider] = None,
        max_steps: int = 15
    ):
        self.manager = manager
        self.provider = provider or get_llm_provider()
        self.max_steps = max_steps
        self.conversation_history: List[Dict[str, Any]] = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        self.pending_confirmation: Optional[Dict[str, Any]] = None

    def run(
        self,
        user_instruction: str,
        on_step: Optional[Callable[[int, str, Dict[str, Any], Dict[str, Any]], None]] = None
    ) -> AgentResult:
        """
        Executes a user instruction through the tool-calling loop.
        Caps execution at max_steps (default 15).
        """
        # Append user message
        self.conversation_history.append({"role": "user", "content": user_instruction})
        tool_traces: List[ToolExecutionTrace] = []

        step = 0
        final_text = ""

        while step < self.max_steps:
            step += 1

            # Request LLM decision
            try:
                response = self.provider.chat(
                    messages=self.conversation_history,
                    tools=TOOL_DEFINITIONS
                )
            except Exception as e:
                return AgentResult(
                    success=False,
                    final_response=f"LLM Error: {str(e)}",
                    steps_taken=step,
                    tool_traces=tool_traces,
                    error=str(e)
                )

            # If LLM replied with content
            if response.content:
                final_text = response.content

            # If no tools called, we've reached a final response
            if not response.tool_calls:
                self.conversation_history.append({"role": "assistant", "content": response.content or ""})
                return AgentResult(
                    success=True,
                    final_response=final_text,
                    steps_taken=step,
                    tool_traces=tool_traces
                )

            # Record assistant turn with tool calls (OpenAI, Anthropic & Gemini compliant)
            assistant_msg = {
                "role": "assistant",
                "content": response.content or "",
                "tool_calls": [
                    {
                        "id": tc.id,
                        "type": "function",
                        "function": {
                            "name": tc.name,
                            "arguments": json.dumps(tc.arguments) if isinstance(tc.arguments, dict) else str(tc.arguments)
                        },
                        "name": tc.name,
                        "arguments": tc.arguments
                    }
                    for tc in response.tool_calls
                ]
            }
            self.conversation_history.append(assistant_msg)

            # Execute each tool call
            for tc in response.tool_calls:
                tool_output = execute_tool(self.manager, tc.name, tc.arguments)

                trace = ToolExecutionTrace(
                    step=step,
                    tool_name=tc.name,
                    arguments=tc.arguments,
                    result=tool_output,
                    error=tool_output.get("error")
                )
                tool_traces.append(trace)

                if on_step:
                    on_step(step, tc.name, tc.arguments, tool_output)

                # Append tool result to history
                self.conversation_history.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "name": tc.name,
                    "content": json.dumps(tool_output, default=str)
                })

                # Check if this requires user confirmation
                if tool_output.get("requires_confirmation"):
                    self.pending_confirmation = {
                        "tool_name": tc.name,
                        "arguments": tc.arguments,
                        "details": tool_output
                    }
                    confirmation_msg = tool_output.get("warning") or tool_output.get("message")
                    conf_final = f"⚠️ Confirmation Required: {confirmation_msg}"
                    self.conversation_history.append({"role": "assistant", "content": conf_final})
                    return AgentResult(
                        success=True,
                        final_response=conf_final,
                        steps_taken=step,
                        tool_traces=tool_traces,
                        pending_confirmation=self.pending_confirmation
                    )

        return AgentResult(
            success=False,
            final_response="Agent reached maximum execution step limit (15) without concluding.",
            steps_taken=step,
            tool_traces=tool_traces,
            error="Max steps exceeded"
        )

    def confirm_and_resume(
        self,
        confirmed: bool = True,
        on_step: Optional[Callable[[int, str, Dict[str, Any], Dict[str, Any]], None]] = None
    ) -> AgentResult:
        """Resumes execution after a user responds to a pending confirmation."""
        if not self.pending_confirmation:
            return AgentResult(
                success=False,
                final_response="No pending action to confirm.",
                steps_taken=0
            )

        pending = self.pending_confirmation
        tool_name = pending["tool_name"]
        stored_args = pending.get("arguments", {})
        self.pending_confirmation = None

        if not confirmed:
            return AgentResult(
                success=True,
                final_response=f"❌ {tool_name} operation cancelled by user. No changes were made.",
                steps_taken=0,
                tool_traces=[]
            )

        # Force confirmed=True on the exact stored arguments
        args = dict(stored_args)
        args["confirmed"] = True

        tool_output = execute_tool(self.manager, tool_name, args)
        trace = ToolExecutionTrace(
            step=1,
            tool_name=tool_name,
            arguments=args,
            result=tool_output,
            error=tool_output.get("error")
        )
        if on_step:
            on_step(1, tool_name, args, tool_output)

        call_id = f"call_confirmed_{tool_name}"
        self.conversation_history.append({
            "role": "assistant",
            "content": f"Proceeding with confirmed {tool_name}.",
            "tool_calls": [
                {
                    "id": call_id,
                    "type": "function",
                    "function": {
                        "name": tool_name,
                        "arguments": json.dumps(args)
                    },
                    "name": tool_name,
                    "arguments": args
                }
            ]
        })
        self.conversation_history.append({
            "role": "tool",
            "tool_call_id": call_id,
            "name": tool_name,
            "content": json.dumps(tool_output, default=str)
        })

        if tool_output.get("error"):
            summary = f"Error executing {tool_name}: {tool_output['error']}"
            success = False
        else:
            summary = tool_output.get("message") or f"Successfully executed {tool_name}."
            success = not tool_output.get("requires_confirmation", False)

        self.conversation_history.append({
            "role": "assistant",
            "content": summary
        })

        return AgentResult(
            success=success,
            final_response=summary,
            steps_taken=1,
            tool_traces=[trace]
        )
