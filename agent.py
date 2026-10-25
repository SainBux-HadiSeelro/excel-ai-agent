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
        max_steps: int = 15,
        max_history_messages: int = 25
    ):
        self.manager = manager
        self.provider = provider or get_llm_provider()
        self.max_steps = max_steps
        self.max_history_messages = max_history_messages
        self.conversation_history: List[Dict[str, Any]] = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        self._current_run_user_idx: Optional[int] = None
        self.pending_confirmation: Optional[Dict[str, Any]] = None

    def _get_current_run_user_idx(self) -> int:
        """Returns the index of the user message that started the current run."""
        if self._current_run_user_idx is not None and 0 <= self._current_run_user_idx < len(self.conversation_history):
            if self.conversation_history[self._current_run_user_idx].get("role") == "user":
                return self._current_run_user_idx
        # Fallback: scan backwards for the latest user message
        for i in range(len(self.conversation_history) - 1, 0, -1):
            if self.conversation_history[i].get("role") == "user":
                return i
        return 1 if len(self.conversation_history) > 1 else 0

    def _format_tool_result_for_history(self, tool_output: Dict[str, Any], max_chars: int = 4000) -> str:
        """Serializes tool output for LLM history, truncating large results to conserve tokens."""
        raw_str = json.dumps(tool_output, default=str)
        if len(raw_str) <= max_chars:
            return raw_str

        # If tool returned a list of rows (e.g. read_range), truncate rows progressively
        if isinstance(tool_output, dict) and "rows" in tool_output and isinstance(tool_output["rows"], list):
            rows = tool_output["rows"]
            total_rows = len(rows)
            low, high = 1, total_rows
            best_output = tool_output
            while low <= high:
                mid = (low + high) // 2
                cand = dict(tool_output)
                cand["rows"] = rows[:mid]
                cand["truncation_note"] = f"...truncated, {total_rows - mid} rows omitted"
                cand_str = json.dumps(cand, default=str)
                if len(cand_str) <= max_chars:
                    best_output = cand
                    low = mid + 1
                else:
                    high = mid - 1
            serialized = json.dumps(best_output, default=str)
            if len(serialized) <= max_chars:
                return serialized

        # Fallback character cap
        omitted = len(raw_str) - max_chars
        return raw_str[:max_chars] + f"\n... [truncated, {omitted} characters omitted]"

    def _trim_conversation_history(self) -> None:
        """Caps conversation history length while preserving system prompt, turn atomicity,
        and ensuring the first non-system message is always a user message.
        Never drops the user message that started the current run.
        """
        if len(self.conversation_history) <= self.max_history_messages:
            return

        system_msg = self.conversation_history[0]
        user_idx = self._get_current_run_user_idx()

        current_run_messages = self.conversation_history[user_idx:]

        # If current run alone meets or exceeds the cap, keep system prompt + all current run messages
        if len(current_run_messages) + 1 >= self.max_history_messages:
            self.conversation_history = [system_msg] + current_run_messages
            self._current_run_user_idx = 1
            return

        # Otherwise, keep as many older messages as will fit under allowed_older
        older_messages = self.conversation_history[1:user_idx]
        allowed_older = self.max_history_messages - 1 - len(current_run_messages)

        # First kept older message must be a user message (so first non-system message is user)
        cutoff = max(0, len(older_messages) - allowed_older)
        kept_older_idx = None
        for idx in range(cutoff, len(older_messages)):
            if older_messages[idx].get("role") == "user":
                kept_older_idx = idx
                break

        if kept_older_idx is not None:
            kept_older = older_messages[kept_older_idx:]
        else:
            kept_older = []

        self.conversation_history = [system_msg] + kept_older + current_run_messages
        self._current_run_user_idx = 1 + len(kept_older)

    def run(
        self,
        user_instruction: str,
        on_step: Optional[Callable[[int, str, Dict[str, Any], Dict[str, Any]], None]] = None
    ) -> AgentResult:
        """
        Executes a user instruction through the tool-calling loop.
        Caps execution at max_steps (default 15).
        """
        # Append user message and record current run index
        self._current_run_user_idx = len(self.conversation_history)
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
                err_msg = str(e)
                if any(k in err_msg.lower() for k in ["not found", "model_not_found", "does not exist", "404", "unknown model", "retired"]):
                    display_msg = f"Model Error: The model '{getattr(self.provider, 'model', 'unknown')}' was not found or has been retired. Please change the Model Name field in the sidebar or environment settings. (Details: {err_msg})"
                else:
                    display_msg = f"LLM Error: {err_msg}"
                return AgentResult(
                    success=False,
                    final_response=display_msg,
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
                self._trim_conversation_history()
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

                # Append tool result to history (truncated for LLM if large)
                tool_history_content = self._format_tool_result_for_history(tool_output)
                self.conversation_history.append({
                    "role": "tool",
                    "tool_call_id": tc.id,
                    "name": tc.name,
                    "content": tool_history_content
                })
                self._trim_conversation_history()

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
                    self._trim_conversation_history()
                    return AgentResult(
                        success=True,
                        final_response=conf_final,
                        steps_taken=step,
                        tool_traces=tool_traces,
                        pending_confirmation=self.pending_confirmation
                    )

        return AgentResult(
            success=False,
            final_response=f"Agent reached maximum execution step limit ({self.max_steps}) without concluding.",
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
            "content": self._format_tool_result_for_history(tool_output)
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
        self._trim_conversation_history()

        return AgentResult(
            success=success,
            final_response=summary,
            steps_taken=1,
            tool_traces=[trace]
        )
