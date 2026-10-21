"""
Multi-provider LLM abstraction for Excel AI Agent.
Supports:
- OpenAI (gpt-4o, gpt-4o-mini, etc.)
- Anthropic (claude-3-5-sonnet, claude-3-5-haiku, etc.)
- Google Gemini (gemini-2.5-flash, gemini-1.5-flash, etc.)
- Mock / Test Provider (for offline testing)
"""

from __future__ import annotations

import os
import json
import uuid
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
import httpx
from dotenv import load_dotenv

load_dotenv()

# Single source of truth for the default Gemini model.
# Override with the GEMINI_MODEL environment variable if needed.
DEFAULT_GEMINI_MODEL = "gemini-2.5-flash"
FALLBACK_GEMINI_MODEL = "gemini-2.5-flash"


@dataclass
class ToolCall:
    """Standardized tool call object returned by LLM."""
    id: str
    name: str
    arguments: Dict[str, Any]


@dataclass
class LLMResponse:
    """Standardized response from any LLM provider."""
    content: Optional[str] = None
    tool_calls: List[ToolCall] = field(default_factory=list)
    raw: Any = None


class BaseLLMProvider(ABC):
    """Abstract base provider interface."""

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        self.api_key = api_key or ""
        self.model = model or ""

    @abstractmethod
    def chat(
        self,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        temperature: float = 0.1
    ) -> LLMResponse:
        """Sends conversation history and available tools to LLM, returns standardized response."""
        pass


def _format_messages_for_openai(messages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Formats conversation history into strictly valid OpenAI chat completion messages."""
    formatted = []
    for m in messages:
        role = m.get("role")
        if role == "assistant" and "tool_calls" in m and m["tool_calls"]:
            tcs = []
            for tc in m["tool_calls"]:
                fn_name = tc.get("name") or tc.get("function", {}).get("name", "")
                fn_args = tc.get("arguments")
                if fn_args is None:
                    fn_args = tc.get("function", {}).get("arguments", "{}")
                if isinstance(fn_args, dict):
                    fn_args = json.dumps(fn_args)
                tcs.append({
                    "id": tc.get("id", str(uuid.uuid4())),
                    "type": "function",
                    "function": {"name": fn_name, "arguments": fn_args}
                })
            item = {"role": "assistant", "content": m.get("content") or None, "tool_calls": tcs}
            formatted.append(item)
        elif role == "tool":
            formatted.append({
                "role": "tool",
                "tool_call_id": m.get("tool_call_id", str(uuid.uuid4())),
                "content": str(m.get("content", ""))
            })
        else:
            formatted.append({
                "role": role,
                "content": str(m.get("content", ""))
            })
    return formatted


class OpenAIProvider(BaseLLMProvider):
    """OpenAI API provider."""

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        key = api_key or os.getenv("OPENAI_API_KEY", "")
        m = model or os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        super().__init__(api_key=key, model=m)
        if not self.api_key:
            raise ValueError("OpenAI API key not found. Set OPENAI_API_KEY in .env or pass api_key.")

    def chat(
        self,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        temperature: float = 0.1
    ) -> LLMResponse:
        formatted_messages = _format_messages_for_openai(messages)
        try:
            import openai
            client = openai.OpenAI(api_key=self.api_key)
            formatted_tools = None
            if tools:
                formatted_tools = [{"type": "function", "function": t} for t in tools]

            response = client.chat.completions.create(
                model=self.model,
                messages=formatted_messages,
                tools=formatted_tools,
                temperature=temperature
            )
            msg = response.choices[0].message
            content = msg.content
            tool_calls = []
            if msg.tool_calls:
                for tc in msg.tool_calls:
                    try:
                        args = json.loads(tc.function.arguments) if isinstance(tc.function.arguments, str) else tc.function.arguments
                    except Exception:
                        args = {}
                    tool_calls.append(ToolCall(
                        id=tc.id or str(uuid.uuid4()),
                        name=tc.function.name,
                        arguments=args
                    ))
            return LLMResponse(content=content, tool_calls=tool_calls, raw=response)
        except ImportError:
            # Fallback to direct HTTP REST call
            return self._chat_http(formatted_messages, tools, temperature)

    def _chat_http(self, messages, tools, temperature):
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        payload = {
            "model": self.model,
            "messages": messages,
            "temperature": temperature
        }
        if tools:
            payload["tools"] = [{"type": "function", "function": t} for t in tools]

        with httpx.Client(timeout=60.0) as client:
            resp = client.post("https://api.openai.com/v1/chat/completions", headers=headers, json=payload)
            if resp.status_code != 200:
                raise RuntimeError(f"OpenAI API Error ({resp.status_code}): {resp.text}")
            data = resp.json()
            choice = data["choices"][0]["message"]
            content = choice.get("content")
            tool_calls = []
            if choice.get("tool_calls"):
                for tc in choice["tool_calls"]:
                    fn = tc.get("function", {})
                    args_str = fn.get("arguments", "{}")
                    try:
                        args = json.loads(args_str) if isinstance(args_str, str) else args_str
                    except Exception:
                        args = {}
                    tool_calls.append(ToolCall(
                        id=tc.get("id", str(uuid.uuid4())),
                        name=fn.get("name", ""),
                        arguments=args
                    ))
            return LLMResponse(content=content, tool_calls=tool_calls, raw=data)


class AnthropicProvider(BaseLLMProvider):
    """Anthropic Claude API provider."""

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        key = api_key or os.getenv("ANTHROPIC_API_KEY", "")
        m = model or os.getenv("ANTHROPIC_MODEL", "claude-3-5-haiku-20241022")
        super().__init__(api_key=key, model=m)
        if not self.api_key:
            raise ValueError("Anthropic API key not found. Set ANTHROPIC_API_KEY in .env or pass api_key.")

    def chat(
        self,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        temperature: float = 0.1
    ) -> LLMResponse:
        system_prompt = ""
        user_assistant_msgs = []

        # Extract system messages for Anthropic format
        for m in messages:
            role = m.get("role")
            if role == "system":
                system_prompt += str(m.get("content", "")) + "\n"
            elif role == "tool":
                # Anthropic tool result block
                user_assistant_msgs.append({
                    "role": "user",
                    "content": [{
                        "type": "tool_result",
                        "tool_use_id": m.get("tool_call_id", str(uuid.uuid4())),
                        "content": str(m.get("content", ""))
                    }]
                })
            elif role == "assistant" and "tool_calls" in m:
                content_blocks = []
                if m.get("content"):
                    content_blocks.append({"type": "text", "text": m["content"]})
                for tc in m["tool_calls"]:
                    content_blocks.append({
                        "type": "tool_use",
                        "id": tc["id"],
                        "name": tc["name"],
                        "input": tc["arguments"]
                    })
                user_assistant_msgs.append({"role": "assistant", "content": content_blocks})
            else:
                user_assistant_msgs.append({
                    "role": role,
                    "content": str(m.get("content", ""))
                })

        formatted_tools = None
        if tools:
            formatted_tools = []
            for t in tools:
                formatted_tools.append({
                    "name": t["name"],
                    "description": t.get("description", ""),
                    "input_schema": t.get("parameters", {"type": "object", "properties": {}})
                })

        try:
            import anthropic
            client = anthropic.Anthropic(api_key=self.api_key)
            kwargs = {
                "model": self.model,
                "max_tokens": 4096,
                "temperature": temperature,
                "messages": user_assistant_msgs
            }
            if system_prompt:
                kwargs["system"] = system_prompt.strip()
            if formatted_tools:
                kwargs["tools"] = formatted_tools

            response = client.messages.create(**kwargs)
            content_text = ""
            tool_calls = []
            for block in response.content:
                if block.type == "text":
                    content_text += block.text
                elif block.type == "tool_use":
                    tool_calls.append(ToolCall(
                        id=block.id,
                        name=block.name,
                        arguments=block.input
                    ))
            return LLMResponse(content=content_text or None, tool_calls=tool_calls, raw=response)
        except ImportError:
            return self._chat_http(system_prompt, user_assistant_msgs, formatted_tools, temperature)

    def _chat_http(self, system_prompt, messages, tools, temperature):
        headers = {
            "x-api-key": self.api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        }
        payload = {
            "model": self.model,
            "max_tokens": 4096,
            "temperature": temperature,
            "messages": messages
        }
        if system_prompt:
            payload["system"] = system_prompt.strip()
        if tools:
            payload["tools"] = tools

        with httpx.Client(timeout=60.0) as client:
            resp = client.post("https://api.anthropic.com/v1/messages", headers=headers, json=payload)
            resp.raise_for_status()
            data = resp.json()
            content_text = ""
            tool_calls = []
            for block in data.get("content", []):
                if block.get("type") == "text":
                    content_text += block.get("text", "")
                elif block.get("type") == "tool_use":
                    tool_calls.append(ToolCall(
                        id=block.get("id", str(uuid.uuid4())),
                        name=block.get("name", ""),
                        arguments=block.get("input", {})
                    ))
            return LLMResponse(content=content_text or None, tool_calls=tool_calls, raw=data)


def _sanitize_schema_for_gemini(schema: Any) -> Any:
    """Recursively removes keys like 'default' that are not permitted by Gemini's Schema validator.
    Ensures properties have explicit types and parameterless objects are omitted cleanly.
    """
    if not isinstance(schema, dict):
        return schema
    allowed_keys = {"type", "format", "description", "nullable", "enum", "properties", "required", "items"}
    clean = {}
    for k, v in schema.items():
        if k in allowed_keys:
            if k == "properties" and isinstance(v, dict):
                clean_props = {}
                for pk, pv in v.items():
                    prop = _sanitize_schema_for_gemini(pv)
                    if isinstance(prop, dict) and "type" not in prop:
                        prop["type"] = "string"
                    clean_props[pk] = prop
                clean[k] = clean_props
            elif k == "items" and isinstance(v, dict):
                clean_items = _sanitize_schema_for_gemini(v)
                if isinstance(clean_items, dict) and "type" not in clean_items:
                    clean_items["type"] = "string"
                clean[k] = clean_items
            else:
                clean[k] = v

    # Gemini requires OBJECT types to have non-empty properties
    if clean.get("type") == "object" and not clean.get("properties"):
        return None

    return clean


class GeminiProvider(BaseLLMProvider):
    """Google Gemini API provider (supports official google-genai and direct REST)."""

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        key = api_key or os.getenv("GEMINI_API_KEY", "") or os.getenv("GOOGLE_API_KEY", "")
        m = model or os.getenv("GEMINI_MODEL", DEFAULT_GEMINI_MODEL)
        super().__init__(api_key=key, model=m)
        if not self.api_key:
            raise ValueError("Gemini API key not found. Set GEMINI_API_KEY in .env or pass api_key.")

    def chat(
        self,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        temperature: float = 0.1
    ) -> LLMResponse:
        return self._chat_http(messages, tools, temperature)

    def _chat_http(self, messages, tools, temperature):
        system_instruction = None
        contents = []

        for m in messages:
            role = m.get("role")
            if role == "system":
                system_instruction = {"parts": [{"text": str(m.get("content", ""))}]}
            elif role == "user":
                contents.append({
                    "role": "user",
                    "parts": [{"text": str(m.get("content", ""))}]
                })
            elif role == "assistant":
                parts = []
                if m.get("content"):
                    parts.append({"text": str(m["content"])})
                if m.get("tool_calls"):
                    for tc in m["tool_calls"]:
                        fn_name = tc.get("name") or tc.get("function", {}).get("name", "")
                        fn_args = tc.get("arguments")
                        if fn_args is None:
                            fn_args = tc.get("function", {}).get("arguments", {})
                        if isinstance(fn_args, str):
                            try:
                                fn_args = json.loads(fn_args)
                            except Exception:
                                fn_args = {}
                        parts.append({
                            "functionCall": {
                                "name": fn_name,
                                "args": fn_args
                            }
                        })
                if parts:
                    contents.append({"role": "model", "parts": parts})
            elif role == "tool":
                # Function response part (Gemini expects role 'user' with functionResponse)
                tool_name = m.get("name", "tool_result")
                raw_content = m.get("content", "")
                try:
                    resp_dict = json.loads(raw_content) if isinstance(raw_content, str) else raw_content
                except Exception:
                    resp_dict = {"result": raw_content}

                func_part = {
                    "functionResponse": {
                        "name": tool_name,
                        "response": {"name": tool_name, "content": resp_dict}
                    }
                }
                # If previous item is user role and already has functionResponse, combine parts
                if contents and contents[-1].get("role") == "user" and any("functionResponse" in p for p in contents[-1].get("parts", [])):
                    contents[-1]["parts"].append(func_part)
                else:
                    contents.append({
                        "role": "user",
                        "parts": [func_part]
                    })

        payload: Dict[str, Any] = {
            "contents": contents,
            "generationConfig": {"temperature": temperature}
        }
        if system_instruction:
            payload["systemInstruction"] = system_instruction

        if tools:
            func_decls = []
            for t in tools:
                decl: Dict[str, Any] = {
                    "name": t["name"],
                    "description": t.get("description", "")
                }
                raw_params = t.get("parameters")
                if raw_params:
                    sanitized_params = _sanitize_schema_for_gemini(raw_params)
                    if sanitized_params and sanitized_params.get("properties"):
                        decl["parameters"] = sanitized_params
                func_decls.append(decl)
            payload["tools"] = [{"functionDeclarations": func_decls}]

        # Prepare models to try (requested model first, then fallback if needed)
        models_to_try = [self.model]
        if self.model != FALLBACK_GEMINI_MODEL:
            models_to_try.append(FALLBACK_GEMINI_MODEL)

        last_error = None
        for current_model in models_to_try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{current_model}:generateContent?key={self.api_key}"
            for attempt in range(2):
                try:
                    with httpx.Client(timeout=30.0) as client:
                        resp = client.post(url, json=payload)

                    if resp.status_code == 200:
                        data = resp.json()
                        candidates = data.get("candidates", [])
                        if not candidates:
                            return LLMResponse(content="", tool_calls=[], raw=data)

                        cand_content = candidates[0].get("content", {})
                        parts = cand_content.get("parts", [])

                        text_parts = []
                        tool_calls = []

                        for p in parts:
                            if "text" in p:
                                text_parts.append(p["text"])
                            elif "functionCall" in p:
                                fc = p["functionCall"]
                                tool_calls.append(ToolCall(
                                    id=str(uuid.uuid4()),
                                    name=fc.get("name", ""),
                                    arguments=fc.get("args", {})
                                ))

                        content_str = "\n".join(text_parts).strip() if text_parts else None
                        return LLMResponse(content=content_str, tool_calls=tool_calls, raw=data)

                    # Transient error (503 Service Unavailable / 429 Rate Limit)
                    if resp.status_code in (503, 429) and attempt < 1:
                        time.sleep(1.0)
                        continue

                    last_error = f"Gemini API Error ({resp.status_code}): {resp.text}"
                    # If 503 or 404, break attempt loop to allow fallback model
                    if resp.status_code in (503, 404):
                        break
                except httpx.RequestError as req_err:
                    last_error = f"Gemini Network Error: {str(req_err)}"
                    if attempt < 1:
                        time.sleep(1.0)
                        continue

        raise RuntimeError(last_error or "Gemini API request failed.")


class MockProvider(BaseLLMProvider):
    """
    Intelligent Mock / Demo Provider for offline testing and grading without API keys.
    Understands typical user instructions in English and Roman Urdu and simulates the tool-calling loop.
    """

    def __init__(self, responses: Optional[List[LLMResponse]] = None):
        super().__init__(api_key="mock-key", model="mock-model")
        self.preset_responses = responses or []
        self._index = 0

    def add_response(self, response: LLMResponse) -> None:
        self.preset_responses.append(response)

    def chat(
        self,
        messages: List[Dict[str, Any]],
        tools: Optional[List[Dict[str, Any]]] = None,
        temperature: float = 0.1
    ) -> LLMResponse:
        # If preset responses exist (e.g. from unit tests), use them
        if self._index < len(self.preset_responses):
            res = self.preset_responses[self._index]
            self._index += 1
            return res

        # Dynamically simulate based on conversation history for the current user turn
        last_user_msg = ""
        last_user_idx = -1
        for idx, m in enumerate(messages):
            if m.get("role") == "user":
                last_user_msg = str(m.get("content", "")).lower()
                last_user_idx = idx

        last_tool_res = None
        if last_user_idx != -1:
            for m in reversed(messages[last_user_idx + 1:]):
                if m.get("role") == "tool":
                    try:
                        last_tool_res = json.loads(m.get("content", "{}"))
                    except Exception:
                        last_tool_res = {}
                    break

        # Check if last user confirmed
        if "confirm" in last_user_msg or "yes" in last_user_msg or "theek hai" in last_user_msg:
            if not last_tool_res:
                if "duplicate" in last_user_msg:
                    return LLMResponse(
                        content="Executing confirmed duplicate removal.",
                        tool_calls=[ToolCall(id=str(uuid.uuid4()), name="remove_duplicates", arguments={"sheet_name": "Sales", "subset_columns": ["Order_ID"], "confirmed": True})]
                    )
                elif "delete" in last_user_msg and "sheet" in last_user_msg:
                    return LLMResponse(
                        content="Executing confirmed sheet deletion.",
                        tool_calls=[ToolCall(id=str(uuid.uuid4()), name="delete_sheet", arguments={"title": "Targets", "confirmed": True})]
                    )
                else:
                    return LLMResponse(
                        content="Operation confirmed, execution in progress.",
                        tool_calls=[]
                    )
            else:
                return LLMResponse(
                    content="Operation confirmed aur kamyabi se execute ho gayi hai.",
                    tool_calls=[]
                )

        # 1. Total sales / calculations
        if any(w in last_user_msg for w in ["total", "sum", "sales nikalo", "calculate", "jama"]):
            if not last_tool_res:
                return LLMResponse(
                    content="Pehele sheet ka structure aur columns inspect kar lete hain.",
                    tool_calls=[ToolCall(id=str(uuid.uuid4()), name="get_sheet_summary", arguments={"sheet_name": "Sales"})]
                )
            elif "columns" in last_tool_res:
                return LLMResponse(
                    content="Ab hum Total Revenue column (H) ka formula `=SUM(H2:H23)` row 25 par likh rahe hain.",
                    tool_calls=[ToolCall(
                        id=str(uuid.uuid4()),
                        name="write_cells",
                        arguments={
                            "sheet_name": "Sales",
                            "updates": [
                                {"cell": "G25", "value": "Total Revenue"},
                                {"cell": "H25", "value": "=SUM(H2:H23)"}
                            ]
                        }
                    )]
                )
            else:
                is_urdu = any(w in last_user_msg for w in ["nikalo", "karo", "batao", "hai"])
                msg = (
                    "Total sales ka formula kamyabi se `=SUM(H2:H23)` cell H25 mein daal diya gaya hai aur verify ho chuka hai."
                    if is_urdu
                    else "Total sales formula `=SUM(H2:H23)` has been successfully written to cell H25 and verified."
                )
                return LLMResponse(content=msg, tool_calls=[])

        # 2. Duplicates
        elif any(w in last_user_msg for w in ["duplicate", "duplicates", "hata do", "remove duplicates", "hatao"]):
            if not last_tool_res:
                return LLMResponse(
                    content="Duplicates check karne ke liye sheet inspect karte hain.",
                    tool_calls=[ToolCall(id=str(uuid.uuid4()), name="remove_duplicates", arguments={"sheet_name": "Sales", "subset_columns": ["Order_ID"], "confirmed": False})]
                )
            elif last_tool_res.get("requires_confirmation"):
                dup_count = last_tool_res.get("duplicate_count", 0)
                is_urdu = any(w in last_user_msg for w in ["hata do", "hatao", "karo"])
                msg = (
                    f"Workbook mein {dup_count} duplicate rows payi gayi hain (Order_ID ke hisaab se). Kya aap inhein hatana chahte hain? Confirm karein."
                    if is_urdu
                    else f"Found {dup_count} duplicate row(s) based on Order_ID. Do you want to proceed with removal? Please confirm."
                )
                return LLMResponse(content=msg, tool_calls=[])
            else:
                is_urdu = any(w in last_user_msg for w in ["hata do", "hatao", "karo"])
                msg = (
                    "Duplicates kamyabi se hata di gayi hain aur sheet update ho chuki hai."
                    if is_urdu
                    else "Duplicate rows have been successfully removed."
                )
                return LLMResponse(content=msg, tool_calls=[])

        # 3. Chart
        elif any(w in last_user_msg for w in ["chart", "graph", "plot", "banao"]):
            if not last_tool_res:
                return LLMResponse(
                    content="Region ke hisaab se chart bananey ke liye Targets sheet ka data use kar rahe hain.",
                    tool_calls=[ToolCall(
                        id=str(uuid.uuid4()),
                        name="add_chart",
                        arguments={
                            "sheet_name": "Targets",
                            "chart_type": "column",
                            "data_range": "B1:B5",
                            "categories_range": "A2:A5",
                            "title": "Region Wise Target Sales",
                            "target_cell": "E2"
                        }
                    )]
                )
            else:
                is_urdu = any(w in last_user_msg for w in ["banao", "karo", "hai"])
                msg = (
                    "Targets sheet mein region ke hisaab se Column Chart cell E2 par bana diya gaya hai."
                    if is_urdu
                    else "A column chart grouped by Region has been successfully added to sheet 'Targets' at cell E2."
                )
                return LLMResponse(content=msg, tool_calls=[])

        # 4. Merge files
        elif any(w in last_user_msg for w in ["merge", "combine", "mila do", "do files"]):
            if not last_tool_res:
                return LLMResponse(
                    content="External workbook `sample_data/regional_extra.xlsx` ko merge kar rahe hain.",
                    tool_calls=[ToolCall(
                        id=str(uuid.uuid4()),
                        name="merge_workbooks",
                        arguments={
                            "source_files": ["sample_data/regional_extra.xlsx"],
                            "mode": "add_sheets"
                        }
                    )]
                )
            else:
                is_urdu = any(w in last_user_msg for w in ["mila do", "karo", "hai"])
                msg = (
                    "Do files kamyabi se merge ho chuki hain aur nayi sheet `Q1_Summary` add kar di gayi hai."
                    if is_urdu
                    else "Files merged successfully. New sheet 'Q1_Summary' was imported into the workbook."
                )
                return LLMResponse(content=msg, tool_calls=[])

        # 5. Clean data / casing / formatting
        elif any(w in last_user_msg for w in ["clean", "saaf", "spaces", "trim", "casing"]):
            if not last_tool_res:
                return LLMResponse(
                    content="Sales sheet mein text casing aur spaces clean kar rahe hain.",
                    tool_calls=[ToolCall(
                        id=str(uuid.uuid4()),
                        name="clean_data",
                        arguments={
                            "sheet_name": "Sales",
                            "columns": ["D"],
                            "trim_spaces": True,
                            "fix_casing": "title"
                        }
                    )]
                )
            else:
                is_urdu = any(w in last_user_msg for w in ["karo", "saaf", "hai"])
                msg = (
                    "Sales Rep column D ka data saaf kar diya gaya hai (spaces trim aur title casing theek ho gayi)."
                    if is_urdu
                    else "Data cleaned successfully: spaces trimmed and casing standardized in Sales Rep column."
                )
                return LLMResponse(content=msg, tool_calls=[])

        # Default fallback
        if not last_tool_res:
            return LLMResponse(
                content="Workbook inspect ki ja rahi hai.",
                tool_calls=[ToolCall(id=str(uuid.uuid4()), name="list_sheets", arguments={})]
            )
        else:
            return LLMResponse(content="Task mukammal ho gaya hai. Sheets check kar sakte hain.", tool_calls=[])


def get_llm_provider(
    provider_name: Optional[str] = None,
    api_key: Optional[str] = None,
    model: Optional[str] = None
) -> BaseLLMProvider:
    """
    Factory function returning the configured LLM provider.
    Checks explicit provider_name -> environment variables -> fallback.
    """
    load_dotenv()
    target_provider = (provider_name or os.getenv("LLM_PROVIDER", "")).lower().strip()

    if not target_provider:
        # Auto-detect from available keys
        if api_key:
            target_provider = "gemini"
        elif os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY"):
            target_provider = "gemini"
        elif os.getenv("OPENAI_API_KEY"):
            target_provider = "openai"
        elif os.getenv("ANTHROPIC_API_KEY"):
            target_provider = "anthropic"
        else:
            target_provider = "gemini"

    if target_provider == "gemini":
        return GeminiProvider(api_key=api_key, model=model)
    elif target_provider == "openai":
        return OpenAIProvider(api_key=api_key, model=model)
    elif target_provider in ("anthropic", "claude"):
        return AnthropicProvider(api_key=api_key, model=model)
    elif target_provider == "mock":
        return MockProvider()
    else:
        raise ValueError(f"Unsupported LLM provider: '{target_provider}'. Choose gemini, openai, or anthropic.")
