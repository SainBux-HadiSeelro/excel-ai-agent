"""
Streamlit Web UI for Excel AI Agent.
Provides:
- Sidebar: File upload, Provider & Model selector, API key input, Session status
- Main Area: Interactive Chat, Quick action pills, Live tool execution traces, Sheet Preview table, Download & Undo buttons
"""

from __future__ import annotations

import os
import sys
import html
import time
import uuid
from pathlib import Path
from typing import List, Optional, Tuple
import streamlit as st
import pandas as pd
from dotenv import load_dotenv

# Ensure local imports work
sys.path.insert(0, str(Path(__file__).parent.resolve()))

from workspace_manager import WorkspaceManager, cleanup_stale_sessions, load_sheet_dataframe, safe_upload_name
from llm import get_llm_provider, GeminiProvider, OpenAIProvider, AnthropicProvider, DEFAULT_GEMINI_MODEL, is_model_error
from agent import ExcelAgent, ToolExecutionTrace, SYSTEM_PROMPT

load_dotenv()


def load_sample_workbook() -> None:
    """Loads the sample_data/sales.xlsx dataset into the active session workspace."""
    sample_path = Path("sample_data/sales.xlsx")
    if not sample_path.exists():
        import generate_sample_data
    st.session_state.manager = WorkspaceManager(session_id=st.session_state.session_id)
    st.session_state.manager.init_from_file(sample_path)
    st.session_state.current_file_name = "sales.xlsx"
    st.session_state.current_file_id = "sample_sales"
    st.session_state.file_loaded = True
    st.session_state.pending_confirmation = None
    if st.session_state.agent:
        st.session_state.agent.manager = st.session_state.manager


def init_new_blank_workbook() -> None:
    """Initializes a new empty workbook in the active session workspace."""
    st.session_state.manager = WorkspaceManager(session_id=st.session_state.session_id)
    st.session_state.manager.init_blank("workbook.xlsx")
    st.session_state.current_file_name = "workbook.xlsx"
    st.session_state.current_file_id = f"blank_{int(time.time() * 1000)}"
    st.session_state.file_loaded = True
    st.session_state.pending_confirmation = None
    if st.session_state.agent:
        st.session_state.agent.manager = st.session_state.manager


def render_file_status_bar_html(filename: str, active_sheets: List[str]) -> str:
    """Renders HTML for active file status bar with all inputs safely HTML-escaped."""
    safe_name = html.escape(str(filename or ""))
    safe_sheets = [html.escape(str(s)) for s in active_sheets]
    sheets_str = ", ".join(safe_sheets)
    count = len(active_sheets)
    return (
        f'<div class="file-status-bar">'
        f'<span>🟢</span>'
        f'<span>Active Workbook: <b>{safe_name}</b> ({count} sheet(s): {sheets_str}) • Isolated Safe Workspace</span>'
        f'</div>'
    )


def render_destructive_confirmation_html(warning_text: str) -> str:
    """Renders HTML for destructive confirmation warning card with content safely HTML-escaped."""
    safe_warning = html.escape(str(warning_text or ""))
    return (
        f'<div class="destructive-box">'
        f'<h4 style="margin:0 0 8px 0; color:#991B1B;">⚠️ Confirmation Required (Destructive Safety Guard)</h4>'
        f'<p style="margin:0; color:#7F1D1D; font-size:0.95rem;">{safe_warning}</p>'
        f'</div>'
    )


st.set_page_config(
    page_title="Excel AI Agent",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="auto"
)

# Custom CSS for ultra-modern 2026 Agentic AI interface
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    code, pre, .trace-card {
        font-family: 'JetBrains Mono', monospace !important;
    }

    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3.5rem;
        max-width: 1240px;
    }
    
    /* Modern Glassmorphic Hero Banner */
    .hero-banner {
        background: linear-gradient(135deg, #090D16 0%, #0F172A 45%, #1E1B4B 100%);
        color: white;
        padding: 2rem 2.4rem;
        border-radius: 20px;
        margin-bottom: 1.4rem;
        border: 1px solid rgba(255, 255, 255, 0.1);
        box-shadow: 0 20px 40px -15px rgba(15, 23, 42, 0.4), 0 0 0 1px rgba(99, 102, 241, 0.15);
        position: relative;
        overflow: hidden;
    }
    .hero-banner::after {
        content: "";
        position: absolute;
        top: -50%;
        right: -10%;
        width: 320px;
        height: 320px;
        background: radial-gradient(circle, rgba(99, 102, 241, 0.25) 0%, rgba(56, 189, 248, 0) 70%);
        pointer-events: none;
    }
    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        letter-spacing: -0.02em;
        margin: 0;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-subtitle {
        font-size: 1.02rem;
        color: #CBD5E1;
        margin-top: 8px;
        margin-bottom: 16px;
        line-height: 1.5;
        max-width: 900px;
    }
    .badge-pill {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 5px 13px;
        background: rgba(255, 255, 255, 0.08);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        color: #F1F5F9;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-right: 8px;
        margin-bottom: 6px;
        border: 1px solid rgba(255, 255, 255, 0.18);
        transition: all 0.2s ease;
    }
    .badge-pill:hover {
        background: rgba(255, 255, 255, 0.15);
        border-color: rgba(99, 102, 241, 0.4);
    }
    
    /* Modern Bento Welcome Cards */
    .welcome-container {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 20px;
        padding: 2.2rem;
        box-shadow: 0 10px 30px -10px rgba(15, 23, 42, 0.06);
        margin-bottom: 1.5rem;
    }
    .bento-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
        gap: 16px;
        margin-top: 1.4rem;
        margin-bottom: 1.4rem;
    }
    .bento-card {
        background: linear-gradient(180deg, #F8FAFC 0%, #F1F5F9 100%);
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 1.25rem;
        transition: transform 0.2s, box-shadow 0.2s, border-color 0.2s;
    }
    .bento-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.08);
        border-color: #CBD5E1;
    }
    .bento-icon {
        font-size: 1.8rem;
        margin-bottom: 10px;
        display: inline-block;
    }
    .bento-title {
        font-weight: 700;
        font-size: 0.98rem;
        color: #0F172A;
        margin-bottom: 6px;
    }
    .bento-desc {
        font-size: 0.84rem;
        color: #64748B;
        line-height: 1.45;
    }

    /* Active File Pill */
    .file-status-bar {
        background: linear-gradient(90deg, #F0FDF4 0%, #ECFDF5 100%);
        border: 1px solid #86EFAC;
        color: #14532D;
        padding: 11px 18px;
        border-radius: 12px;
        font-size: 0.92rem;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 1.2rem;
        box-shadow: 0 2px 8px rgba(34, 197, 94, 0.06);
    }

    /* Tool traces styling */
    .trace-card {
        background: #0F172A;
        color: #E2E8F0;
        border-left: 3px solid #6366F1;
        border-radius: 8px;
        padding: 10px 14px;
        margin: 6px 0;
        font-size: 0.82rem;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    }
    
    /* Destructive Confirmation Box */
    .destructive-box {
        background: #FFF1F2;
        border: 1.5px solid #FDA4AF;
        border-radius: 14px;
        padding: 1.3rem;
        margin: 1.2rem 0;
        box-shadow: 0 4px 16px rgba(225, 29, 72, 0.08);
    }

    /* Floating Corner Chatbot Widget */
    .st-key-corner_chatbot {
        position: fixed !important;
        bottom: 24px !important;
        right: 24px !important;
        width: 450px !important;
        max-width: calc(100vw - 32px) !important;
        height: 650px !important;
        max-height: calc(100vh - 44px) !important;
        background: #FFFFFF !important;
        border-radius: 24px !important;
        box-shadow: 0 25px 65px -12px rgba(15, 23, 42, 0.35), 0 0 0 1px rgba(226, 232, 240, 0.9) !important;
        z-index: 999999 !important;
        display: flex !important;
        flex-direction: column !important;
        padding: 0 !important;
        overflow: hidden !important;
        animation: cornerChatFadeIn 0.25s cubic-bezier(0.16, 1, 0.3, 1) !important;
    }

    @keyframes cornerChatFadeIn {
        from {
            opacity: 0;
            transform: translateY(20px) scale(0.96);
        }
        to {
            opacity: 1;
            transform: translateY(0) scale(1);
        }
    }

    .st-key-corner_chatbot [data-testid="stVerticalBlock"] {
        gap: 0.3rem !important;
    }

    /* Corner Chatbot Header */
    .corner-header-container {
        background: linear-gradient(135deg, #090D16 0%, #0F172A 60%, #1E1B4B 100%);
        padding: 14px 18px;
        color: #FFFFFF;
        border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.12);
    }

    .chat-avatar-badge {
        width: 38px;
        height: 38px;
        border-radius: 12px;
        background: linear-gradient(135deg, #3B82F6 0%, #8B5CF6 100%);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.25rem;
        box-shadow: 0 4px 16px rgba(59, 130, 246, 0.45);
    }

    .pulsing-green-dot {
        width: 8px;
        height: 8px;
        background-color: #22C55E;
        border-radius: 50%;
        display: inline-block;
        box-shadow: 0 0 8px #22C55E;
        animation: pulseDot 2s infinite;
    }

    @keyframes pulseDot {
        0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0.7); }
        70% { transform: scale(1.15); box-shadow: 0 0 0 6px rgba(34, 197, 94, 0); }
        100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(34, 197, 94, 0); }
    }

    /* Welcome Card Inside Corner Chat */
    .corner-welcome-card {
        background: linear-gradient(135deg, #F8FAFC 0%, #EFF6FF 100%);
        border: 1px solid #DBEAFE;
        border-radius: 16px;
        padding: 14px 16px;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(37, 99, 235, 0.04);
    }

    .corner-welcome-pill-list {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-top: 8px;
    }

    .corner-task-badge {
        display: inline-block;
        font-size: 0.74rem;
        font-weight: 600;
        color: #1E40AF;
        background: #DBEAFE;
        padding: 4px 10px;
        border-radius: 9999px;
        border: 1px solid #BFDBFE;
    }

    /* Chat Messages styling */
    div[data-testid="stChatMessage"] {
        border-radius: 14px !important;
        padding: 9px 13px !important;
        margin-bottom: 6px !important;
        border: 1px solid #F1F5F9 !important;
    }

    /* Quick Action Pills in Corner Chat */
    .st-key-corner_chatbot .stButton > button {
        border-radius: 9999px !important;
        min-height: 32px !important;
        height: 32px !important;
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        padding: 2px 10px !important;
        background: #F8FAFC !important;
        color: #334155 !important;
        border: 1px solid #E2E8F0 !important;
        transition: all 0.18s ease !important;
    }

    .st-key-corner_chatbot .stButton > button:hover {
        background: #EFF6FF !important;
        color: #1D4ED8 !important;
        border-color: #93C5FD !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.12) !important;
    }

    /* Floating Corner Trigger Pill (When Minimized) */
    .st-key-corner_chat_trigger {
        position: fixed !important;
        bottom: 24px !important;
        right: 24px !important;
        z-index: 999999 !important;
    }

    .st-key-corner_chat_trigger button {
        border-radius: 9999px !important;
        padding: 13px 26px !important;
        background: linear-gradient(135deg, #090D16 0%, #0F172A 40%, #2563EB 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 0.98rem !important;
        letter-spacing: 0.2px !important;
        box-shadow: 0 14px 34px -4px rgba(37, 99, 235, 0.5), 0 4px 14px rgba(15, 23, 42, 0.25) !important;
        border: 1.5px solid rgba(255, 255, 255, 0.4) !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    .st-key-corner_chat_trigger button:hover {
        transform: translateY(-3px) scale(1.03) !important;
        box-shadow: 0 18px 45px -4px rgba(37, 99, 235, 0.65), 0 6px 18px rgba(15, 23, 42, 0.35) !important;
        border-color: rgba(255, 255, 255, 0.7) !important;
    }

    /* Modern Metric Cards */
    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 12px 16px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.02);
    }
    div[data-testid="stMetricValue"] {
        font-weight: 700 !important;
        color: #0F172A !important;
    }

    @media (max-width: 768px) {
        .st-key-corner_chatbot {
            bottom: 6px !important;
            right: 6px !important;
            left: 6px !important;
            width: auto !important;
            height: 82vh !important;
            border-radius: 18px !important;
        }
        .hero-banner {
            padding: 1.3rem 1.1rem !important;
            border-radius: 14px !important;
        }
        .hero-title {
            font-size: 1.4rem !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if "session_id" not in st.session_state:
    st.session_state.session_id = f"session_{int(time.time() * 1000)}_{uuid.uuid4().hex[:8]}"
if "cleaned_stale_sessions" not in st.session_state:
    cleanup_stale_sessions(active_session_ids=[st.session_state.session_id])
    st.session_state.cleaned_stale_sessions = True
if "manager" not in st.session_state:
    st.session_state.manager = WorkspaceManager(session_id=st.session_state.session_id)
if "agent" not in st.session_state:
    st.session_state.agent = None
if "messages" not in st.session_state:
    st.session_state.messages = []
if "traces" not in st.session_state:
    st.session_state.traces = []
if "file_loaded" not in st.session_state:
    st.session_state.file_loaded = False
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = ""
if "pending_confirmation" not in st.session_state:
    st.session_state.pending_confirmation = None
if "chat_open" not in st.session_state:
    st.session_state.chat_open = False

# Sidebar Configuration
with st.sidebar:
    st.markdown("### 📊 **Excel AI Control Panel**")
    st.caption("Autonomous Agentic Copilot • English & Roman Urdu")

    st.markdown("#### 1. 🤖 Intelligence Engine")
    has_any_key = any(
        os.getenv(k) for k in ["GEMINI_API_KEY", "GOOGLE_API_KEY", "OPENAI_API_KEY", "ANTHROPIC_API_KEY"]
    )
    env_provider = os.getenv("LLM_PROVIDER", "").lower()
    provider_map = {"gemini": 0, "openai": 1, "anthropic": 2, "claude": 2, "mock": 3}
    if env_provider in provider_map:
        default_provider_idx = provider_map[env_provider]
    else:
        default_provider_idx = 0 if has_any_key else 3

    provider_choice = st.selectbox(
        "Provider",
        options=["Gemini", "OpenAI", "Anthropic", "Offline Demo (Mock)"],
        index=default_provider_idx
    )

    default_models = {
        "Gemini": DEFAULT_GEMINI_MODEL,
        "OpenAI": "gpt-6-luna",
        "Anthropic": "claude-haiku-5-5",
        "Offline Demo (Mock)": "mock-model"
    }

    model_choice = st.text_input(
        "Model Name",
        value=os.getenv(f"{provider_choice.upper().split()[0]}_MODEL", default_models.get(provider_choice, "default"))
    )

    env_key = os.getenv(f"{provider_choice.upper().split()[0]}_API_KEY", "")
    if provider_choice == "Gemini" and not env_key:
        env_key = os.getenv("GOOGLE_API_KEY", "")
    api_key_input = st.text_input(
        f"{provider_choice} API Key",
        value=env_key if "Offline" not in provider_choice else "mock-key",
        type="password",
        help="Reads automatically from .env if set. For Offline Demo, leave as default."
    )

    if "Offline" not in provider_choice and not api_key_input:
        st.info("💡 **No API key?** Switch Provider to **'Offline Demo (Mock)'** above to test instantly without an API key.")

    st.markdown("---")
    st.markdown("#### 2. 📂 Workbook Sandbox")
    uploaded_file = st.file_uploader(
        "Upload Excel (.xlsx) or CSV",
        type=["xlsx", "csv"]
    )

    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        use_sample = st.button("📁 Load Sample", use_container_width=True)
    with col_sb2:
        create_blank = st.button("📄 Blank Sheet", use_container_width=True)

    if uploaded_file is not None:
        safe_name = safe_upload_name(uploaded_file.name)
        file_id = f"{safe_name}_{uploaded_file.size}"
        if st.session_state.get("current_file_id") != file_id:
            session_upload_dir = Path("temp_uploads") / st.session_state.session_id
            session_upload_dir.mkdir(parents=True, exist_ok=True)
            temp_path = session_upload_dir / safe_name
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            st.session_state.manager = WorkspaceManager(session_id=st.session_state.session_id)
            st.session_state.manager.init_from_file(temp_path)
            st.session_state.current_file_name = safe_name
            st.session_state.current_file_id = file_id
            st.session_state.file_loaded = True
            st.session_state.messages.append({
                "role": "assistant",
                "content": f"📂 Uploaded workbook `{safe_name}` load kar liya gaya hai! Ab aap is par koi bhi task de sakte hain.",
                "traces": []
            })
            st.session_state.traces = []
            st.session_state.pending_confirmation = None
            if st.session_state.agent:
                st.session_state.agent.manager = st.session_state.manager
            st.toast(f"✅ Loaded: {safe_name}", icon="📁")
            st.rerun()

    elif use_sample:
        load_sample_workbook()
        st.session_state.messages.append({
            "role": "assistant",
            "content": "📁 Sample workbook `sales.xlsx` load kar liya gaya hai! Ab aap koi bhi task de sakte hain.",
            "traces": []
        })
        st.toast("✅ Loaded sample_data/sales.xlsx", icon="📁")
        st.rerun()

    elif create_blank:
        init_new_blank_workbook()
        st.session_state.messages.append({
            "role": "assistant",
            "content": "📄 Nayi blank workbook `workbook.xlsx` create kardi gayi hai! Ab aap columns aur data add karne ka task de sakte hain.",
            "traces": []
        })
        st.toast("✅ Created workbook.xlsx", icon="📄")
        st.rerun()

    # Load agent instance
    current_config = f"{provider_choice}_{model_choice}_{api_key_input}"
    if st.session_state.get("agent_config") != current_config or st.session_state.agent is None:
        if api_key_input or "Offline" in provider_choice:
            try:
                if "Offline" in provider_choice:
                    from llm import MockProvider
                    llm = MockProvider()
                elif provider_choice == "Gemini":
                    llm = GeminiProvider(api_key=api_key_input, model=model_choice)
                elif provider_choice == "OpenAI":
                    llm = OpenAIProvider(api_key=api_key_input, model=model_choice)
                else:
                    llm = AnthropicProvider(api_key=api_key_input, model=model_choice)
                st.session_state.agent = ExcelAgent(manager=st.session_state.manager, provider=llm)
                st.session_state.agent_config = current_config
            except Exception as e:
                err_str = str(e)
                if is_model_error(err_str):
                    st.sidebar.error(f"Provider Error: {err_str}\n\n💡 **Tip:** Model not found or retired. Please update the **Model Name** field above.")
                else:
                    st.sidebar.error(f"Provider Error: {err_str}")

    st.markdown("---")
    # Status & Undo Controls
    history_count = st.session_state.manager.get_history_count()
    st.markdown(f"**🔄 Reversible Snapshots:** `{history_count}` state(s)")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        if st.button("↩️ Undo", use_container_width=True, disabled=(history_count == 0)):
            undo_res = st.session_state.manager.undo()
            if undo_res["success"]:
                st.toast(undo_res["message"], icon="↩️")
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": f"↩️ **Undo:** {undo_res['message']}",
                    "traces": []
                })
                st.rerun()
            else:
                st.warning(undo_res["message"])

    with col_u2:
        if st.session_state.file_loaded:
            curr_path = st.session_state.manager.get_current_file_path()
            with open(curr_path, "rb") as f:
                st.download_button(
                    label="💾 Export XLSX",
                    data=f.read(),
                    file_name=f"updated_{st.session_state.current_file_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    st.markdown("---")
    chat_view_mode = st.radio(
        "Chatbot Interface Mode",
        options=["Corner Floating Chatbot (Recommended)", "Full Page Tabs"],
        index=0,
        help="Dock the copilot in the bottom-right corner or view full-page tabs."
    )

    st.markdown("---")
    if st.button("🗑️ Reset Chat History", use_container_width=True):
        st.session_state.messages = []
        st.session_state.traces = []
        st.session_state.pending_confirmation = None
        if st.session_state.agent:
            st.session_state.agent.conversation_history = [
                {"role": "system", "content": SYSTEM_PROMPT}
            ]
        st.toast("Chat history cleared!", icon="🗑️")
        st.rerun()


def run_agent_task(prompt_to_run: str) -> None:
    """Executes a user task through the ExcelAgent, updating workbook and chat state."""
    # 1. Auto-initialize workbook if none loaded yet
    if not st.session_state.file_loaded:
        prompt_lower = prompt_to_run.lower()
        if any(k in prompt_lower for k in ["sample", "demo", "sales"]):
            load_sample_workbook()
        else:
            init_new_blank_workbook()

    # 2. Ensure agent instance is ready
    if st.session_state.agent is None:
        if "Offline" in provider_choice or not api_key_input:
            from llm import MockProvider
            llm = MockProvider()
        elif provider_choice == "Gemini":
            llm = GeminiProvider(api_key=api_key_input, model=model_choice)
        elif provider_choice == "OpenAI":
            llm = OpenAIProvider(api_key=api_key_input, model=model_choice)
        else:
            llm = AnthropicProvider(api_key=api_key_input, model=model_choice)
        st.session_state.agent = ExcelAgent(manager=st.session_state.manager, provider=llm)

    # 3. Add user prompt to chat
    st.session_state.messages.append({"role": "user", "content": prompt_to_run})

    live_traces = []
    def streamlit_on_step(step, tool_name, args, result):
        live_traces.append(ToolExecutionTrace(
            step=step,
            tool_name=tool_name,
            arguments=args,
            result=result
        ))

    with st.spinner("Analyzing spreadsheet & executing task..."):
        agent_res = st.session_state.agent.run(prompt_to_run, on_step=streamlit_on_step)

    # 4. Handle confirmation state
    if agent_res.pending_confirmation:
        st.session_state.pending_confirmation = agent_res.pending_confirmation

    # 5. Save to history
    st.session_state.messages.append({
        "role": "assistant",
        "content": agent_res.final_response,
        "traces": agent_res.tool_traces
    })
    st.rerun()


@st.cache_data(max_entries=50, ttl=3600)
def load_preview_cached(file_path_str: str, mtime_ns: int, size: int, sheet_name: Optional[str] = None, max_rows: int = 500) -> Tuple[pd.DataFrame, str, bool]:
    """Caches preview dataframe load keyed by (file_path, mtime_ns, size)."""
    return load_sheet_dataframe(file_path_str, sheet_name=sheet_name, max_rows=max_rows)


# Top Hero Banner with Glassmorphic Elements
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">
        <span>📊 Excel AI Copilot</span>
    </div>
    <div class="hero-subtitle">
        Autonomous multilingual spreadsheet intelligence engine. Formulates native Excel formulas, purges duplicates, normalizes messy datasets, and embeds charts via conversational Roman Urdu & English!
    </div>
    <div>
        <span class="badge-pill">⚡ Autonomous Engine Live</span>
        <span class="badge-pill">🗣️ Roman Urdu & English</span>
        <span class="badge-pill">🛡️ Zero-Risk Sandbox</span>
        <span class="badge-pill">🧮 Pure-Python Formula AST</span>
        <span class="badge-pill">↩️ Reversible Undo Stack</span>
    </div>
</div>
""", unsafe_allow_html=True)

# Render View Based on User Preference
if chat_view_mode == "Corner Floating Chatbot (Recommended)":
    # MAIN CANVAS: Full Spreadsheet Workspace View
    if st.session_state.file_loaded:
        try:
            active_sheets = st.session_state.manager.get_sheet_names()
        except Exception:
            active_sheets = []
        st.markdown(
            render_file_status_bar_html(st.session_state.current_file_name, active_sheets),
            unsafe_allow_html=True
        )

        col_tb1, col_tb2, col_tb3, col_tb4 = st.columns([3, 1, 1, 2])
        with col_tb1:
            selected_sheet = st.selectbox("Select Active Worksheet", options=active_sheets, key="main_sheet_sel")
        with col_tb2:
            if st.button("↩️ Undo", key="canvas_undo", use_container_width=True, disabled=(st.session_state.manager.get_history_count() == 0)):
                undo_res = st.session_state.manager.undo()
                if undo_res["success"]:
                    st.toast(undo_res["message"], icon="↩️")
                    st.session_state.messages.append({"role": "assistant", "content": f"↩️ **Undo:** {undo_res['message']}", "traces": []})
                    st.rerun()
                else:
                    st.warning(undo_res["message"])
        with col_tb3:
            curr_path = st.session_state.manager.get_current_file_path()
            with open(curr_path, "rb") as f:
                st.download_button(
                    label="💾 Export XLSX",
                    data=f.read(),
                    file_name=f"updated_{st.session_state.current_file_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="canvas_download"
                )
        with col_tb4:
            toggle_label = "💬 Open AI Copilot" if not st.session_state.chat_open else "💬 Minimize Copilot"
            if st.button(toggle_label, key="canvas_toggle_chat", use_container_width=True):
                st.session_state.chat_open = not st.session_state.chat_open
                st.rerun()

        try:
            curr_path = st.session_state.manager.get_current_file_path()
            curr_stat = Path(curr_path).stat()
            df, active_s, is_oversize = load_preview_cached(
                str(Path(curr_path).resolve()),
                curr_stat.st_mtime_ns,
                curr_stat.st_size,
                sheet_name=selected_sheet,
                max_rows=500
            )

            if is_oversize:
                st.info("ℹ️ **Large workbook (>20,000 cells):** Pure-Python formula evaluation is skipped to preserve responsiveness. Showing raw formulas and basic operations.")

            df = df.convert_dtypes()
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype("string")

            # Modern Metrics row
            mcol1, mcol2, mcol3, mcol4 = st.columns(4)
            with mcol1:
                st.metric("Total Rows", f"{len(df):,}")
            with mcol2:
                st.metric("Total Columns", f"{len(df.columns)}")
            with mcol3:
                st.metric("Current Sheet", selected_sheet)
            with mcol4:
                st.metric("Undo Snapshots", f"{st.session_state.manager.get_history_count()}")

            st.dataframe(df, use_container_width=True, height=520)
        except Exception as e:
            st.error(f"Error reading sheet data: {str(e)}")

    else:
        st.markdown("""
        <div class="welcome-container">
            <h3 style="margin-top:0; color:#0F172A; font-weight:800; font-size:1.4rem;">
                👋 Welcome to the Autonomous Excel AI Workspace
            </h3>
            <p style="color:#64748B; font-size:0.98rem; line-height:1.5; margin-bottom:1rem;">
                Apni spreadsheet upload karein ya sample data load karein. Hamara autonomous agent aapki <b>Roman Urdu</b> ya <b>English</b> instructions ko deterministic OpenXML operations mein tabdeel karega!
            </p>
            <div class="bento-grid">
                <div class="bento-card">
                    <span class="bento-icon">🧮</span>
                    <div class="bento-title">Dynamic Formula Synthesis</div>
                    <div class="bento-desc">Static numbers ki bajaye live Excel formulas (<code>=SUM</code>, <code>=IF</code>, <code>=AVERAGE</code>) generate karta hai.</div>
                </div>
                <div class="bento-card">
                    <span class="bento-icon">🛡️</span>
                    <div class="bento-title">Safe Reversible Sandbox</div>
                    <div class="bento-desc">Aapki asal file safe rehti hai; har modification se pehle automatic snapshot create hota hai.</div>
                </div>
                <div class="bento-card">
                    <span class="bento-icon">🧹</span>
                    <div class="bento-title">Intelligent Data Cleaning</div>
                    <div class="bento-desc">Extra spaces trim karna, text casing theek karna, aur duplicate rows ko purge karna.</div>
                </div>
                <div class="bento-card">
                    <span class="bento-icon">📈</span>
                    <div class="bento-title">OpenXML Native Visualizations</div>
                    <div class="bento-desc">Worksheets mein direct Bar, Column, Line aur Pie charts embed karta hai.</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_ws1, col_ws2 = st.columns(2)
        with col_ws1:
            if st.button("📁 Load Sample Data (sales.xlsx)", key="canvas_load_sample", type="primary", use_container_width=True):
                load_sample_workbook()
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": "📁 Sample workbook `sales.xlsx` load kar liya gaya hai! Ab aap koi bhi task de sakte hain.",
                    "traces": []
                })
                st.rerun()
        with col_ws2:
            if st.button("📄 Create Blank Workbook", key="canvas_create_blank", use_container_width=True):
                init_new_blank_workbook()
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": "📄 Nayi blank workbook `workbook.xlsx` create kardi gayi hai!",
                    "traces": []
                })
                st.rerun()

    # DOCKED CORNER CHATBOT WIDGET
    corner_quick_selected = None
    if st.session_state.chat_open:
        with st.container(key="corner_chatbot"):
            active_name = html.escape(str(st.session_state.current_file_name or "Ready"))
            col_ch1, col_ch2 = st.columns([5, 1])
            with col_ch1:
                st.markdown(
                    f"""
                    <div style="display:flex; align-items:center; gap:12px; padding:12px 14px 6px 14px;">
                        <div class="chat-avatar-badge">🤖</div>
                        <div>
                            <div style="font-weight:700; font-size:0.98rem; color:#0F172A; line-height:1.2;">
                                Excel AI Copilot
                            </div>
                            <div style="font-size:0.75rem; color:#64748B; display:flex; align-items:center; gap:6px; margin-top:2px;">
                                <span class="pulsing-green-dot"></span>
                                <span>Active • <b>{active_name}</b></span>
                            </div>
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            with col_ch2:
                st.markdown("<div style='padding-top:12px;'>", unsafe_allow_html=True)
                if st.button("➖", key="minimize_corner_chat", help="Minimize corner chatbot"):
                    st.session_state.chat_open = False
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

            # Scrollable chat message area
            with st.container(height=370):
                if not st.session_state.messages:
                    st.markdown(
                        """
                        <div class="corner-welcome-card">
                            <div style="font-size:1.02rem; font-weight:700; color:#0F172A; margin-bottom:4px;">
                                👋 Assalam-o-Alaikum!
                            </div>
                            <p style="font-size:0.85rem; color:#475569; margin:0 0 8px 0; line-height:1.45;">
                                Main aapka <b>Excel AI Copilot</b> hoon. Jo task aap denge, main spreadsheet par live deterministic execution karta rahoonga!
                            </p>
                            <div class="corner-welcome-pill-list">
                                <span class="corner-task-badge">⚡ Live Formulas</span>
                                <span class="corner-task-badge">🧹 Data Cleaning</span>
                                <span class="corner-task-badge">📈 Native Charts</span>
                                <span class="corner-task-badge">🗣️ English & Roman Urdu</span>
                            </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                for msg in st.session_state.messages:
                    with st.chat_message(msg["role"]):
                        st.markdown(msg["content"])
                        if "traces" in msg and msg["traces"]:
                            with st.expander(f"🛠️ Agent Execution Traces ({len(msg['traces'])})"):
                                for t in msg["traces"]:
                                    st.markdown(f"**Step {t.step}: `{t.tool_name}`**")
                                    st.json({"arguments": t.arguments, "result": t.result})

                # Destructive confirmation inside corner chat
                if st.session_state.pending_confirmation:
                    conf = st.session_state.pending_confirmation
                    warning_text = conf.get('details', {}).get('warning', 'Do you confirm this action?')
                    st.markdown(render_destructive_confirmation_html(warning_text), unsafe_allow_html=True)
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        if st.button("✅ Confirm", key="corner_confirm_btn", type="primary", use_container_width=True):
                            with st.spinner("Executing confirmed action..."):
                                res = st.session_state.agent.confirm_and_resume(confirmed=True)
                            st.session_state.pending_confirmation = None
                            st.session_state.messages.append({
                                "role": "assistant",
                                "content": res.final_response,
                                "traces": res.tool_traces
                            })
                            st.rerun()
                    with cc2:
                        if st.button("❌ Cancel", key="corner_cancel_btn", use_container_width=True):
                            res = st.session_state.agent.confirm_and_resume(confirmed=False)
                            st.session_state.pending_confirmation = None
                            st.session_state.messages.append({
                                "role": "assistant",
                                "content": res.final_response,
                                "traces": []
                            })
                            st.rerun()

            # Compact Quick Action Tasks (4 pills)
            st.markdown("<p style='font-size:0.75rem; font-weight:700; color:#64748B; margin:4px 0 2px 8px; letter-spacing:0.5px;'>⚡ QUICK PROMPTS:</p>", unsafe_allow_html=True)
            qp1, qp2, qp3, qp4 = st.columns(4)
            with qp1:
                if st.button("📊 Total", key="qp_corner_total", use_container_width=True):
                    corner_quick_selected = "total sales nikalo"
            with qp2:
                if st.button("🧹 Dups", key="qp_corner_dups", use_container_width=True):
                    corner_quick_selected = "Order_ID ke hisaab se duplicates hata do"
            with qp3:
                if st.button("📈 Chart", key="qp_corner_chart", use_container_width=True):
                    corner_quick_selected = "Targets sheet mein region ke hisaab se column chart banao"
            with qp4:
                if st.button("✨ Clean", key="qp_corner_clean", use_container_width=True):
                    corner_quick_selected = "Sales Rep column D ke extra spaces hatao aur title case kardo"

            # Corner Chat Input
            corner_inp = st.chat_input("Task likhein (e.g. 'total sales nikalo')...", key="corner_chat_input")
            to_execute = corner_quick_selected or corner_inp
            if to_execute:
                run_agent_task(to_execute)

    else:
        # Floating corner trigger pill
        with st.container(key="corner_chat_trigger"):
            if st.button("✨ Ask Excel AI Copilot", key="open_corner_chat", help="Open Excel AI Chatbot"):
                st.session_state.chat_open = True
                st.rerun()

else:
    # ALTERNATIVE VIEW: Full Page Tabs
    tab_preview, tab_chat = st.tabs(["📑 Live Sheet Preview", "💬 AI Copilot"])

    with tab_preview:
        if not st.session_state.file_loaded:
            st.info("👈 Abhi koi workbook loaded nahi hai. Nayi sheet banane ke liye chat mein task dein ya sample data load karein.")
        else:
            sheet_names = st.session_state.manager.get_sheet_names()
            selected_sheet = st.selectbox("Select Sheet to View", options=sheet_names, key="tab_sheet_view_sel")
            try:
                curr_path = st.session_state.manager.get_current_file_path()
                curr_stat = Path(curr_path).stat()
                df, active_s, is_oversize = load_preview_cached(
                    str(Path(curr_path).resolve()),
                    curr_stat.st_mtime_ns,
                    curr_stat.st_size,
                    sheet_name=selected_sheet,
                    max_rows=500
                )
                df = df.convert_dtypes()
                for col in df.columns:
                    if df[col].dtype == object:
                        df[col] = df[col].astype("string")
                st.dataframe(df, use_container_width=True, height=480)
            except Exception as e:
                st.error(f"Error reading sheet data: {str(e)}")

    with tab_chat:
        if st.session_state.file_loaded:
            try:
                active_sheets = st.session_state.manager.get_sheet_names()
            except Exception:
                active_sheets = []
            st.markdown(
                render_file_status_bar_html(st.session_state.current_file_name, active_sheets),
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div class="file-status-bar" style="background:#F8FAFC; border:1px solid #CBD5E1; color:#334155;">'
                '<span>💡</span>'
                '<span>No active workbook loaded • Type any task below, or click a Quick Prompt to start!</span>'
                '</div>',
                unsafe_allow_html=True
            )

        st.markdown("<p style='font-size:0.85rem; font-weight:700; color:#64748B; margin-bottom:6px;'>⚡ QUICK PROMPTS:</p>", unsafe_allow_html=True)
        qrow1_c1, qrow1_c2 = st.columns(2)
        tab_quick_selected = None

        if not st.session_state.file_loaded:
            with qrow1_c1:
                if st.button("📁 Load Sample Data (sales.xlsx)", key="tab_load_sample", use_container_width=True):
                    load_sample_workbook()
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": "📁 Sample workbook `sales.xlsx` load kar liya gaya hai! Ab aap koi bhi task de sakte hain.",
                        "traces": []
                    })
                    st.rerun()
            with qrow1_c2:
                if st.button("📄 Nayi Blank Sheet Banao", key="tab_create_blank", use_container_width=True):
                    init_new_blank_workbook()
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": "📄 Nayi blank workbook `workbook.xlsx` create kardi gayi hai! Ab aap columns aur data add karne ka task de sakte hain.",
                        "traces": []
                    })
                    st.rerun()
        else:
            with qrow1_c1:
                if st.button("📊 Total sales nikalo", key="tab_quick_total", use_container_width=True):
                    tab_quick_selected = "total sales nikalo"
            with qrow1_c2:
                if st.button("🧹 Duplicates hata do", key="tab_quick_dups", use_container_width=True):
                    tab_quick_selected = "Order_ID ke hisaab se duplicates hata do"

        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if "traces" in msg and msg["traces"]:
                    with st.expander(f"🛠️ Agent Execution Traces ({len(msg['traces'])} action(s))"):
                        for t in msg["traces"]:
                            st.markdown(f"**Step {t.step}: `{t.tool_name}`**")
                            st.json({"arguments": t.arguments, "result": t.result})

        if st.session_state.pending_confirmation:
            conf = st.session_state.pending_confirmation
            warning_text = conf.get('details', {}).get('warning', 'Do you confirm this action?')
            st.markdown(render_destructive_confirmation_html(warning_text), unsafe_allow_html=True)
            col_c1, col_c2 = st.columns([1, 1])
            with col_c1:
                if st.button("✅ Confirm & Execute", key="tab_confirm_btn", type="primary", use_container_width=True):
                    with st.spinner("Executing confirmed action..."):
                        res = st.session_state.agent.confirm_and_resume(confirmed=True)
                    st.session_state.pending_confirmation = None
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": res.final_response,
                        "traces": res.tool_traces
                    })
                    st.rerun()
            with col_c2:
                if st.button("❌ Cancel Action", key="tab_cancel_btn", use_container_width=True):
                    res = st.session_state.agent.confirm_and_resume(confirmed=False)
                    st.session_state.pending_confirmation = None
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": res.final_response,
                        "traces": []
                    })
                    st.rerun()

        tab_chat_input = st.chat_input("Ask Excel AI (e.g. 'total sales nikalo')...", key="tab_chat_input")
        tab_prompt = tab_quick_selected or tab_chat_input
        if tab_prompt:
            run_agent_task(tab_prompt)
