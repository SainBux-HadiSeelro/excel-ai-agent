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

# Custom CSS for modern polished interface
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .main .block-container {
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }
    
    /* Hero Banner */
    .hero-banner {
        background: linear-gradient(135deg, #0F2027 0%, #203A43 50%, #2C5364 100%);
        color: white;
        padding: 1.8rem 2.2rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px rgba(15, 32, 39, 0.15);
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 700;
        margin: 0;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hero-subtitle {
        font-size: 1rem;
        color: #E2E8F0;
        margin-top: 6px;
        margin-bottom: 12px;
    }
    .badge-pill {
        display: inline-block;
        padding: 4px 12px;
        background: rgba(255, 255, 255, 0.15);
        backdrop-filter: blur(8px);
        color: #F8FAFC;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 500;
        margin-right: 8px;
        border: 1px solid rgba(255, 255, 255, 0.2);
    }
    
    /* Feature Welcome Cards */
    .welcome-container {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 2rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03);
        margin-bottom: 1.5rem;
    }
    .feature-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.2rem;
        height: 100%;
        transition: transform 0.2s, box-shadow 0.2s;
    }
    .feature-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(0,0,0,0.06);
    }
    .feature-icon {
        font-size: 1.8rem;
        margin-bottom: 8px;
    }
    .feature-title {
        font-weight: 600;
        font-size: 1rem;
        color: #1E293B;
        margin-bottom: 4px;
    }
    .feature-desc {
        font-size: 0.85rem;
        color: #64748B;
        line-height: 1.4;
    }

    /* Active File Pill */
    .file-status-bar {
        background: #F0FDF4;
        border: 1px solid #BBF7D0;
        color: #166534;
        padding: 10px 16px;
        border-radius: 10px;
        font-size: 0.9rem;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 1.2rem;
    }

    /* Tool traces styling */
    .trace-card {
        background: #F8FAFC;
        border-left: 3px solid #3B82F6;
        border-radius: 6px;
        padding: 8px 12px;
        margin: 6px 0;
        font-family: monospace;
        font-size: 0.85rem;
    }
    
    /* Destructive Confirmation Box */
    .destructive-box {
        background: #FEF2F2;
        border: 1.5px solid #FCA5A5;
        border-radius: 12px;
        padding: 1.2rem;
        margin: 1.2rem 0;
    }

    /* Mobile Responsiveness & Viewport Optimization */
    @media (max-width: 768px) {
        .main .block-container {
            padding-top: 1rem !important;
            padding-left: 0.75rem !important;
            padding-right: 0.75rem !important;
            padding-bottom: 2.5rem !important;
        }
        .hero-banner {
            padding: 1.2rem 1rem !important;
            border-radius: 12px !important;
            margin-bottom: 1rem !important;
        }
        .hero-title {
            font-size: 1.35rem !important;
            gap: 6px !important;
        }
        .hero-subtitle {
            font-size: 0.85rem !important;
            margin-top: 4px !important;
            margin-bottom: 10px !important;
            line-height: 1.3 !important;
        }
        .badge-pill {
            display: inline-block !important;
            font-size: 0.7rem !important;
            padding: 3px 8px !important;
            margin-bottom: 4px !important;
            margin-right: 4px !important;
        }
        .welcome-container {
            padding: 1.2rem 1rem !important;
            border-radius: 12px !important;
            margin-bottom: 1rem !important;
        }
        .feature-card {
            padding: 0.9rem !important;
            margin-bottom: 8px !important;
        }
        .file-status-bar {
            font-size: 0.8rem !important;
            padding: 8px 10px !important;
            line-height: 1.35 !important;
            word-break: break-word !important;
        }
        .destructive-box {
            padding: 1rem !important;
            border-radius: 10px !important;
        }
        .stButton > button {
            min-height: 44px !important;
            font-size: 0.86rem !important;
            padding: 8px 10px !important;
            border-radius: 8px !important;
        }
        div[data-testid="stChatMessage"] {
            padding: 8px 10px !important;
        }
        div[data-testid="stChatMessage"] p {
            font-size: 0.9rem !important;
            line-height: 1.4 !important;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.2rem !important;
        }
        div[data-testid="stMetricLabel"] {
            font-size: 0.78rem !important;
        }
        div[data-testid="stDataFrame"] {
            overflow-x: auto !important;
        }
    }

    /* Floating Corner Chatbot Widget ("kone m laga rhe") */
    .st-key-corner_chatbot {
        position: fixed !important;
        bottom: 20px !important;
        right: 20px !important;
        width: 440px !important;
        max-width: calc(100vw - 32px) !important;
        height: 610px !important;
        max-height: calc(100vh - 40px) !important;
        background: #FFFFFF !important;
        border-radius: 16px !important;
        box-shadow: 0 20px 50px rgba(15, 23, 42, 0.25), 0 4px 15px rgba(15, 23, 42, 0.1) !important;
        border: 1.5px solid #CBD5E1 !important;
        z-index: 999999 !important;
        display: flex !important;
        flex-direction: column !important;
        padding: 12px 14px !important;
        overflow: hidden !important;
    }

    .st-key-corner_chatbot [data-testid="stVerticalBlock"] {
        gap: 0.35rem !important;
    }

    .st-key-corner_chat_trigger {
        position: fixed !important;
        bottom: 24px !important;
        right: 24px !important;
        z-index: 999999 !important;
    }

    .st-key-corner_chat_trigger button {
        border-radius: 30px !important;
        padding: 12px 24px !important;
        background: linear-gradient(135deg, #0F2027 0%, #203A43 50%, #2C5364 100%) !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        box-shadow: 0 8px 24px rgba(15, 32, 39, 0.35) !important;
        border: 1.5px solid rgba(255, 255, 255, 0.3) !important;
        transition: transform 0.2s, box-shadow 0.2s !important;
    }

    .st-key-corner_chat_trigger button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 28px rgba(15, 32, 39, 0.45) !important;
    }

    @media (max-width: 768px) {
        .st-key-corner_chatbot {
            bottom: 8px !important;
            right: 8px !important;
            left: 8px !important;
            width: auto !important;
            height: 75vh !important;
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
    st.session_state.chat_open = True

# Sidebar Configuration
with st.sidebar:
    st.markdown("### 📊 **Excel AI Settings**")
    st.caption("Multilingual Spreadsheet Agent • English & Roman Urdu")

    st.markdown("#### 1. 🤖 Provider & Model")
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
    st.markdown("#### 2. 📂 Workbook Actions")
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
    st.markdown(f"**🔄 Undo History:** `{history_count}` snapshot(s)")

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
                    label="💾 Download",
                    data=f.read(),
                    file_name=f"updated_{st.session_state.current_file_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    st.markdown("---")
    chat_view_mode = st.radio(
        "Chatbot Position",
        options=["Corner Floating Chatbot (Recommended)", "Full Page Tabs"],
        index=0,
        help="Dock the chatbot in the bottom-right corner or view full-page tabs."
    )

    st.markdown("---")
    if st.button("🗑️ Clear Chat History", use_container_width=True):
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


# Top Hero Banner
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">
        <span>📊 Excel AI Workspace</span>
    </div>
    <div class="hero-subtitle">
        Interactive spreadsheet canvas with an autonomous AI chatbot docked in the corner. Ask in English or Roman Urdu!
    </div>
    <div>
        <span class="badge-pill">💬 Corner Chatbot Active</span>
        <span class="badge-pill">🗣️ English & Roman Urdu</span>
        <span class="badge-pill">🛡️ Non-Destructive Sandbox</span>
        <span class="badge-pill">⚡ Live Excel Formulas</span>
        <span class="badge-pill">↩️ Instant Undo</span>
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
            selected_sheet = st.selectbox("Select Sheet to View", options=active_sheets, key="main_sheet_sel")
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
                    label="💾 Download",
                    data=f.read(),
                    file_name=f"updated_{st.session_state.current_file_name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="canvas_download"
                )
        with col_tb4:
            if st.button("💬 Toggle Corner Chatbot", key="canvas_toggle_chat", use_container_width=True):
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

            # Metrics row
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
            <h3 style="margin-top:0; color:#1E293B;">👋 Welcome to Excel AI Workspace!</h3>
            <p style="color:#64748B; font-size:0.95rem;">
                Apni <b>.xlsx</b> ya <b>.csv</b> spreadsheet upload karein, sample data load karein, ya <b>kone mein mojood AI Chatbot</b> se seedha koi bhi task dein!
            </p>
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

    # DOCKED CORNER CHATBOT WIDGET ("kone m laga rhe")
    corner_quick_selected = None
    if st.session_state.chat_open:
        with st.container(key="corner_chatbot"):
            col_ch1, col_ch2 = st.columns([5, 1])
            with col_ch1:
                st.markdown(
                    "<div style='display:flex; align-items:center; gap:8px;'>"
                    "<b>💬 Excel AI Chatbot</b>"
                    "<span style='font-size:0.75rem; color:#16A34A; background:#DCFCE7; padding:2px 8px; border-radius:10px; font-weight:600;'>● Online</span>"
                    "</div>",
                    unsafe_allow_html=True
                )
            with col_ch2:
                if st.button("➖", key="minimize_corner_chat", help="Minimize corner chatbot"):
                    st.session_state.chat_open = False
                    st.rerun()

            # Scrollable chat message area
            with st.container(height=360):
                if not st.session_state.messages:
                    st.markdown(
                        "👋 **Assalam-o-Alaikum!** Main aapka Excel AI Chatbot hoon.\n\n"
                        "Aap jo bhi task yahan kone se denge, main spreadsheet par execute karta rahoonga:\n\n"
                        "- 📊 *'total sales nikalo'*\n"
                        "- 🧹 *'Order_ID ke hisaab se duplicates hata do'*\n"
                        "- 📈 *'region ke hisaab se chart banao'*\n"
                        "- 📄 *'nayi Targets sheet banao'*"
                    )

                for msg in st.session_state.messages:
                    with st.chat_message(msg["role"]):
                        st.markdown(msg["content"])
                        if "traces" in msg and msg["traces"]:
                            with st.expander(f"🛠️ Tool Traces ({len(msg['traces'])})"):
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

            # Compact Quick Action Tasks
            qp1, qp2, qp3 = st.columns(3)
            with qp1:
                if st.button("📊 Total", key="qp_corner_total", use_container_width=True):
                    corner_quick_selected = "total sales nikalo"
            with qp2:
                if st.button("🧹 Dups", key="qp_corner_dups", use_container_width=True):
                    corner_quick_selected = "Order_ID ke hisaab se duplicates hata do"
            with qp3:
                if st.button("📈 Chart", key="qp_corner_chart", use_container_width=True):
                    corner_quick_selected = "Targets sheet mein region ke hisaab se column chart banao"

            # Corner Chat Input
            corner_inp = st.chat_input("Task likhein (e.g. 'total sales nikalo')...", key="corner_chat_input")
            to_execute = corner_quick_selected or corner_inp
            if to_execute:
                run_agent_task(to_execute)

    else:
        # Floating corner trigger pill
        with st.container(key="corner_chat_trigger"):
            if st.button("💬 Excel AI Chatbot", key="open_corner_chat", help="Open Excel AI Chatbot"):
                st.session_state.chat_open = True
                st.rerun()

else:
    # ALTERNATIVE VIEW: Full Page Tabs
    tab_chat, tab_preview = st.tabs(["💬 AI Chatbot", "📑 Live Sheet Preview"])

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

        st.markdown("<p style='font-size:0.85rem; font-weight:600; color:#64748B; margin-bottom:6px;'>⚡ QUICK PROMPTS:</p>", unsafe_allow_html=True)
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
                    with st.expander(f"🛠️ Tool Execution Traces ({len(msg['traces'])} action(s))"):
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
