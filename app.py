"""
Streamlit Web UI for Excel AI Agent.
Provides:
- Sidebar: File upload, Provider & Model selector, API key input, Session status
- Main Area: Interactive Chat, Live tool execution traces, Sheet Preview table, Download & Undo buttons
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
import streamlit as st
import pandas as pd
from dotenv import load_dotenv

# Ensure local imports work
sys.path.insert(0, str(Path(__file__).parent.resolve()))

from workspace_manager import WorkspaceManager
from llm import get_llm_provider, GeminiProvider, OpenAIProvider, AnthropicProvider, DEFAULT_GEMINI_MODEL
from agent import ExcelAgent, ToolExecutionTrace

load_dotenv()

st.set_page_config(
    page_title="Excel AI Agent",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for polished interface
st.markdown("""
<style>
    .reportview-container .main .block-container{
        padding-top: 1.5rem;
    }
    .tool-pill {
        display: inline-block;
        padding: 2px 8px;
        background-color: #EBF3FB;
        color: #1F4E78;
        border-radius: 4px;
        font-family: monospace;
        font-size: 0.85em;
        margin: 2px;
    }
    .destructive-box {
        border: 2px solid #E06C75;
        border-radius: 8px;
        padding: 12px;
        background-color: #FFF5F5;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if "manager" not in st.session_state:
    st.session_state.manager = WorkspaceManager()
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

# Sidebar Configuration
with st.sidebar:
    st.title("📊 Excel AI Agent")
    st.caption("Natural language spreadsheet manipulation (English & Roman Urdu)")

    st.subheader("1. LLM Settings")
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
        "OpenAI": "gpt-4o-mini",
        "Anthropic": "claude-3-5-haiku-20241022",
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
        st.info("💡 **No API key?** Switch Provider to **'Offline Demo (Mock)'** above to test without an API key.")

    st.markdown("---")
    st.subheader("2. Workbook Upload")
    uploaded_file = st.file_uploader(
        "Choose an Excel (.xlsx) or CSV file",
        type=["xlsx", "csv"]
    )

    use_sample = st.button("📁 Load Sample Data (sales.xlsx)", use_container_width=True)

    if uploaded_file is not None:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if st.session_state.get("current_file_id") != file_id:
            # Save temp copy and initialize workspace
            temp_dir = Path("temp_uploads")
            temp_dir.mkdir(exist_ok=True)
            temp_path = temp_dir / uploaded_file.name
            with open(temp_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            st.session_state.manager = WorkspaceManager()
            st.session_state.manager.init_from_file(temp_path)
            st.session_state.current_file_name = uploaded_file.name
            st.session_state.current_file_id = file_id
            st.session_state.file_loaded = True
            st.session_state.messages = []
            st.session_state.traces = []
            st.session_state.pending_confirmation = None
            st.session_state.agent = None
            st.success(f"Loaded: {uploaded_file.name}")

    elif use_sample:
        if st.session_state.get("current_file_id") != "sample_sales":
            sample_path = Path("sample_data/sales.xlsx")
            if not sample_path.exists():
                import generate_sample_data
            st.session_state.manager = WorkspaceManager()
            st.session_state.manager.init_from_file(sample_path)
            st.session_state.current_file_name = "sales.xlsx"
            st.session_state.current_file_id = "sample_sales"
            st.session_state.file_loaded = True
            st.session_state.messages = []
            st.session_state.traces = []
            st.session_state.pending_confirmation = None
            st.session_state.agent = None
            st.success("Loaded sample_data/sales.xlsx")

    # Load agent if keys available or offline demo (only instantiate when needed/changed)
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
                st.sidebar.error(f"Provider Error: {str(e)}")

    st.markdown("---")
    # Status & Undo Controls
    history_count = st.session_state.manager.get_history_count()
    st.write(f"🔄 Undoable steps available: **{history_count}**")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        if st.button("↩️ Undo", use_container_width=True, disabled=(history_count == 0)):
            undo_res = st.session_state.manager.undo()
            if undo_res["success"]:
                st.success(undo_res["message"])
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


# Main Content Layout: Two Tabs (Chat & Data Preview)
tab_chat, tab_preview = st.tabs(["💬 AI Assistant", "📑 Sheet Preview"])

# TAB 1: Chat and Tool Traces
with tab_chat:
    if not st.session_state.file_loaded:
        st.info("👈 Please upload an Excel/CSV file or click 'Load Sample Data' in the sidebar to begin.")
    else:
        st.caption(f"Currently working on: **{st.session_state.current_file_name}** (Isolated Workspace)")

        # Display Chat History
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if "traces" in msg and msg["traces"]:
                    with st.expander(f"🛠️ Tool Traces ({len(msg['traces'])} steps)"):
                        for t in msg["traces"]:
                            st.markdown(f"**Step {t.step}: `{t.tool_name}`**")
                            st.json({"args": t.arguments, "result": t.result})

        # Destructive Confirmation Card (if any pending)
        if st.session_state.pending_confirmation:
            conf = st.session_state.pending_confirmation
            st.warning(f"⚠️ **Destructive Action Confirmation Needed**: {conf.get('details', {}).get('warning', 'Proceed?')}")
            col_c1, col_c2 = st.columns([1, 1])
            with col_c1:
                if st.button("✅ Confirm & Execute", type="primary", use_container_width=True):
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
                if st.button("❌ Cancel Action", use_container_width=True):
                    res = st.session_state.agent.confirm_and_resume(confirmed=False)
                    st.session_state.pending_confirmation = None
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": "Action cancelled by user.",
                        "traces": []
                    })
                    st.rerun()

        # Chat Input
        user_prompt = st.chat_input("Ask Excel AI (e.g. 'total sales nikalo', 'duplicates hata do', 'region ka chart banao')...")

        if user_prompt:
            if not st.session_state.agent:
                st.error("⚠️ Please supply a valid API key in the sidebar, or select **Offline Demo (Mock)** in Provider to test without an API key.")
            else:
                # Add user prompt to chat
                st.session_state.messages.append({"role": "user", "content": user_prompt})
                with st.chat_message("user"):
                    st.markdown(user_prompt)

                with st.chat_message("assistant"):
                    trace_placeholder = st.empty()
                    live_traces = []

                    def streamlit_on_step(step, tool_name, args, result):
                        live_traces.append(ToolExecutionTrace(
                            step=step,
                            tool_name=tool_name,
                            arguments=args,
                            result=result
                        ))
                        trace_placeholder.markdown(
                            f"⚡ *Calling tool `{tool_name}` (Step {step})...*"
                        )

                    with st.spinner("Analyzing spreadsheet and executing..."):
                        agent_res = st.session_state.agent.run(user_prompt, on_step=streamlit_on_step)

                    trace_placeholder.empty()

                    # Render final response
                    if not agent_res.success and agent_res.error:
                        st.error(f"❌ {agent_res.final_response}")
                    else:
                        st.markdown(agent_res.final_response)

                    if agent_res.tool_traces:
                        with st.expander(f"🛠️ Tool Traces ({len(agent_res.tool_traces)} steps)"):
                            for t in agent_res.tool_traces:
                                st.markdown(f"**Step {t.step}: `{t.tool_name}`**")
                                st.json({"args": t.arguments, "result": t.result})

                    # Handle confirmation state
                    if agent_res.pending_confirmation:
                        st.session_state.pending_confirmation = agent_res.pending_confirmation

                    # Save to history
                    st.session_state.messages.append({
                        "role": "assistant",
                        "content": agent_res.final_response,
                        "traces": agent_res.tool_traces
                    })

                    st.rerun()


# TAB 2: Sheet Preview
with tab_preview:
    if not st.session_state.file_loaded:
        st.info("No workbook loaded yet.")
    else:
        sheet_names = st.session_state.manager.get_sheet_names()
        selected_sheet = st.selectbox("Select Sheet to View", options=sheet_names)

        col_meta1, col_meta2 = st.columns(2)
        try:
            df, active_s = st.session_state.manager.get_sheet_dataframe(sheet_name=selected_sheet, max_rows=500)
            # Normalize mixed-type object columns so Streamlit/Arrow can serialize them safely
            df = df.convert_dtypes()
            for col in df.columns:
                if df[col].dtype == object:
                    df[col] = df[col].astype("string")
            with col_meta1:
                st.metric("Total Visible Rows", f"{len(df):,}")
            with col_meta2:
                st.metric("Total Columns", f"{len(df.columns)}")

            st.dataframe(df, use_container_width=True, height=450)
        except Exception as e:
            st.error(f"Error reading sheet data: {str(e)}")
