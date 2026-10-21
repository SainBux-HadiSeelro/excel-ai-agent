"""Tests for ExcelAgent loop and error self-correction."""

from agent import ExcelAgent
from llm import MockProvider, ToolCall, LLMResponse


def test_agent_one_step_tool_call(sample_manager):
    # Setup mock provider: calls list_sheets, then gives final response
    mock_llm = MockProvider([
        LLMResponse(
            content="I will check available sheets first.",
            tool_calls=[ToolCall(id="tc-1", name="list_sheets", arguments={})]
        ),
        LLMResponse(
            content="I found the sheets: Sales and Targets.",
            tool_calls=[]
        )
    ])

    agent = ExcelAgent(manager=sample_manager, provider=mock_llm)
    result = agent.run("What sheets do we have?")

    assert result.success is True
    assert result.steps_taken == 2
    assert len(result.tool_traces) == 1
    assert result.tool_traces[0].tool_name == "list_sheets"
    assert "Sales" in result.final_response


def test_agent_destructive_confirmation_pause(sample_manager):
    # LLM attempts to delete sheet without confirmed=True
    mock_llm = MockProvider([
        LLMResponse(
            content="I will delete sheet Targets.",
            tool_calls=[ToolCall(id="tc-del", name="delete_sheet", arguments={"title": "Targets", "confirmed": False})]
        )
    ])

    agent = ExcelAgent(manager=sample_manager, provider=mock_llm)
    result = agent.run("Delete Targets sheet")

    # Agent must halt and request confirmation
    assert result.pending_confirmation is not None
    assert "Confirmation Required" in result.final_response


def test_agent_error_recovery(sample_manager):
    # LLM makes an erroneous tool call, receives error, then self-corrects
    mock_llm = MockProvider([
        LLMResponse(
            content="Attempting invalid sheet read",
            tool_calls=[ToolCall(id="tc-err", name="read_range", arguments={"sheet_name": "NonExistent"})]
        ),
        LLMResponse(
            content="Correcting sheet name to Sales",
            tool_calls=[ToolCall(id="tc-ok", name="read_range", arguments={"sheet_name": "Sales", "max_rows": 2})]
        ),
        LLMResponse(
            content="Completed successfully.",
            tool_calls=[]
        )
    ])

    agent = ExcelAgent(manager=sample_manager, provider=mock_llm)
    result = agent.run("Read sales rows")

    assert result.success is True
    assert result.steps_taken == 3
    assert len(result.tool_traces) == 2
    assert result.tool_traces[0].error is not None
    assert result.tool_traces[1].error is None
