"""Tests for format_range and add_chart."""

from tools.formatting import format_range
from tools.visualization import add_chart


def test_format_range(sample_manager):
    res = format_range(
        sample_manager,
        sheet_name="Sales",
        cell_range="A1:I1",
        bold=True,
        bg_color="2F5597",
        font_color="FFFFFF",
        border="thin",
        align_horizontal="center",
        auto_fit_columns=True
    )
    assert res["success"] is True
    assert res["formatted_cells_count"] == 9
    assert res["auto_fit_applied"] is True


def test_add_chart_bar_and_pie(sample_manager):
    res_bar = add_chart(
        sample_manager,
        sheet_name="Targets",
        chart_type="column",
        data_range="B1:B5",
        categories_range="A2:A5",
        title="Regional Targets",
        target_cell="E2"
    )
    assert res_bar["success"] is True

    res_pie = add_chart(
        sample_manager,
        sheet_name="Targets",
        chart_type="pie",
        data_range="B1:B5",
        categories_range="A2:A5",
        title="Targets Distribution",
        target_cell="E15"
    )
    assert res_pie["success"] is True
