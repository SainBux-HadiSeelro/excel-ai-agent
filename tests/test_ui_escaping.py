"""Tests for HTML escaping of user- and file-derived strings in UI."""

from app import render_file_status_bar_html, render_destructive_confirmation_html


def test_sheet_name_html_escaping():
    # A sheet named "<b>x</b>" must be escaped so it renders as text rather than raw bold HTML
    html_out = render_file_status_bar_html(
        filename="sales.xlsx",
        active_sheets=["Sales", "<b>x</b>"]
    )
    assert "&lt;b&gt;x&lt;/b&gt;" in html_out
    # Ensure there is no raw unescaped <b>x</b> (only the <b>sales.xlsx</b> wrapper)
    assert "<b>x</b>" not in html_out


def test_filename_html_escaping():
    html_out = render_file_status_bar_html(
        filename="<script>alert(1)</script>.xlsx",
        active_sheets=["Sheet1"]
    )
    assert "&lt;script&gt;alert(1)&lt;/script&gt;.xlsx" in html_out
    assert "<script>" not in html_out


def test_destructive_warning_html_escaping():
    html_out = render_destructive_confirmation_html(
        warning_text="Are you sure you want to delete sheet '<b>MySheet</b>' & row <1>?"
    )
    assert "&lt;b&gt;MySheet&lt;/b&gt;" in html_out
    assert "&amp;" in html_out
    assert "&lt;1&gt;" in html_out
    assert "<b>MySheet</b>" not in html_out
