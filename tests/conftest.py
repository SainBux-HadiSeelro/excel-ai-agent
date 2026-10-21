"""Pytest fixtures for Excel AI Agent tests."""

import pytest
import shutil
from pathlib import Path
from workspace_manager import WorkspaceManager
import generate_sample_data


@pytest.fixture
def sample_manager(tmp_path):
    """Provides a fresh WorkspaceManager initialized with a copy of sample_data/sales.xlsx."""
    sample_file = Path("sample_data/sales.xlsx")
    if not sample_file.exists():
        # Generate if needed
        import runpy
        runpy.run_path("generate_sample_data.py")

    manager = WorkspaceManager(workspace_root=str(tmp_path / "workspace"))
    manager.init_from_file(sample_file)
    return manager
