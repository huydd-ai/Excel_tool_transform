"""Tests for project registry: @register_project, _discover_projects."""
import importlib
import os
import sys
import textwrap

import pytest

from excel_to_airtest import AirtestGenerator, _PROJECT_REGISTRY, register_project


@pytest.fixture(autouse=True)
def clean_registry():
    """Isolate registry state between tests."""
    before = dict(_PROJECT_REGISTRY)
    yield
    _PROJECT_REGISTRY.clear()
    _PROJECT_REGISTRY.update(before)


def test_register_project_adds_to_registry():
    @register_project("testgame")
    class TestGen(AirtestGenerator):
        pass
    assert "testgame" in _PROJECT_REGISTRY
    assert _PROJECT_REGISTRY["testgame"] is TestGen


def test_register_project_lowercases_name():
    @register_project("TestGame2")
    class TestGen2(AirtestGenerator):
        pass
    assert "testgame2" in _PROJECT_REGISTRY


def test_register_project_returns_class_unchanged():
    @register_project("testgame3")
    class TestGen3(AirtestGenerator):
        DEFAULT_APP_PACKAGE = "com.test"
    assert TestGen3.DEFAULT_APP_PACKAGE == "com.test"


def test_register_project_overwrites_same_name():
    @register_project("dup")
    class Gen1(AirtestGenerator):
        pass
    @register_project("dup")
    class Gen2(AirtestGenerator):
        pass
    assert _PROJECT_REGISTRY["dup"] is Gen2


def test_discover_projects_skips_missing_dir():
    """No error when projects/ dir does not exist."""
    from excel_to_airtest import _discover_projects
    _discover_projects(projects_dir="/nonexistent/path/xyz")  # must not raise


def test_discover_projects_imports_project_file(tmp_path):
    """A .py file in projects_dir with @register_project gets loaded."""
    from excel_to_airtest import _discover_projects
    proj_dir = tmp_path / "projects"
    proj_dir.mkdir()
    (proj_dir / "__init__.py").write_text("")
    (proj_dir / "dyntest.py").write_text(textwrap.dedent("""
        import sys, os
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from excel_to_airtest import AirtestGenerator, register_project

        @register_project("dyntest_proj")
        class DynTestGen(AirtestGenerator):
            DEFAULT_APP_PACKAGE = "com.dyntest"
    """))
    _discover_projects(projects_dir=str(proj_dir))
    assert "dyntest_proj" in _PROJECT_REGISTRY


def test_discover_projects_skips_dunder_files(tmp_path):
    proj_dir = tmp_path / "projects"
    proj_dir.mkdir()
    (proj_dir / "__init__.py").write_text("")
    (proj_dir / "__helpers.py").write_text("raise RuntimeError('should not import')")
    from excel_to_airtest import _discover_projects
    _discover_projects(projects_dir=str(proj_dir))  # must not raise


def test_classname_from_name():
    from excel_to_airtest import _classname_from_name
    assert _classname_from_name("pixon") == "Pixon"
    assert _classname_from_name("my_game") == "MyGame"
    assert _classname_from_name("my-game") == "MyGame"
    assert _classname_from_name("my game") == "MyGame"
