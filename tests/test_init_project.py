"""Tests for --init-project scaffold."""
import os
import py_compile

import pytest

from excel_to_airtest import AirtestGenerator, _PROJECT_REGISTRY, _scaffold_project, register_project


@pytest.fixture(autouse=True)
def clean_registry():
    before = dict(_PROJECT_REGISTRY)
    yield
    _PROJECT_REGISTRY.clear()
    _PROJECT_REGISTRY.update(before)


def test_scaffold_creates_file(tmp_path):
    projects_dir = tmp_path / "projects"
    projects_dir.mkdir()
    _scaffold_project("newgame", projects_dir=str(projects_dir))
    assert (projects_dir / "newgame.py").exists()


def test_scaffold_file_compiles(tmp_path):
    projects_dir = tmp_path / "projects"
    projects_dir.mkdir()
    _scaffold_project("newgame", projects_dir=str(projects_dir))
    py_compile.compile(str(projects_dir / "newgame.py"), doraise=True)


def test_scaffold_contains_register_decorator(tmp_path):
    projects_dir = tmp_path / "projects"
    projects_dir.mkdir()
    _scaffold_project("newgame", projects_dir=str(projects_dir))
    content = (projects_dir / "newgame.py").read_text()
    assert '@register_project("newgame")' in content


def test_scaffold_class_name_pascal_case(tmp_path):
    projects_dir = tmp_path / "projects"
    projects_dir.mkdir()
    _scaffold_project("my_game", projects_dir=str(projects_dir))
    content = (projects_dir / "my_game.py").read_text()
    assert "MyGameGenerator" in content


def test_scaffold_contains_package_placeholder(tmp_path):
    projects_dir = tmp_path / "projects"
    projects_dir.mkdir()
    _scaffold_project("newgame", projects_dir=str(projects_dir))
    content = (projects_dir / "newgame.py").read_text()
    assert "com.example.newgame" in content


def test_scaffold_raises_if_file_exists(tmp_path):
    projects_dir = tmp_path / "projects"
    projects_dir.mkdir()
    _scaffold_project("newgame", projects_dir=str(projects_dir))
    with pytest.raises(FileExistsError):
        _scaffold_project("newgame", projects_dir=str(projects_dir))
