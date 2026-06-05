import pytest
import core_engine

def test_add_project():
    core_engine.add_project("Test Project 1")
    projects = core_engine.get_projects()
    assert "Test Project 1" in projects

def test_add_duplicate_project():
    core_engine.add_project("Duplicate Project")
    with pytest.raises(ValueError, match="already exists"):
        core_engine.add_project("Duplicate Project")

def test_add_empty_project():
    with pytest.raises(ValueError, match="cannot be empty"):
        core_engine.add_project("")

def test_update_project():
    core_engine.add_project("Old Name")
    core_engine.update_project("Old Name", "New Name")
    projects = core_engine.get_projects()
    assert "Old Name" not in projects
    assert "New Name" in projects

def test_update_nonexistent_project():
    with pytest.raises(ValueError, match="not found"):
        core_engine.update_project("Missing", "New Name")

def test_delete_project():
    core_engine.add_project("To Delete")
    core_engine.delete_project("To Delete")
    projects = core_engine.get_projects()
    assert "To Delete" not in projects

def test_delete_nonexistent_project():
    with pytest.raises(ValueError, match="not found"):
        core_engine.delete_project("Missing")

def test_set_and_get_default_project():
    core_engine.add_project("Proj A")
    core_engine.add_project("Proj B")
    core_engine.set_default_project("Proj B")
    assert core_engine.get_default_project() == "Proj B"
    
    core_engine.set_default_project("Proj A")
    assert core_engine.get_default_project() == "Proj A"
