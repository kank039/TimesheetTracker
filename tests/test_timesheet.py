import pytest
import core_engine

@pytest.fixture
def setup_data():
    core_engine.add_project("Test Project")
    yield "Test Project", "Dev-Bug Fixing"

def test_add_timesheet_entry_success(setup_data):
    project, activity = setup_data
    # Use a specific weekday date (e.g. 2026-06-05 is Friday)
    date_str = "2026-06-05"
    
    core_engine.add_timesheet_entry(project, activity, date_str, 2, 30, "Fixed a bug")
    entries = core_engine.get_timesheet_entries_for_day(date_str)
    
    assert len(entries) == 1
    assert entries[0]["project"] == project
    assert entries[0]["activity"] == activity
    assert entries[0]["hours"] == 2
    assert entries[0]["minutes"] == 30
    assert entries[0]["description"] == "Fixed a bug"
    
    # Check total hours
    assert core_engine.get_logged_hours_for_day(date_str) == 2.5
    assert core_engine.get_logged_minutes_for_day(date_str) == 150

def test_add_timesheet_entry_weekend(setup_data):
    project, activity = setup_data
    # 2026-06-06 is Saturday
    date_str = "2026-06-06"
    with pytest.raises(ValueError, match="Cannot log hours on a weekend"):
        core_engine.add_timesheet_entry(project, activity, date_str, 2, 0, "Weekend work")

def test_add_timesheet_entry_invalid_project(setup_data):
    _, activity = setup_data
    with pytest.raises(ValueError, match="Invalid Project"):
        core_engine.add_timesheet_entry("Bad Project", activity, "2026-06-05", 2, 0, "Desc")

def test_add_timesheet_entry_invalid_activity(setup_data):
    project, _ = setup_data
    with pytest.raises(ValueError, match="Invalid Activity"):
        core_engine.add_timesheet_entry(project, "Bad Activity", "2026-06-05", 2, 0, "Desc")

def test_add_timesheet_entry_exceeds_max_entry_hours(setup_data):
    project, activity = setup_data
    # max hours per entry is 3
    with pytest.raises(ValueError, match="Each entry hours must be between 0 and 3"):
        core_engine.add_timesheet_entry(project, activity, "2026-06-05", 4, 0, "Desc")

def test_add_timesheet_entry_exceeds_max_daily_hours(setup_data):
    project, activity = setup_data
    date_str = "2026-06-05"
    # Max daily hours is 8. Add 3, 3, 3 -> should fail on 3rd
    core_engine.add_timesheet_entry(project, activity, date_str, 3, 0, "Desc")
    core_engine.add_timesheet_entry(project, activity, date_str, 3, 0, "Desc")
    with pytest.raises(ValueError, match="Exceeds 8 hr daily limit"):
        core_engine.add_timesheet_entry(project, activity, date_str, 3, 0, "Desc")

def test_replace_timesheet_entries_for_day(setup_data):
    project, activity = setup_data
    date_str = "2026-06-05"
    
    entries = [
        {"project": project, "activity": activity, "hours": 3, "minutes": 0, "description": "T1"},
        {"project": project, "activity": activity, "hours": 3, "minutes": 0, "description": "T2"},
        {"project": project, "activity": activity, "hours": 2, "minutes": 0, "description": "T3"},
    ]
    
    core_engine.replace_timesheet_entries_for_day(date_str, entries)
    
    db_entries = core_engine.get_timesheet_entries_for_day(date_str)
    assert len(db_entries) == 3
    assert db_entries[0]["description"] == "T1"
    assert db_entries[2]["description"] == "T3"
    assert core_engine.get_logged_hours_for_day(date_str) == 8.0

def test_replace_timesheet_entries_wrong_total(setup_data):
    project, activity = setup_data
    date_str = "2026-06-05"
    
    entries = [
        {"project": project, "activity": activity, "hours": 3, "minutes": 0, "description": "T1"},
        {"project": project, "activity": activity, "hours": 3, "minutes": 0, "description": "T2"},
    ] # Total 6 hours
    
    with pytest.raises(ValueError, match="must total exactly 8 hours"):
        core_engine.replace_timesheet_entries_for_day(date_str, entries)
