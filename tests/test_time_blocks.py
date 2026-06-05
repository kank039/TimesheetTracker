import pytest
import core_engine

@pytest.fixture
def setup_data():
    core_engine.add_project("TB Project")
    yield "TB Project", "Dev-Team Discussion"

def test_add_and_get_time_block(setup_data):
    project, activity = setup_data
    core_engine.add_time_block("Morning Standup", project, activity, 0, 30, "Daily sync", "0,1,2,3,4")
    
    blocks = core_engine.get_time_blocks()
    assert len(blocks) == 1
    assert blocks[0]["name"] == "Morning Standup"
    assert blocks[0]["hours"] == 0
    assert blocks[0]["minutes"] == 30
    assert blocks[0]["enabled"] is True

def test_add_invalid_time_block(setup_data):
    project, activity = setup_data
    with pytest.raises(ValueError, match="Block name cannot be empty"):
        core_engine.add_time_block("", project, activity, 1, 0, "Desc")
        
    with pytest.raises(ValueError, match="Invalid project"):
        core_engine.add_time_block("T", "Bad Proj", activity, 1, 0, "Desc")

    with pytest.raises(ValueError, match="Invalid activity"):
        core_engine.add_time_block("T", project, "Bad Act", 1, 0, "Desc")

def test_update_time_block(setup_data):
    project, activity = setup_data
    core_engine.add_time_block("Old Name", project, activity, 1, 0, "Desc", "0")
    blocks = core_engine.get_time_blocks()
    block_id = blocks[0]["id"]
    
    core_engine.update_time_block(block_id, "New Name", project, activity, 2, 0, "New Desc", "1,2")
    blocks2 = core_engine.get_time_blocks()
    assert blocks2[0]["name"] == "New Name"
    assert blocks2[0]["hours"] == 2
    assert blocks2[0]["days_of_week"] == "1,2"

def test_delete_time_block(setup_data):
    project, activity = setup_data
    core_engine.add_time_block("To Delete", project, activity, 1, 0, "Desc", "0")
    blocks = core_engine.get_time_blocks()
    
    core_engine.delete_time_block(blocks[0]["id"])
    assert len(core_engine.get_time_blocks()) == 0

def test_insert_time_blocks_for_day(setup_data):
    project, activity = setup_data
    # "0,1,2,3,4" means Monday to Friday
    core_engine.add_time_block("Daily Sync", project, activity, 1, 0, "Sync", "0,1,2,3,4")
    
    # 2026-06-05 is Friday
    date_str = "2026-06-05"
    inserted = core_engine.insert_time_blocks_for_day(date_str)
    assert inserted == 1
    
    entries = core_engine.get_timesheet_entries_for_day(date_str)
    assert len(entries) == 1
    assert entries[0]["description"] == "Sync"
    
    # Inserting again on same day should skip
    inserted2 = core_engine.insert_time_blocks_for_day(date_str)
    assert inserted2 == 0
    
    # Inserting on weekend should skip
    # 2026-06-06 is Saturday
    inserted3 = core_engine.insert_time_blocks_for_day("2026-06-06")
    assert inserted3 == 0
