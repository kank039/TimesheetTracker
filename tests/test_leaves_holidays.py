import pytest
import core_engine

def test_seed_company_holidays():
    # setup_database from conftest already seeds holidays
    leaves = core_engine.get_all_leaves()
    # It should have the company holidays
    assert len(leaves) > 0
    # verify it has some expected holiday
    has_new_year = any(l["date"] == "2026-01-01" and l["type"] == "Company Holiday" for l in leaves)
    assert has_new_year

def test_add_and_remove_personal_leave():
    core_engine.add_leave("2026-07-15", "Personal Leave", "Doctor appointment")
    leaves = core_engine.get_all_leaves(2026)
    personal = [l for l in leaves if l["date"] == "2026-07-15"]
    assert len(personal) == 1
    assert personal[0]["type"] == "Personal Leave"
    assert personal[0]["name"] == "Doctor appointment"
    
    # Now remove it
    core_engine.remove_leave("2026-07-15")
    leaves_after = core_engine.get_all_leaves(2026)
    personal_after = [l for l in leaves_after if l["date"] == "2026-07-15"]
    assert len(personal_after) == 0

def test_remove_company_holiday_fails():
    with pytest.raises(ValueError, match="Cannot remove a company holiday"):
        core_engine.remove_leave("2026-01-26")

def test_is_leave_or_holiday():
    assert core_engine.is_leave_or_holiday("2026-01-26") is True
    assert core_engine.is_leave_or_holiday("2026-07-15") is False
    
    core_engine.add_leave("2026-07-15")
    assert core_engine.is_leave_or_holiday("2026-07-15") is True
