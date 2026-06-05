import os
import tempfile
import sqlite3
import pytest
import core_engine

@pytest.fixture(autouse=True)
def setup_test_db(tmp_path):
    """
    Override the database path for tests so we don't modify the real database.
    This fixture runs automatically before every test.
    """
    test_db = tmp_path / "test_timesheet_brain.db"
    # We patch the core_engine.DB_NAME to point to our temp test database
    core_engine.DB_NAME = str(test_db)
    
    # Run the setup database logic
    core_engine.setup_database()
    
    yield
    
    # Teardown: no explicit cleanup needed for tmp_path, pytest handles it.
