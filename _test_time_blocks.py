"""Quick test: auto-insert on a clean future date."""
from core_engine import (
    setup_database, get_projects, get_time_blocks, add_time_block,
    delete_time_block, insert_time_blocks_for_day, get_timesheet_entries_for_day,
)

setup_database()

# Clean up
for b in get_time_blocks():
    delete_time_block(b["id"])

# Add a block
add_time_block("SCRUM Test", get_projects()[0], "Dev-Scrum Meetings", 0, 30, "Test standup", "0,1,2,3,4")

# Insert for a clean future weekday (2026-06-08 is a Monday)
n = insert_time_blocks_for_day("2026-06-08")
print(f"Inserted: {n}")
entries = get_timesheet_entries_for_day("2026-06-08")
print(f"Entries for 2026-06-08: {len(entries)}")
for e in entries:
    print(f"  {e['activity']} - {e['hours']}h {e['minutes']}m - {e['description']}")

# Idempotency
n2 = insert_time_blocks_for_day("2026-06-08")
print(f"Re-insert: {n2} (should be 0)")

# Clean up
for b in get_time_blocks():
    delete_time_block(b["id"])

# Clean up test entries
import sqlite3
from core_engine import DB_NAME
conn = sqlite3.connect(DB_NAME)
conn.execute("DELETE FROM timesheet WHERE log_date = '2026-06-08'")
conn.execute("DELETE FROM time_block_insertions WHERE insert_date = '2026-06-08'")
conn.commit()
conn.close()

print("\nDone!")
