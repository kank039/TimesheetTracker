import sqlite3
import datetime
import calendar
import os
import sys
import configparser

# ==========================================
# CONFIGURATION & DYNAMIC PATH ROUTING
# ==========================================
APP_NAME = "TimesheetTracker"
CONFIG_FILE = "tracker_config.ini"

if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

config = configparser.ConfigParser()
CURRENT_MODE = "Standard"
IS_FIRST_RUN = False

def determine_paths():
    global CURRENT_MODE, IS_FIRST_RUN
    
    local_config_path = os.path.join(base_dir, CONFIG_FILE)
    
    # 1. Manual Portable Override
    if os.path.exists(local_config_path):
        config.read(local_config_path)
        if config.has_section('Settings') and config.get('Settings', 'Mode', fallback='') == 'Portable':
            CURRENT_MODE = "Portable"
            return os.path.join(base_dir, "timesheet_brain.db")

    # 2. Try Standard Mode (%APPDATA%)
    if sys.platform == "win32":
        try:
            app_data_dir = os.path.join(os.environ["APPDATA"], APP_NAME)
            os.makedirs(app_data_dir, exist_ok=True)
            
            test_file = os.path.join(app_data_dir, '.write_test')
            with open(test_file, 'w') as f: f.write('test')
            os.remove(test_file)
            
            CURRENT_MODE = "Standard"
            appdata_config_path = os.path.join(app_data_dir, CONFIG_FILE)
            
            if not os.path.exists(appdata_config_path):
                IS_FIRST_RUN = True
                config['Settings'] = {'Mode': 'Standard'}
                with open(appdata_config_path, 'w') as configfile:
                    config.write(configfile)

            return os.path.join(app_data_dir, "timesheet_brain.db")
        except (PermissionError, OSError):
            CURRENT_MODE = "Portable"
    else:
        CURRENT_MODE = "Portable"

    # 3. Forced Portable Mode Fallback
    if not os.path.exists(local_config_path):
        IS_FIRST_RUN = True
        try:
            config['Settings'] = {'Mode': 'Portable'}
            with open(local_config_path, 'w') as configfile:
                config.write(configfile)
        except PermissionError:
            pass 

    return os.path.join(base_dir, "timesheet_brain.db")

DB_NAME = determine_paths()

# Strict HR Parameters
VALID_PROJECTS = ["RAM - Project Transcend 2026"]
VALID_ACTIVITIES = [#["Dev-Client Interaction", "Dev-Bug Fixing", "Dev-R & D"]
    "Dev-Bug Fixing",
    "Dev-Change Request",
    "Dev-Client Interaction",
    "Dev-Client Visit",
    "Dev-Code review",
    "Dev-Coding and Implementation",
    "Dev-Daily Standup Meeting",
    "Dev-Database Design",
    "Dev-DB Administration",
    "Dev-Documentation",
    "Dev-Efforts Estimation",
    "Dev-Guiding Colleague",
    "Dev-Investigation issue/scenario",
    "Dev-Others",
    "Dev-Performance Tuning",
    "Dev-R & D",
    "Dev-Reports Design and Devlopment",
    "Dev-Requirement understanding/analysis",
    "Dev-Retrospective Meeting",
    "Dev-Scrum Meetings",
    "Dev-Sprint Planning Meetings",
    "Dev-Story Grooming",
    "Dev-Support",
    "Dev-System and Architecture Design",
    "Dev-Team Discussion",
    "Dev-Team Review",
    "Dev-Technical Discussion",
    "Dev-Training",
    ]
MAX_HOURS_PER_DAY = 8
MAX_HOURS_PER_ENTRY = 3

# Company-wide holidays (same for everyone)
COMPANY_HOLIDAYS = [
    ("2026-01-01", "New Year's Day"),
    ("2026-01-26", "Republic Day"),
    ("2026-03-17", "Holi"),
    ("2026-04-02", "Ram Navami"),
    ("2026-04-03", "Good Friday"),
    ("2026-05-01", "May Day"),
    ("2026-06-26", "Eid-ul-Adha (Bakrid)"),
    ("2026-08-15", "Independence Day"),
    ("2026-08-25", "Janmashtami"),
    ("2026-10-02", "Gandhi Jayanti"),
    ("2026-10-20", "Dussehra"),
    ("2026-11-09", "Diwali"),
    ("2026-11-10", "Diwali (Day 2)"),
    ("2026-12-25", "Christmas"),
]

# ==========================================
# DATABASE & LOGIC
# ==========================================

def setup_database():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS timesheet (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project TEXT NOT NULL,
            activity TEXT NOT NULL,
            log_date TEXT NOT NULL,
            hours INTEGER NOT NULL,
            minutes INTEGER DEFAULT 0,
            tag TEXT,
            description TEXT NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS recent_activities (
            activity TEXT PRIMARY KEY,
            last_used_at INTEGER NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS leaves (
            leave_date TEXT PRIMARY KEY,
            leave_type TEXT NOT NULL,
            leave_name TEXT DEFAULT ''
        )
    ''')
    # Add leave_name column if upgrading from an older schema
    try:
        cursor.execute("ALTER TABLE leaves ADD COLUMN leave_name TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass  # column already exists
    conn.commit()
    conn.close()
    seed_company_holidays()


def seed_company_holidays():
    """Insert company holidays into the leaves table (idempotent)."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    for date_str, name in COMPANY_HOLIDAYS:
        cursor.execute(
            'INSERT OR IGNORE INTO leaves (leave_date, leave_type, leave_name) VALUES (?, ?, ?)',
            (date_str, "Company Holiday", name),
        )
    conn.commit()
    conn.close()

def is_weekend(date_str):
    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.weekday() >= 5

def is_leave_or_holiday(date_str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT leave_type FROM leaves WHERE leave_date = ?", (date_str,))
    result = cursor.fetchone()
    conn.close()
    return result is not None

def add_leave(date_str, leave_type="Personal Leave", leave_name=""):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        'INSERT OR REPLACE INTO leaves (leave_date, leave_type, leave_name) VALUES (?, ?, ?)',
        (date_str, leave_type, leave_name),
    )
    conn.commit()
    conn.close()


def remove_leave(date_str):
    """Remove a personal leave. Refuses to delete company holidays."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT leave_type FROM leaves WHERE leave_date = ?", (date_str,))
    row = cursor.fetchone()
    if row and row[0] == "Company Holiday":
        conn.close()
        raise ValueError("Cannot remove a company holiday.")
    cursor.execute("DELETE FROM leaves WHERE leave_date = ?", (date_str,))
    conn.commit()
    conn.close()


def get_all_leaves(year=None):
    """Return all leaves, optionally filtered by year."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    if year:
        prefix = f"{year}-"
        cursor.execute(
            "SELECT leave_date, leave_type, leave_name FROM leaves WHERE leave_date LIKE ? ORDER BY leave_date ASC",
            (prefix + "%",),
        )
    else:
        cursor.execute("SELECT leave_date, leave_type, leave_name FROM leaves ORDER BY leave_date ASC")
    results = [
        {"date": row[0], "type": row[1], "name": row[2] or ""}
        for row in cursor.fetchall()
    ]
    conn.close()
    return results


def get_company_holidays():
    """Return the hardcoded company holidays list."""
    return list(COMPANY_HOLIDAYS)

def record_recent_activity(activity):
    if activity not in VALID_ACTIVITIES:
        return

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        '''
        INSERT OR REPLACE INTO recent_activities (activity, last_used_at)
        VALUES (?, ?)
        ''',
        (activity, int(datetime.datetime.now().timestamp())),
    )
    conn.commit()
    conn.close()

def get_recent_activities(limit=5):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        '''
        SELECT activity
        FROM recent_activities
        ORDER BY last_used_at DESC
        LIMIT ?
        ''',
        (limit,),
    )
    activities = [row[0] for row in cursor.fetchall()]
    conn.close()
    return activities

def get_logged_hours_for_day(date_str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT SUM(hours * 60 + minutes) FROM timesheet WHERE log_date = ?", (date_str,))
    result = cursor.fetchone()[0]
    conn.close()
    total_minutes = result if result else 0
    return total_minutes / 60.0


def get_logged_minutes_for_day(date_str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT SUM(hours * 60 + minutes) FROM timesheet WHERE log_date = ?", (date_str,))
    result = cursor.fetchone()[0]
    conn.close()
    return int(result) if result else 0

def get_remaining_hours_for_day(date_str):
    return max(0, MAX_HOURS_PER_DAY - get_logged_hours_for_day(date_str))


def get_timesheet_entries_for_day(date_str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        '''
        SELECT id, project, activity, log_date, hours, minutes, description
        FROM timesheet
        WHERE log_date = ?
        ORDER BY id ASC
        ''',
        (date_str,),
    )
    entries = [
        {
            "id": row[0],
            "project": row[1],
            "activity": row[2],
            "log_date": row[3],
            "hours": row[4],
            "minutes": row[5],
            "description": row[6],
        }
        for row in cursor.fetchall()
    ]
    conn.close()
    return entries

def get_unlogged_hours(date_str, current_time=None):
    if is_weekend(date_str) or is_leave_or_holiday(date_str):
        return 0

    if current_time is None:
        current_time = datetime.datetime.now()
    
    expected_hours = 0
    if current_time.hour >= 19: 
        expected_hours = 8
    elif current_time.hour > 10:
        expected_hours = min(current_time.hour - 10, MAX_HOURS_PER_DAY)

    logged_hours = get_logged_hours_for_day(date_str)
    return max(0, expected_hours - logged_hours)

def add_timesheet_entry(project, activity, date_str, hours, minutes=0, description=""):
    if is_weekend(date_str):
        raise ValueError("Cannot log hours on a weekend.")
    if project not in VALID_PROJECTS:
        raise ValueError("Invalid Project selected.")
    if activity not in VALID_ACTIVITIES:
        raise ValueError("Invalid Activity selected.")
    if minutes < 0 or minutes > 59:
        raise ValueError("Minutes must be between 0 and 59.")
    if hours < 0 or hours > MAX_HOURS_PER_ENTRY:
        raise ValueError(f"Each entry hours must be between 0 and {MAX_HOURS_PER_ENTRY}.")
    if hours == 0 and minutes == 0:
        raise ValueError("Cannot add an empty time entry.")
    if hours == MAX_HOURS_PER_ENTRY and minutes > 0:
        raise ValueError(f"Max {MAX_HOURS_PER_ENTRY} hours per entry.")

    total_today_minutes = get_logged_minutes_for_day(date_str)
    added_minutes = int(hours) * 60 + int(minutes)
    if total_today_minutes + added_minutes > MAX_HOURS_PER_DAY * 60:
        remaining_min = MAX_HOURS_PER_DAY * 60 - total_today_minutes
        rem_h = remaining_min // 60
        rem_m = remaining_min % 60
        raise ValueError(f"Exceeds {MAX_HOURS_PER_DAY} hr daily limit. Remaining: {rem_h}h {rem_m}m.")

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO timesheet (project, activity, log_date, hours, minutes, tag, description)
        VALUES (?, ?, ?, ?, ?, '', ?)
    ''', (project, activity, date_str, int(hours), int(minutes), description))
    conn.commit()
    conn.close()


def replace_timesheet_entries_for_day(date_str, entries):
    if is_weekend(date_str):
        raise ValueError("Cannot log hours on a weekend.")

    if not entries:
        raise ValueError("Add at least one task for the day.")

    total_minutes = 0
    normalized_entries = []

    for entry in entries:
        project = entry.get("project", "")
        activity = entry.get("activity", "")
        hours = int(entry.get("hours", 0))
        minutes = int(entry.get("minutes", 0))
        description = entry.get("description", "").strip()

        if project not in VALID_PROJECTS:
            raise ValueError("Invalid Project selected.")
        if activity not in VALID_ACTIVITIES:
            raise ValueError("Invalid Activity selected.")
        if hours < 0:
            raise ValueError("Each record must have non-negative hours.")
        if minutes < 0 or minutes > 59:
            raise ValueError("Minutes must be between 0 and 59.")
        if hours == 0 and minutes == 0:
            raise ValueError("Each record must have a positive time value.")
        if hours > MAX_HOURS_PER_ENTRY or (hours == MAX_HOURS_PER_ENTRY and minutes > 0):
            raise ValueError(f"Max {MAX_HOURS_PER_ENTRY} hours per entry.")
        if not description:
            raise ValueError("Description cannot be empty.")

        total_minutes += hours * 60 + minutes
        normalized_entries.append((project, activity, hours, minutes, description))

    if total_minutes != MAX_HOURS_PER_DAY * 60:
        raise ValueError(f"Entries for the day must total exactly {MAX_HOURS_PER_DAY} hours.")

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM timesheet WHERE log_date = ?", (date_str,))
        for project, activity, hours, minutes, description in normalized_entries:
            cursor.execute(
                '''
                INSERT INTO timesheet (project, activity, log_date, hours, minutes, tag, description)
                VALUES (?, ?, ?, ?, ?, '', ?)
                ''',
                (project, activity, date_str, hours, minutes, description),
            )
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

def get_last_working_day(year, month):
    last_day = calendar.monthrange(year, month)[1]
    last_date = datetime.date(year, month, last_day)
    if last_date.weekday() == 6: last_date -= datetime.timedelta(days=2)
    elif last_date.weekday() == 5: last_date -= datetime.timedelta(days=1)
    return last_date.strftime("%Y-%m-%d")

def is_month_end_freeze(date_str):
    date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    return date_str == get_last_working_day(date_obj.year, date_obj.month)