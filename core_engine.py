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
        CREATE TABLE IF NOT EXISTS leaves (
            leave_date TEXT PRIMARY KEY,
            leave_type TEXT NOT NULL
        )
    ''')
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

def add_leave(date_str, leave_type="Personal Leave"):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('INSERT OR REPLACE INTO leaves (leave_date, leave_type) VALUES (?, ?)', (date_str, leave_type))
    conn.commit()
    conn.close()

def get_logged_hours_for_day(date_str):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT SUM(hours) FROM timesheet WHERE log_date = ?", (date_str,))
    result = cursor.fetchone()[0]
    conn.close()
    return result if result else 0

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

def add_timesheet_entry(project, activity, date_str, hours, description):
    if is_weekend(date_str): raise ValueError("Cannot log hours on a weekend.")
    if project not in VALID_PROJECTS: raise ValueError("Invalid Project selected.")
    if activity not in VALID_ACTIVITIES: raise ValueError("Invalid Activity selected.")
    if hours > MAX_HOURS_PER_ENTRY: raise ValueError(f"Max {MAX_HOURS_PER_ENTRY} hours per entry.")
    
    total_today = get_logged_hours_for_day(date_str)
    if total_today + hours > MAX_HOURS_PER_DAY:
        raise ValueError(f"Exceeds {MAX_HOURS_PER_DAY} hr daily limit. Remaining: {MAX_HOURS_PER_DAY - total_today} hrs.")

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO timesheet (project, activity, log_date, hours, minutes, tag, description)
        VALUES (?, ?, ?, ?, 0, '', ?)
    ''', (project, activity, date_str, hours, description))
    conn.commit()
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