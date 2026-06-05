# 🤝 Contributing to Timesheet Tracker

Welcome! This guide gets you from zero to a running development environment and explains how the codebase is structured so you can make changes confidently.

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Cloning & Environment Setup](#2-cloning--environment-setup)
3. [Running from Source](#3-running-from-source)
4. [Project Architecture](#4-project-architecture)
5. [Core Engine Reference](#5-core-engine-reference)
6. [GUI Architecture](#6-gui-architecture)
7. [Database & Schema Migrations](#7-database--schema-migrations)
8. [Adding a New Feature — Step-by-Step](#8-adding-a-new-feature--step-by-step)
9. [Manual Smoke Tests](#9-manual-smoke-tests)
10. [Building the Executable](#10-building-the-executable)
11. [Building the Installer](#11-building-the-installer)
12. [Style & Conventions](#12-style--conventions)

---

## 1. Prerequisites

| Tool | Version | Notes |
|---|---|---|
| **Python** | ≥ 3.14 | As declared in `pyproject.toml` and `.python-version` |
| **uv** | Latest | Replaces pip for dependency management |
| **Git** | Any | For cloning and branching |
| **Inno Setup** (optional) | 6.x | Only needed to rebuild the `.exe` installer |
| **PyInstaller** (optional) | ≥ 6.20 | Only needed to recompile the `.exe` |

> **Why uv?**  
> This project uses [uv](https://docs.astral.sh/uv/) for fast, reproducible installs. The `uv.lock` file pins every transitive dependency. You can use plain `pip` in a pinch but `uv sync` is the canonical workflow.

---

## 2. Cloning & Environment Setup

```bash
# Clone the repo
git clone <your-remote-url>
cd ts

# Install uv if you don't have it yet
pip install uv

# Create a virtual environment and install all dependencies from uv.lock
uv sync
```

This creates a `.venv/` folder and installs:
- `PyQt6` — Qt bindings for the GUI and tray
- `pandas` — DataFrame used for the Excel export
- `openpyxl` — Excel writer backend
- `pyinstaller` — compiles to `.exe`

---

## 3. Running from Source

```bash
# Start the app (uses the single-instance lock, same as the .exe)
uv run python main.py

# Or, skip the lock for quick iteration (bypasses single-instance guard)
uv run python gui_app.py
```

On first run the app will:
1. Auto-detect `%APPDATA%` write access and choose Standard or Portable mode.
2. Create the SQLite database and seed company holidays.
3. Show a first-run popup confirming the mode.

The tray icon appears in the Windows notification area. You can right-click it for the full menu or left-click it to open the Holidays/Leaves/Time Blocks panel.

> **Tip:** If you want a fresh database during development, delete `%APPDATA%\TimesheetTracker\timesheet_brain.db` (Standard mode) or `timesheet_brain.db` next to the script (Portable mode).

---

## 4. Project Architecture

```
main.py                  ← entry point
  └─ single_instance.py  ← QLockFile guard (one process at a time)
  └─ gui_app.py          ← all PyQt6 code + background daemon
       └─ core_engine.py ← all business logic, DB, validation helpers
```

The dependency flow is strictly **one-way**: `gui_app` imports from `core_engine`; `core_engine` has no GUI dependency.

### Thread model
There are **no threads**. Everything runs on the Qt main event loop:
- A `QTimer` fires every 60 seconds → `check_background_tasks()` → decides whether to show a prompt.
- All DB writes happen synchronously in the main thread.

---

## 5. Core Engine Reference

[`core_engine.py`](core_engine.py) is split into logical sections — each is marked with a header comment.

### Configuration & Path Routing (`determine_paths`)
Runs at module import time. Decides where the SQLite database lives:

| Priority | Condition | Result |
|---|---|---|
| 1 | `tracker_config.ini` exists next to the exe with `Mode=Portable` | Portable |
| 2 | `%APPDATA%` is writable | Standard |
| 3 | Fallback | Portable |

Sets the module-level globals `DB_NAME`, `CURRENT_MODE`, and `IS_FIRST_RUN`.

### HR Parameters
```python
MAX_HOURS_PER_DAY   = 8    # absolute daily ceiling
MAX_HOURS_PER_ENTRY = 3    # per-row ceiling (portal rule)
VALID_ACTIVITIES    = [...]  # exhaustive list; any entry not in this list is rejected
COMPANY_HOLIDAYS    = [...]  # hardcoded (date_str, name) tuples for 2026
```

### Key Functions

| Function | Description |
|---|---|
| `setup_database()` | Creates all tables if they don't exist; seeds holidays and default project; applies schema migrations. **Always idempotent.** |
| `add_timesheet_entry(...)` | Validates and inserts a single entry. Raises `ValueError` on any violation. |
| `replace_timesheet_entries_for_day(date_str, entries)` | Atomically replaces all entries for a day — used by the "Edit Old Day" dialog. Enforces 8 h total. |
| `insert_time_blocks_for_day(date_str)` | Inserts enabled time blocks as real timesheet rows. Idempotent (checks `time_block_insertions`). Skips weekends, leaves, and days where the block's `days_of_week` doesn't match. |
| `get_unlogged_hours(date_str)` | Returns hours that should have been logged based on the 10 AM – 7 PM window minus what's already logged. |
| `get_last_working_day(year, month)` | Returns the last Monday–Friday of the month as a `YYYY-MM-DD` string. |
| `is_month_end_freeze(date_str)` | Returns `True` if today is the last working day of the month. |
| `add_project / update_project / delete_project` | CRUD for projects. `delete_project` refuses if entries reference it; `update_project` cascades to `timesheet`. |
| `add_time_block / update_time_block / delete_time_block / toggle_time_block` | CRUD for recurring time blocks. `add_time_block` checks that the new total blocked minutes don't exceed 8 h. |

---

## 6. GUI Architecture

[`gui_app.py`](gui_app.py) contains the following major components:

### Helper functions (module-level)
| Function | Description |
|---|---|
| `setup_persistence()` | Writes the Registry run key (or Task Scheduler task) so the app auto-starts on login. |
| `export_timesheet(parent, prompt_for_path)` | Reads the DB, reshapes with pandas, and saves as `.xlsx`. |
| `show_box / ask_yes_no` | Thin wrappers over `QMessageBox` used throughout. |

### Dialogs
| Class | Purpose |
|---|---|
| `ManualLogDialog` | The primary "log time" dialog. Auto-splits large entries into ≤ 3 h chunks. Recent activities are shown at the top of the activity dropdown. |
| `OldDayEditorDialog` | Date-picker + dynamic row list for re-editing any past day. Uses `DayEntryRow` widgets. Validates total == 8 h before enabling Save. |
| `DayEntryRow` | Reusable row widget (project, activity, hours, minutes, description, remove button). |
| `ProjectManagerDialog` | Add / rename / delete projects. Rename uses an inline `QLineEdit` that appears on click. |
| `HolidayManagerWindow` | Three-panel window (company holidays, personal leaves, time blocks). Opened by left-clicking the tray icon. |

### `TimesheetController` (the "backend")
Owns the `QApplication`, `QSystemTrayIcon`, and `QTimer`. Responsible for:
- Initialising the tray icon and context menu.
- Running `check_background_tasks()` every minute.
- Calling `insert_time_blocks_for_day()` both on startup and in the timer loop.
- Dispatching first-run notices.

---

## 7. Database & Schema Migrations

The `setup_database()` function uses **non-destructive `ALTER TABLE`** guards to add new columns to existing databases:

```python
try:
    cursor.execute("ALTER TABLE leaves ADD COLUMN leave_name TEXT DEFAULT ''")
except sqlite3.OperationalError:
    pass  # column already exists — safe to ignore
```

**When you add a new column to an existing table:**
1. Add it to the `CREATE TABLE` statement (for fresh installs).
2. Add a matching `ALTER TABLE ... ADD COLUMN` guard below the `conn.commit()` in `setup_database()`.
3. Never `DROP` or rename a column — just add new ones.

---

## 8. Adding a New Feature — Step-by-Step

This example walks through adding a new feature end-to-end.

### Example: Add a "Notes" field to time entries

**Step 1 — Core engine: DB schema**

In `setup_database()`, add the column to the `CREATE TABLE`:
```python
# inside the CREATE TABLE IF NOT EXISTS timesheet (...) block
notes TEXT DEFAULT ''
```
And add the migration guard:
```python
try:
    cursor.execute("ALTER TABLE timesheet ADD COLUMN notes TEXT DEFAULT ''")
except sqlite3.OperationalError:
    pass
```

**Step 2 — Core engine: logic**

Update `add_timesheet_entry` and `replace_timesheet_entries_for_day` to accept and store `notes`.

**Step 3 — GUI: expose the field**

Add a `QLineEdit` for "Notes" in `ManualLogDialog` and `DayEntryRow`. Wire it up the same way `description` is wired.

**Step 4 — Test manually**

```bash
uv run python _test_seed.py   # confirm DB setup still works
uv run python main.py         # exercise the UI
```

**Step 5 — Export**

Update `export_timesheet()` to include the new column in the `SELECT` and in `df.columns`.

---

## 9. Manual Smoke Tests

There are two standalone test scripts you can run without any test framework:

```bash
# Test leave / holiday seeding and CRUD
uv run python _test_seed.py

# Test time block insertion logic
uv run python _test_time_blocks.py
```

Both scripts print pass/fail lines to stdout. They use the real database (whichever mode is active), so they're integration tests rather than unit tests.

> **Heads up:** These scripts do real writes. Run them against a dev copy of the database, or clean up afterwards.

---

## 10. Building the Executable

The project ships a `TimesheetTracker.spec` for reproducible PyInstaller builds.

```bash
# Compile to dist/TimesheetTracker/TimesheetTracker.exe
uv run pyinstaller TimesheetTracker.spec
```

The spec file already includes:
- `--noconsole` — no terminal window
- `--onedir` — outputs a folder (better cold-start performance than `--onefile`)
- `--icon=app.ico`
- `--version-file=version.txt`

To update the app version, edit `version.txt` and bump the version number in `TimesheetTracker_Setup.iss`.

---

## 11. Building the Installer

The installer is built with [Inno Setup](https://jrsoftware.org/isinfo.php) (free, Windows only).

1. Install Inno Setup 6.x.
2. Update the `OutputDir` and `Source` paths in `TimesheetTracker_Setup.iss` if you've checked out to a different directory.
3. Compile:
   ```bash
   iscc TimesheetTracker_Setup.iss
   ```
   This outputs `TimesheetTracker_Setup.exe` to the project root.

The installer:
- Copies all files from `dist\TimesheetTracker\` to `Program Files\TimesheetTracker\`.
- Requires only user-level privileges (`PrivilegesRequired=lowest`).
- Creates a Start Menu shortcut.
- Optionally launches the app immediately after install.

---

## 12. Style & Conventions

| Convention | Rule |
|---|---|
| **Validation** | All business-rule validation lives in `core_engine.py`. GUI code must not duplicate validation logic — call engine functions and catch `ValueError`. |
| **DB access** | Every function opens and closes its own connection. No shared connection objects. |
| **Error handling** | Engine functions raise `ValueError` with human-readable messages. GUI catches them and shows a `QMessageBox`. |
| **Idempotency** | Operations that run automatically (block insertion, holiday seeding) must be safe to call multiple times. Use `INSERT OR IGNORE` and `INSERT OR REPLACE` appropriately. |
| **Qt style** | UI theming is done via `setStyleSheet` with `QSS` inline strings defined as class-level `STYLE` constants. Avoid ad-hoc inline styles on individual widgets. |
| **No threads** | Keep everything on the main Qt event loop. If you need background work, use `QTimer.singleShot`. |
| **Imports** | `core_engine` → no GUI imports. `gui_app` → imports from `core_engine` only. `main.py` → imports from `gui_app` only. |
| **Type hints** | Not currently used. Don't add them unless you're prepared to add them consistently throughout. |

---

*Happy logging! If something's unclear, open an issue or check the inline comments in `core_engine.py`.*
