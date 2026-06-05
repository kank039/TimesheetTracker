# ⏰ Timesheet Tracker

> A lightweight, intelligent, and strictly-compliant background utility that automates daily timesheet logging — eliminating friction while adhering to HR portal constraints.

**Author:** Kshitij  
**Status:** Production Ready (`v4.0`)

---

## ✨ Key Features

| Feature | Description |
|---|---|
| 🧠 **Smart Accumulator** | Tracks unlogged hours between 10 AM – 7 PM. Misses a prompt? It gracefully groups accumulated hours and asks you to allocate them later. |
| 🛡️ **Portal-Compliant Validation** | Hardcoded rules prevent portal rejection: max 3 h per row, max 8 h per day, minutes are enforced. |
| 🎉 **Auto-Sleep ("Go Home" Detection)** | Once 8 hours are logged, the daemon permanently sleeps for the rest of the day and shows a victory message. |
| 🚨 **Month-End Freeze Alerts** | Calculates the last working weekday of the month and fires high-priority alerts at 10 AM and auto-exports at 6 PM. |
| 🌴 **Out-of-Office (OOF) Mode** | Mark any day as leave from the tray menu — silences all reminders until tomorrow. |
| 📅 **Holiday & Leave Manager** | Full UI panel showing company holidays, personal leaves, and time blocks. Open with a single left-click on the tray icon. |
| 📦 **Project Manager** | Add, rename, and delete projects directly from the tray. Renaming cascades to all historical entries. |
| ⏱️ **Time Blocks** | Define recurring auto-entries (e.g. "Daily SCRUM — 30 min, Mon–Fri"). Automatically inserted as real timesheet rows each qualifying day — idempotent and respects weekends/holidays. |
| 📊 **One-Click Export** | Generates a perfectly-formatted Excel (`.xlsx`) matching the RAM – Project Transcend 2026 template. Choose where to save via a Qt file dialog. |
| 🔒 **Single-Instance Guard** | A `QLockFile` prevents multiple copies of the app from running simultaneously. |
| 🖥️ **PyQt6 UI** | Native Qt dialogs and tray integration — no Tkinter flicker, no Windows Explorer save-dialog bugs. |

---

## 🛠️ Modes of Operation

The app automatically adapts to your Windows environment and IT permissions.

### 1. Standard Mode *(default)*
Data is stored in `%APPDATA%\TimesheetTracker\` — clean Desktop/Downloads, protected data even if you move the `.exe`.

### 2. Portable Mode *(USB-friendly)*
If `%APPDATA%` is blocked by IT, or you drop a `tracker_config.ini` next to the `.exe`, the app switches to Portable Mode and stores everything beside the executable.

**Force Portable Mode** by creating `tracker_config.ini`:
```ini
[Settings]
Mode=Portable
```

---

## 🚀 Installation & Usage

### Running the Installer
1. Download and run `TimesheetTracker_Setup.exe`.
2. On first launch a popup tells you whether Standard or Portable mode is active.
3. The app registers itself in Windows Startup (Registry → Task Scheduler fallback) so it launches on every login.
4. Look for the **Blue Square icon** in the system tray (near the clock/Wi-Fi).

### System Tray Controls

Right-click the tray icon to access:

| Menu Item | What it does |
|---|---|
| **Manual Log → Log Today** | Force-open the log dialog for the current day. Time blocks are pre-applied before calculating remaining time. |
| **Manual Log → Edit Old Day** | Date-picker editor for past entries. Must still total 8 h; no single row may exceed 3 h. |
| **Mark Today as Leave (OOF)** | Marks today as Personal Leave — silences all reminders until tomorrow. |
| **Manage Projects** | Add / rename / delete projects. A project with existing entries cannot be deleted. |
| **Export** | Opens a save-file dialog and writes the Excel export. Does **not** clear the database. |
| **Quit** | Force-closes the background daemon and hides the tray icon. |

Left-click the tray icon to open the **Holidays, Leaves & Time Blocks** panel.

---

## ⏱️ Time Blocks

Time blocks are recurring timesheet entries that are automatically inserted at the start of each qualifying day.

- Each block has: **Name**, **Project**, **Activity**, **Duration** (h / m), **Description**, and **Days of Week** (Mon–Fri checkboxes).
- Blocks are only inserted on working days (weekdays that are not leaves or company holidays).
- A block is skipped if today's total logged time would exceed 8 h.
- Insertions are tracked in `time_block_insertions` — re-running is idempotent.
- You can **enable / disable** a block without deleting it, using the toggle button.

---

## 🗃️ Database Schema

The SQLite database (`timesheet_brain.db`) lives in `%APPDATA%\TimesheetTracker\` (Standard) or beside the `.exe` (Portable).

| Table | Purpose |
|---|---|
| `timesheet` | All logged time entries |
| `recent_activities` | Last-used activities for smart ordering in the UI |
| `leaves` | Personal leaves **and** company holidays (`leave_type` differentiates them) |
| `projects` | User-managed project list |
| `time_blocks` | Recurring block definitions |
| `time_block_insertions` | Idempotency guard — records which blocks were already inserted on which dates |

---

## 📁 File Structure

```
TimesheetTracker/
├── main.py                     # Entry point — acquires single-instance lock, launches GUI
├── single_instance.py          # QLockFile-based single-instance guard
├── core_engine.py              # All business logic, DB, validation, and date helpers
├── gui_app.py                  # PyQt6 UI: tray, dialogs, and the background daemon loop
├── app.ico                     # Application icon
├── version.txt                 # PyInstaller version resource file
├── pyproject.toml              # uv / pip project manifest with dependencies
├── TimesheetTracker.spec       # PyInstaller spec file for reproducible builds
├── TimesheetTracker_Setup.iss  # Inno Setup installer script
└── tests/                      # Automated pytests suite
```

---

## 🔧 Building from Source

> See [CONTRIBUTING.md](CONTRIBUTING.md) for the full developer setup guide.

**Quick build:**

```bash
# 1. Install dependencies (using uv — recommended)
uv sync

# 2. Run from source
uv run python main.py

# 3. Compile to a single .exe
uv run pyinstaller TimesheetTracker.spec

uv run pyinstaller --noconsole --onefile --name "TimesheetTracker" --icon=app.ico --version-file=version.txt main.py

# 4. Package the installer (requires Inno Setup installed)
iscc TimesheetTracker_Setup.iss
```

---

## 📝 Changelog

### [v4.0] — Project Manager & Time Block Day Selector
- **Added:** `ProjectManagerDialog` — full CRUD UI for projects (add, rename, delete) accessible from the tray.
- **Added:** `days_of_week` field on time blocks — per-block day-of-week checkboxes (Mon–Fri).
- **Added:** `update_time_block` engine function for editing existing blocks.
- **Changed:** Time block insertion now respects the block's `days_of_week` — a SCRUM block set to Mon–Fri won't fire on a Wednesday if it's not in the list.
- **Fixed:** Schema migration — PRAGMA checks safely add new columns to existing databases without hiding true SQL errors.

### [v1.2.0] — PyQt6 UI Rewrite
- **Changed:** Replaced Tkinter and pystray with a single PyQt6 event loop.
- **Changed:** Export now opens a Qt save dialog so you can choose the destination.
- **Changed:** Export success and error dialogs use Qt-native modal windows.

### [v1.0.0] — Production Release
- **Added:** Multi-tier Windows startup persistence (Registry → Task Scheduler fallback).
- **Added:** Dynamic path routing for SQLite (Standard `%APPDATA%` vs. Portable mode).
- **Added:** `tracker_config.ini` generation for explicit environment overrides.
- **Added:** First-run initialization popups to inform users of their installation mode.

### [v0.9.0] — Quality of Life Update
- **Added:** System tray integration for silent background operation.
- **Added:** "Mark Today as Leave" feature to bypass the 8-hour daily quota.
- **Added:** Custom Qt dialogs that force themselves to the front of the screen.
- **Changed:** Background daemon now runs in the main Qt event loop to prevent UI locking.

### [v0.8.0] — Smart Accumulation & Victory Conditions
- **Added:** "Go Home" auto-sleep detection — stops prompting once 8 hours are logged.
- **Added:** Expected hours calculation based on a strict 10 AM – 7 PM operating window.
- **Added:** Month-end portal freeze detection (`get_last_working_day`).

### [v0.5.0] — Core Engine
- **Added:** SQLite integration for robust local data storage.
- **Added:** Strict data validation (weekends blocked, max 3 h/row, forced 0 minutes).
- **Added:** Automated CSV mapping with dates formatted as `dd-MM-yyyy`.

### [v0.1.0] — Initial Prototype
- **Added:** Basic CLI flow and Pandas export logic.
- **Added:** Rigid column headers (`Project`, `Activity`, `Date`, etc.) mapped to the RAM – Project Transcend 2026 template.