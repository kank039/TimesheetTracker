⏰ Timesheet Tracker (v1.0)
A lightweight, intelligent, and strictly-compliant background utility designed to automate timesheet logging. Built specifically to eliminate the friction of daily time tracking while adhering strictly to HR portal constraints.

Author: Kshitij

Status: Production Ready (v1.0.0)

✨ Key Features
🧠 Smart Accumulator: Tracks unlogged hours automatically between 10:00 AM and 7:00 PM. If you miss a prompt due to an emergency, it gracefully groups your missing hours and asks you to allocate them later.

🛡️ Portal-Compliant Validation: Hardcoded rules prevent portal rejection.

Maximum 3 hours per row.

Maximum 8 hours per day.

Minutes are strictly locked to 0.

Weekends are automatically ignored.

🎉 Auto-Sleep ("Go Home" Detection): The moment you hit 8 logged hours for the day, the background daemon permanently sleeps for the rest of the day and displays a victory message.

🚨 Month-End Freeze Alerts: Automatically calculates the final working weekday of the month and triggers high-priority alerts at 10:00 AM and 7:00 PM to ensure you submit before the portal locks.

🌴 Out of Office (OOF) Mode: Use the Qt system tray menu to mark a day as "Personal Leave," completely silencing all reminders until tomorrow.

📊 One-Click Export: Generates a perfectly formatted Excel (.xlsx) file matching the rigid RAM - Project Transcend 2026 template and saves it next to the app automatically.

🖥️ PySide6 UI: Uses native Qt dialogs and tray integration to avoid the Windows Explorer save-window issues seen in the Tkinter build.

🛠️ Modes of Operation
Timesheet Tracker automatically adapts to your Windows environment and IT permissions.

1. Standard Mode (Default)
If you run the application normally, it will securely install its SQLite database (timesheet_brain.db) and configuration file into your hidden Windows %APPDATA%\TimesheetTracker folder. This keeps your Desktop/Downloads folder clean and protects your data even if you delete or move the .exe.

2. Portable Mode (USB Friendly)
If your IT environment blocks %APPDATA%, or if you want to carry the app on a USB drive, the tracker falls back to Portable Mode. It will generate the database in the exact same folder as the .exe.

Power User Tip: You can forcefully trigger Portable Mode by creating a file named tracker_config.ini next to the .exe containing the following lines:

Ini, TOML
[Settings]
Mode=Portable
🚀 Installation & Usage
Running the Executable
Download TimesheetTracker.exe.

Double-click to run. On the very first run, a popup will inform you whether it installed in Standard or Portable mode.

The application will silently add itself to your Windows Startup (via Registry or Task Scheduler) so you never have to launch it manually again.

Look for the Blue Square Icon in your System Tray (near the clock/Wi-Fi).

System Tray Controls
Right-click the tray icon to access:

Manual Log:
Log Today: Force a log entry for the current day (great for bulk-logging 8 hours at once).
Edit Old Day: Open a date-based editor for past entries so you can fix spelling mistakes or rebalance hours.

Old-day edits must still total 8 hours for the selected date, and no single row can exceed 3 hours.

Mark Today as Leave (OOF): Silence the tracker for sick days or PTO.

Export: Generate the Excel (.xlsx) file for the HR portal without clearing the database, so older entries remain available for future exports.

Quit: Force close the background daemon.

Building from Source
If you wish to modify the code and compile it yourself:

Bash
# 1. Install dependencies
pip install pandas PySide6 pyinstaller openpyxl

# 2. Compile to a single .exe
pyinstaller --noconsole --onefile --name "TimesheetTracker" --icon=app.ico --version-file=version.txt main.py
<!-- pyinstaller --noconsole --onefile --name "TimesheetTracker" main.py -->
📝 Changelog
[v1.2.0] - PySide6 UI Rewrite
Changed: Replaced Tkinter and pystray with a single PySide6 event loop.

Changed: Export now saves directly to the executable folder as an Excel file without opening the Windows Explorer save dialog.

Changed: Export success and error dialogs now use Qt-native modal windows.

[v1.0.0] - Production Release
Added: Multi-tier Windows startup persistence (Registry injection with Task Scheduler fallback).

Added: Dynamic routing for SQLite database (Standard %APPDATA% vs. Portable mode).

Added: tracker_config.ini generation for explicit environment overrides.

Added: First-run initialization popups to inform users of their installation mode.

[v0.9.0] - The "Quality of Life" Update
Added: System Tray integration for silent background operation.

Added: "Mark Today as Leave" feature to bypass the 8-hour expected daily quota.

Added: Custom GUI dialogs that force themselves to the front of the screen.

Changed: Background daemon now runs in the main Qt event loop to prevent UI locking.

[v0.8.0] - Smart Accumulation & Victory Conditions
Added: "Go Home" auto-sleep detection. The script stops prompting entirely once 8 hours are logged.

Added: Expected hours calculation based on a strict 10:00 AM to 7:00 PM operating window.

Added: Month-end portal freeze detection logic (get_last_working_day).

[v0.5.0] - Core Engine Implementation
Added: SQLite integration for robust local data storage.

Added: Strict data validation (Weekends blocked, Max 3 hours per row limit, forced 0 minutes).

Added: Automated CSV mapping to format dates rigidly to dd-MM-yyyy.

[v0.1.0] - Initial Prototype
Added: Basic CLI flow and Pandas export logic.

Added: Rigid column headers (Project, Activity, Date, etc.) mapped to the RAM - Project Transcend 2026 template.