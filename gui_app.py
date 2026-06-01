import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pystray
from PIL import Image, ImageDraw
import threading
import time
import datetime
import pandas as pd
import sqlite3
import os
import sys
import winreg
import subprocess

from core_engine import (
    setup_database, get_unlogged_hours, get_logged_hours_for_day,
    add_timesheet_entry, is_month_end_freeze, add_leave,
    VALID_PROJECTS, VALID_ACTIVITIES, MAX_HOURS_PER_ENTRY, MAX_HOURS_PER_DAY,
    DB_NAME, CURRENT_MODE, IS_FIRST_RUN
)

# ==========================================
# WINDOWS STARTUP PERSISTENCE
# ==========================================
def setup_persistence():
    if sys.platform != 'win32': return
    
    app_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
    app_name = "TimesheetTracker"

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Run', 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
        winreg.CloseKey(key)
        return
    except Exception:
        pass

    try:
        cmd = f'schtasks /create /tn "{app_name}" /tr "{app_path}" /sc onlogon /f'
        subprocess.run(cmd, shell=True, capture_output=True, text=True)
    except Exception:
        pass

# ==========================================
# GUI FORMS & EXPORT
# ==========================================
def show_logging_form(unlogged_hours, today_str):
    root = tk.Tk()
    root.title("Timesheet Reminder")
    root.geometry("400x350")
    root.attributes('-topmost', True)
    root.resizable(False, False)

    ttk.Label(root, text=f"You have {unlogged_hours} unlogged hours.", font=("Arial", 12, "bold")).pack(pady=10)
    
    ttk.Label(root, text="Select Activity:").pack(anchor="w", padx=20)
    activity_var = tk.StringVar(value=VALID_ACTIVITIES[1])
    ttk.Combobox(root, textvariable=activity_var, values=VALID_ACTIVITIES, state="readonly", width=40).pack(pady=5)

    max_allowed = min(MAX_HOURS_PER_ENTRY, unlogged_hours)
    ttk.Label(root, text=f"Hours (Max {max_allowed}):").pack(anchor="w", padx=20)
    hours_var = tk.IntVar(value=max_allowed)
    ttk.Spinbox(root, from_=1, to=max_allowed, textvariable=hours_var, width=10).pack(pady=5, anchor="w", padx=20)

    ttk.Label(root, text="Description:").pack(anchor="w", padx=20)
    desc_text = tk.Text(root, height=4, width=42)
    desc_text.pack(pady=5)

    def submit_entry():
        act, hrs, desc = activity_var.get(), hours_var.get(), desc_text.get("1.0", tk.END).strip()
        if not desc:
            messagebox.showwarning("Warning", "Description cannot be empty.", parent=root)
            return
        try:
            add_timesheet_entry(VALID_PROJECTS[0], act, today_str, hrs, desc)
            total_now = get_logged_hours_for_day(today_str)
            if total_now >= MAX_HOURS_PER_DAY:
                messagebox.showinfo("Done for the day!", "🎉 Hurray you have worked enough go home now!", parent=root)
            else:
                messagebox.showinfo("Success", f"Logged {hrs} hours successfully! Total today: {total_now}/8", parent=root)
            root.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=root)

    ttk.Button(root, text="Submit Log", command=submit_entry).pack(pady=15)
    root.mainloop()

def export_gui_flow():
    if not os.path.exists(DB_NAME): return

    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT project, activity, log_date, hours, minutes, tag, description FROM timesheet ORDER BY log_date ASC", conn)
    conn.close()

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.update_idletasks()
    
    if df.empty:
        messagebox.showinfo("Export", "Timesheet database is currently empty.", parent=root)
        root.destroy()
        return

    df.columns = ['Project', 'Activity', 'Date(dd-MM-yyyy)', 'Time Spent(hh)', 'Time Spent (mm)', 'Tag', 'Description']
    df['Date(dd-MM-yyyy)'] = pd.to_datetime(df['Date(dd-MM-yyyy)']).dt.strftime('%d-%m-%Y')

    file_path = filedialog.asksaveasfilename(
        parent=root,
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Timesheet As",
        confirmoverwrite=True
    )
    if file_path:
        if not file_path.lower().endswith(".xlsx"):
            file_path = f"{file_path}.xlsx"
        df.to_excel(file_path, index=False)
        if messagebox.askyesno("Cleanup", "Export successful!\n\nDo you want to clear the old data (Fresh start)?", parent=root):
            conn = sqlite3.connect(DB_NAME)
            conn.execute("DELETE FROM timesheet")
            conn.commit()
            conn.close()
            messagebox.showinfo("Cleanup", "Database cleared for the next cycle.", parent=root)
    root.destroy()

# ==========================================
# BACKGROUND & SYSTEM TRAY
# ==========================================
def background_tracker():
    while True:
        now = datetime.datetime.now()
        today_str = now.strftime("%Y-%m-%d")

        if now.minute == 0 and 10 <= now.hour <= 19:
            unlogged = get_unlogged_hours(today_str)
            
            if is_month_end_freeze(today_str):
                if now.hour == 10:
                    threading.Thread(target=lambda: messagebox.showwarning("ALERT", "Portal freezes EOD today!")).start()
                elif now.hour == 19:
                    threading.Thread(target=export_gui_flow).start()
                    time.sleep(60) 
                    continue

            if unlogged > 0:
                show_logging_form(unlogged, today_str)
        time.sleep(60)

def create_image():
    image = Image.new('RGB', (64, 64), color=(0, 120, 215))
    draw = ImageDraw.Draw(image)
    draw.rectangle((16, 16, 48, 48), fill=(255, 255, 255))
    return image

def on_manual_log(icon, item):
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    unlogged = get_unlogged_hours(today_str)
    if unlogged > 0: show_logging_form(unlogged, today_str)
    else:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("All Good!", "You have 0 unlogged hours for today.", parent=root)
        root.destroy()

def on_mark_leave(icon, item):
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    root = tk.Tk()
    root.withdraw()
    if messagebox.askyesno("Out of Office", "Taking today off?\n\nThis will mute timesheet reminders for the rest of the day.", parent=root):
        add_leave(today_str)
        messagebox.showinfo("Rest Up!", "Today is marked as leave. The tracker is muted until tomorrow.", parent=root)
    root.destroy()

def on_export(icon, item): export_gui_flow()
def on_quit(icon, item): icon.stop(); os._exit(0)

def main():
    if IS_FIRST_RUN:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        msg = f"Welcome to Timesheet Tracker!\n\nYou are running in {CURRENT_MODE} MODE.\n\n"
        if CURRENT_MODE == "Portable":
            msg += "Database/Config are saved in the application folder. Ideal for USB drives."
            messagebox.showwarning("First Run Setup", msg, parent=root)
        else:
            msg += "Database is securely saved in your hidden Windows %APPDATA% folder."
            messagebox.showinfo("First Run Setup", msg, parent=root)
        root.destroy()

    setup_database()
    setup_persistence()
    threading.Thread(target=background_tracker, daemon=True).start()

    menu = pystray.Menu(
        pystray.MenuItem("Manual Log", on_manual_log),
        pystray.MenuItem("Mark Today as Leave (OOF)", on_mark_leave),
        pystray.MenuItem("Export & Clean", on_export),
        pystray.MenuItem("Quit", on_quit)
    )
    pystray.Icon("TimesheetTracker", create_image(), "Timesheet Tracker", menu).run()

if __name__ == "__main__":
    main()