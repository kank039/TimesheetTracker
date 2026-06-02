import datetime
import os
import sqlite3
import subprocess
import sys

import pandas as pd
import winreg
from PySide6.QtCore import QDate, QTimer, Qt
from PySide6.QtGui import QAction, QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QDateEdit,
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QFileDialog,
    QScrollArea,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QSystemTrayIcon,
    QVBoxLayout,
    QFrame,
    QWidget,
)

from core_engine import (
    DB_NAME,
    CURRENT_MODE,
    IS_FIRST_RUN,
    MAX_HOURS_PER_DAY,
    MAX_HOURS_PER_ENTRY,
    VALID_ACTIVITIES,
    VALID_PROJECTS,
    add_leave,
    add_timesheet_entry,
    get_timesheet_entries_for_day,
    get_logged_hours_for_day,
    get_recent_activities,
    get_remaining_hours_for_day,
    get_unlogged_hours,
    is_month_end_freeze,
    record_recent_activity,
    replace_timesheet_entries_for_day,
    setup_database,
)


def setup_persistence():
    if sys.platform != "win32":
        return

    app_path = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    app_name = "TimesheetTracker"

    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE,
        )
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


def get_export_folder():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def create_app_icon():
    pixmap = QPixmap(64, 64)
    pixmap.fill(QColor("#0078d7"))
    painter = QPainter(pixmap)
    painter.fillRect(16, 16, 32, 32, QColor("white"))
    painter.end()
    return QIcon(pixmap)


def show_box(parent, icon, title, text):
    box = QMessageBox(parent)
    box.setIcon(icon)
    box.setWindowTitle(title)
    box.setText(text)
    box.setStandardButtons(QMessageBox.StandardButton.Ok)
    box.exec()


def ask_yes_no(parent, title, text):
    result = QMessageBox.question(
        parent,
        title,
        text,
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        QMessageBox.StandardButton.No,
    )
    return result == QMessageBox.StandardButton.Yes


def export_timesheet(parent=None, prompt_for_path=True):
    if not os.path.exists(DB_NAME):
        show_box(parent, QMessageBox.Warning, "Export", "Timesheet database not found.")
        return None

    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query(
        "SELECT project, activity, log_date, hours, minutes, tag, description FROM timesheet ORDER BY log_date ASC",
        conn,
    )
    conn.close()

    if df.empty:
        show_box(parent, QMessageBox.Information, "Export", "Timesheet database is currently empty.")
        return None

    df.columns = [
        "Project",
        "Activity",
        "Date(dd-MM-yyyy)",
        "Time Spent(hh)",
        "Time Spent (mm)",
        "Tag",
        "Description",
    ]
    df["Date(dd-MM-yyyy)"] = pd.to_datetime(df["Date(dd-MM-yyyy)"]).dt.strftime("%d-%m-%Y")

    export_name = f"TimesheetTracker_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_path = os.path.join(get_export_folder(), export_name)

    if prompt_for_path:
        selected_path, _ = QFileDialog.getSaveFileName(
            parent,
            "Save Export As",
            export_path,
            "Excel Workbook (*.xlsx)",
        )
        if not selected_path:
            return None
        if not selected_path.lower().endswith(".xlsx"):
            selected_path += ".xlsx"
        export_path = selected_path

    try:
        df.to_excel(export_path, index=False)
        show_box(parent, QMessageBox.Information, "Export", f"Export successful!\n\nSaved to:\n{export_path}")
        return export_path
    except Exception as exc:
        show_box(parent, QMessageBox.Critical, "Export Failed", f"Could not export the file.\n\n{exc}")
        return None


class ManualLogDialog(QDialog):
    def __init__(self, remaining_hours, today_str, parent=None):
        super().__init__(parent)
        self.today_str = today_str
        self.remaining_hours = remaining_hours

        self.setWindowTitle("Manual Log")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)

        info_label = QLabel(f"You have {remaining_hours} hours remaining to log today.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        form = QFormLayout()

        self.activity_combo = QComboBox()
        self.populate_activity_combo()
        form.addRow("Activity:", self.activity_combo)

        self.hours_spin = QSpinBox()
        max_allowed = remaining_hours
        self.hours_spin.setRange(1, max_allowed)
        self.hours_spin.setValue(max_allowed)
        form.addRow("Hours:", self.hours_spin)

        self.desc_edit = QTextEdit()
        self.desc_edit.setPlaceholderText("Enter a short description")
        self.desc_edit.setFixedHeight(110)
        form.addRow("Description:", self.desc_edit)

        layout.addLayout(form)

        button_row = QHBoxLayout()
        button_row.addStretch(1)

        submit_button = QPushButton("Submit Log")
        submit_button.clicked.connect(self.submit_entry)
        button_row.addWidget(submit_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_row.addWidget(cancel_button)

        layout.addLayout(button_row)

    def populate_activity_combo(self):
        recent_activities = [activity for activity in get_recent_activities() if activity in VALID_ACTIVITIES]
        remaining_activities = [activity for activity in VALID_ACTIVITIES if activity not in recent_activities]

        if recent_activities:
            self.activity_combo.addItems(recent_activities)
            if remaining_activities:
                self.activity_combo.insertSeparator(self.activity_combo.count())

        self.activity_combo.addItems(remaining_activities)

        if self.activity_combo.count() > 0:
            self.activity_combo.setCurrentIndex(0)

    def submit_entry(self):
        activity = self.activity_combo.currentText().strip()
        hours = int(self.hours_spin.value())
        description = self.desc_edit.toPlainText().strip()

        if not description:
            show_box(self, QMessageBox.Warning, "Warning", "Description cannot be empty.")
            return

        try:
            remaining_hours = get_remaining_hours_for_day(self.today_str)
            if hours > remaining_hours:
                show_box(self, QMessageBox.Warning, "Warning", f"Only {remaining_hours} hours remain for today.")
                return

            hours_left = hours
            while hours_left > 0:
                chunk = min(MAX_HOURS_PER_ENTRY, hours_left)
                add_timesheet_entry(VALID_PROJECTS[0], activity, self.today_str, chunk, description)
                hours_left -= chunk

            record_recent_activity(activity)
            total_now = get_logged_hours_for_day(self.today_str)
            if total_now >= MAX_HOURS_PER_DAY:
                show_box(self, QMessageBox.Information, "Done for the day!", "You have worked enough for today.")
            else:
                show_box(self, QMessageBox.Information, "Success", f"Logged {hours} hours successfully! Total today: {total_now}/8")
            self.accept()
        except Exception as exc:
            show_box(self, QMessageBox.Critical, "Error", str(exc))


class DayEntryRow(QWidget):
    def __init__(self, entry=None, on_change=None, on_remove=None, parent=None):
        super().__init__(parent)
        self.on_change = on_change
        self.on_remove = on_remove

        layout = QGridLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setHorizontalSpacing(8)
        layout.setVerticalSpacing(6)

        self.activity_combo = QComboBox()
        self.activity_combo.addItems(VALID_ACTIVITIES)
        layout.addWidget(QLabel("Activity"), 0, 0)
        layout.addWidget(self.activity_combo, 1, 0)

        self.hours_spin = QSpinBox()
        self.hours_spin.setRange(1, MAX_HOURS_PER_ENTRY)
        layout.addWidget(QLabel("Hours"), 0, 1)
        layout.addWidget(self.hours_spin, 1, 1)

        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Task description")
        layout.addWidget(QLabel("Description"), 0, 2)
        layout.addWidget(self.description_edit, 1, 2)

        self.remove_button = QPushButton("Remove")
        self.remove_button.clicked.connect(self.handle_remove)
        layout.addWidget(self.remove_button, 1, 3)

        self.activity_combo.currentIndexChanged.connect(self.emit_change)
        self.hours_spin.valueChanged.connect(self.emit_change)
        self.description_edit.textChanged.connect(self.emit_change)

        if entry:
            self.set_entry(entry)
        else:
            self.hours_spin.setValue(1)

    def set_entry(self, entry):
        activity = entry.get("activity", "")
        hours = int(entry.get("hours", 1))
        description = entry.get("description", "")

        activity_index = self.activity_combo.findText(activity)
        if activity_index >= 0:
            self.activity_combo.setCurrentIndex(activity_index)
        self.hours_spin.setValue(max(1, min(MAX_HOURS_PER_ENTRY, hours)))
        self.description_edit.setText(description)

    def get_entry(self):
        return {
            "project": VALID_PROJECTS[0],
            "activity": self.activity_combo.currentText().strip(),
            "hours": int(self.hours_spin.value()),
            "description": self.description_edit.text().strip(),
        }

    def emit_change(self):
        if self.on_change:
            self.on_change()

    def handle_remove(self):
        if self.on_remove:
            self.on_remove(self)


class OldDayEditorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Edit Old Day")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumWidth(860)
        self.loading_rows = False
        self.rows = []

        layout = QVBoxLayout(self)

        help_label = QLabel(f"Edit a past day by rebalancing the rows. The day must total {MAX_HOURS_PER_DAY} hours and each row must stay at {MAX_HOURS_PER_ENTRY} hours or less.")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Date:"))

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setDate(QDate.currentDate().addDays(-1))
        self.date_edit.setMaximumDate(QDate.currentDate())
        self.date_edit.dateChanged.connect(self.load_selected_date)
        top_row.addWidget(self.date_edit)
        top_row.addStretch(1)
        layout.addLayout(top_row)

        self.summary_label = QLabel("")
        layout.addWidget(self.summary_label)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll_container = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_container)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(10)
        self.scroll_layout.addStretch(1)
        self.scroll.setWidget(self.scroll_container)
        layout.addWidget(self.scroll)

        button_row = QHBoxLayout()
        self.add_row_button = QPushButton("Add Task")
        self.add_row_button.clicked.connect(self.add_blank_row)
        button_row.addWidget(self.add_row_button)
        button_row.addStretch(1)

        self.submit_button = QPushButton("Save Day")
        self.submit_button.clicked.connect(self.submit_day)
        button_row.addWidget(self.submit_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_row.addWidget(cancel_button)

        layout.addLayout(button_row)

        self.load_selected_date()

    def current_date_str(self):
        return self.date_edit.date().toString("yyyy-MM-dd")

    def clear_rows(self):
        while self.rows:
            row = self.rows.pop()
            row.setParent(None)
            row.deleteLater()

    def add_blank_row(self):
        self.add_entry_row()
        self.update_summary()

    def add_entry_row(self, entry=None):
        row = DayEntryRow(entry=entry, on_change=self.update_summary, on_remove=self.remove_row)
        insert_at = self.scroll_layout.count() - 1
        self.scroll_layout.insertWidget(insert_at, row)
        self.rows.append(row)
        return row

    def remove_row(self, row):
        if len(self.rows) <= 1:
            show_box(self, QMessageBox.Warning, "Edit Old Day", "Keep at least one task row for the day.")
            return

        self.rows.remove(row)
        row.setParent(None)
        row.deleteLater()
        self.update_summary()

    def load_selected_date(self):
        date_str = self.current_date_str()
        self.loading_rows = True
        try:
            self.clear_rows()
            entries = get_timesheet_entries_for_day(date_str)
            if not entries:
                self.add_entry_row()
            else:
                for entry in entries:
                    self.add_entry_row(entry)
        finally:
            self.loading_rows = False

        self.update_summary()

    def collect_entries(self):
        return [row.get_entry() for row in self.rows]

    def is_valid(self):
        entries = self.collect_entries()
        if not entries:
            return False, "Add at least one task."

        total_hours = 0
        for entry in entries:
            if entry["activity"] not in VALID_ACTIVITIES:
                return False, "Select a valid activity for every row."
            if not entry["description"]:
                return False, "Description cannot be empty."
            if entry["hours"] < 1 or entry["hours"] > MAX_HOURS_PER_ENTRY:
                return False, f"Each row must be between 1 and {MAX_HOURS_PER_ENTRY} hours."
            total_hours += entry["hours"]

        if total_hours != MAX_HOURS_PER_DAY:
            return False, f"Total must equal {MAX_HOURS_PER_DAY} hours."

        return True, ""

    def update_summary(self):
        if self.loading_rows:
            return

        entries = self.collect_entries()
        total_hours = sum(entry["hours"] for entry in entries)
        valid, message = self.is_valid()
        self.summary_label.setText(
            f"Total hours: {total_hours} / {MAX_HOURS_PER_DAY}"
            + (f"  |  {message}" if not valid and message else "")
        )
        self.submit_button.setEnabled(valid)

    def submit_day(self):
        date_str = self.current_date_str()
        entries = self.collect_entries()

        try:
            replace_timesheet_entries_for_day(date_str, entries)
            for activity in dict.fromkeys(entry["activity"] for entry in entries):
                record_recent_activity(activity)
            show_box(self, QMessageBox.Information, "Saved", f"Updated {date_str} successfully.")
            self.accept()
        except Exception as exc:
            show_box(self, QMessageBox.Critical, "Edit Old Day", str(exc))


class TimesheetController(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.tray_icon = None
        self.last_tick_key = None
        self.init_tray()
        self.init_timer()
        self.show_first_run_notice()

    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(create_app_icon(), self.app)
        self.tray_icon.setToolTip("Timesheet Tracker")

        menu = QMenu()

        manual_log_menu = QMenu("Manual Log", self)

        log_today_action = QAction("Log Today", self)
        log_today_action.triggered.connect(self.show_manual_log)
        manual_log_menu.addAction(log_today_action)

        edit_old_day_action = QAction("Edit Old Day", self)
        edit_old_day_action.triggered.connect(self.show_old_day_editor)
        manual_log_menu.addAction(edit_old_day_action)

        menu.addMenu(manual_log_menu)

        leave_action = QAction("Mark Today as Leave (OOF)", self)
        leave_action.triggered.connect(self.mark_today_leave)
        menu.addAction(leave_action)

        menu.addSeparator()

        export_action = QAction("Export", self)
        export_action.triggered.connect(self.export_now)
        menu.addAction(export_action)

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit_app)
        menu.addAction(quit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.show()

    def init_timer(self):
        self.timer = QTimer(self)
        self.timer.setInterval(60 * 1000)
        self.timer.timeout.connect(self.check_background_tasks)
        self.timer.start()

    def show_first_run_notice(self):
        if not IS_FIRST_RUN:
            return

        message = f"Welcome to Timesheet Tracker!\n\nYou are running in {CURRENT_MODE} mode.\n"
        if CURRENT_MODE == "Portable":
            show_box(self, QMessageBox.Warning, "First Run Setup", message + "\nDatabase and config are stored in the application folder.")
        else:
            show_box(self, QMessageBox.Information, "First Run Setup", message + "\nDatabase is stored in your Windows AppData folder.")

    def show_manual_log(self):
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        remaining_hours = get_remaining_hours_for_day(today_str)
        if remaining_hours <= 0:
            show_box(self, QMessageBox.Information, "All Good!", "You have 0 hours remaining for today.")
            return

        dialog = ManualLogDialog(remaining_hours, today_str, self)
        dialog.exec()

    def show_old_day_editor(self):
        dialog = OldDayEditorDialog(self)
        dialog.exec()

    def mark_today_leave(self):
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        if ask_yes_no(
            self,
            "Out of Office",
            "Taking today off?\n\nThis will mute timesheet reminders for the rest of the day.",
        ):
            add_leave(today_str)
            show_box(self, QMessageBox.Information, "Rest Up!", "Today is marked as leave. The tracker is muted until tomorrow.")

    def export_now(self):
        export_timesheet(self, prompt_for_path=True)

    def quit_app(self):
        self.tray_icon.hide()
        QApplication.instance().quit()

    def check_background_tasks(self):
        now = datetime.datetime.now()
        if now.minute != 0 or not (10 <= now.hour <= 19):
            return

        tick_key = f"{now.strftime('%Y-%m-%d')}:{now.hour}:{now.minute}"
        if self.last_tick_key == tick_key:
            return
        self.last_tick_key = tick_key

        today_str = now.strftime("%Y-%m-%d")
        if is_month_end_freeze(today_str):
            if now.hour == 10:
                show_box(self, QMessageBox.Warning, "ALERT", "Portal freezes EOD today!")
            elif now.hour == 19:
                export_timesheet(self, prompt_for_path=False)
                return

        unlogged_hours = get_unlogged_hours(today_str)
        if unlogged_hours > 0:
            dialog = ManualLogDialog(unlogged_hours, today_str, self)
            dialog.exec()


def main():
    setup_database()
    setup_persistence()

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setApplicationName("TimesheetTracker")

    controller = TimesheetController(app)
    app.controller = controller

    sys.exit(app.exec())


if __name__ == "__main__":
    main()