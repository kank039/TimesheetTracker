import datetime
import os
import sqlite3
import subprocess
import sys

import pandas as pd
import winreg
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QAction, QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QSystemTrayIcon,
    QVBoxLayout,
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
    get_logged_hours_for_day,
    get_unlogged_hours,
    is_month_end_freeze,
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


def export_timesheet(parent=None):
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

    export_folder = get_export_folder()
    export_name = f"TimesheetTracker_Export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    export_path = os.path.join(export_folder, export_name)

    try:
        df.to_excel(export_path, index=False)
        show_box(parent, QMessageBox.Information, "Export", f"Export successful!\n\nSaved to:\n{export_path}")
        return export_path
    except Exception as exc:
        show_box(parent, QMessageBox.Critical, "Export Failed", f"Could not export the file.\n\n{exc}")
        return None


class ManualLogDialog(QDialog):
    def __init__(self, unlogged_hours, today_str, parent=None):
        super().__init__(parent)
        self.today_str = today_str
        self.unlogged_hours = unlogged_hours

        self.setWindowTitle("Timesheet Reminder")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumWidth(420)

        layout = QVBoxLayout(self)

        info_label = QLabel(f"You have {unlogged_hours} unlogged hours.")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)

        form = QFormLayout()

        self.activity_combo = QComboBox()
        self.activity_combo.addItems(VALID_ACTIVITIES)
        self.activity_combo.setCurrentIndex(1 if len(VALID_ACTIVITIES) > 1 else 0)
        form.addRow("Activity:", self.activity_combo)

        self.hours_spin = QSpinBox()
        max_allowed = min(MAX_HOURS_PER_ENTRY, unlogged_hours)
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

    def submit_entry(self):
        activity = self.activity_combo.currentText().strip()
        hours = int(self.hours_spin.value())
        description = self.desc_edit.toPlainText().strip()

        if not description:
            show_box(self, QMessageBox.Warning, "Warning", "Description cannot be empty.")
            return

        try:
            add_timesheet_entry(VALID_PROJECTS[0], activity, self.today_str, hours, description)
            total_now = get_logged_hours_for_day(self.today_str)
            if total_now >= MAX_HOURS_PER_DAY:
                show_box(self, QMessageBox.Information, "Done for the day!", "You have worked enough for today.")
            else:
                show_box(self, QMessageBox.Information, "Success", f"Logged {hours} hours successfully! Total today: {total_now}/8")
            self.accept()
        except Exception as exc:
            show_box(self, QMessageBox.Critical, "Error", str(exc))


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

        manual_log_action = QAction("Manual Log", self)
        manual_log_action.triggered.connect(self.show_manual_log)
        menu.addAction(manual_log_action)

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
        unlogged_hours = get_unlogged_hours(today_str)
        if unlogged_hours <= 0:
            show_box(self, QMessageBox.Information, "All Good!", "You have 0 unlogged hours for today.")
            return

        dialog = ManualLogDialog(unlogged_hours, today_str, self)
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
        export_timesheet(self)

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
                export_timesheet(self)
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