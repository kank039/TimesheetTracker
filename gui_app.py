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
    add_leave,
    add_project,
    add_timesheet_entry,
    delete_project,
    get_all_leaves,
    get_company_holidays,
    get_projects,
    get_timesheet_entries_for_day,
    get_logged_hours_for_day,
    get_recent_activities,
    get_remaining_hours_for_day,
    get_logged_minutes_for_day,
    get_unlogged_hours,
    is_month_end_freeze,
    record_recent_activity,
    remove_leave,
    replace_timesheet_entries_for_day,
    setup_database,
    update_project,
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

        self.project_combo = QComboBox()
        projects = get_projects()
        self.project_combo.addItems(projects)
        if len(projects) == 1:
            self.project_combo.setCurrentIndex(0)
            self.project_combo.setEnabled(False)
        form.addRow("Project:", self.project_combo)

        self.activity_combo = QComboBox()
        self.populate_activity_combo()
        form.addRow("Activity:", self.activity_combo)

        # remaining_hours may be float (hours). Convert to minutes for precise handling.
        remaining_minutes = int(round(remaining_hours * 60))

        self.hours_spin = QSpinBox()
        max_allowed_hours = remaining_minutes // 60
        self.hours_spin.setRange(0, max_allowed_hours)
        self.hours_spin.setValue(max_allowed_hours)
        form.addRow("Hours:", self.hours_spin)

        self.minutes_spin = QSpinBox()
        self.minutes_spin.setRange(0, 59)
        default_minutes = remaining_minutes % 60
        self.minutes_spin.setValue(default_minutes)
        form.addRow("Minutes:", self.minutes_spin)

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
        minutes = int(self.minutes_spin.value())
        description = self.desc_edit.toPlainText().strip()

        if not description:
            show_box(self, QMessageBox.Warning, "Warning", "Description cannot be empty.")
            return

        try:
            remaining_minutes = MAX_HOURS_PER_DAY * 60 - get_logged_minutes_for_day(self.today_str)
            total_minutes = hours * 60 + minutes
            if total_minutes <= 0:
                show_box(self, QMessageBox.Warning, "Warning", "Specify a positive time to log.")
                return
            if total_minutes > remaining_minutes:
                rem_h = remaining_minutes // 60
                rem_m = remaining_minutes % 60
                show_box(self, QMessageBox.Warning, "Warning", f"Only {rem_h}h {rem_m}m remain for today.")
                return

            minutes_left = total_minutes
            while minutes_left > 0:
                chunk = min(MAX_HOURS_PER_ENTRY * 60, minutes_left)
                chunk_hours = chunk // 60
                chunk_minutes = chunk % 60
                add_timesheet_entry(self.project_combo.currentText().strip(), activity, self.today_str, chunk_hours, chunk_minutes, description)
                minutes_left -= chunk

            record_recent_activity(activity)
            total_now_min = get_logged_minutes_for_day(self.today_str)
            if total_now_min >= MAX_HOURS_PER_DAY * 60:
                show_box(self, QMessageBox.Information, "Done for the day!", "You have worked enough for today.")
            else:
                now_h = total_now_min // 60
                now_m = total_now_min % 60
                show_box(self, QMessageBox.Information, "Success", f"Logged {hours}h {minutes}m successfully! Total today: {now_h}h {now_m}m / {MAX_HOURS_PER_DAY}h")
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

        self.project_combo = QComboBox()
        projects = get_projects()
        self.project_combo.addItems(projects)
        if len(projects) == 1:
            self.project_combo.setCurrentIndex(0)
            self.project_combo.setEnabled(False)
        layout.addWidget(QLabel("Project"), 0, 0)
        layout.addWidget(self.project_combo, 1, 0)

        self.activity_combo = QComboBox()
        self.activity_combo.addItems(VALID_ACTIVITIES)
        layout.addWidget(QLabel("Activity"), 0, 1)
        layout.addWidget(self.activity_combo, 1, 1)

        self.hours_spin = QSpinBox()
        self.hours_spin.setRange(0, MAX_HOURS_PER_ENTRY)
        layout.addWidget(QLabel("Hours"), 0, 2)
        layout.addWidget(self.hours_spin, 1, 2)

        self.minutes_spin = QSpinBox()
        self.minutes_spin.setRange(0, 59)
        layout.addWidget(QLabel("Minutes"), 0, 3)
        layout.addWidget(self.minutes_spin, 1, 3)

        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Task description")
        layout.addWidget(QLabel("Description"), 0, 4)
        layout.addWidget(self.description_edit, 1, 4)

        self.remove_button = QPushButton("Remove")
        self.remove_button.clicked.connect(self.handle_remove)
        layout.addWidget(self.remove_button, 1, 5)

        self.project_combo.currentIndexChanged.connect(self.emit_change)
        self.activity_combo.currentIndexChanged.connect(self.emit_change)
        self.hours_spin.valueChanged.connect(self.emit_change)
        self.minutes_spin.valueChanged.connect(self.emit_change)
        self.description_edit.textChanged.connect(self.emit_change)

        if entry:
            self.set_entry(entry)
        else:
            self.hours_spin.setValue(1)

    def set_entry(self, entry):
        project = entry.get("project", "")
        activity = entry.get("activity", "")
        hours = int(entry.get("hours", 0))
        minutes = int(entry.get("minutes", 0))
        description = entry.get("description", "")

        project_index = self.project_combo.findText(project)
        if project_index >= 0:
            self.project_combo.setCurrentIndex(project_index)
        activity_index = self.activity_combo.findText(activity)
        if activity_index >= 0:
            self.activity_combo.setCurrentIndex(activity_index)
        self.hours_spin.setValue(max(0, min(MAX_HOURS_PER_ENTRY, hours)))
        self.minutes_spin.setValue(max(0, min(59, minutes)))
        self.description_edit.setText(description)

    def get_entry(self):
        return {
            "project": self.project_combo.currentText().strip(),
            "activity": self.activity_combo.currentText().strip(),
            "hours": int(self.hours_spin.value()),
            "minutes": int(self.minutes_spin.value()),
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
        total_minutes = 0
        valid_projects = get_projects()
        for entry in entries:
            if entry["project"] not in valid_projects:
                return False, "Select a valid project for every row."
            if entry["activity"] not in VALID_ACTIVITIES:
                return False, "Select a valid activity for every row."
            if not entry["description"]:
                return False, "Description cannot be empty."
            hours = int(entry.get("hours", 0))
            minutes = int(entry.get("minutes", 0))
            if hours < 0 or hours > MAX_HOURS_PER_ENTRY:
                return False, f"Each row hours must be between 0 and {MAX_HOURS_PER_ENTRY}."
            if minutes < 0 or minutes > 59:
                return False, "Minutes must be between 0 and 59."
            if hours == 0 and minutes == 0:
                return False, "Each row must have a positive time value."
            total_minutes += hours * 60 + minutes

        if total_minutes != MAX_HOURS_PER_DAY * 60:
            return False, f"Total must equal {MAX_HOURS_PER_DAY} hours."

        return True, ""

    def update_summary(self):
        if self.loading_rows:
            return

        entries = self.collect_entries()
        total_minutes = sum(entry.get("hours", 0) * 60 + entry.get("minutes", 0) for entry in entries)
        th = total_minutes // 60
        tm = total_minutes % 60
        valid, message = self.is_valid()
        self.summary_label.setText(
            f"Total: {th}h {tm}m / {MAX_HOURS_PER_DAY}h"
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


class HolidayManagerWindow(QWidget):
    """Main GUI window opened on tray left-click: shows company holidays & personal leaves."""

    STYLE = """
        HolidayManagerWindow {
            background: #1e1e2e;
        }
        QLabel#windowTitle {
            color: #cdd6f4;
            font-size: 20px;
            font-weight: 700;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#sectionTitle {
            color: #89b4fa;
            font-size: 14px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
            padding: 4px 0;
        }
        QLabel#subtitle {
            color: #6c7086;
            font-size: 11px;
            font-family: 'Segoe UI', sans-serif;
        }
        QFrame#card {
            background: #313244;
            border: 1px solid #45475a;
            border-radius: 10px;
            padding: 14px;
        }
        QLabel#holidayDate {
            color: #a6adc8;
            font-size: 12px;
            font-family: 'Segoe UI', monospace;
            min-width: 90px;
        }
        QLabel#holidayName {
            color: #cdd6f4;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#holidayPast {
            color: #585b70;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#holidayDatePast {
            color: #585b70;
            font-size: 12px;
            font-family: 'Segoe UI', monospace;
            min-width: 90px;
        }
        QLabel#holidayUpcoming {
            color: #a6e3a1;
            font-size: 12px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#holidayDateUpcoming {
            color: #a6e3a1;
            font-size: 12px;
            font-weight: 600;
            font-family: 'Segoe UI', monospace;
            min-width: 90px;
        }
        QLabel#badge {
            background: #89b4fa;
            color: #1e1e2e;
            font-size: 10px;
            font-weight: 700;
            border-radius: 4px;
            padding: 2px 6px;
        }
        QLabel#badgePast {
            background: #45475a;
            color: #6c7086;
            font-size: 10px;
            font-weight: 700;
            border-radius: 4px;
            padding: 2px 6px;
        }
        QLabel#leaveDate {
            color: #cdd6f4;
            font-size: 12px;
            font-family: 'Segoe UI', monospace;
            min-width: 90px;
        }
        QLabel#leaveName {
            color: #cdd6f4;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#removeBtn {
            background: transparent;
            color: #f38ba8;
            border: 1px solid #f38ba8;
            border-radius: 4px;
            font-size: 11px;
            padding: 2px 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#removeBtn:hover {
            background: #f38ba8;
            color: #1e1e2e;
        }
        QPushButton#addBtn {
            background: #89b4fa;
            color: #1e1e2e;
            border: none;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            padding: 6px 18px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#addBtn:hover {
            background: #b4d0fb;
        }
        QDateEdit {
            background: #45475a;
            color: #cdd6f4;
            border: 1px solid #585b70;
            border-radius: 6px;
            padding: 4px 8px;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLineEdit#leaveReasonEdit {
            background: #45475a;
            color: #cdd6f4;
            border: 1px solid #585b70;
            border-radius: 6px;
            padding: 4px 8px;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QScrollArea {
            border: none;
            background: transparent;
        }
        QScrollBar:vertical {
            background: #313244;
            width: 8px;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical {
            background: #585b70;
            border-radius: 4px;
            min-height: 30px;
        }
        QScrollBar::handle:vertical:hover {
            background: #6c7086;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        QLabel#emptyState {
            color: #6c7086;
            font-size: 12px;
            font-style: italic;
            font-family: 'Segoe UI', sans-serif;
            padding: 20px;
        }
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Timesheet Tracker — Holidays & Leaves")
        self.setMinimumSize(780, 520)
        self.resize(820, 560)
        self.setStyleSheet(self.STYLE)
        self.setWindowIcon(create_app_icon())

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        # Title bar
        title = QLabel("📅  Holidays & Leaves")
        title.setObjectName("windowTitle")
        root.addWidget(title)

        sub = QLabel("Company holidays are fixed. You can add personal leaves to skip logging on those days.")
        sub.setObjectName("subtitle")
        sub.setWordWrap(True)
        root.addWidget(sub)

        # Two-panel horizontal layout
        panels = QHBoxLayout()
        panels.setSpacing(14)

        # ── Left panel: Company Holidays ──
        left_card = QFrame()
        left_card.setObjectName("card")
        left_layout = QVBoxLayout(left_card)
        left_layout.setSpacing(6)

        section_title = QLabel("🏢  Company Holidays (2026)")
        section_title.setObjectName("sectionTitle")
        left_layout.addWidget(section_title)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        self.company_scroll_content = QWidget()
        self.company_scroll_layout = QVBoxLayout(self.company_scroll_content)
        self.company_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.company_scroll_layout.setSpacing(4)
        left_scroll.setWidget(self.company_scroll_content)
        left_layout.addWidget(left_scroll)

        panels.addWidget(left_card, stretch=1)

        # ── Right panel: Personal Leaves ──
        right_card = QFrame()
        right_card.setObjectName("card")
        right_layout = QVBoxLayout(right_card)
        right_layout.setSpacing(8)

        section_title2 = QLabel("🧘  My Personal Leaves")
        section_title2.setObjectName("sectionTitle")
        right_layout.addWidget(section_title2)

        # Add-leave form
        form_row = QHBoxLayout()
        form_row.setSpacing(8)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_edit.setDate(QDate.currentDate())
        form_row.addWidget(self.date_edit)

        self.reason_edit = QLineEdit()
        self.reason_edit.setObjectName("leaveReasonEdit")
        self.reason_edit.setPlaceholderText("Reason (optional)")
        form_row.addWidget(self.reason_edit, stretch=1)

        add_btn = QPushButton("+ Add Leave")
        add_btn.setObjectName("addBtn")
        add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_btn.clicked.connect(self.add_personal_leave)
        form_row.addWidget(add_btn)

        right_layout.addLayout(form_row)

        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #45475a;")
        right_layout.addWidget(sep)

        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        self.personal_scroll_content = QWidget()
        self.personal_scroll_layout = QVBoxLayout(self.personal_scroll_content)
        self.personal_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.personal_scroll_layout.setSpacing(4)
        right_scroll.setWidget(self.personal_scroll_content)
        right_layout.addWidget(right_scroll)

        panels.addWidget(right_card, stretch=1)

        root.addLayout(panels, stretch=1)

        self.populate()

    # ── Data population ──

    def populate(self):
        self._populate_company_holidays()
        self._populate_personal_leaves()

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.setParent(None)
                widget.deleteLater()

    def _populate_company_holidays(self):
        self._clear_layout(self.company_scroll_layout)
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        holidays = get_company_holidays()
        next_upcoming_found = False

        for date_str, name in holidays:
            row = QHBoxLayout()
            row.setSpacing(8)

            is_past = date_str < today
            is_next = (not is_past and not next_upcoming_found)

            date_label = QLabel(date_str)
            name_label = QLabel(name)

            if is_past:
                date_label.setObjectName("holidayDatePast")
                name_label.setObjectName("holidayPast")
            elif is_next:
                date_label.setObjectName("holidayDateUpcoming")
                name_label.setObjectName("holidayUpcoming")
                next_upcoming_found = True
            else:
                date_label.setObjectName("holidayDate")
                name_label.setObjectName("holidayName")

            row.addWidget(date_label)
            row.addWidget(name_label, stretch=1)

            if is_next:
                badge = QLabel("NEXT")
                badge.setObjectName("badge")
                row.addWidget(badge)
            elif is_past:
                badge = QLabel("PAST")
                badge.setObjectName("badgePast")
                row.addWidget(badge)

            container = QWidget()
            container.setLayout(row)
            self.company_scroll_layout.addWidget(container)

        self.company_scroll_layout.addStretch(1)

    def _populate_personal_leaves(self):
        self._clear_layout(self.personal_scroll_layout)
        all_leaves = get_all_leaves()
        personal = [l for l in all_leaves if l["type"] == "Personal Leave"]

        if not personal:
            empty = QLabel("No personal leaves set yet.")
            empty.setObjectName("emptyState")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.personal_scroll_layout.addWidget(empty)
        else:
            for leave in personal:
                row = QHBoxLayout()
                row.setSpacing(8)

                date_label = QLabel(leave["date"])
                date_label.setObjectName("leaveDate")
                row.addWidget(date_label)

                reason = leave.get("name", "") or "Personal Leave"
                name_label = QLabel(reason)
                name_label.setObjectName("leaveName")
                row.addWidget(name_label, stretch=1)

                remove_btn = QPushButton("Remove")
                remove_btn.setObjectName("removeBtn")
                remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                date_val = leave["date"]
                remove_btn.clicked.connect(lambda checked, d=date_val: self.remove_personal_leave(d))
                row.addWidget(remove_btn)

                container = QWidget()
                container.setLayout(row)
                self.personal_scroll_layout.addWidget(container)

        self.personal_scroll_layout.addStretch(1)

    # ── Actions ──

    def add_personal_leave(self):
        date_str = self.date_edit.date().toString("yyyy-MM-dd")
        reason = self.reason_edit.text().strip()

        # Check if it's a weekend
        from core_engine import is_weekend
        if is_weekend(date_str):
            show_box(self, QMessageBox.Icon.Warning, "Cannot Add", "That date is a weekend — no logging happens anyway.")
            return

        # Check if it's already a company holiday
        company_dates = [d for d, _ in get_company_holidays()]
        if date_str in company_dates:
            show_box(self, QMessageBox.Icon.Information, "Already a Holiday", "That date is already a company holiday.")
            return

        # Check if already added
        existing = get_all_leaves()
        if any(l["date"] == date_str and l["type"] == "Personal Leave" for l in existing):
            show_box(self, QMessageBox.Icon.Information, "Already Added", "You already have a leave on that date.")
            return

        add_leave(date_str, "Personal Leave", reason or "Personal Leave")
        self.reason_edit.clear()
        self._populate_personal_leaves()

    def remove_personal_leave(self, date_str):
        try:
            remove_leave(date_str)
            self._populate_personal_leaves()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Cannot Remove", str(exc))


class ProjectManagerDialog(QDialog):
    """Dialog to add, rename, and delete projects."""

    STYLE = """
        ProjectManagerDialog {
            background: #1e1e2e;
        }
        QLabel#dialogTitle {
            color: #cdd6f4;
            font-size: 18px;
            font-weight: 700;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#subtitle {
            color: #6c7086;
            font-size: 11px;
            font-family: 'Segoe UI', sans-serif;
        }
        QFrame#card {
            background: #313244;
            border: 1px solid #45475a;
            border-radius: 10px;
            padding: 14px;
        }
        QLabel#projectName {
            color: #cdd6f4;
            font-size: 13px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLineEdit#projectInput {
            background: #45475a;
            color: #cdd6f4;
            border: 1px solid #585b70;
            border-radius: 6px;
            padding: 6px 10px;
            font-size: 13px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#addBtn {
            background: #a6e3a1;
            color: #1e1e2e;
            border: none;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            padding: 6px 18px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#addBtn:hover {
            background: #c6f0c4;
        }
        QPushButton#renameBtn {
            background: transparent;
            color: #89b4fa;
            border: 1px solid #89b4fa;
            border-radius: 4px;
            font-size: 11px;
            padding: 2px 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#renameBtn:hover {
            background: #89b4fa;
            color: #1e1e2e;
        }
        QPushButton#deleteBtn {
            background: transparent;
            color: #f38ba8;
            border: 1px solid #f38ba8;
            border-radius: 4px;
            font-size: 11px;
            padding: 2px 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#deleteBtn:hover {
            background: #f38ba8;
            color: #1e1e2e;
        }
        QLabel#emptyState {
            color: #6c7086;
            font-size: 12px;
            font-style: italic;
            font-family: 'Segoe UI', sans-serif;
            padding: 20px;
        }
        QScrollArea {
            border: none;
            background: transparent;
        }
        QScrollBar:vertical {
            background: #313244;
            width: 8px;
            border-radius: 4px;
        }
        QScrollBar::handle:vertical {
            background: #585b70;
            border-radius: 4px;
            min-height: 30px;
        }
        QScrollBar::handle:vertical:hover {
            background: #6c7086;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Projects")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumSize(520, 380)
        self.resize(560, 420)
        self.setStyleSheet(self.STYLE)
        self.setWindowIcon(create_app_icon())

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        title = QLabel("📁  Manage Projects")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)

        sub = QLabel("Add, rename, or remove projects. Projects with timesheet entries cannot be deleted.")
        sub.setObjectName("subtitle")
        sub.setWordWrap(True)
        layout.addWidget(sub)

        # Add project form
        form_row = QHBoxLayout()
        form_row.setSpacing(8)

        self.name_edit = QLineEdit()
        self.name_edit.setObjectName("projectInput")
        self.name_edit.setPlaceholderText("New project name")
        form_row.addWidget(self.name_edit, stretch=1)

        add_btn = QPushButton("+ Add Project")
        add_btn.setObjectName("addBtn")
        add_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_btn.clicked.connect(self.add_new_project)
        form_row.addWidget(add_btn)

        layout.addLayout(form_row)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #45475a;")
        layout.addWidget(sep)

        # Project list
        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(6)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(6)
        scroll.setWidget(self.scroll_content)
        card_layout.addWidget(scroll)

        layout.addWidget(card, stretch=1)

        # Close button
        close_row = QHBoxLayout()
        close_row.addStretch(1)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_row.addWidget(close_btn)
        layout.addLayout(close_row)

        self.populate()

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.setParent(None)
                widget.deleteLater()

    def populate(self):
        self._clear_layout(self.scroll_layout)
        projects = get_projects()

        if not projects:
            empty = QLabel("No projects defined. Add one above.")
            empty.setObjectName("emptyState")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.scroll_layout.addWidget(empty)
        else:
            for name in projects:
                row = QHBoxLayout()
                row.setSpacing(8)

                label = QLabel(name)
                label.setObjectName("projectName")
                row.addWidget(label, stretch=1)

                rename_btn = QPushButton("Rename")
                rename_btn.setObjectName("renameBtn")
                rename_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                rename_btn.clicked.connect(lambda checked, n=name: self.rename_project(n))
                row.addWidget(rename_btn)

                delete_btn = QPushButton("Delete")
                delete_btn.setObjectName("deleteBtn")
                delete_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                delete_btn.clicked.connect(lambda checked, n=name: self.remove_project(n))
                row.addWidget(delete_btn)

                container = QWidget()
                container.setLayout(row)
                self.scroll_layout.addWidget(container)

        self.scroll_layout.addStretch(1)

    def add_new_project(self):
        name = self.name_edit.text().strip()
        if not name:
            show_box(self, QMessageBox.Icon.Warning, "Add Project", "Project name cannot be empty.")
            return
        try:
            add_project(name)
            self.name_edit.clear()
            self.populate()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Add Project", str(exc))

    def rename_project(self, old_name):
        from PySide6.QtWidgets import QInputDialog
        new_name, ok = QInputDialog.getText(self, "Rename Project", f"New name for '{old_name}':", text=old_name)
        if not ok or not new_name.strip():
            return
        try:
            update_project(old_name, new_name.strip())
            self.populate()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Rename Project", str(exc))

    def remove_project(self, name):
        if not ask_yes_no(self, "Delete Project", f"Delete project '{name}'?\n\nThis cannot be undone."):
            return
        try:
            delete_project(name)
            self.populate()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Delete Project", str(exc))


class TimesheetController(QWidget):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.tray_icon = None
        self.holiday_window = None
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

        projects_action = QAction("Manage Projects", self)
        projects_action.triggered.connect(self.show_project_manager)
        menu.addAction(projects_action)

        menu.addSeparator()

        export_action = QAction("Export", self)
        export_action.triggered.connect(self.export_now)
        menu.addAction(export_action)

        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit_app)
        menu.addAction(quit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
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
        remaining_minutes = MAX_HOURS_PER_DAY * 60 - get_logged_minutes_for_day(today_str)
        if remaining_minutes <= 0:
            show_box(self, QMessageBox.Information, "All Good!", "You have 0 hours remaining for today.")
            return

        dialog = ManualLogDialog(remaining_minutes / 60.0, today_str, self)
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

    def show_project_manager(self):
        dialog = ProjectManagerDialog(self)
        dialog.exec()

    def show_holiday_window(self):
        if self.holiday_window is None:
            self.holiday_window = HolidayManagerWindow()
        self.holiday_window.populate()
        self.holiday_window.show()
        self.holiday_window.raise_()
        self.holiday_window.activateWindow()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.show_holiday_window()

    def quit_app(self):
        if self.holiday_window:
            self.holiday_window.close()
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
            elif now.hour == 18:
                export_timesheet(self, prompt_for_path=False)
                return

        unlogged_hours = get_unlogged_hours(today_str)
        if unlogged_hours > 0:
            dialog = ManualLogDialog(unlogged_hours, today_str, self)
            dialog.exec()


def main(lock=None):
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