import datetime
import os
import sqlite3
import subprocess
import sys

import winreg
from openpyxl import Workbook
from PyQt6.QtCore import QDate, QTimer, Qt
from PyQt6.QtGui import QAction, QColor, QIcon, QPainter, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QDateEdit,
    QCheckBox,
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
    QRadioButton,
    QScrollArea,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTextEdit,
    QSystemTrayIcon,
    QVBoxLayout,
    QFrame,
    QWidget,
)

from core_engine import (
    DB_NAME,
    CURRENT_MODE,
    DAY_NAMES,
    IS_FIRST_RUN,
    MAX_HOURS_PER_DAY,
    MAX_HOURS_PER_ENTRY,
    VALID_ACTIVITIES,
    add_leave,
    add_project,
    add_time_block,
    add_timesheet_entry,
    delete_project,
    delete_time_block,
    force_delete_project,
    get_all_leaves,
    get_company_holidays,
    get_dates_with_project_entries,
    get_default_project,
    get_projects,
    get_time_blocks,
    get_timesheet_entries_for_day,
    get_total_blocked_minutes,
    get_logged_hours_for_day,
    get_recent_activities,
    get_remaining_hours_for_day,
    get_logged_minutes_for_day,
    get_unlogged_hours,
    insert_time_blocks_for_day,
    is_month_end_freeze,
    reassign_project_entries,
    record_recent_activity,
    remove_leave,
    replace_timesheet_entries_for_day,
    set_default_project,
    setup_database,
    toggle_time_block,
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


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)


def create_app_icon():
    icon_path = get_resource_path("app.ico")
    if os.path.exists(icon_path):
        return QIcon(icon_path)
    
    # Fallback to drawn icon if app.ico is missing
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
    cursor = conn.cursor()
    cursor.execute("SELECT project, activity, log_date, hours, minutes, tag, description FROM timesheet ORDER BY log_date ASC")
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        show_box(parent, QMessageBox.Information, "Export", "Timesheet database is currently empty.")
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = "Timesheet Export"
    
    headers = [
        "Project",
        "Activity",
        "Date(dd-MM-yyyy)",
        "Time Spent(hh)",
        "Time Spent (mm)",
        "Tag",
        "Description",
    ]
    ws.append(headers)

    for r in rows:
        project, activity, log_date, hours, minutes, tag, description = r
        try:
            date_obj = datetime.datetime.strptime(log_date, "%Y-%m-%d")
            formatted_date = date_obj.strftime("%d-%m-%Y")
        except ValueError:
            formatted_date = log_date
            
        ws.append([project, activity, formatted_date, hours, minutes, tag, description])

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
        wb.save(export_path)
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
    def __init__(self, entry=None, on_change=None, on_remove=None,
                 default_project=None, use_default=False, parent=None):
        super().__init__(parent)
        self.on_change = on_change
        self.on_remove = on_remove
        self._default_project = default_project

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.card = QFrame()
        self.card.setFrameShape(QFrame.Shape.StyledPanel)
        card_layout = QVBoxLayout(self.card)
        card_layout.setContentsMargins(6, 4, 6, 4)
        card_layout.setSpacing(4)

        # Row 1: Project + Activity + Remove
        row1 = QHBoxLayout()
        row1.setSpacing(6)

        self.project_combo = QComboBox()
        self.project_combo.setMinimumWidth(120)
        projects = get_projects()
        self.project_combo.addItems(projects)
        if len(projects) == 1:
            self.project_combo.setCurrentIndex(0)
            self.project_combo.setEnabled(False)
        self.project_label = QLabel("Project:")
        row1.addWidget(self.project_label)
        row1.addWidget(self.project_combo)

        self.activity_combo = QComboBox()
        self.activity_combo.addItems(VALID_ACTIVITIES)
        act_label = QLabel("Activity:")
        row1.addWidget(act_label)
        row1.addWidget(self.activity_combo, stretch=1)

        self.remove_button = QPushButton("✕")
        self.remove_button.setFixedSize(26, 26)
        self.remove_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.remove_button.clicked.connect(self.handle_remove)
        row1.addWidget(self.remove_button)

        card_layout.addLayout(row1)

        # Row 2: Hours + Minutes + Description
        row2 = QHBoxLayout()
        row2.setSpacing(6)

        hr_label = QLabel("Hr:")
        row2.addWidget(hr_label)
        self.hours_spin = QSpinBox()
        self.hours_spin.setRange(0, MAX_HOURS_PER_ENTRY)
        # removed fixed width so numbers are visible
        row2.addWidget(self.hours_spin)

        min_label = QLabel("Min:")
        row2.addWidget(min_label)
        self.minutes_spin = QSpinBox()
        self.minutes_spin.setRange(0, 59)
        # removed fixed width so numbers are visible
        row2.addWidget(self.minutes_spin)

        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Task description…")
        row2.addWidget(self.description_edit, stretch=1)

        card_layout.addLayout(row2)
        outer.addWidget(self.card)

        # Connect signals
        self.project_combo.currentIndexChanged.connect(self.emit_change)
        self.activity_combo.currentIndexChanged.connect(self.emit_change)
        self.hours_spin.valueChanged.connect(self.emit_change)
        self.minutes_spin.valueChanged.connect(self.emit_change)
        self.description_edit.textChanged.connect(self.emit_change)

        # Apply default project visibility
        if use_default and default_project:
            self._apply_default(True, default_project)

        if entry:
            self.set_entry(entry)
        else:
            self.hours_spin.setValue(1)

    def _apply_default(self, use, project_name):
        """Show/hide project controls based on default project mode."""
        self.project_label.setVisible(not use)
        self.project_combo.setVisible(not use)
        if use and project_name:
            idx = self.project_combo.findText(project_name)
            if idx >= 0:
                self.project_combo.setCurrentIndex(idx)

    def set_use_default(self, use, project_name):
        """Toggle default project mode on this row."""
        self._default_project = project_name
        self._apply_default(use, project_name)

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
    """Main GUI window opened on tray left-click: shows company holidays, personal leaves & time blocks."""

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
        /* ── Manage Projects styles ── */
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
        QPushButton#defaultBtn {
            background: transparent;
            color: #f9e2af;
            border: 1px solid #f9e2af;
            border-radius: 4px;
            font-size: 11px;
            padding: 2px 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#defaultBtn:hover {
            background: #f9e2af;
            color: #1e1e2e;
        }
        /* ── Time Blocks panel styles ── */
        QLineEdit#blockInput {
            background: #45475a;
            color: #cdd6f4;
            border: 1px solid #585b70;
            border-radius: 6px;
            padding: 4px 8px;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QComboBox {
            background: #45475a;
            color: #cdd6f4;
            border: 1px solid #585b70;
            border-radius: 6px;
            padding: 4px 8px;
            font-size: 11px;
            font-family: 'Segoe UI', sans-serif;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px;
            border-left: none;
        }
        QComboBox QAbstractItemView {
            background: #313244;
            color: #cdd6f4;
            selection-background-color: #45475a;
            border: 1px solid #585b70;
            border-radius: 4px;
        }
        QSpinBox {
            background: #45475a;
            color: #cdd6f4;
            border: 1px solid #585b70;
            border-radius: 6px;
            padding: 4px 6px;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#blockName {
            color: #cdd6f4;
            font-size: 12px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#blockDetail {
            color: #a6adc8;
            font-size: 11px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#blockDuration {
            color: #f9e2af;
            font-size: 11px;
            font-weight: 600;
            font-family: 'Segoe UI', monospace;
        }
        QLabel#blockDays {
            color: #89b4fa;
            font-size: 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        QLabel#blockDisabled {
            color: #585b70;
            font-size: 12px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#toggleBtn {
            background: transparent;
            color: #a6e3a1;
            border: 1px solid #a6e3a1;
            border-radius: 4px;
            font-size: 10px;
            padding: 2px 8px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#toggleBtn:hover {
            background: #a6e3a1;
            color: #1e1e2e;
        }
        QPushButton#toggleBtnOff {
            background: transparent;
            color: #585b70;
            border: 1px solid #585b70;
            border-radius: 4px;
            font-size: 10px;
            padding: 2px 8px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#toggleBtnOff:hover {
            background: #585b70;
            color: #cdd6f4;
        }
        QLabel#blockTotal {
            color: #f9e2af;
            font-size: 12px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
            padding: 6px 0;
        }
        QCheckBox {
            color: #a6adc8;
            font-size: 11px;
            font-family: 'Segoe UI', sans-serif;
            spacing: 6px;
        }
        QLabel#formLabel {
            color: #a6adc8;
            font-size: 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        /* ── Day Entries editor styles ── */
        QLabel#entrySummary {
            color: #f9e2af;
            font-size: 13px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
            padding: 4px 0;
        }
        QLabel#entrySummaryError {
            color: #f38ba8;
            font-size: 13px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
            padding: 4px 0;
        }
        QLabel#entryHint {
            color: #6c7086;
            font-size: 11px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#saveBtn {
            background: #a6e3a1;
            color: #1e1e2e;
            border: none;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            padding: 6px 18px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#saveBtn:hover {
            background: #c6f0c4;
        }
        QPushButton#saveBtn:disabled {
            background: #45475a;
            color: #6c7086;
        }
        QPushButton#addTaskBtn {
            background: transparent;
            color: #89b4fa;
            border: 1px solid #89b4fa;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            padding: 6px 14px;
            font-family: 'Segoe UI', sans-serif;
        }
        QPushButton#addTaskBtn:hover {
            background: #89b4fa;
            color: #1e1e2e;
        }
        QFrame#entryRow {
            background: #1e1e2e;
            border: 1px solid #45475a;
            border-radius: 6px;
            padding: 8px;
        }
        /* ── Tab widget styles ── */
        QTabWidget::pane {
            border: 1px solid #45475a;
            border-radius: 8px;
            background: transparent;
            top: -1px;
        }
        QTabBar::tab {
            background: #313244;
            color: #a6adc8;
            border: 1px solid #45475a;
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            padding: 8px 24px;
            margin-right: 4px;
            font-size: 13px;
            font-weight: 600;
            font-family: 'Segoe UI', sans-serif;
        }
        QTabBar::tab:selected {
            background: #1e1e2e;
            color: #89b4fa;
            border-color: #89b4fa;
            border-bottom: 2px solid #1e1e2e;
        }
        QTabBar::tab:hover:!selected {
            background: #45475a;
            color: #cdd6f4;
        }
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Timesheet Tracker — Holidays, Leaves & Time Blocks")
        self.setMinimumSize(580, 520)
        self.resize(780, 600)
        # self.setStyleSheet(self.STYLE)
        self.setWindowIcon(create_app_icon())

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        # Title bar
        title = QLabel("📅  Holidays, Leaves & Time Blocks")
        title.setObjectName("windowTitle")
        root.addWidget(title)

        sub = QLabel("Company holidays are fixed. Add personal leaves or recurring time blocks to auto-fill your daily timesheet.")
        sub.setObjectName("subtitle")
        sub.setWordWrap(True)
        root.addWidget(sub)

        # ── Tabbed layout ──
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(False)

        # ═══════════════════════════════════════════
        # TAB 1: Holidays & Leaves
        # ═══════════════════════════════════════════
        holidays_tab = QWidget()
        holidays_layout = QVBoxLayout(holidays_tab)
        holidays_layout.setContentsMargins(8, 12, 8, 8)
        holidays_layout.setSpacing(12)

        # ── Company Holidays section ──
        company_card = QFrame()
        company_card.setObjectName("card")
        company_card_layout = QVBoxLayout(company_card)
        company_card_layout.setSpacing(6)

        section_title = QLabel("🏢  Company Holidays (2026)")
        section_title.setObjectName("sectionTitle")
        company_card_layout.addWidget(section_title)

        self.company_scroll = QScrollArea()
        self.company_scroll.setWidgetResizable(True)
        self.company_scroll_content = QWidget()
        self.company_scroll_layout = QVBoxLayout(self.company_scroll_content)
        self.company_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.company_scroll_layout.setSpacing(4)
        self.company_scroll.setWidget(self.company_scroll_content)
        company_card_layout.addWidget(self.company_scroll)

        holidays_layout.addWidget(company_card, stretch=1)

        # ── Personal Leaves section ──
        personal_card = QFrame()
        personal_card.setObjectName("card")
        personal_card_layout = QVBoxLayout(personal_card)
        personal_card_layout.setSpacing(8)

        section_title2 = QLabel("🧘  My Personal Leaves")
        section_title2.setObjectName("sectionTitle")
        personal_card_layout.addWidget(section_title2)

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

        personal_card_layout.addLayout(form_row)

        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        personal_card_layout.addWidget(sep)

        self.personal_scroll = QScrollArea()
        self.personal_scroll.setWidgetResizable(True)
        self.personal_scroll_content = QWidget()
        self.personal_scroll_layout = QVBoxLayout(self.personal_scroll_content)
        self.personal_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.personal_scroll_layout.setSpacing(4)
        self.personal_scroll.setWidget(self.personal_scroll_content)
        personal_card_layout.addWidget(self.personal_scroll)

        holidays_layout.addWidget(personal_card, stretch=1)

        self.tabs.addTab(holidays_tab, "🏖  Holidays && Leaves")

        # ═══════════════════════════════════════════
        # TAB 2: Time Blocking
        # ═══════════════════════════════════════════
        blocks_tab = QWidget()
        blocks_layout = QVBoxLayout(blocks_tab)
        blocks_layout.setContentsMargins(8, 12, 8, 8)
        blocks_layout.setSpacing(12)

        blocks_card = QFrame()
        blocks_card.setObjectName("card")
        blocks_card_layout = QVBoxLayout(blocks_card)
        blocks_card_layout.setSpacing(6)

        section_title3 = QLabel("⏱  Time Blocks")
        section_title3.setObjectName("sectionTitle")
        blocks_card_layout.addWidget(section_title3)

        # -- Add block form --
        block_form = QVBoxLayout()
        block_form.setSpacing(4)

        # Row 1 removed: Name is no longer used

        # Row 2: Project + Activity
        proj_act_row = QHBoxLayout()
        proj_act_row.setSpacing(6)

        proj_lbl = QLabel("Project")
        proj_lbl.setObjectName("formLabel")
        proj_act_row.addWidget(proj_lbl)
        self.block_project_combo = QComboBox()
        self.block_project_combo.setObjectName("blockCombo")
        self.block_project_combo.addItems(get_projects())
        proj_act_row.addWidget(self.block_project_combo, stretch=1)

        act_lbl = QLabel("Activity")
        act_lbl.setObjectName("formLabel")
        proj_act_row.addWidget(act_lbl)
        self.block_activity_combo = QComboBox()
        self.block_activity_combo.setObjectName("blockCombo")
        self.block_activity_combo.addItems(VALID_ACTIVITIES)
        proj_act_row.addWidget(self.block_activity_combo, stretch=1)

        block_form.addLayout(proj_act_row)

        # Row 3: Hours + Minutes + Description
        time_desc_row = QHBoxLayout()
        time_desc_row.setSpacing(6)

        hr_lbl = QLabel("Hr")
        hr_lbl.setObjectName("formLabel")
        time_desc_row.addWidget(hr_lbl)
        self.block_hours_spin = QSpinBox()
        self.block_hours_spin.setObjectName("blockSpin")
        self.block_hours_spin.setRange(0, MAX_HOURS_PER_ENTRY)
        self.block_hours_spin.setValue(0)
        time_desc_row.addWidget(self.block_hours_spin)

        min_lbl = QLabel("Min")
        min_lbl.setObjectName("formLabel")
        time_desc_row.addWidget(min_lbl)
        self.block_minutes_spin = QSpinBox()
        self.block_minutes_spin.setObjectName("blockSpin")
        self.block_minutes_spin.setRange(0, 59)
        self.block_minutes_spin.setValue(30)
        time_desc_row.addWidget(self.block_minutes_spin)

        self.block_desc_edit = QLineEdit()
        self.block_desc_edit.setObjectName("blockInput")
        self.block_desc_edit.setPlaceholderText("Description")
        time_desc_row.addWidget(self.block_desc_edit, stretch=1)

        block_form.addLayout(time_desc_row)

        # Row 4: Day-of-week checkboxes + Add button
        days_btn_row = QHBoxLayout()
        days_btn_row.setSpacing(4)

        self.day_checkboxes = []
        for i, day_name in enumerate(DAY_NAMES[:5]):  # Mon-Fri
            cb = QCheckBox(day_name)
            cb.setObjectName("dayCheck")
            cb.setChecked(True)
            self.day_checkboxes.append((i, cb))
            days_btn_row.addWidget(cb)

        days_btn_row.addStretch(1)

        add_block_btn = QPushButton("+ Add Block")
        add_block_btn.setObjectName("addBtn")
        add_block_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_block_btn.clicked.connect(self.add_time_block)
        days_btn_row.addWidget(add_block_btn)

        block_form.addLayout(days_btn_row)

        blocks_card_layout.addLayout(block_form)

        # Separator
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setFrameShadow(QFrame.Shadow.Sunken)
        blocks_card_layout.addWidget(sep2)

        # Block list scroll
        block_scroll = QScrollArea()
        block_scroll.setWidgetResizable(True)
        self.block_scroll_content = QWidget()
        self.block_scroll_layout = QVBoxLayout(self.block_scroll_content)
        self.block_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.block_scroll_layout.setSpacing(6)
        block_scroll.setWidget(self.block_scroll_content)
        blocks_card_layout.addWidget(block_scroll)

        # Total footer
        self.block_total_label = QLabel("")
        self.block_total_label.setObjectName("blockTotal")
        blocks_card_layout.addWidget(self.block_total_label)

        blocks_layout.addWidget(blocks_card, stretch=1)

        self.tabs.addTab(blocks_tab, "⏱  Time Blocking")

        # ═══════════════════════════════════════════
        # TAB 3: Day Entries (View & Edit)
        # ═══════════════════════════════════════════
        entries_tab = QWidget()
        entries_layout = QVBoxLayout(entries_tab)
        entries_layout.setContentsMargins(8, 12, 8, 8)
        entries_layout.setSpacing(10)

        entries_card = QFrame()
        entries_card.setObjectName("card")
        entries_card_layout = QVBoxLayout(entries_card)
        entries_card_layout.setSpacing(8)

        section_title4 = QLabel("📋  View & Edit Day Entries")
        section_title4.setObjectName("sectionTitle")
        entries_card_layout.addWidget(section_title4)

        hint_label = QLabel(f"Pick a date to view logged entries. Edit rows and save — the day must total {MAX_HOURS_PER_DAY} hours.")
        hint_label.setObjectName("entryHint")
        hint_label.setWordWrap(True)
        entries_card_layout.addWidget(hint_label)

        # Date picker row
        entry_date_row = QHBoxLayout()
        entry_date_row.setSpacing(8)

        date_lbl = QLabel("Date:")
        date_lbl.setObjectName("formLabel")
        entry_date_row.addWidget(date_lbl)

        self.entry_date_edit = QDateEdit()
        self.entry_date_edit.setCalendarPopup(True)
        self.entry_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.entry_date_edit.setDate(QDate.currentDate())
        self.entry_date_edit.setMaximumDate(QDate.currentDate())
        self.entry_date_edit.dateChanged.connect(self._load_entries_for_date)
        entry_date_row.addWidget(self.entry_date_edit)

        entry_date_row.addStretch(1)

        # Default project checkbox
        self.use_default_check = QCheckBox("Use default project")
        self.use_default_check.setObjectName("dayCheck")
        default_proj = get_default_project()
        self.use_default_check.setEnabled(default_proj is not None)
        self.use_default_check.setChecked(default_proj is not None)
        if not default_proj:
            self.use_default_check.setToolTip("Set a default project in Manage Projects first")
        else:
            self.use_default_check.setToolTip(f"Default: {default_proj}")
        self.use_default_check.toggled.connect(self._toggle_default_project)
        entry_date_row.addWidget(self.use_default_check)

        entries_card_layout.addLayout(entry_date_row)

        # Summary label
        self.entry_summary_label = QLabel("")
        self.entry_summary_label.setObjectName("entrySummary")
        entries_card_layout.addWidget(self.entry_summary_label)

        # Separator
        sep3 = QFrame()
        sep3.setFrameShape(QFrame.Shape.HLine)
        sep3.setFrameShadow(QFrame.Shadow.Sunken)
        entries_card_layout.addWidget(sep3)

        # Scroll area for entry rows
        entry_scroll = QScrollArea()
        entry_scroll.setWidgetResizable(True)
        self.entry_scroll_content = QWidget()
        self.entry_scroll_layout = QVBoxLayout(self.entry_scroll_content)
        self.entry_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.entry_scroll_layout.setSpacing(8)
        self.entry_scroll_layout.addStretch(1)
        entry_scroll.setWidget(self.entry_scroll_content)
        entries_card_layout.addWidget(entry_scroll)

        # Button row: Add Task + Save Day
        entry_btn_row = QHBoxLayout()
        entry_btn_row.setSpacing(8)

        add_task_btn = QPushButton("+ Add Task")
        add_task_btn.setObjectName("addTaskBtn")
        add_task_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_task_btn.clicked.connect(self._add_blank_entry_row)
        entry_btn_row.addWidget(add_task_btn)

        entry_btn_row.addStretch(1)

        self.entry_save_btn = QPushButton("💾  Save Day")
        self.entry_save_btn.setObjectName("saveBtn")
        self.entry_save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.entry_save_btn.clicked.connect(self._save_day_entries)
        self.entry_save_btn.setEnabled(False)
        entry_btn_row.addWidget(self.entry_save_btn)

        entries_card_layout.addLayout(entry_btn_row)

        entries_layout.addWidget(entries_card, stretch=1)

        self.tabs.addTab(entries_tab, "📋  Day Entries")

        # ═══════════════════════════════════════════
        # TAB 4: Manage Projects
        # ═══════════════════════════════════════════
        projects_tab = QWidget()
        projects_layout = QVBoxLayout(projects_tab)
        projects_layout.setContentsMargins(8, 12, 8, 8)
        projects_layout.setSpacing(12)

        # Add project form
        proj_form_row = QHBoxLayout()
        proj_form_row.setSpacing(8)

        self.project_name_edit = QLineEdit()
        self.project_name_edit.setObjectName("projectInput")
        self.project_name_edit.setPlaceholderText("New project name")
        proj_form_row.addWidget(self.project_name_edit, stretch=1)

        add_proj_btn = QPushButton("+ Add Project")
        add_proj_btn.setObjectName("addBtn")
        add_proj_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        add_proj_btn.clicked.connect(self.add_new_project)
        proj_form_row.addWidget(add_proj_btn)

        projects_layout.addLayout(proj_form_row)

        # Separator
        sep4 = QFrame()
        sep4.setFrameShape(QFrame.Shape.HLine)
        sep4.setFrameShadow(QFrame.Shadow.Sunken)
        projects_layout.addWidget(sep4)

        # Project list
        proj_card = QFrame()
        proj_card.setObjectName("card")
        proj_card_layout = QVBoxLayout(proj_card)
        proj_card_layout.setSpacing(6)

        proj_scroll = QScrollArea()
        proj_scroll.setWidgetResizable(True)
        self.proj_scroll_content = QWidget()
        self.proj_scroll_layout = QVBoxLayout(self.proj_scroll_content)
        self.proj_scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.proj_scroll_layout.setSpacing(6)
        proj_scroll.setWidget(self.proj_scroll_content)
        proj_card_layout.addWidget(proj_scroll)

        projects_layout.addWidget(proj_card, stretch=1)
        self.tabs.addTab(projects_tab, "📁  Manage Projects")

        # ── Internal state for entry editor ──
        self.entry_rows = []
        self.entry_loading = False

        root.addWidget(self.tabs, stretch=1)

        self.populate()
        self._load_entries_for_date()

    # ── Data population ──

    def populate(self):
        self._populate_company_holidays()
        self._populate_personal_leaves()
        self._populate_time_blocks()
        self._populate_projects()

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
        target_widget = None

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
            
            if is_next:
                target_widget = container

        self.company_scroll_layout.addStretch(1)

        if target_widget:
            def _scroll_company():
                try:
                    spacing = self.company_scroll_layout.spacing()
                    # Scroll so that exactly 1 past item (approx target_widget.height() + spacing) is visible above the NEXT item
                    val = max(0, target_widget.y() - target_widget.height() - spacing)
                    self.company_scroll.verticalScrollBar().setValue(val)
                except RuntimeError:
                    pass
            QTimer.singleShot(100, _scroll_company)

    def _populate_personal_leaves(self):
        self._clear_layout(self.personal_scroll_layout)
        all_leaves = get_all_leaves()
        personal = sorted([l for l in all_leaves if l["type"] == "Personal Leave"], key=lambda x: x["date"])
        target_widget = None

        if not personal:
            empty = QLabel("No personal leaves set yet.")
            empty.setObjectName("emptyState")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.personal_scroll_layout.addWidget(empty)
        else:
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            next_upcoming_found = False

            for leave in personal:
                row = QHBoxLayout()
                row.setSpacing(8)

                date_str = leave["date"]
                is_past = date_str < today
                is_next = (not is_past and not next_upcoming_found)

                date_label = QLabel(date_str)
                if is_past:
                    date_label.setObjectName("holidayDatePast")
                elif is_next:
                    date_label.setObjectName("holidayDateUpcoming")
                else:
                    date_label.setObjectName("leaveDate")
                row.addWidget(date_label)

                reason = leave.get("name", "") or "Personal Leave"
                name_label = QLabel(reason)
                if is_past:
                    name_label.setObjectName("holidayPast")
                elif is_next:
                    name_label.setObjectName("holidayUpcoming")
                else:
                    name_label.setObjectName("leaveName")
                row.addWidget(name_label, stretch=1)
                
                if is_next:
                    badge = QLabel("NEXT")
                    badge.setObjectName("badge")
                    row.addWidget(badge)
                    next_upcoming_found = True
                elif is_past:
                    badge = QLabel("PAST")
                    badge.setObjectName("badgePast")
                    row.addWidget(badge)

                remove_btn = QPushButton("Remove")
                remove_btn.setObjectName("removeBtn")
                remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                date_val = leave["date"]
                remove_btn.clicked.connect(lambda checked, d=date_val: self.remove_personal_leave(d))
                row.addWidget(remove_btn)

                container = QWidget()
                container.setLayout(row)
                self.personal_scroll_layout.addWidget(container)

                if is_next:
                    target_widget = container

        self.personal_scroll_layout.addStretch(1)

        if target_widget:
            def _scroll_personal():
                try:
                    spacing = self.personal_scroll_layout.spacing()
                    # Scroll so that exactly 1 past item is visible above the NEXT item
                    val = max(0, target_widget.y() - target_widget.height() - spacing)
                    self.personal_scroll.verticalScrollBar().setValue(val)
                except RuntimeError:
                    pass
            QTimer.singleShot(100, _scroll_personal)

    def _populate_time_blocks(self):
        self._clear_layout(self.block_scroll_layout)
        blocks = get_time_blocks()

        if not blocks:
            empty = QLabel("No time blocks defined yet.")
            empty.setObjectName("emptyState")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.block_scroll_layout.addWidget(empty)
        else:
            for block in blocks:
                block_widget = QFrame()
                block_widget.setFrameShape(QFrame.Shape.StyledPanel)
                block_inner = QVBoxLayout(block_widget)
                block_inner.setContentsMargins(8, 4, 8, 4)
                block_inner.setSpacing(2)

                # Top row: name + duration
                top_row = QHBoxLayout()
                top_row.setSpacing(6)

                name_label = QLabel(block["description"])
                name_label.setObjectName("blockName" if block["enabled"] else "blockDisabled")
                top_row.addWidget(name_label)

                top_row.addStretch(1)

                duration_h = block["hours"]
                duration_m = block["minutes"]
                dur_text = f"{duration_h}h {duration_m}m" if duration_m else f"{duration_h}h"
                dur_label = QLabel(dur_text)
                dur_label.setObjectName("blockDuration" if block["enabled"] else "blockDisabled")
                top_row.addWidget(dur_label)

                block_inner.addLayout(top_row)

                # Middle row: activity + days
                mid_row = QHBoxLayout()
                mid_row.setSpacing(6)

                act_label = QLabel(block["activity"])
                act_label.setObjectName("blockDetail" if block["enabled"] else "blockDisabled")
                mid_row.addWidget(act_label)

                mid_row.addStretch(1)

                # Format days of week
                day_indices = [int(d.strip()) for d in block["days_of_week"].split(",") if d.strip()]
                day_labels = [DAY_NAMES[d] for d in day_indices if d < len(DAY_NAMES)]
                days_text = ", ".join(day_labels) if day_labels else "No days"
                days_label = QLabel(days_text)
                days_label.setObjectName("blockDays" if block["enabled"] else "blockDisabled")
                mid_row.addWidget(days_label)

                block_inner.addLayout(mid_row)

                # Bottom row: toggle + remove
                btn_row = QHBoxLayout()
                btn_row.setSpacing(6)
                btn_row.addStretch(1)

                toggle_btn = QPushButton("Enabled" if block["enabled"] else "Disabled")
                toggle_btn.setObjectName("toggleBtn" if block["enabled"] else "toggleBtnOff")
                toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                bid = block["id"]
                is_on = block["enabled"]
                toggle_btn.clicked.connect(lambda checked, b=bid, e=is_on: self.toggle_block(b, e))
                btn_row.addWidget(toggle_btn)

                remove_btn = QPushButton("Remove")
                remove_btn.setObjectName("removeBtn")
                remove_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                remove_btn.clicked.connect(lambda checked, b=bid: self.remove_block(b))
                btn_row.addWidget(remove_btn)

                block_inner.addLayout(btn_row)

                self.block_scroll_layout.addWidget(block_widget)

        self.block_scroll_layout.addStretch(1)

        # Update total
        total_min = get_total_blocked_minutes()
        t_h = total_min // 60
        t_m = total_min % 60
        self.block_total_label.setText(f"Total: {t_h}h {t_m}m / {MAX_HOURS_PER_DAY}h reserved daily")

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

    def add_time_block(self):
        project = self.block_project_combo.currentText().strip()
        activity = self.block_activity_combo.currentText().strip()
        hours = self.block_hours_spin.value()
        minutes = self.block_minutes_spin.value()
        description = self.block_desc_edit.text().strip()
        name = description  # Use description as the name in the database

        # Collect selected days
        selected_days = [str(idx) for idx, cb in self.day_checkboxes if cb.isChecked()]
        if not selected_days:
            show_box(self, QMessageBox.Icon.Warning, "Cannot Add", "Select at least one day of the week.")
            return
        days_of_week = ",".join(selected_days)

        try:
            add_time_block(name, project, activity, hours, minutes, description, days_of_week)
            # Clear form
            self.block_hours_spin.setValue(0)
            self.block_minutes_spin.setValue(30)
            self.block_desc_edit.clear()
            for _, cb in self.day_checkboxes:
                cb.setChecked(True)
            self._populate_time_blocks()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Cannot Add Block", str(exc))

    def toggle_block(self, block_id, currently_enabled):
        try:
            toggle_time_block(block_id, not currently_enabled)
            self._populate_time_blocks()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Toggle Failed", str(exc))

    def remove_block(self, block_id):
        if not ask_yes_no(self, "Remove Time Block", "Remove this time block?\n\nThis cannot be undone."):
            return
        try:
            delete_time_block(block_id)
            self._populate_time_blocks()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Cannot Remove", str(exc))

    # ── Day Entries tab actions ──

    def _entry_date_str(self):
        return self.entry_date_edit.date().toString("yyyy-MM-dd")

    def _clear_entry_rows(self):
        while self.entry_rows:
            row = self.entry_rows.pop()
            row.setParent(None)
            row.deleteLater()

    def _add_entry_row(self, entry=None):
        default_proj = get_default_project()
        use = self.use_default_check.isChecked() and default_proj is not None
        row = DayEntryRow(
            entry=entry,
            on_change=self._update_entry_summary,
            on_remove=self._remove_entry_row,
            default_project=default_proj,
            use_default=use,
        )
        insert_at = self.entry_scroll_layout.count() - 1  # before the stretch
        self.entry_scroll_layout.insertWidget(insert_at, row)
        self.entry_rows.append(row)
        return row

    def _add_blank_entry_row(self):
        self._add_entry_row()
        self._update_entry_summary()

    def _remove_entry_row(self, row):
        if len(self.entry_rows) <= 1:
            show_box(self, QMessageBox.Warning, "Day Entries", "Keep at least one task row.")
            return
        self.entry_rows.remove(row)
        row.setParent(None)
        row.deleteLater()
        self._update_entry_summary()

    def _toggle_default_project(self, checked):
        """Toggle default project on all existing rows (option A: override all)."""
        default_proj = get_default_project()
        for row in self.entry_rows:
            row.set_use_default(checked and default_proj is not None, default_proj)
        self._update_entry_summary()

    def _load_entries_for_date(self, _date=None):
        date_str = self._entry_date_str()
        self.entry_loading = True
        try:
            self._clear_entry_rows()
            entries = get_timesheet_entries_for_day(date_str)
            if not entries:
                self._add_entry_row()
            else:
                for entry in entries:
                    self._add_entry_row(entry)
        finally:
            self.entry_loading = False
        self._update_entry_summary()

    def _collect_entries(self):
        return [row.get_entry() for row in self.entry_rows]

    def _validate_entries(self):
        entries = self._collect_entries()
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
                return False, f"Each row must be ≤ {MAX_HOURS_PER_ENTRY}h."
            if minutes < 0 or minutes > 59:
                return False, "Minutes must be 0–59."
            if hours == 0 and minutes == 0:
                return False, "Each row needs a positive time."
            total_minutes += hours * 60 + minutes
        if total_minutes != MAX_HOURS_PER_DAY * 60:
            return False, f"Total must equal {MAX_HOURS_PER_DAY}h."
        return True, ""

    def _update_entry_summary(self):
        if self.entry_loading:
            return
        entries = self._collect_entries()
        total_minutes = sum(e.get("hours", 0) * 60 + e.get("minutes", 0) for e in entries)
        th = total_minutes // 60
        tm = total_minutes % 60
        valid, message = self._validate_entries()
        if valid:
            self.entry_summary_label.setText(f"Total: {th}h {tm}m / {MAX_HOURS_PER_DAY}h  ✓ Ready to save")
            self.entry_summary_label.setObjectName("entrySummary")
        else:
            self.entry_summary_label.setText(f"Total: {th}h {tm}m / {MAX_HOURS_PER_DAY}h  —  {message}")
            self.entry_summary_label.setObjectName("entrySummaryError")
        # Force style refresh after objectName change
        self.entry_summary_label.setStyleSheet(self.entry_summary_label.styleSheet())
        self.entry_summary_label.style().unpolish(self.entry_summary_label)
        self.entry_summary_label.style().polish(self.entry_summary_label)
        self.entry_save_btn.setEnabled(valid)

    def _save_day_entries(self):
        date_str = self._entry_date_str()
        entries = self._collect_entries()
        try:
            replace_timesheet_entries_for_day(date_str, entries)
            for activity in dict.fromkeys(e["activity"] for e in entries):
                record_recent_activity(activity)
            show_box(self, QMessageBox.Information, "Saved", f"Updated {date_str} successfully.")
            self._load_entries_for_date()
        except Exception as exc:
            show_box(self, QMessageBox.Critical, "Save Failed", str(exc))

    # ── Manage Projects tab actions ──

    def _populate_projects(self):
        self._clear_layout(self.proj_scroll_layout)
        projects = get_projects()

        if not projects:
            empty = QLabel("No projects defined. Add one above.")
            empty.setObjectName("emptyState")
            empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.proj_scroll_layout.addWidget(empty)
        else:
            default_proj = get_default_project()
            for name in projects:
                row = QHBoxLayout()
                row.setSpacing(8)

                label_text = f"{name}  [DEFAULT]" if name == default_proj else name
                label = QLabel(label_text)
                label.setObjectName("projectName")
                if name == default_proj:
                    pass
                row.addWidget(label, stretch=1)

                if name != default_proj:
                    default_btn = QPushButton("Set Default")
                    default_btn.setObjectName("defaultBtn")
                    default_btn.setCursor(Qt.CursorShape.PointingHandCursor)
                    default_btn.clicked.connect(lambda checked, n=name: self.set_default(n))
                    row.addWidget(default_btn)

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
                self.proj_scroll_layout.addWidget(container)

        self.proj_scroll_layout.addStretch(1)

    def _sync_project_updates(self):
        """Called after project changes to ensure all tabs are in sync."""
        self._populate_projects()
        self._populate_time_blocks()
        self._load_entries_for_date()

    def add_new_project(self):
        name = self.project_name_edit.text().strip()
        if not name:
            show_box(self, QMessageBox.Icon.Warning, "Add Project", "Project name cannot be empty.")
            return
        try:
            add_project(name)
            self.project_name_edit.clear()
            self._sync_project_updates()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Add Project", str(exc))

    def rename_project(self, old_name):
        from PyQt6.QtWidgets import QInputDialog
        new_name, ok = QInputDialog.getText(self, "Rename Project", f"New name for '{old_name}':", text=old_name)
        if not ok or not new_name.strip():
            return
        try:
            update_project(old_name, new_name.strip())
            self._sync_project_updates()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Rename Project", str(exc))

    def set_default(self, name):
        try:
            set_default_project(name)
            self._sync_project_updates()
        except ValueError as exc:
            show_box(self, QMessageBox.Icon.Warning, "Set Default", str(exc))

    def remove_project(self, name):
        try:
            delete_project(name)
            self._sync_project_updates()
        except ValueError as exc:
            if "entries reference it" in str(exc):
                dates = get_dates_with_project_entries(name)
                other_projects = [p for p in get_projects() if p != name]
                dialog = ProjectDeleteDialog(name, dates, other_projects, self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    action, value = dialog.get_action()
                    if action == "reassign":
                        reassign_project_entries(name, value)
                        force_delete_project(name)
                        self._sync_project_updates()
                    elif action == "rename":
                        if not value:
                            show_box(self, QMessageBox.Icon.Warning, "Rename", "Name cannot be empty.")
                            return
                        try:
                            update_project(name, value)
                            self._sync_project_updates()
                        except ValueError as e:
                            show_box(self, QMessageBox.Icon.Warning, "Rename", str(e))
                    elif action == "review":
                        self.navigate_to_date(value)
            else:
                show_box(self, QMessageBox.Icon.Warning, "Delete Project", str(exc))

    def navigate_to_date(self, date_str):
        """Navigate the Day Entries tab to a specific date (used by project deletion review)."""
        self.tabs.setCurrentIndex(2)  # Day Entries tab
        qdate = QDate.fromString(date_str, "yyyy-MM-dd")
        if qdate.isValid():
            self.entry_date_edit.setDate(qdate)


class ProjectDeleteDialog(QDialog):
    STYLE = """
        ProjectDeleteDialog { background: #1e1e2e; }
        QLabel { color: #cdd6f4; font-size: 13px; font-family: 'Segoe UI', sans-serif; }
        QLabel#dialogTitle { font-size: 16px; font-weight: 700; }
        QRadioButton { color: #cdd6f4; font-size: 13px; font-family: 'Segoe UI', sans-serif; }
        QRadioButton::indicator { width: 14px; height: 14px; border-radius: 7px; border: 1px solid #585b70; background: transparent; }
        QRadioButton::indicator:checked { background: #89b4fa; border: 1px solid #89b4fa; }
        QComboBox { background: #45475a; color: #cdd6f4; border: 1px solid #585b70; border-radius: 4px; padding: 4px 8px; }
        QComboBox::drop-down { subcontrol-origin: padding; subcontrol-position: top right; width: 20px; border-left: none; }
        QComboBox QAbstractItemView { background: #313244; color: #cdd6f4; selection-background-color: #45475a; border: 1px solid #585b70; border-radius: 4px; }
        QLineEdit { background: #45475a; color: #cdd6f4; border: 1px solid #585b70; border-radius: 4px; padding: 4px 8px; }
        QPushButton { background: #45475a; color: #cdd6f4; border: none; border-radius: 6px; padding: 6px 16px; font-weight: 600; }
        QPushButton:hover { background: #585b70; }
        QPushButton#primaryBtn { background: #f38ba8; color: #1e1e2e; }
        QPushButton#primaryBtn:hover { background: #fab387; }
    """

    def __init__(self, project_name, dates, other_projects, parent=None):
        super().__init__(parent)
        self.project_name = project_name
        self.dates = dates
        self.other_projects = other_projects
        
        self.setWindowTitle("Delete Project")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        # self.setStyleSheet(self.STYLE)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        
        title = QLabel(f"Project '{project_name}' is in use")
        title.setObjectName("dialogTitle")
        layout.addWidget(title)
        
        sub = QLabel(f"There are {len(dates)} days with timesheet entries referencing this project. You cannot delete it directly. Choose an action:")
        sub.setWordWrap(True)
        layout.addWidget(sub)
        
        # Options
        self.radio_reassign = QRadioButton("Reassign entries to existing project")
        self.combo_projects = QComboBox()
        self.combo_projects.addItems(other_projects)
        self.combo_projects.setEnabled(bool(other_projects))
        self.radio_reassign.setEnabled(bool(other_projects))
        
        self.radio_rename = QRadioButton("Rename project entirely")
        self.edit_rename = QLineEdit()
        self.edit_rename.setPlaceholderText("New project name")
        
        self.radio_review = QRadioButton("Review entries day-by-day")
        review_lbl = QLabel(f"(Starts at {dates[0]})")
        # review_lbl.setStyleSheet("color: #6c7086; font-size: 11px;")
        
        # Add to layout
        opt1 = QHBoxLayout()
        opt1.addWidget(self.radio_reassign)
        opt1.addWidget(self.combo_projects)
        opt1.addStretch()
        
        opt2 = QHBoxLayout()
        opt2.addWidget(self.radio_rename)
        opt2.addWidget(self.edit_rename)
        opt2.addStretch()
        
        opt3 = QHBoxLayout()
        opt3.addWidget(self.radio_review)
        opt3.addWidget(review_lbl)
        opt3.addStretch()
        
        layout.addLayout(opt1)
        layout.addLayout(opt2)
        layout.addLayout(opt3)
        
        # Auto-select
        if other_projects:
            self.radio_reassign.setChecked(True)
        else:
            self.radio_rename.setChecked(True)
            
        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)
        
        self.proceed_btn = QPushButton("Proceed")
        self.proceed_btn.setObjectName("primaryBtn")
        self.proceed_btn.clicked.connect(self.accept)
        btn_row.addWidget(self.proceed_btn)
        
        layout.addLayout(btn_row)
        
    def get_action(self):
        if self.radio_reassign.isChecked():
            return "reassign", self.combo_projects.currentText()
        elif self.radio_rename.isChecked():
            return "rename", self.edit_rename.text().strip()
        elif self.radio_review.isChecked():
            return "review", self.dates[0]
        return None, None


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

        # Auto-insert time blocks as soon as the app starts (2s delay for init)
        QTimer.singleShot(2000, self._startup_insert_blocks)

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

    def _startup_insert_blocks(self):
        """Called once shortly after app launch to insert today's time blocks immediately."""
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        insert_time_blocks_for_day(today_str)

    def show_manual_log(self):
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")

        # Ensure time blocks are inserted before calculating remaining hours
        insert_time_blocks_for_day(today_str)

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
        self.show_holiday_window()
        if self.holiday_window:
            self.holiday_window.tabs.setCurrentIndex(3)

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

        # Auto-insert time blocks for today (idempotent — skips if already inserted)
        insert_time_blocks_for_day(today_str)

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
    app.setWindowIcon(create_app_icon())

    controller = TimesheetController(app)
    app.controller = controller

    sys.exit(app.exec())


if __name__ == "__main__":
    main()