import os
import tempfile

from PySide6.QtCore import QLockFile


def acquire_single_instance_lock(app_name="TimesheetTracker"):
    lock_path = os.path.join(tempfile.gettempdir(), f"{app_name}.lock")
    lock_file = QLockFile(lock_path)

    if not lock_file.tryLock(0):
        return None

    return lock_file