"""
Bolles Schedule – All Classes Reminder Application
--------------------------------------------------

This script monitors *all* class periods each day and prompts the
student five minutes before the end of every class block.  It is
particularly useful for students who want gentle nudges for every
course rather than only a specific period.

The behaviour is similar to the single‑period version in
``student_app.py``.  When the reminder window appears the user may
choose whether or not they have homework.  Selecting “Yes” will open
Outlook and create a calendar appointment for the next occurrence of
that period.  The reminder on the appointment is set to one hour
before the next class begins.

Class times mirror the Upper School schedule found on the Bolles
website【846292384622532†L351-L364】.  Wednesday timings are
slightly different【846292384622532†L351-L379】.  See the module
documentation of ``student_app.py`` for a detailed explanation of the
time blocks.
"""

from __future__ import annotations

import json
import os
import threading
import time
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Dict, Optional, Tuple, Any

# Import shared schedule utilities from the single‑period application.  These
# functions expose the rotation (letter days), period ordering and time slot
# definitions.
import student_app

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    tk = None  # type: ignore
    messagebox = None  # type: ignore

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    win32com = None  # type: ignore
    pythoncom = None  # type: ignore

# Import tray icon dependencies if available
try:
    from PIL import Image
    import pystray
except ImportError:
    Image = None  # type: ignore
    pystray = None  # type: ignore


class AllClassesApp:
    """Application that monitors all class periods each day.

    This class can honour a user‑selected lunch option so that
    midday times reflect whether the student has first or second lunch.
    """

    def __init__(self, check_interval: int = 30, lunch_option: Optional[str] = None) -> None:
        self.check_interval = check_interval
        # Use provided lunch option or fall back to the global setting in student_app
        self.lunch_option = lunch_option or student_app.LUNCH_OPTION
        self.running = True
        self.thread: Optional[threading.Thread] = None
        # Track last reminder per period to avoid multiple prompts within the
        # same day.  Keys are (date, period_number) and values indicate
        # whether the reminder has already fired today.
        self.triggered: Dict[Tuple[date, int], bool] = {}

    def start(self) -> None:
        """Start the background monitoring thread without blocking.

        The caller is responsible for keeping the main thread alive and
        shutting down the application by clearing ``running`` when
        appropriate.
        """
        if self.thread is None or not self.thread.is_alive():
            self.thread = threading.Thread(target=self._run_loop, daemon=True)
            self.thread.start()

    def _run_loop(self) -> None:
        while self.running:
            now = datetime.now()
            # Reset triggers at midnight
            self.triggered = {
                (d, p): fired
                for (d, p), fired in self.triggered.items()
                if d == now.date()
            }
            if now.weekday() < 5:  # Monday–Friday
                # Determine the letter day and period ordering
                letter = student_app.get_letter_day(now.date())
                order = student_app.PERIOD_ORDER.get(letter, [])
                # Pass lunch option so that midday times reflect the student's schedule
                time_slots = student_app.get_time_slots(now.weekday() == 2, self.lunch_option)
                # Iterate through the five time slots and associated periods
                for idx, class_time in enumerate(time_slots):
                    if idx >= len(order):
                        break
                    period_number = order[idx]
                    reminder_time = class_time.reminder_time(5)
                    key = (now.date(), period_number)
                    if (
                        now.time().hour == reminder_time.hour
                        and now.time().minute == reminder_time.minute
                        and not self.triggered.get(key, False)
                    ):
                        self.triggered[key] = True
                        self.show_reminder(now.date(), period_number, class_time)
                        # Avoid immediate re‑prompt
                        time.sleep(60)
                        break
            time.sleep(self.check_interval)

    def show_reminder(self, class_date: date, period_index: int, class_time: student_app.ClassTime) -> None:
        """
        Display the reminder window for the given period.

        If the user confirms they have homework, look up the next meeting
        of the same period (respecting the lunch option) and open an
        Outlook appointment window.
        """
        if tk is None:
            return
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        result = messagebox.askyesno(
            title="Homework Reminder",
            message=f"Class period {period_index} is ending soon.  Do you have homework?",
            parent=root,
        )
        root.destroy()
        if result:
            # Compute the next occurrence of this period using the rotation
            next_info = student_app.get_next_class(period_index, class_date, self.lunch_option)
            if next_info is not None:
                next_date, next_time = next_info
                start_dt = datetime.combine(next_date, next_time.start)
                end_dt = datetime.combine(next_date, next_time.end)
                self.create_outlook_event(period_index, start_dt, end_dt)

    def create_outlook_event(self, period_index: int, start_dt: datetime, end_dt: datetime) -> None:
        """Create a calendar appointment in Outlook for the given period.

        A COM initialisation is attempted on the current thread to
        ensure that Outlook automation works reliably.  If pywin32 is
        unavailable or an exception occurs, the error is ignored so that
        the reminder application continues running.
        """
        if win32com is None or pythoncom is None:
            return
        try:
            pythoncom.CoInitialize()
            outlook = win32com.Dispatch('Outlook.Application')
            appt = outlook.CreateItem(1)
            appt.Start = start_dt.strftime("%m/%d/%Y %H:%M")
            appt.End = end_dt.strftime("%m/%d/%Y %H:%M")
            appt.Subject = f"Homework – Period {period_index}"
            appt.Body = "Enter your assignment details here."
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 60
            appt.Display(True)
        except Exception:
            pass


def load_config(path: Path) -> Dict[str, Any]:
    """Read JSON configuration or return an empty dict."""
    if path.exists():
        try:
            with path.open('r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(path: Path, data: Dict[str, Any]) -> None:
    """Write configuration data to disk."""
    with path.open('w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


def main() -> None:
    """Entry point for the all classes reminder application.

    On first run the user is prompted to select their lunch pattern (first
    or second).  The choice is persisted in a configuration file so
    subsequent executions use the same setting.  The lunch option is
    passed to the scheduler so that midday times are calculated
    correctly.
    """
    base_dir = Path(__file__).resolve().parent
    config_path = base_dir / "all_classes_config.json"
    config: Dict[str, Any] = load_config(config_path)
    lunch_option: Optional[str] = config.get("lunch_option")
    if lunch_option not in {"1", "2"}:
        # Ask the user for lunch option
        lunch_option = student_app.ask_lunch()
        config["lunch_option"] = lunch_option
        save_config(config_path, config)
    # Set global lunch option so that schedule functions pick it up
    student_app.set_lunch_option(lunch_option)
    # Copy the current script/executable into the Startup folder so
    # students do not need to manually copy it.  The name reflects
    # the Skaphysics branding.
    student_app.ensure_startup_copy('Skaphysics Homework Reminder')
    app = AllClassesApp(lunch_option=lunch_option)
    # Start the monitoring thread (non‑blocking)
    app.start()
    # Display a tray icon with quit option
    student_app.setup_tray_icon(app)  # reuse helper from student_app
    # Keep the main thread alive until the app stops running
    try:
        while app.running:
            time.sleep(1)
    except KeyboardInterrupt:
        app.running = False


if __name__ == "__main__":
    main()