"""
Skaphysics Middle School Schedule Reminder Application (updated)
----------------------------------------------------------------

This module provides reminder functionality for the Bolles Middle School
schedule with lunch-specific prompts.  On first run, students select
their grade (6, 7 or 8) and whether they take first or second lunch.
Those preferences are saved in a configuration file.  When a class
block is about to end, the application asks if the student has
homework.  If yes, it creates an Outlook appointment titled
“Nth Period HW due” (where N is the ordinal position of the block) with
a 60‑minute reminder and then opens the Schoology calendar in the
default browser.  The Outlook appointment body is left blank so
students record assignment details only in Schoology.

Class times for the middle school are derived from the 2025‑26
schedule posted on the Bolles website【810082011620849†L330-L345】【810082011620849†L347-L368】.

Lunch windows vary by grade and week.  On regular days (most of the
year) grade 8 eats at 12:05–12:25, grade 6 at 12:25–12:40, and
grade 7 at 12:40–13:00【810082011620849†L330-L345】.  Some weeks have special
schedules, e.g. Oct 6–10 and Nov 17–21, when lunch times start
earlier【810082011620849†L347-L368】.  The tables below encode these ranges
and can be customised easily.
"""

from __future__ import annotations

import json
import os
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime, date, timedelta, time as dttime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

try:
    import tkinter as tk  # type: ignore
    from tkinter import simpledialog, messagebox  # type: ignore
except ImportError:
    tk = None  # type: ignore
    simpledialog = None  # type: ignore
    messagebox = None  # type: ignore

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    win32com = None  # type: ignore
    pythoncom = None  # type: ignore

import webbrowser

try:
    from PIL import Image  # type: ignore
    import pystray  # type: ignore
except ImportError:
    Image = None  # type: ignore
    pystray = None  # type: ignore

# Reuse upper school rotation utilities for letter day calculation and
# period ordering.  The middle school follows the same seven‑day
# rotation (A–G) as the upper school.
from student_app import (
    LETTERS,
    PERIOD_ORDER,
    get_letter_day as upper_get_letter_day,
    int_to_ordinal,
    ensure_startup_copy,
    setup_tray_icon,
    SCHOOLOGY_CALENDAR_URL,
)  # type: ignore

# For the middle school, we still need the period input prompt.  Reuse a
# simplified version of the ask_period helper from the original
# middle school module.
def ask_period() -> int:
    """
    Prompt the user to enter a class period (1–7) using a simple
    dialog.  Loops until a valid number is entered.  If the user
    cancels, exit the program.
    """
    if tk is None or simpledialog is None or messagebox is None:
        raise RuntimeError("Tkinter is required to prompt for the period")
    root = tk.Tk()
    root.withdraw()
    while True:
        result = simpledialog.askstring(
            "Class Period", "Enter your class period number (1–7):"
        )
        if result is None:
            messagebox.showinfo("Setup", "Setup cancelled.")
            root.destroy()
            sys.exit(0)
        try:
            value = int(result)
            if 1 <= value <= 7:
                root.destroy()
                return value
        except ValueError:
            pass
        messagebox.showerror("Invalid input", "Please enter a number between 1 and 7.")


def get_letter_day(current_date: date) -> str:
    """Return the letter day for the given date using the upper school helper."""
    return upper_get_letter_day(current_date)


@dataclass
class ClassTime:
    start: dttime
    end: dttime

    def reminder_time(self, minutes_before: int = 5) -> dttime:
        dt_end = datetime.combine(date.today(), self.end)
        return (dt_end - timedelta(minutes=minutes_before)).time()


# ---------------------------------------------------------------------------
# Lunch windows configuration
#
# Each entry maps a grade number to a start/end time (in minutes since
# midnight).  These windows represent the lunch break for that grade
# during the mid‑day block.  Students are prompted to indicate if they
# take first or second lunch, but the class schedule itself does not
# vary between first and second lunch; rather, these windows are
# provided for completeness and future customisation.
# ---------------------------------------------------------------------------

@dataclass
class MSLunchWindow:
    start_min: int
    end_min: int


def hm(h: int, m: int) -> int:
    """Convert hours and minutes to minutes since midnight."""
    return h * 60 + m


# Regular days lunch windows (Mon/Tue/Thu/Fri and most weeks).  Times
# from the Bolles 2025‑26 middle school planner【810082011620849†L330-L345】.
MS_LUNCH_WINDOWS_REGULAR: Dict[int, MSLunchWindow] = {
    8: MSLunchWindow(hm(12, 5), hm(12, 25)),   # 12:05–12:25
    6: MSLunchWindow(hm(12, 25), hm(12, 40)),  # 12:25–12:40
    7: MSLunchWindow(hm(12, 40), hm(13, 0)),   # 12:40–13:00
}

# Alternative lunch windows for special weeks.  Each override is
# identified by a date range and provides a separate mapping.
MS_LUNCH_WINDOWS_ALT_OCT6: Dict[int, MSLunchWindow] = {
    8: MSLunchWindow(hm(11, 55), hm(12, 15)),   # 11:55–12:15
    6: MSLunchWindow(hm(12, 15), hm(12, 30)),   # 12:15–12:30
    7: MSLunchWindow(hm(12, 30), hm(12, 45)),   # 12:30–12:45
}

MS_LUNCH_WINDOWS_ALT_NOV17: Dict[int, MSLunchWindow] = {
    8: MSLunchWindow(hm(11, 45), hm(12, 5)),    # 11:45–12:05
    6: MSLunchWindow(hm(12, 5), hm(12, 20)),    # 12:05–12:20
    7: MSLunchWindow(hm(12, 20), hm(12, 40)),   # 12:20–12:40
}

# Each override is a tuple of (start_month, start_day) to (end_month, end_day) and the
# corresponding lunch window mapping.  You can extend this list as
# needed for other special schedules.
MS_LUNCH_OVERRIDES: List[Tuple[Tuple[int, int], Tuple[int, int], Dict[int, MSLunchWindow]]] = [
    ((10, 6), (10, 10), MS_LUNCH_WINDOWS_ALT_OCT6),   # Oct 6–10
    ((11, 17), (11, 21), MS_LUNCH_WINDOWS_ALT_NOV17), # Nov 17–21
]


def get_ms_lunch_window_for(date_obj: date, grade: int) -> MSLunchWindow:
    """
    Return the lunch window for the given date and grade.  If the date
    falls within one of the override ranges, use that; otherwise
    return the regular mapping.  If a grade is not found, fall back
    to a default (grade 8) window.
    """
    m, d = date_obj.month, date_obj.day
    for (start_md, end_md, table) in MS_LUNCH_OVERRIDES:
        (sm, sd), (em, ed) = start_md, end_md
        in_range = (
            (m > sm or (m == sm and d >= sd)) and
            (m < em or (m == em and d <= ed))
        )
        if in_range and grade in table:
            return table[grade]
    return MS_LUNCH_WINDOWS_REGULAR.get(grade, MS_LUNCH_WINDOWS_REGULAR[8])


# ---------------------------------------------------------------------------
# Schedule definitions
#
# The middle school has different class times on Wednesday and
# Thursday compared to Monday/Tuesday/Friday.  Each function returns a
# list of ``ClassTime`` objects representing the five blocks for a
# particular day of the week.
# ---------------------------------------------------------------------------

def get_schedule_for_day(day_of_week: int) -> List[ClassTime]:
    """
    Return a list of class times for the given weekday index (0=Mon … 6=Sun).
    The middle school schedule varies by day.  Wednesday starts late;
    Thursday starts early.  On other days (Mon/Tue/Fri) the schedule
    follows a standard pattern【810082011620849†L330-L345】【810082011620849†L347-L368】.
    """
    if day_of_week == 2:  # Wednesday
        return [
            ClassTime(dttime(8, 55), dttime(9, 50)),
            ClassTime(dttime(9, 55), dttime(10, 50)),
            ClassTime(dttime(11, 10), dttime(12, 5)),
            ClassTime(dttime(13, 5), dttime(14, 0)),
            ClassTime(dttime(14, 5), dttime(15, 0)),
        ]
    elif day_of_week == 3:  # Thursday
        return [
            ClassTime(dttime(8, 30), dttime(9, 25)),
            ClassTime(dttime(9, 30), dttime(10, 25)),
            ClassTime(dttime(11, 10), dttime(12, 5)),
            ClassTime(dttime(13, 5), dttime(14, 0)),
            ClassTime(dttime(14, 5), dttime(15, 0)),
        ]
    else:  # Monday, Tuesday, Friday
        return [
            ClassTime(dttime(8, 45), dttime(9, 40)),
            ClassTime(dttime(9, 45), dttime(10, 40)),
            ClassTime(dttime(11, 10), dttime(12, 5)),
            ClassTime(dttime(13, 5), dttime(14, 0)),
            ClassTime(dttime(14, 5), dttime(15, 0)),
        ]


def get_next_class_ms(period: int, from_date: date) -> Optional[Tuple[date, ClassTime]]:
    """
    Find the next date and time slot when ``period`` meets in the middle school.

    Starting one day after ``from_date``, search forward (skipping
    weekends) until the letter rotation includes the requested period.
    Return a tuple of (date, ClassTime).  This function reuses the
    upper school period ordering and letter rotation to determine
    which periods meet on each letter day.
    """
    next_date = from_date + timedelta(days=1)
    while True:
        # Skip weekends
        if next_date.weekday() >= 5:
            next_date += timedelta(days=1)
            continue
        letter = get_letter_day(next_date)
        order = PERIOD_ORDER.get(letter, [])
        if period in order:
            idx = order.index(period)
            slots = get_schedule_for_day(next_date.weekday())
            if idx < len(slots):
                return next_date, slots[idx]
        next_date += timedelta(days=1)
    return None


# ---------------------------------------------------------------------------
# User prompts for grade and lunch option
#
# These helper functions present dialogs to the student on first run
# to capture their grade level and lunch preference.  The results are
# stored in the configuration file for subsequent launches.
# ---------------------------------------------------------------------------

def prompt_for_ms_grade_and_lunch() -> Dict[str, object]:
    """
    Prompt the user for their grade (6/7/8) and lunch option.

    Students indicate their grade and whether they take first or
    second lunch.  The returned dictionary has keys ``grade`` (int)
    and ``lunch_choice`` ("first" or "second").  If no lunch split
    applies (e.g. single lunch day), lunch_choice defaults to "first".
    """
    if tk is None or simpledialog is None or messagebox is None:
        raise RuntimeError("Tkinter is required for grade and lunch prompts")
    root = tk.Tk()
    root.withdraw()
    grade: Optional[int] = None
    # Ask for grade until valid
    while grade not in (6, 7, 8):
        res = simpledialog.askstring(
            "Grade Level", "Enter your grade (6, 7, or 8):"
        )
        if res is None:
            messagebox.showinfo("Setup", "Setup cancelled.")
            root.destroy()
            sys.exit(0)
        try:
            g = int(res)
            if g in (6, 7, 8):
                grade = g
        except ValueError:
            pass
        if grade is None:
            messagebox.showerror("Invalid input", "Please enter 6, 7, or 8.")
    # Ask whether the day uses split lunch
    uses_split = messagebox.askyesno(
        "Lunch Schedule",
        "Does your grade use split lunch (first/second lunch)?",
    )
    lunch_choice = "first"
    if uses_split:
        # Ask for first or second lunch
        while lunch_choice.lower() not in ("first", "second"):
            ans = simpledialog.askstring(
                "Lunch Choice",
                "Do you have first or second lunch? (type 'first' or 'second')",
            )
            if ans is None:
                messagebox.showinfo("Setup", "Setup cancelled.")
                root.destroy()
                sys.exit(0)
            if ans:
                lunch_choice = ans.strip().lower()
        # Standardise
        lunch_choice = lunch_choice.lower()
    root.destroy()
    return {"grade": grade, "lunch_choice": lunch_choice}


class StudentReminderApp:
    """Monitor a single class period for middle school."""

    def __init__(self, period: int, check_interval: int = 30) -> None:
        self.period = period
        self.check_interval = check_interval
        self.running = True
        self.thread: Optional[threading.Thread] = None

    def start(self) -> None:
        """Start the background monitoring thread without blocking."""
        if self.thread is None or not self.thread.is_alive():
            self.thread = threading.Thread(target=self._run_loop, daemon=True)
            self.thread.start()

    def _run_loop(self) -> None:
        while self.running:
            now = datetime.now()
            # Only operate on weekdays
            if now.weekday() < 5:
                letter = get_letter_day(now.date())
                order = PERIOD_ORDER.get(letter, [])
                # If the student's period meets today, determine its time slot
                if self.period in order:
                    idx = order.index(self.period)
                    schedule = get_schedule_for_day(now.weekday())
                    if idx < len(schedule):
                        class_time = schedule[idx]
                        reminder = class_time.reminder_time(5)
                        if (
                            now.time().hour == reminder.hour
                            and now.time().minute == reminder.minute
                        ):
                            self.show_reminder(now.date(), class_time)
                            # Wait a minute to avoid duplicate prompts
                            time.sleep(60)
                            continue
            time.sleep(self.check_interval)

    def show_reminder(self, class_date: date, class_time: ClassTime) -> None:
        if tk is None:
            return
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        result = messagebox.askyesno(
            title="Homework Reminder",
            message="Do you have homework for tonight?",
            parent=root,
        )
        root.destroy()
        if result:
            # Determine ordinal of the current class block (1st–5th)
            try:
                schedule = get_schedule_for_day(class_date.weekday())
                ordinal = int_to_ordinal(schedule.index(class_time) + 1)
            except Exception:
                ordinal = None
            # Compute the next class occurrence using the rotation
            next_info = get_next_class_ms(self.period, class_date)
            if next_info is not None:
                next_date, next_time = next_info
                start_dt = datetime.combine(next_date, next_time.start)
                end_dt = datetime.combine(next_date, next_time.end)
                subject = f"{ordinal} Period HW due" if ordinal else None
                self.create_outlook_event(start_dt, end_dt, subject=subject)

    def create_outlook_event(
        self,
        start_dt: datetime,
        end_dt: datetime,
        subject: Optional[str] = None,
    ) -> None:
        """
        Create a calendar appointment with a 60‑minute reminder.

        The subject defaults to "Homework" if not provided.  The body
        is left blank, and after displaying the appointment the
        Schoology calendar is opened.  Any Outlook errors are ignored.
        """
        if win32com is None or pythoncom is None:
            return
        try:
            pythoncom.CoInitialize()
            outlook = win32com.Dispatch('Outlook.Application')
            appt = outlook.CreateItem(1)
            appt.Start = start_dt.strftime("%m/%d/%Y %H:%M")
            appt.End = end_dt.strftime("%m/%d/%Y %H:%M")
            appt.Subject = subject or "Homework"
            appt.Body = ""
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 60
            appt.Display(True)
            try:
                webbrowser.open_new_tab(SCHOOLOGY_CALENDAR_URL)
            except Exception:
                pass
        except Exception:
            pass


class AllClassesReminderApp:
    """Monitor every period each day for middle school."""

    def __init__(self, check_interval: int = 30) -> None:
        self.check_interval = check_interval
        self.running = True
        self.thread: Optional[threading.Thread] = None
        # Track whether a reminder has already fired for (date, period)
        self.triggered: Dict[Tuple[date, int], bool] = {}

    def start(self) -> None:
        if self.thread is None or not self.thread.is_alive():
            self.thread = threading.Thread(target=self._run_loop, daemon=True)
            self.thread.start()

    def _run_loop(self) -> None:
        while self.running:
            now = datetime.now()
            # Discard triggers from previous days
            self.triggered = {
                (d, p): fired
                for (d, p), fired in self.triggered.items()
                if d == now.date()
            }
            if now.weekday() < 5:
                letter = get_letter_day(now.date())
                order = PERIOD_ORDER.get(letter, [])
                schedule = get_schedule_for_day(now.weekday())
                for idx, class_time in enumerate(schedule):
                    if idx >= len(order):
                        break
                    period_number = order[idx]
                    reminder = class_time.reminder_time(5)
                    key = (now.date(), period_number)
                    if (
                        now.time().hour == reminder.hour
                        and now.time().minute == reminder.minute
                        and not self.triggered.get(key, False)
                    ):
                        self.triggered[key] = True
                        self.show_reminder(now.date(), period_number, class_time)
                        # Wait a minute to avoid duplicate prompts
                        time.sleep(60)
                        break
            time.sleep(self.check_interval)

    def show_reminder(
        self,
        class_date: date,
        period_index: int,
        class_time: ClassTime,
    ) -> None:
        if tk is None:
            return
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        result = messagebox.askyesno(
            title="Homework Reminder",
            message=(
                f"Class period {period_index} is ending soon.  "
                "Do you have homework?"
            ),
            parent=root,
        )
        root.destroy()
        if result:
            try:
                schedule = get_schedule_for_day(class_date.weekday())
                ordinal = int_to_ordinal(schedule.index(class_time) + 1)
            except Exception:
                ordinal = None
            next_info = get_next_class_ms(period_index, class_date)
            if next_info is not None:
                next_date, next_time = next_info
                start_dt = datetime.combine(next_date, next_time.start)
                end_dt = datetime.combine(next_date, next_time.end)
                subject = f"{ordinal} Period HW due" if ordinal else None
                self.create_outlook_event(period_index, start_dt, end_dt, subject=subject)

    def create_outlook_event(
        self,
        period_index: int,
        start_dt: datetime,
        end_dt: datetime,
        subject: Optional[str] = None,
    ) -> None:
        if win32com is None or pythoncom is None:
            return
        try:
            pythoncom.CoInitialize()
            outlook = win32com.Dispatch('Outlook.Application')
            appt = outlook.CreateItem(1)
            appt.Start = start_dt.strftime("%m/%d/%Y %H:%M")
            appt.End = end_dt.strftime("%m/%d/%Y %H:%M")
            appt.Subject = subject if subject else f"Homework – Period {period_index}"
            appt.Body = ""
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 60
            appt.Display(True)
            try:
                webbrowser.open_new_tab(SCHOOLOGY_CALENDAR_URL)
            except Exception:
                pass
        except Exception:
            pass


def run_student_app() -> None:
    """
    Entry point for the single-period middle school reminder.  It
    prompts for the student's period, grade and lunch preference on
    first run, starts the monitoring thread, and installs itself into
    the Startup folder so it runs at login.
    """
    base_dir = Path(__file__).resolve().parent
    config_path = base_dir / "middle_school_config.json"
    config: Dict[str, object] = {}
    if config_path.exists():
        try:
            with config_path.open('r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception:
            config = {}
    period = config.get('period')  # type: ignore
    grade = config.get('grade')    # type: ignore
    lunch_choice = config.get('lunch_choice')  # type: ignore
    # Prompt for period if missing
    if period is None:
        period = ask_period()
        config['period'] = period
    # Prompt for grade/lunch if missing
    if grade is None or lunch_choice is None:
        info = prompt_for_ms_grade_and_lunch()
        config['grade'] = info['grade']
        config['lunch_choice'] = info['lunch_choice']
    # Save configuration
    with config_path.open('w', encoding='utf-8') as f:
        json.dump(config, f, indent=2)
    # Start the reminder app
    app = StudentReminderApp(period=int(period))
    # Copy script/executable into startup folder (reuse helper from student_app)
    ensure_startup_copy('Skaphysics Homework Reminder')
    # Start monitoring thread (non-blocking)
    app.start()
    # Create tray icon to allow quitting (reuse helper from upper school)
    setup_tray_icon(app)
    # Keep main thread alive until user quits
    try:
        while app.running:
            time.sleep(1)
    except KeyboardInterrupt:
        app.running = False


def run_all_classes_app() -> None:
    """Entry point for monitoring all periods in the middle school."""
    base_dir = Path(__file__).resolve().parent
    config_path = base_dir / "middle_school_config.json"
    # Ensure grade/lunch information is recorded even if unused
    config: Dict[str, object] = {}
    if config_path.exists():
        try:
            with config_path.open('r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception:
            config = {}
    grade = config.get('grade')    # type: ignore
    lunch_choice = config.get('lunch_choice')  # type: ignore
    if grade is None or lunch_choice is None:
        info = prompt_for_ms_grade_and_lunch()
        config['grade'] = info['grade']
        config['lunch_choice'] = info['lunch_choice']
        with config_path.open('w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)
    app = AllClassesReminderApp()
    ensure_startup_copy('Skaphysics Homework Reminder')
    app.start()
    setup_tray_icon(app)
    try:
        while app.running:
            time.sleep(1)
    except KeyboardInterrupt:
        app.running = False


if __name__ == '__main__':
    run_student_app()