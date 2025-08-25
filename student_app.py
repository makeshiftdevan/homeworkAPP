"""
Bolles Schedule Student Application
----------------------------------

This script implements a lightweight reminder application for students at
the Bolles School.  When first executed it prompts the student to enter
the period number (1–7) during which they attend physics.  From that
point forward the application runs quietly in the background and
displays an on‑top reminder a few minutes before each class ends.

If the student indicates they have homework, the script launches
Microsoft Outlook (Office 365) and pre‑creates a calendar appointment
for the next occurrence of the same period.  The reminder on the
appointment is set to trigger one hour before the next class.  The
student is still responsible for filling in the details of the
assignment before saving the event.

Two variants of this script exist:

  * ``student_app.py`` – monitors only the student’s selected period.
  * ``all_classes_app.py`` – monitors every class period (1–5) each
    day and prompts for each.

The schedule definitions are deliberately simple.  They reflect the
times published on the Bolles schedule pages for the 2025‑26 academic
year.  Upper School class blocks on Monday, Tuesday, Thursday and
Friday run 8:30–9:30, 9:35–10:35, 11:25–12:25, 13:10–14:10 and
14:15–15:15【846292384622532†L351-L364】.  On Wednesdays the start time is
shifted to 9:15 am and 10:20 am for the first two blocks but otherwise
retains the same structure【846292384622532†L351-L379】.  Middle School
times differ slightly and are defined in ``middle_school_app.py``.

The application stores its configuration in a JSON file named
``student_config.json`` located in the same directory as the script.

This code is designed for Windows and uses ``win32com.client`` to
interact with Outlook.  If ``pywin32`` is not installed the calendar
integration will silently fail, and the reminder window will still
appear.
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
from typing import Dict, List, Tuple, Optional

try:
    import tkinter as tk
    from tkinter import simpledialog, messagebox
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

# Attempt to import pystray and PIL for system tray icon support.  If
# unavailable the application will still run but without a tray icon.
try:
    from PIL import Image
    import pystray
except ImportError:
    Image = None  # type: ignore
    pystray = None  # type: ignore


@dataclass
class ClassTime:
    """Representation of a single class period timing."""

    start: dttime
    end: dttime

    def reminder_time(self, minutes_before: int = 5) -> dttime:
        """Return the time at which the reminder should fire."""
        dt_end = datetime.combine(date.today(), self.end)
        dt_rem = dt_end - timedelta(minutes=minutes_before)
        return dt_rem.time()


def load_config(config_path: Path) -> Dict[str, Any]:
    """Load the configuration from disk or return an empty dictionary.

    The configuration may contain values of different types (e.g., int
    for the period and str for the lunch option), so the return type
    allows any value.
    """
    if config_path.exists():
        try:
            with config_path.open("r", encoding="utf-8") as f:
                data: Dict[str, Any] = json.load(f)
                return data
        except Exception:
            return {}
    return {}


def save_config(config_path: Path, data: Dict[str, Any]) -> None:
    """Save configuration data to disk."""
    with config_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def ensure_startup_copy(app_name: str) -> None:
    """Copy the running script or executable into the Windows Startup folder.

    This function determines the current process path and, if an
    identically named file does not already exist in the user's
    Startup directory, copies it there.  Errors during the copy are
    silently ignored.

    Parameters
    ----------
    app_name: str
        Base filename (without extension) to use when copying into
        startup.  The extension is derived from the current executable
        (``.exe`` if running as a frozen PyInstaller bundle, otherwise
        ``.py``).
    """
    try:
        # Determine the source path of the currently executing file.
        src_path = Path(sys.argv[0]).resolve()
        # Determine the desired destination filename and path.
        extension = src_path.suffix
        # Map .pyz to .py for interactive runs (e.g., zipapp); treat other
        # unknown extensions as .py
        if extension.lower() not in {'.exe', '.py'}:
            extension = '.py'
        dest_filename = app_name + extension
        startup_dir = Path(os.environ.get('APPDATA', '')) / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs' / 'Startup'
        dest_path = startup_dir / dest_filename
        # If the destination file does not exist and the startup directory
        # exists, copy the file.  Silently ignore errors.
        if not dest_path.exists() and startup_dir.exists():
            import shutil
            shutil.copy2(src_path, dest_path)
    except Exception:
        # Ignore all failures; absence of the copy merely means the app
        # will not start automatically.
        pass


def setup_tray_icon(app: "ReminderApp") -> None:
    """Create and display a system tray icon for the homework reminder.

    If ``pystray`` or ``PIL`` is not available the function does
    nothing.  The tray icon uses the ``skaphysics_icon.png`` image
    located alongside this script.  Right‑clicking the icon will
    display a menu with a Quit option that stops the reminder
    application and removes the icon.

    Parameters
    ----------
    app: ReminderApp
        The running reminder application.  Its ``running`` flag will
        be cleared when the user chooses to quit from the tray menu.
    """
    if pystray is None or Image is None:
        return
    try:
        icon_path = Path(__file__).resolve().parent / 'skaphysics_icon.png'
        # Load and resize the image for the system tray (32x32 is typical)
        image = Image.open(icon_path)
        # Ensure the image has an alpha channel (required by some systems)
        if image.mode != 'RGBA':
            image = image.convert('RGBA')
        # Resize to 64x64 for high DPI clarity; pystray will downsample
        image = image.resize((64, 64), Image.LANCZOS)
        # Define an action for the Quit menu item
        def on_quit(icon: pystray.Icon, item: pystray.MenuItem) -> None:
            app.running = False
            try:
                icon.visible = False
                icon.stop()
            except Exception:
                pass
        menu = pystray.Menu(pystray.MenuItem('Quit', on_quit))
        tray_icon = pystray.Icon(name='Skaphysics Homework Reminder', icon=image, title='Skaphysics Homework Reminder', menu=menu)
        # Start the icon in a separate thread so it doesn't block
        tray_icon.run_detached()
    except Exception:
        # Fail silently if tray icon cannot be created
        pass


def ask_period() -> int:
    """Prompt the user for their physics period using a simple dialog."""
    if tk is None or simpledialog is None:
        raise RuntimeError("Tkinter is required to prompt for the period")
    root = tk.Tk()
    root.withdraw()
    while True:
        result = simpledialog.askstring(
            "Class Period",
            "Enter your physics period number (1‑7):",
            parent=root,
        )
        if result is None:
            # user cancelled; exit gracefully
            sys.exit(0)
        try:
            value = int(result)
            if 1 <= value <= 7:
                root.destroy()
                return value
        except ValueError:
            pass
        messagebox.showerror("Invalid input", "Please enter a number between 1 and 7.")

def ask_lunch() -> str:
    """Prompt the user to select their lunch option (first or second).

    Returns
    -------
    str
        "1" if the user has first lunch (midday class after lunch),
        "2" if the user has second lunch (midday class before lunch).
    """
    if tk is None or simpledialog is None:
        # If we cannot prompt graphically, fall back to the global default
        return LUNCH_OPTION
    root = tk.Tk()
    root.withdraw()
    while True:
        result = simpledialog.askstring(
            "Lunch Option",
            "Select your lunch pattern:\n1 – first lunch (class 12:05–13:05)\n2 – second lunch (class 11:25–12:25)",
            parent=root,
        )
        if result is None:
            # user cancelled; use default and exit
            root.destroy()
            return LUNCH_OPTION
        cleaned = result.strip()
        if cleaned in {"1", "2"}:
            root.destroy()
            return cleaned
        # invalid input – show an error and continue loop
        messagebox.showerror("Invalid input", "Please enter 1 or 2.")


def get_time_slots(is_wednesday: bool, lunch_option: Optional[str] = None) -> List[ClassTime]:
    """
    Return the five class time slots for a given day of the week.

    The Bolles Upper School operates on a rotating seven‑day schedule
    where five of seven periods meet each day.  While the *order* of
    periods depends on the letter day, the times of each block are
    fixed.  Monday, Tuesday, Thursday and Friday follow one schedule
    and Wednesdays follow another【846292384622532†L351-L364】.  On
    Monday/Tuesday/Thursday/Friday the blocks run 8:30–9:30,
    9:35–10:35, 11:25–12:25 (or 12:05–13:05 for first lunch),
    13:10–14:10 and 14:15–15:15.  On Wednesdays the first two blocks
    shift to 9:15–10:15 and 10:20–11:20 but the remaining three blocks
    stay the same【846292384622532†L351-L379】.

    Parameters
    ----------
    is_wednesday: bool
        True if the day is a Wednesday (index 2), otherwise False.
    lunch_option: Optional[str], optional
        Either "1" or "2" to indicate first or second lunch
        respectively.  If ``None``, the module‑level ``LUNCH_OPTION``
        value is used.  This option determines whether the midday
        class (third block) meets before or after lunch.

    Returns
    -------
    List[ClassTime]
        A list of five ``ClassTime`` objects corresponding to the
        start and end times of the day's blocks.
    """
    selected = lunch_option or LUNCH_OPTION
    if selected == "1":
        midday_start = dttime(12, 5)
        midday_end = dttime(13, 5)
    else:
        midday_start = dttime(11, 25)
        midday_end = dttime(12, 25)
    if not is_wednesday:
        return [
            ClassTime(dttime(8, 30), dttime(9, 30)),
            ClassTime(dttime(9, 35), dttime(10, 35)),
            ClassTime(midday_start, midday_end),
            ClassTime(dttime(13, 10), dttime(14, 10)),
            ClassTime(dttime(14, 15), dttime(15, 15)),
        ]
    else:
        return [
            ClassTime(dttime(9, 15), dttime(10, 15)),
            ClassTime(dttime(10, 20), dttime(11, 20)),
            ClassTime(midday_start, midday_end),
            ClassTime(dttime(13, 10), dttime(14, 10)),
            ClassTime(dttime(14, 15), dttime(15, 15)),
        ]

# -----------------------------------------------------------------------------
# Letter day rotation
# The school uses a seven‑letter rotation (A–G) where each letter
# specifies which five of the seven periods meet that day.  This mapping
# derives from the 2025–26 planner.  For example, an A day schedules
# periods 1–5 in order, a B day schedules periods 6,7,1,2,3 and so on.

LETTERS = ["A", "B", "C", "D", "E", "F", "G"]

# Mapping from letter day to the sequence of periods that meet (1‑7).
PERIOD_ORDER: Dict[str, List[int]] = {
    "A": [1, 2, 3, 4, 5],
    "B": [6, 7, 1, 2, 3],
    "C": [4, 5, 6, 7, 1],
    "D": [2, 3, 4, 5, 6],
    "E": [7, 1, 2, 3, 4],
    "F": [5, 6, 7, 1, 2],
    "G": [3, 4, 5, 6, 7],
}

# First day of the academic year (A day)
FIRST_DAY: date = date(2025, 8, 14)

# -----------------------------------------------------------------------------
# Lunch option configuration
#
# Some classes have different midday schedules depending on whether the
# teacher has first or second lunch.  Students with first lunch attend
# their midday class after lunch (12:05–13:05) while those with second
# lunch meet before lunch (11:25–12:25).  This global variable controls
# which pattern is used when calculating class times.  It is set in
# ``main`` based on user input or configuration.  The default value is
# "2" (second lunch) for backward compatibility.
LUNCH_OPTION: str = "2"

def set_lunch_option(option: str) -> None:
    """Update the global lunch option used by schedule calculations.

    Parameters
    ----------
    option: str
        Must be either "1" (first lunch) or "2" (second lunch).  Any
        other value is ignored.
    """
    global LUNCH_OPTION
    if option in {"1", "2"}:
        LUNCH_OPTION = option

def get_letter_day(current_date: date) -> str:
    """Return the letter day (A–G) corresponding to ``current_date``.

    The rotation begins on ``FIRST_DAY`` which is defined as an A day.
    Only weekdays (Monday–Friday) advance the rotation; weekends do not.
    If ``current_date`` is before the rotation start, the function
    returns "A" to provide a safe default.
    """
    if current_date <= FIRST_DAY:
        return LETTERS[0]
    # count weekdays between FIRST_DAY and current_date
    weekday_count = 0
    day = FIRST_DAY
    while day < current_date:
        day += timedelta(days=1)
        if day.weekday() < 5:
            weekday_count += 1
    index = weekday_count % len(LETTERS)
    return LETTERS[index]

def get_next_class(period: int, from_date: date, lunch_option: Optional[str] = None) -> Optional[Tuple[date, ClassTime]]:
    """
    Find the next date and time slot when ``period`` meets.

    Starting one day after ``from_date``, search forward (skipping
    weekends) until the letter day rotation includes the requested
    period.  Return a tuple of (date, ClassTime) for the next class.
    If no class is found (should not happen), return ``None``.

    Parameters
    ----------
    period: int
        The period number (1–7) to search for.
    from_date: datetime.date
        The date after which to search for the next class.
    lunch_option: Optional[str], optional
        Overrides the module‑level ``LUNCH_OPTION`` for calculating
        midday times.  See ``get_time_slots`` for details.
    """
    next_date = from_date + timedelta(days=1)
    while True:
        # Skip weekends (Saturday=5, Sunday=6)
        if next_date.weekday() >= 5:
            next_date += timedelta(days=1)
            continue
        letter = get_letter_day(next_date)
        order = PERIOD_ORDER.get(letter, [])
        if period in order:
            idx = order.index(period)
            slots = get_time_slots(next_date.weekday() == 2, lunch_option)
            if idx < len(slots):
                return next_date, slots[idx]
        next_date += timedelta(days=1)
    return None


class ReminderApp:
    """Main application class that monitors class times and shows reminders."""

    def __init__(self, period: int, check_interval: int = 30, lunch_option: Optional[str] = None) -> None:
        """
        Initialize the reminder application.

        Parameters
        ----------
        period: int
            The student's physics period to monitor.
        check_interval: int, optional
            Interval in seconds between checks of the current time.  A
            lower value results in more timely reminders but uses more
            CPU.
        lunch_option: Optional[str], optional
            The lunch option to use when calculating class times.  If
            ``None``, the module‑level ``LUNCH_OPTION`` is used.  See
            ``get_time_slots`` for details.
        """
        self.period = period
        self.check_interval = check_interval
        # use provided lunch_option or fall back to global setting
        self.lunch_option = lunch_option or LUNCH_OPTION
        self.running = True
        self.thread: Optional[threading.Thread] = None

    def start(self) -> None:
        """Start the background monitoring thread.

        This method spawns a daemon thread to execute the main loop and
        returns immediately.  The caller is responsible for keeping
        Python alive (e.g., via ``while app.running``) and for shutting
        down the application by clearing the ``running`` flag.
        """
        # Avoid starting multiple threads if already running
        if self.thread is None or not self.thread.is_alive():
            self.thread = threading.Thread(target=self._run_loop, daemon=True)
            self.thread.start()

    def _run_loop(self) -> None:
        """Main loop that checks the time and triggers reminders."""
        while self.running:
            now = datetime.now()
            # Only consider weekdays (0=Mon, 4=Fri)
            if now.weekday() < 5:
                # Determine the current letter day and the order of periods
                letter = get_letter_day(now.date())
                order = PERIOD_ORDER.get(letter, [])
                # If the student's selected period meets today, find its time slot
                if self.period in order:
                    slot_index = order.index(self.period)
                    time_slots = get_time_slots(now.weekday() == 2, self.lunch_option)
                    if slot_index < len(time_slots):
                        class_time = time_slots[slot_index]
                        reminder_time = class_time.reminder_time(5)
                        # Check if the current time matches the reminder time
                        if (
                            now.time().hour == reminder_time.hour
                            and now.time().minute == reminder_time.minute
                        ):
                            self.show_reminder(now.date(), class_time)
                            # Avoid duplicate prompts within the same minute
                            time.sleep(60)
                            continue
            time.sleep(self.check_interval)

    def show_reminder(self, class_date: date, class_time: ClassTime) -> None:
        """
        Display the reminder window.

        When the user clicks “Yes” the application searches for the next
        occurrence of the selected period using the letter rotation and
        creates an Outlook appointment for that class.  The student can
        then enter the assignment details.  If the user clicks “No” the
        reminder simply closes.
        """
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
            # Find the next date/time when this period meets
            next_info = get_next_class(self.period, class_date, self.lunch_option)
            if next_info is not None:
                next_date, next_time = next_info
                start_dt = datetime.combine(next_date, next_time.start)
                end_dt = datetime.combine(next_date, next_time.end)
                self.create_outlook_event(start_dt, end_dt)

    def create_outlook_event(self, start_dt: datetime, end_dt: datetime) -> None:
        """Create a calendar appointment in Outlook with a 60‑minute reminder.

        This method attempts to initialise COM for the current thread
        before interacting with Outlook.  If ``pywin32`` is not
        available or an exception occurs, the error is swallowed so
        that the reminder application continues to run without
        interruption.
        """
        if win32com is None or pythoncom is None:
            return
        try:
            # Each thread that uses COM must initialize it.  Without
            # calling CoInitialize the Dispatch call may silently fail.
            pythoncom.CoInitialize()
            outlook = win32com.Dispatch('Outlook.Application')
            appt = outlook.CreateItem(1)  # 1=olAppointmentItem
            appt.Start = start_dt.strftime("%m/%d/%Y %H:%M")
            appt.End = end_dt.strftime("%m/%d/%Y %H:%M")
            appt.Subject = "Homework"
            appt.Body = "Enter your assignment details here."
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = 60
            # Display the appointment window so the student can modify it
            appt.Display(True)
        except Exception:
            pass


def main() -> None:
    # Determine the directory of the running script
    base_dir = Path(__file__).resolve().parent
    config_path = base_dir / "student_config.json"
    config: Dict[str, Any] = load_config(config_path)
    # Retrieve existing configuration values
    period: Optional[int] = config.get("period")
    lunch_option: Optional[str] = config.get("lunch_option")  # may be None
    # Prompt for period if not already stored
    if period is None:
        period = ask_period()
        config["period"] = period
    # Prompt for lunch option if not already stored
    # Default to first lunch ("1") for physics by design
    if lunch_option not in {"1", "2"}:
        # Suggest first lunch as default but allow user choice
        lunch_option = ask_lunch()
        config["lunch_option"] = lunch_option
    # Persist updated configuration
    save_config(config_path, config)
    # Set the global lunch option so that helper functions pick it up
    set_lunch_option(lunch_option)
    # Copy the executable or script into the Startup folder so the app
    # launches automatically after reboot.  The name reflects the
    # Skaphysics branding.
    ensure_startup_copy('Skaphysics Homework Reminder')
    # Create the reminder application instance with explicit lunch_option
    app = ReminderApp(period=period, lunch_option=lunch_option)
    # Start the background monitoring thread (non‑blocking)
    app.start()
    # Create a tray icon with a quit option so the user can close the app
    setup_tray_icon(app)
    # Keep the main thread alive until the reminder is stopped
    try:
        while app.running:
            time.sleep(1)
    except KeyboardInterrupt:
        app.running = False


if __name__ == "__main__":
    main()