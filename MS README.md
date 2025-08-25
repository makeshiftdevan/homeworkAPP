Skaphysics Middle School Update
==============================

This update introduces grade and lunch selection to the middle school
variant of the **Skaphysics Homework Reminder**.  On first run,
students are asked for their class period *and* their grade (6, 7 or 8)
and whether they take first or second lunch.  These preferences are
stored in a `middle_school_config.json` file so the program does not
prompt again on subsequent launches.

After acknowledging that they have homework, the program now
automatically creates an Outlook calendar appointment titled
“Nth Period HW due” with a one‑hour reminder and leaves the body
blank.  It then opens the Schoology calendar in the default web
browser so that students can enter assignment details once, avoiding
duplication between Outlook and Schoology.

Installation and Packaging
-------------------------

1. **Compile the executable** using PyInstaller.  Run the following
   command in the folder containing `middle_school_app.py`:

   ```
   pyinstaller --onefile --windowed --icon=skaphysics_icon.png middle_school_app.py --name "Skaphysics Middle School Reminder"
   ```

   This produces a `dist/Skaphysics Middle School Reminder.exe` file.

2. **Create the installer** with Inno Setup using the provided
   `SkaphysicsHomeworkReminder.iss` script.  Open it in Inno Setup and
   set `#define SourceExe` to the path of your compiled middle school
   executable.  The script copies the program into the user’s Startup
   folder and adds a Start Menu entry so it launches automatically
   each time the student logs in.

3. **Distribute** the installer to students.  When they run it,
   the setup silently installs the reminder.  On first run the
   student will be asked for their grade, lunch and class period.

Customising Lunch Times
-----------------------

The lunch window tables are defined at the top of
`middle_school_app.py` in the `MS_LUNCH_WINDOWS_*` constants.  They
capture the start and end times (in minutes since midnight) for each
grade on regular days and for special weeks (e.g. Oct 6–10 and
Nov 17–21).  If your schedule changes, update these tables or add
entries to `MS_LUNCH_OVERRIDES` with the appropriate date ranges.

Schoology URL
-------------

By default the program opens `https://app.schoology.com/calendar` after
creating the Outlook appointment.  If your school uses a subdomain
such as `https://bolles.schoology.com/calendar`, change the
`SCHOOLOGY_CALENDAR_URL` constant in `student_app.py`.  All modules
import this constant, so the change takes effect across the entire
project.

Security Considerations
-----------------------

To minimise Windows Defender false positives, the Inno Setup script
uses ZIP compression and runs at the lowest privilege.  For further
reduction of security warnings, consider purchasing a code‑signing
certificate and signing both the executable and installer.  If
Defender still flags the installer, submit it to Microsoft as a
false positive for review.