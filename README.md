Skaphysics Homework Reminder
============================

This folder contains the source code and build artefacts for the **Skaphysics Homework Reminder** utility.  The project is designed for students at the Bolles School and provides lightweight pop‑up reminders before each class ends.  When a student acknowledges that they have homework, an Outlook appointment is created for the next meeting of that class period with a one‑hour reminder.  A system‑tray icon makes it easy to quit the program.

The code here can be packaged for distribution in two steps:

1. **Build an executable** from the Python script using *PyInstaller*.
2. **Compile an installer** using *Inno Setup* which copies the executable into the user’s Windows Startup folder so that it runs automatically when the student logs in.

Prerequisites
-------------

1. Install Python 3.11 or newer from [python.org](https://www.python.org/).  When installing, tick “Add Python to PATH”.
2. Install the Python dependencies:

```
pip install pyinstaller pywin32 pystray pillow
```

Building the executable
-----------------------

From a command prompt in this folder, run the following command to produce a single‑file Windows executable for the main reminder app:

```
pyinstaller --onefile --windowed --icon=skaphysics_icon.png student_app.py --name "Skaphysics Homework Reminder"
```

This command creates a `dist/` subdirectory containing **`Skaphysics Homework Reminder.exe`**.  The `--windowed` flag suppresses the console window; the `--icon` flag embeds the tray icon you supplied.

You can repeat the same command for the other variants, substituting `all_classes_app.py` or `middle_school_app.py` and adjusting the `--name` parameter accordingly.

Packaging with Inno Setup
-------------------------

The provided Inno Setup script (`SkaphysicsHomeworkReminder.iss`) automates installation.  It places the compiled executable into the user’s Startup folder so that the program launches automatically after reboot.  To create the installer:

1. Install [Inno Setup](https://jrsoftware.org/isinfo.php) on a Windows machine.
2. Open the file **`SkaphysicsHomeworkReminder.iss`** in the Inno Setup IDE.
3. Adjust the `#define SourceExe` directive near the top of the script to point to the path of your compiled `Skaphysics Homework Reminder.exe` file.  For example:

   ```
   #define SourceExe "C:\\Users\\YourName\\Path\\to\\dist\\Skaphysics Homework Reminder.exe"
   ```

4. Press *Compile* in the IDE.  Inno Setup will generate an installer EXE, typically named `Skaphysics Homework Reminder Installer.exe`.

Distributing to students
------------------------

Share the compiled installer (`Skaphysics Homework Reminder Installer.exe`) with your students.  When they run the installer, it silently copies the reminder program into their Startup folder and adds a Start Menu entry.  On first run the program will prompt them for their physics period and lunch option.  Thereafter it runs quietly in the background and can be exited by right‑clicking its tray icon and selecting **Quit**.

Code signing
------------

To avoid security warnings on Windows, consider obtaining a code‑signing certificate from a trusted provider (e.g., DigiCert, Sectigo).  Once you have the certificate, sign both the compiled executable and the installer using Microsoft’s `signtool`:

```
signtool sign /f yourcertificate.pfx /p password /tr http://timestamp.digicert.com /td sha256 /fd sha256 "Skaphysics Homework Reminder.exe"
signtool sign /f yourcertificate.pfx /p password /tr http://timestamp.digicert.com /td sha256 /fd sha256 "Skaphysics Homework Reminder Installer.exe"
```

Signing the files reduces the likelihood of anti‑virus false positives and shows your name as the verified publisher.

Troubleshooting
---------------

If students run into issues where nothing seems to happen, note that the reminder only triggers five minutes before the end of a class block.  To test quickly, launch the program a few minutes before a block ends or temporarily change the `minutes_before` parameter in the source code (search for `reminder_time(5)` and change `5` to `1`).

For more details on the schedule times and rotation logic, see the comments within `student_app.py`, `all_classes_app.py`, and `middle_school_app.py`.