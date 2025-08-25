; Skaphysics Homework Reminder Inno Setup script
;
; This installer places the compiled reminder executable into the
; current user's Startup folder so that it launches automatically
; when they log into Windows.  It also adds a Start Menu shortcut.
;
; Before compiling this script you must define the SourceExe
; constant to point to the compiled executable produced by
; PyInstaller.  See the accompanying README for instructions.

#define SourceExe "C:\Users\SkapetisD\OneDrive - The Bolles School\Apps\Homework Alert\dist\Skaphysics Homework Reminder.exe"

[Setup]
AppName=Skaphysics Homework Reminder
AppVersion=1.0
AppPublisher=Skaphysics
DefaultDirName={userappdata}\SkaphysicsHomeworkReminder
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=no
OutputBaseFilename=Skaphysics Homework Reminder Installer
; Use ZIP compression instead of LZMA to reduce anti‑virus heuristics.
Compression=zip
SolidCompression=no
; Ensure the installer runs with the current user's privileges.  This
; avoids UAC prompts and allows writing into per‑user locations like
; {userstartup} without generating a warning.
PrivilegesRequired=lowest

[Files]
; Copy the compiled EXE into the Startup folder.  If an older
; version exists it will be overwritten silently.
Source: "{#SourceExe}"; DestDir: "{userstartup}"; DestName: "Skaphysics Homework Reminder.exe"; Flags: ignoreversion overwritereadonly

[Icons]
; Add a Start Menu shortcut (optional).  Note that the program
; itself resides in the Startup folder, so this shortcut points
; there as well.
Name: "{userprograms}\Skaphysics Homework Reminder"; Filename: "{userstartup}\Skaphysics Homework Reminder.exe"; WorkingDir: "{userstartup}"

[Run]
; Launch the reminder immediately after installation finishes.  This
; allows the student to configure their period and lunch option.
Filename: "{userstartup}\Skaphysics Homework Reminder.exe"; Description: "Start Skaphysics Homework Reminder"; Flags: postinstall nowait