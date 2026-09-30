[Setup]
AppId={{B8AB3AE5-B4B6-4B77-A17C-347CB9F564E1}
AppName=Payroll Desk
AppVersion=1.0.0
AppPublisher=Payroll Desk
DefaultDirName={localappdata}\Programs\PayrollDesk
DefaultGroupName=Payroll Desk
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
OutputDir=..\dist
OutputBaseFilename=PayrollDesk-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\PayrollDesk.exe
CloseApplications=yes
SetupLogging=yes

[Files]
Source: "..\dist\PayrollDesk\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Payroll Desk"; Filename: "{app}\PayrollDesk.exe"
Name: "{autodesktop}\Payroll Desk"; Filename: "{app}\PayrollDesk.exe"

[Run]
Filename: "{app}\PayrollDesk.exe"; Description: "Open Payroll Desk"; Flags: nowait postinstall skipifsilent
