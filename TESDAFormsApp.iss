; Inno Setup Script for TESDA Forms App
; Generated for .NET 10.0 Windows Forms Application

#define MyAppName "TESDA Forms App"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "TESDA"
#define MyAppExeName "TESDAFormsApp.exe"
#define SourcePath "bin\Release\net10.0-windows\win-x64\publish"

[Setup]
AppId={{90E8F08E-2F3A-4B1C-8D7E-5F9A6B2C1E3D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName=C:\TESDAForms
DefaultGroupName={#MyAppName}
OutputDir=bin\Release\Installers
OutputBaseFilename=TESDAFormsApp-{#MyAppVersion}-Installer
Compression=lzma
SolidCompression=yes
WizardStyle=modern
LicenseFile=
InfoBeforeFile=
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
SetupIconFile=Resources\tesda-logo.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
ChangesEnvironment=no

; Require Windows 10 or later
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Copy all files from the publish folder
Source: "{#SourcePath}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Explicitly copy Templates folder
Source: "{#SourcePath}\Templates\*"; DestDir: "{app}\Templates"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\tesda-logo.ico"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\tesda-logo.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: dirifempty; Name: "{app}"
