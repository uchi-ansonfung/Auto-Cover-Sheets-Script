; Inno Setup script — Windows full installer for Automatic Exhibit Cover Sheets
;
; Expects a staged payload produced by scripts/build_windows_full.ps1:
;   dist\windows-full\
;     coversheets.exe
;     tesseract\...
;     ghostscript\...
;
; Build:
;   iscc /DMyAppVersion=0.8.3 /DPayloadDir=..\..\dist\windows-full installer\windows\coversheets.iss

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#ifndef PayloadDir
  #define PayloadDir "..\..\dist\windows-full"
#endif

#define MyAppName "Automatic Exhibit Cover Sheets"
#define MyAppPublisher "Anson Fung"
#define MyAppURL "https://github.com/uchi-ansonfung/Auto-Cover-Sheets-Script"
#define MyAppExeName "coversheets.exe"

[Setup]
AppId={{A7C0E5D1-9B42-4F6E-8C31-5D6E7F8091A2}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={localappdata}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; Per-user install: no admin elevation (easier for non-tech users / locked PCs).
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\..\dist
OutputBaseFilename=coversheets-{#MyAppVersion}-windows-x64-setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
SetupLogging=yes
CloseApplications=yes
RestartApplications=no
; Info shown on finished page.
InfoAfterFile=

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main app (PyInstaller one-file with pikepdf + ocrmypdf baked in).
Source: "{#PayloadDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Bundled OCR engines (required for the OCR checkbox).
Source: "{#PayloadDir}\tesseract\*"; DestDir: "{app}\tesseract"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#PayloadDir}\ghostscript\*"; DestDir: "{app}\ghostscript"; Flags: ignoreversion recursesubdirs createallsubdirs
; Optional readme next to the app.
Source: "{#PayloadDir}\README-INSTALLED.txt"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Messages]
WelcomeLabel2=This will install [name/ver] on your computer.%n%nThis full build includes Linearize support and OCR (English) so you do not need to install Python, Tesseract, or Ghostscript yourself.%n%nIt is recommended that you close all other applications before continuing.
FinishedLabel=Setup has finished installing [name] on your computer.%n%nTo add cover sheets: open the app, choose Open Folder or Add PDFs, edit labels if needed, then click Generate Cover Sheets.
