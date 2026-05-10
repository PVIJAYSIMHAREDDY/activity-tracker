; Inno Setup script for Daily Activity Tracker
; Run with: ISCC.exe build\windows_installer.iss

#define AppName      "Daily Activity Tracker"
#define AppVersion   "1.0.0"
#define AppPublisher "Activity Tracker"
#define AppExeName   "ActivityTracker.exe"
#define AppId        "{{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"

[Setup]
AppId={#AppId}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\ActivityTracker
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=build
OutputBaseFilename=ActivityTracker_Setup
SetupIconFile=..\icon.ico
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
UninstallDisplayIcon={app}\ActivityTracker.exe
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";   Description: "{cm:CreateDesktopIcon}";   GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startupicon";   Description: "Launch at Windows startup"; GroupDescription: "Startup:"; Flags: unchecked

[Files]
; All PyInstaller output files
Source: "dist\ActivityTracker\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#AppName}";        Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall";          Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; Launch at startup (optional)
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
  ValueType: string; ValueName: "{#AppName}"; \
  ValueData: """{app}\{#AppExeName}"""; Flags: uninsdeletevalue; Tasks: startupicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; \
  Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "taskkill"; Parameters: "/f /im ActivityTracker.exe"; Flags: runhidden; RunOnceId: "KillApp"

[Code]
// Check if WebView2 runtime is installed (required for the app window)
function IsWebView2Installed: Boolean;
var
  Version: String;
begin
  Result := RegQueryStringValue(HKLM,
    'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
    'pv', Version) and (Version <> '');
  if not Result then
    Result := RegQueryStringValue(HKCU,
      'Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
      'pv', Version) and (Version <> '');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then begin
    if not IsWebView2Installed then begin
      MsgBox('Tip: For the best experience, install the Microsoft Edge WebView2 Runtime from:' + #13#10 +
             'https://developer.microsoft.com/en-us/microsoft-edge/webview2/' + #13#10 + #13#10 +
             '(It may already be installed on Windows 11 or if you have Microsoft Office/Edge)',
             mbInformation, MB_OK);
    end;
  end;
end;
