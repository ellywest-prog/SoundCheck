[Setup]
AppName=Soundcheck
AppVersion=1.0.1
AppPublisher=Keski
DefaultDirName={autopf}\soundcheck
DefaultGroupName=Soundcheck
UninstallDisplayIcon={app}\Soundcheck.exe
Compression=lzma2
SolidCompression=yes
OutputDir=dist
OutputBaseFilename=SoundcheckSetup
SetupIconFile=icon.ico
DisableDirPage=no
UsePreviousAppDir=no
DisableProgramGroupPage=yes
PrivilegesRequired=admin

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\Soundcheck\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Soundcheck"; Filename: "{app}\Soundcheck.exe"
Name: "{autodesktop}\Soundcheck"; Filename: "{app}\Soundcheck.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Soundcheck.exe"; Description: "{cm:LaunchProgram,Soundcheck}"; Flags: nowait postinstall skipifsilent

[Code]
const
  HWND_TOPMOST = -1;
  HWND_NOTOPMOST = -2;
  SWP_NOSIZE = 1;
  SWP_NOMOVE = 2;

function SetWindowPos(hWnd: Integer; hWndInsertAfter: Integer; X, Y, cx, cy: Integer; uFlags: Cardinal): Boolean;
  external 'SetWindowPos@user32.dll stdcall';

procedure InitializeWizard();
begin
  // Kurulum penceresini en öne zorla getir
  SetWindowPos(WizardForm.Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
  // Sürekli en üstte kalmaması için normale döndür
  SetWindowPos(WizardForm.Handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE or SWP_NOSIZE);
end;
