#ifndef MyAppVersion
  #define MyAppVersion "0.1.0"
#endif

#ifndef DistDir
  #define DistDir "..\\dist\\win-x64"
#endif

#ifndef InstallerOut
  #define InstallerOut "..\\dist\\installer"
#endif

#define MyAppName "DesktopAgent"
#define MyAppPublisher "DesktopAgent"
#define MyAppExeName "start-desktopagent.cmd"

[Setup]
AppId={{A0A1C0D7-410B-4B45-8E76-BD8617B8F19A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
Compression=lzma
SolidCompression=yes
WizardStyle=modern
OutputDir={#InstallerOut}
OutputBaseFilename=DesktopAgent-Setup-{#MyAppVersion}
UninstallDisplayIcon={app}\tray\DesktopAgent.Tray.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"
Name: "startonfinish"; Description: "Start DesktopAgent after setup"; GroupDescription: "Post-install:"
Name: "installffmpeg"; Description: "Install FFmpeg (required for screen recording)"; GroupDescription: "Optional components:"; Check: IsWingetAvailable

[Files]
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\DesktopAgent\Start DesktopAgent"; Filename: "{app}\tray\DesktopAgent.Tray.exe"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autoprograms}\DesktopAgent\Stop DesktopAgent"; Filename: "{app}\stop-desktopagent.cmd"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autoprograms}\DesktopAgent\DesktopAgent Tray"; Filename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autodesktop}\DesktopAgent"; Filename: "{app}\tray\DesktopAgent.Tray.exe"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\tray\DesktopAgent.Tray.exe"; Description: "Start DesktopAgent"; Flags: postinstall nowait skipifsilent; Tasks: startonfinish
Filename: "{cmd}"; Parameters: "/C winget install -e --id Gyan.FFmpeg --accept-package-agreements --accept-source-agreements"; Flags: runhidden waituntilterminated skipifsilent; Tasks: installffmpeg

[UninstallRun]
Filename: "{app}\stop-desktopagent.cmd"; Flags: runhidden

[Code]
function CommandExists(const Command: string): Boolean;
var
  ResultCode: Integer;
begin
  Result :=
    Exec(ExpandConstant('{cmd}'),
         '/C where ' + Command,
         '',
         SW_HIDE,
         ewWaitUntilTerminated,
         ResultCode)
    and (ResultCode = 0);
end;

function IsWingetAvailable: Boolean;
begin
  Result := CommandExists('winget');
end;

function IsFFmpegAvailable: Boolean;
begin
  Result := CommandExists('ffmpeg');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    if not IsFFmpegAvailable then
    begin
      MsgBox(
        'FFmpeg was not detected. Screen recording features require FFmpeg.'#13#10#13#10 +
        'Install manually with:'#13#10 +
        '  winget install -e --id Gyan.FFmpeg' #13#10#13#10 +
        'Then restart DesktopAgent.',
        mbInformation,
        MB_OK);
    end;
  end;
end;
