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

[Files]
Source: "{#DistDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autoprograms}\DesktopAgent\Start DesktopAgent"; Filename: "{app}\start-desktopagent.cmd"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autoprograms}\DesktopAgent\Stop DesktopAgent"; Filename: "{app}\stop-desktopagent.cmd"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autoprograms}\DesktopAgent\DesktopAgent Tray"; Filename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autodesktop}\DesktopAgent"; Filename: "{app}\start-desktopagent.cmd"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\start-desktopagent.cmd"; Description: "Start DesktopAgent"; Flags: postinstall nowait skipifsilent; Tasks: startonfinish

[UninstallRun]
Filename: "{app}\stop-desktopagent.cmd"; Flags: runhidden
