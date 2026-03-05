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
Name: "{autoprograms}\DesktopAgent\Start DesktopAgent"; Filename: "{app}\tray\DesktopAgent.Tray.exe"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autoprograms}\DesktopAgent\Stop DesktopAgent"; Filename: "{app}\stop-desktopagent.cmd"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autoprograms}\DesktopAgent\DesktopAgent Tray"; Filename: "{app}\tray\DesktopAgent.Tray.exe"
Name: "{autodesktop}\DesktopAgent"; Filename: "{app}\tray\DesktopAgent.Tray.exe"; IconFilename: "{app}\tray\DesktopAgent.Tray.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\tray\DesktopAgent.Tray.exe"; Description: "Start DesktopAgent"; Flags: postinstall nowait skipifsilent; Tasks: startonfinish
Filename: "{cmd}"; Parameters: "/C winget install -e --id Gyan.FFmpeg --accept-package-agreements --accept-source-agreements"; Flags: runhidden waituntilterminated skipifsilent; Check: ShouldInstallFfmpegPlugin
Filename: "{cmd}"; Parameters: "/C winget install -e --id UB-Mannheim.TesseractOCR --accept-package-agreements --accept-source-agreements"; Flags: runhidden waituntilterminated skipifsilent; Check: ShouldInstallOcrPrimaryPackage
Filename: "{cmd}"; Parameters: "/C winget install -e --id tesseract-ocr.tesseract --accept-package-agreements --accept-source-agreements"; Flags: runhidden waituntilterminated skipifsilent; Check: ShouldInstallOcrFallbackPackage

[UninstallRun]
Filename: "{app}\stop-desktopagent.cmd"; Flags: runhidden

[Code]
var
  PluginsPage: TWizardPage;
  PluginIntroLabel: TNewStaticText;
  FfmpegCheckBox: TNewCheckBox;
  FfmpegDetailsLabel: TNewStaticText;
  OcrCheckBox: TNewCheckBox;
  OcrDetailsLabel: TNewStaticText;
  InstallFfmpegPlugin: Boolean;
  InstallOcrPlugin: Boolean;
  CachedWingetAvailable: Boolean;
  CachedFfmpegAvailable: Boolean;
  CachedTesseractAvailable: Boolean;
  CachedTesseractPrimaryAvailable: Boolean;
  CachedTesseractFallbackAvailable: Boolean;

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

function IsTesseractAvailable: Boolean;
begin
  Result := CommandExists('tesseract');
end;

function CanOfferFfmpegInstall: Boolean;
begin
  Result := CachedWingetAvailable and (not CachedFfmpegAvailable);
end;

function CanOfferOcrInstall: Boolean;
begin
  Result := CachedWingetAvailable and (not CachedTesseractAvailable);
end;

function WingetPackageExists(const PackageId: string): Boolean;
var
  ResultCode: Integer;
begin
  Result :=
    Exec(ExpandConstant('{cmd}'),
         '/C winget show -e --id ' + PackageId,
         '',
         SW_HIDE,
         ewWaitUntilTerminated,
         ResultCode)
    and (ResultCode = 0);
end;

function IsTesseractPackagePrimaryAvailable: Boolean;
begin
  Result := CachedTesseractPrimaryAvailable;
end;

function IsTesseractPackageFallbackAvailable: Boolean;
begin
  Result := CachedTesseractFallbackAvailable;
end;

function ShouldInstallFfmpegPlugin: Boolean;
begin
  Result := InstallFfmpegPlugin and CanOfferFfmpegInstall;
end;

function ShouldInstallOcrPlugin: Boolean;
begin
  Result := InstallOcrPlugin and CanOfferOcrInstall;
end;

function ShouldInstallOcrPrimaryPackage: Boolean;
begin
  Result := ShouldInstallOcrPlugin and IsTesseractPackagePrimaryAvailable;
end;

function ShouldInstallOcrFallbackPackage: Boolean;
begin
  Result := ShouldInstallOcrPlugin
    and (not IsTesseractPackagePrimaryAvailable)
    and IsTesseractPackageFallbackAvailable;
end;

procedure DetectPluginEnvironment;
begin
  CachedWingetAvailable := IsWingetAvailable;
  CachedFfmpegAvailable := IsFFmpegAvailable;
  CachedTesseractAvailable := IsTesseractAvailable;
  CachedTesseractPrimaryAvailable := False;
  CachedTesseractFallbackAvailable := False;

  if CachedWingetAvailable and (not CachedTesseractAvailable) then
  begin
    CachedTesseractPrimaryAvailable := WingetPackageExists('UB-Mannheim.TesseractOCR');
    if not CachedTesseractPrimaryAvailable then
    begin
      CachedTesseractFallbackAvailable := WingetPackageExists('tesseract-ocr.tesseract');
    end;
  end;
end;

procedure InitializeWizard;
begin
  InstallFfmpegPlugin := ExpandConstant('{param:installffmpeg|0}') = '1';
  InstallOcrPlugin := ExpandConstant('{param:installocr|0}') = '1';

  DetectPluginEnvironment;

  PluginsPage := CreateCustomPage(
    wpSelectTasks,
    'Optional Plugins',
    'Install extra local tooling for advanced DesktopAgent features.');

  PluginIntroLabel := TNewStaticText.Create(PluginsPage);
  PluginIntroLabel.Parent := PluginsPage.Surface;
  PluginIntroLabel.Left := ScaleX(0);
  PluginIntroLabel.Top := ScaleY(0);
  PluginIntroLabel.Width := PluginsPage.SurfaceWidth;
  PluginIntroLabel.Height := ScaleY(34);
  PluginIntroLabel.AutoSize := False;
  PluginIntroLabel.Caption :=
    'Choose optional plugins to install with DesktopAgent. ' +
    'You can also install/enable them later from the app Configuration > Utilities.';
  PluginIntroLabel.WordWrap := True;

  FfmpegCheckBox := TNewCheckBox.Create(PluginsPage);
  FfmpegCheckBox.Parent := PluginsPage.Surface;
  FfmpegCheckBox.Left := ScaleX(0);
  FfmpegCheckBox.Top := ScaleY(46);
  FfmpegCheckBox.Width := PluginsPage.SurfaceWidth;
  FfmpegCheckBox.Caption := 'Install FFmpeg plugin (screen recording)';
  FfmpegCheckBox.Checked := InstallFfmpegPlugin and CanOfferFfmpegInstall;
  FfmpegCheckBox.Enabled := CanOfferFfmpegInstall;

  FfmpegDetailsLabel := TNewStaticText.Create(PluginsPage);
  FfmpegDetailsLabel.Parent := PluginsPage.Surface;
  FfmpegDetailsLabel.Left := ScaleX(22);
  FfmpegDetailsLabel.Top := ScaleY(68);
  FfmpegDetailsLabel.Width := PluginsPage.SurfaceWidth - ScaleX(22);
  FfmpegDetailsLabel.Height := ScaleY(28);
  FfmpegDetailsLabel.AutoSize := False;
  if not CachedWingetAvailable then
    FfmpegDetailsLabel.Caption := 'winget not found: FFmpeg plugin install is unavailable.'
  else if CachedFfmpegAvailable then
    FfmpegDetailsLabel.Caption := 'FFmpeg already detected on this machine.'
  else
    FfmpegDetailsLabel.Caption := 'Enables screen capture/recording workflows (installed via winget: Gyan.FFmpeg).';
  FfmpegDetailsLabel.WordWrap := True;

  OcrCheckBox := TNewCheckBox.Create(PluginsPage);
  OcrCheckBox.Parent := PluginsPage.Surface;
  OcrCheckBox.Left := ScaleX(0);
  OcrCheckBox.Top := ScaleY(110);
  OcrCheckBox.Width := PluginsPage.SurfaceWidth;
  OcrCheckBox.Caption := 'Install OCR plugin (Tesseract for vision fallback)';
  OcrCheckBox.Checked := InstallOcrPlugin and CanOfferOcrInstall;
  OcrCheckBox.Enabled := CanOfferOcrInstall;

  OcrDetailsLabel := TNewStaticText.Create(PluginsPage);
  OcrDetailsLabel.Parent := PluginsPage.Surface;
  OcrDetailsLabel.Left := ScaleX(22);
  OcrDetailsLabel.Top := ScaleY(132);
  OcrDetailsLabel.Width := PluginsPage.SurfaceWidth - ScaleX(22);
  OcrDetailsLabel.Height := ScaleY(42);
  OcrDetailsLabel.AutoSize := False;
  if not CachedWingetAvailable then
    OcrDetailsLabel.Caption := 'winget not found: OCR plugin install is unavailable.'
  else if CachedTesseractAvailable then
    OcrDetailsLabel.Caption := 'Tesseract already detected on this machine.'
  else if CachedTesseractPrimaryAvailable then
    OcrDetailsLabel.Caption := 'Uses winget package UB-Mannheim.TesseractOCR.'
  else if CachedTesseractFallbackAvailable then
    OcrDetailsLabel.Caption := 'Uses fallback winget package tesseract-ocr.tesseract.'
  else
    OcrDetailsLabel.Caption := 'No known Tesseract winget package detected. You can install it manually later.';
  OcrDetailsLabel.WordWrap := True;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if Assigned(PluginsPage) and (CurPageID = PluginsPage.ID) then
  begin
    InstallFfmpegPlugin := Assigned(FfmpegCheckBox) and FfmpegCheckBox.Checked and FfmpegCheckBox.Enabled;
    InstallOcrPlugin := Assigned(OcrCheckBox) and OcrCheckBox.Checked and OcrCheckBox.Enabled;
  end;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
  if Assigned(PluginsPage) and (PageID = PluginsPage.ID) then
  begin
    Result := WizardSilent or ((not CanOfferFfmpegInstall) and (not CanOfferOcrInstall));
  end;
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

    if not IsTesseractAvailable then
    begin
      MsgBox(
        'Tesseract OCR was not detected. Vision OCR fallback requires Tesseract.'#13#10#13#10 +
        'Install manually with one of:'#13#10 +
        '  winget install -e --id UB-Mannheim.TesseractOCR'#13#10 +
        '  winget install -e --id tesseract-ocr.tesseract'#13#10#13#10 +
        'Then enable OCR in DesktopAgent Configuration and restart.',
        mbInformation,
        MB_OK);
    end;
  end;
end;
