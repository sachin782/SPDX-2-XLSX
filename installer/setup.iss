; SPDX JSON to Excel Converter - Windows Installer
; Publisher: Sachin Rawat
;
; Installs prerequisites (Microsoft Visual C++ Redistributable) first,
; then installs the main application executable.

#define MyAppName "SPDX JSON to Excel Converter"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Sachin Rawat"
#define MyAppExeName "SPDX-Excel-Converter.exe"

[Setup]
AppId={{8F4E2A91-6C3D-4B7E-9F12-3D5A8C1E4B60}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\SPDX Excel Converter
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist\installer
OutputBaseFilename=SPDX-Excel-Converter-Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Prerequisite: VC++ Redistributable (bundled; installed before app files via BeforeInstall)
Source: "redist\vc_redist.x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall dontcopy

; Main application (installed after prerequisites)
Source: "..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion; BeforeInstall: InstallPrerequisites

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
function IsVCRedistInstalled: Boolean;
var
  Installed: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64', 'Installed', Installed) then
  begin
    if Installed = 1 then
    begin
      Result := True;
      Exit;
    end;
  end;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64', 'Installed', Installed) then
  begin
    if Installed = 1 then
      Result := True;
  end;
end;

function VCRedistNeedsInstall: Boolean;
begin
  Result := not IsVCRedistInstalled;
end;

procedure InstallPrerequisites;
var
  ResultCode: Integer;
begin
  if VCRedistNeedsInstall then
  begin
    WizardForm.StatusLabel.Caption := 'Installing Microsoft Visual C++ Redistributable...';
    ExtractTemporaryFile('vc_redist.x64.exe');
    if not Exec(
      ExpandConstant('{tmp}\vc_redist.x64.exe'),
      '/install /quiet /norestart',
      '',
      SW_SHOW,
      ewWaitUntilTerminated,
      ResultCode) then
    begin
      MsgBox(
        'Failed to install Microsoft Visual C++ Redistributable.' + #13#10 +
        'The application may not run correctly without it.',
        mbError, MB_OK);
    end;
  end;
  WizardForm.StatusLabel.Caption := 'Installing SPDX JSON to Excel Converter...';
end;
