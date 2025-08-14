#define MyAppName       "My Office Add-In"
#define MyAppVersion    "1.0.0.4"
#define MyPublisher     "YourCompany"
#define VstoFile        "my-addin.vsto"
#define AppFilesDir     "Application Files\my-addin_1_0_0_4"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyPublisher}
DefaultDirName={autopf}\MyOfficeAddIn
DefaultGroupName=My Office Add-In
OutputDir=Output
OutputBaseFilename=MyOfficeAddIn_Setup
Compression=lzma2
SolidCompression=yes
DisableProgramGroupPage=yes
DisableWelcomePage=no
PrivilegesRequired=admin
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64
SetupLogging=yes

[Files]
; ClickOnce payload
Source: "{#VstoFile}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AppFilesDir}\*"; DestDir: "{app}\{#AppFilesDir}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Run]
; Install via the system VSTO installer (more reliable than publisher setup.exe)
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/install ""{app}\{#VstoFile}"" /silent"; Flags: runhidden

[UninstallRun]
Filename: "{code:GetVstoInstallerPath}"; Parameters: "/uninstall ""{app}\{#VstoFile}"" /silent"; Flags: runhidden

[Code]
function GetVstoInstallerPath(Param: string): string;
var
  p64, p86: string;
begin
  p64 := ExpandConstant('{pf}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe');   { Program Files (x64) }
  p86 := ExpandConstant('{pf32}\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe'); { Program Files (x86) }
  if FileExists(p64) then
    Result := p64
  else if FileExists(p86) then
    Result := p86
  else
    Result := '';  { runtime not present }
end;

function IsVstoInstalled: Boolean;
var
  v: Cardinal;
  s: string;
begin
  { Fast path: VSTOInstaller.exe exists }
  Result := (GetVstoInstallerPath('') <> '');
  if Result then Exit;

  { Fallback: check registry – both views, v4R and v4; value can be DWORD or string }
  if RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Install', v) and (v = 1) then begin Result := True; Exit; end;
  if RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4',  'Install', v) and (v = 1) then begin Result := True; Exit; end;
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Install', s) and (s = '1') then begin Result := True; Exit; end;
  if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\VSTO Runtime Setup\v4',  'Install', s) and (s = '1') then begin Result := True; Exit; end;

  if RegQueryDWordValue(HKLM,   'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Install', v) and (v = 1) then begin Result := True; Exit; end;
  if RegQueryDWordValue(HKLM,   'SOFTWARE\Microsoft\VSTO Runtime Setup\v4',  'Install', v) and (v = 1) then begin Result := True; Exit; end;
  if RegQueryStringValue(HKLM,  'SOFTWARE\Microsoft\VSTO Runtime Setup\v4R', 'Install', s) and (s = '1') then begin Result := True; Exit; end;
  if RegQueryStringValue(HKLM,  'SOFTWARE\Microsoft\VSTO Runtime Setup\v4',  'Install', s) and (s = '1') then begin Result := True; Exit; end;
end;

function InitializeSetup(): Boolean;
var
  path: string;
begin
  if not IsVstoInstalled then
  begin
    MsgBox('This add-in requires the Microsoft Visual Studio Tools for Office Runtime. Please install it first.', mbError, MB_OK);
    Result := False;
    Exit;
  end;

  { Extra safety: if for some reason detection says installed but we cannot find the exe, bail gracefully }
  path := GetVstoInstallerPath('');
  if path = '' then
  begin
    MsgBox('VSTO runtime seems installed but VSTOInstaller.exe was not found. Please reinstall the VSTO runtime.', mbError, MB_OK);
    Result := False;
    Exit;
  end;

  Result := True;
end;
