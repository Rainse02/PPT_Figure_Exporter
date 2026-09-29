#ifndef AppVersion
  #define AppVersion "1.0.0.0"
#endif

#define AssemblyVersion "1.0.0.0"
#define AddinGuid "A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76"

; The .NET runtime is a separate prerequisite and is never embedded.

[Setup]
AppId={{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}
AppName=PPT Figure Exporter
AppVersion={#AppVersion}
AppPublisher=PPT Figure Exporter
DefaultDirName={localappdata}\Programs\PPTFigureExporter
DefaultGroupName=PPT Figure Exporter
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
MinVersion=10.0.10240
OutputDir=..\dist
OutputBaseFilename=PPTFigureExporter_Setup
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
Uninstallable=yes
CloseApplications=no
RestartIfNeededByRun=no
SetupLogging=yes
AllowNoIcons=yes

[Files]
Source: "..\dist\app\PPTFigureExporter.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\app\figure-export-core.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\dist\app\*.config"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\dist\app\build-manifest.json"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\dist\app\README.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\dist\licenses\*"; DestDir: "{app}\licenses"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist

[Registry]
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}"; ValueType: none; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "PPTFigureExporter.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "PPTFigureExporter, Version={#AssemblyVersion}, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "file:///{app}/PPTFigureExporter.dll"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\ProgID"; ValueType: string; ValueName: ""; ValueData: "PPTFigureExporter.Connect"
Root: HKCU; Subkey: "Software\Classes\CLSID\{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}\VersionIndependentProgID"; ValueType: string; ValueName: ""; ValueData: "PPTFigureExporter.Connect"
Root: HKCU; Subkey: "Software\Classes\PPTFigureExporter.Connect"; ValueType: string; ValueName: ""; ValueData: "PPT Figure Exporter COM Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\PPTFigureExporter.Connect\CLSID"; ValueType: string; ValueName: ""; ValueData: "{{A82ECF4B-81E5-462B-8C4D-3FA2B5BD1E76}"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTFigureExporter.Connect"; ValueType: none; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTFigureExporter.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "PPT Figure Exporter"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTFigureExporter.Connect"; ValueType: string; ValueName: "Description"; ValueData: "Export a selected PowerPoint figure to publication-ready formats."
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\PPTFigureExporter.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3

[Code]
const
  TH32CS_SNAPPROCESS = $00000002;
  INVALID_HANDLE_VALUE = -1;

type
  TProcessEntry32 = record
    dwSize: DWORD;
    cntUsage: DWORD;
    th32ProcessID: DWORD;
    th32DefaultHeapID: DWORD;
    th32ModuleID: DWORD;
    cntThreads: DWORD;
    th32ParentProcessID: DWORD;
    pcPriClassBase: Longint;
    dwFlags: DWORD;
    szExeFile: array[0..259] of Char;
  end;

function CreateToolhelp32Snapshot(dwFlags, th32ProcessID: DWORD): THandle;
  external 'CreateToolhelp32Snapshot@kernel32.dll stdcall';
function Process32First(hSnapshot: THandle; var lppe: TProcessEntry32): BOOL;
  external 'Process32FirstW@kernel32.dll stdcall';
function Process32Next(hSnapshot: THandle; var lppe: TProcessEntry32): BOOL;
  external 'Process32NextW@kernel32.dll stdcall';
function CloseHandle(hObject: THandle): BOOL;
  external 'CloseHandle@kernel32.dll stdcall';

function GetProcessImageName(const Entry: TProcessEntry32): String;
var
  Index: Integer;
begin
  Result := '';
  for Index := 0 to 259 do
  begin
    if Entry.szExeFile[Index] = #0 then
      Break;
    Result := Result + Entry.szExeFile[Index];
  end;
end;

function IsPowerPointRunning: Boolean;
var
  Snapshot: THandle;
  Entry: TProcessEntry32;
begin
  Result := False;
  Snapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if Snapshot = INVALID_HANDLE_VALUE then
    Exit;

  try
    Entry.dwSize := SizeOf(Entry);
    if Process32First(Snapshot, Entry) then
      repeat
        if CompareText(GetProcessImageName(Entry), 'POWERPNT.EXE') = 0 then
        begin
          Result := True;
          Break;
        end;
      until not Process32Next(Snapshot, Entry);
  finally
    CloseHandle(Snapshot);
  end;
end;

function GetPowerPointPath: string;
var
  OfficeRoot: string;
  Candidate: string;
begin
  Result := '';

  if RegQueryStringValue(HKLM64,
    'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
    'InstallationPath', OfficeRoot) then
  begin
    Candidate := AddBackslash(OfficeRoot) + 'root\Office16\POWERPNT.EXE';
    if FileExists(Candidate) then
    begin
      Result := Candidate;
      Exit;
    end;

    Candidate := AddBackslash(OfficeRoot) + 'Office16\POWERPNT.EXE';
    if FileExists(Candidate) then
    begin
      Result := Candidate;
      Exit;
    end;
  end;

  if RegQueryStringValue(HKLM32,
    'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
    'InstallationPath', OfficeRoot) then
  begin
    Candidate := AddBackslash(OfficeRoot) + 'root\Office16\POWERPNT.EXE';
    if FileExists(Candidate) then
    begin
      Result := Candidate;
      Exit;
    end;

    Candidate := AddBackslash(OfficeRoot) + 'Office16\POWERPNT.EXE';
    if FileExists(Candidate) then
    begin
      Result := Candidate;
      Exit;
    end;
  end;

  Candidate := ExpandConstant('{autopf}\Microsoft Office\root\Office16\POWERPNT.EXE');
  if FileExists(Candidate) then
  begin
    Result := Candidate;
    Exit;
  end;

  Candidate := ExpandConstant('{commonpf32}\Microsoft Office\root\Office16\POWERPNT.EXE');
  if FileExists(Candidate) then
    Result := Candidate;
end;

function GetBinaryType(lpApplicationName: string; var BinaryType: DWORD): Boolean;
  external 'GetBinaryTypeW@kernel32.dll stdcall';

function IsPowerPointBinaryX64(const FileName: string): Boolean;
var
  BinaryType: DWORD;
begin
  BinaryType := 0;
  Result := False;
  if GetBinaryType(FileName, BinaryType) then
    Result := BinaryType = 6;
end;

function HasX64PowerPoint: Boolean;
var
  PowerPointPath: string;
  Platform: string;
  ComServer: string;
  PlatformConfirmsX64: Boolean;
begin
  Result := False;
  PowerPointPath := GetPowerPointPath;
  if (PowerPointPath = '') or not IsPowerPointBinaryX64(PowerPointPath) then
    Exit;

  PlatformConfirmsX64 := False;
  if RegQueryStringValue(HKLM64,
    'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
    'Platform', Platform) then
    PlatformConfirmsX64 := CompareText(Platform, 'x64') = 0;

  if not PlatformConfirmsX64 and RegQueryStringValue(HKLM32,
    'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
    'Platform', Platform) then
    PlatformConfirmsX64 := CompareText(Platform, 'x64') = 0;

  if not PlatformConfirmsX64 then
    if RegQueryStringValue(HKCR,
      'CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32',
      '', ComServer) then
      PlatformConfirmsX64 := Pos('POWERPNT.EXE', Uppercase(ComServer)) > 0;

  Result := PlatformConfirmsX64;
end;
function IsDotNet48Installed: Boolean;
var
  ReleaseValue: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM64,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release', ReleaseValue) then
    Result := ReleaseValue >= 528040;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  NeedsRestart := False;
  if not IsDotNet48Installed then
    Result := 'Microsoft .NET Framework 4.8 or later is required.' + #13#10 +
      'Install the separate Microsoft Runtime package, restart Windows if requested, ' +
      'then run this setup again.' + #13#10 +
      'Download: https://dotnet.microsoft.com/download/dotnet-framework/net48';
end;

function InitializeSetup: Boolean;
begin
  Result := False;
  if IsPowerPointRunning then
  begin
    MsgBox('PowerPoint is running. Close POWERPNT.EXE before installing or upgrading PPT Figure Exporter.', mbError, MB_OK);
    Exit;
  end;

  if not HasX64PowerPoint then
  begin
    MsgBox('A 64-bit PowerPoint installation is required. This installer supports Windows 10/11 with PowerPoint x64.', mbError, MB_OK);
    Exit;
  end;

  Result := True;
end;

function InitializeUninstall: Boolean;
begin
  Result := True;
  if IsPowerPointRunning then
  begin
    MsgBox('PowerPoint is running. Close POWERPNT.EXE before uninstalling PPT Figure Exporter.', mbError, MB_OK);
    Result := False;
  end;
end;



