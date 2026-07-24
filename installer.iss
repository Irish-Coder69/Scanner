[Setup]
AppName=Document Scanner
AppVersion=1.0.13
DefaultDirName={autopf}\Document Scanner
DefaultGroupName=Document Scanner
OutputDir=installer_output
OutputBaseFilename=DocumentScannerSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
CloseApplications=no

[Code]
procedure CloseRunningScanner;
var
	ResultCode: Integer;
begin
	Exec(ExpandConstant('{cmd}'), '/C taskkill /IM "Document Scanner.exe" /F /T >nul 2>&1', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
	if CurStep = ssInstall then
		CloseRunningScanner;
end;

[Files]
Source: "dist\Document Scanner.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "scanner_icon.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.json"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Icons]
Name: "{group}\Document Scanner"; Filename: "{app}\Document Scanner.exe"; IconFilename: "{app}\scanner_icon.ico"
Name: "{autodesktop}\Document Scanner"; Filename: "{app}\Document Scanner.exe"; Tasks: desktopicon; IconFilename: "{app}\scanner_icon.ico"
