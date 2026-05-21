; AppVersion is supplied by build_installer.bat via ISCC /DAppVersion=...
; build_installer.bat extracts it from VERSION in DataViewerEnterprise.pro,
; the single source of truth.
#ifndef AppVersion
  #define AppVersion "0.0.0-dev"
#endif

[Setup]
AppName=DataViewer Enterprise
AppVersion={#AppVersion}
AppVerName=DataViewer Enterprise v{#AppVersion}
AppPublisher=SDR
DefaultDirName={autopf}\DataViewer Enterprise
DefaultGroupName=DataViewer Enterprise
OutputDir=dist
OutputBaseFilename=DataViewer-setup
SetupIconFile=resources\images\ccell_icon.ico
Compression=lzma2/ultra64
SolidCompression=no
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayName=DataViewer Enterprise
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

; If a running DataViewer.exe is detected, prompt to close it (with the
; option to force-close). After install, restart any apps we closed so
; the user lands on the new version automatically. AppMutex is checked
; against a named mutex set by main.cpp so detection works even if the
; .exe was renamed.
CloseApplications=force
RestartApplications=yes
AppMutex=DataViewerEnterprise_SingleInstance

[Code]
const
  // Production defaults — match the smoorecig office NAS that runs the
  // shared Postgres container. Sysadmins can override at install time
  // via the wizard. If the NAS moves or the port changes, bump these
  // constants and rebuild; rolling out new defaults is one file edit.
  DEFAULT_DB_HOST = '192.168.222.10';
  DEFAULT_DB_PORT = '5433';
  // v2.0.3: bake the shared NAS password in so internal users don't
  // have to type it at install time. The password is still encrypted
  // per-user on disk (machine+user-bound XOR via ConfigLoader). The
  // installer wizard still SHOWS the password field — users can override
  // if their sysadmin gave them a different value — but the default is
  // the production one. Rotate this constant if the NAS password
  // changes.
  DEFAULT_DB_PASSWORD = 'SDR2026@2026!@#';

var
  DbPasswordPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  DbPasswordPage := CreateInputQueryPage(wpUserInfo,
    'Database Connection',
    'Confirm the PostgreSQL database connection details',
    'These defaults match the team''s production NAS database. ' +
    'Override the host or port only if your sysadmin gave you different values. ' +
    'The password is required and is what your sysadmin set when first deploying the NAS container.');
  DbPasswordPage.Add('Database host:',  False);
  DbPasswordPage.Add('Database port:',  False);
  DbPasswordPage.Add('Database password:', True);

  // Pre-populate the defaults — host, port, AND password (v2.0.3) — so
  // most users can hit Next without typing anything. Sysadmins with a
  // non-default deployment can still override before clicking Next.
  DbPasswordPage.Values[0] := DEFAULT_DB_HOST;
  DbPasswordPage.Values[1] := DEFAULT_DB_PORT;
  DbPasswordPage.Values[2] := DEFAULT_DB_PASSWORD;
end;

procedure WriteDbConf;
var
  ConfDir, ConfPath, EncPwd, HostVal, PortVal: AnsiString;
  ResultCode: Integer;
begin
  // Per-user config — must work without admin elevation since end users
  // typically don't have local admin rights. Tied to the Windows user
  // profile, which matches the AES password key (machine+user bound).
  ConfDir := ExpandConstant('{localappdata}') + '\DataViewer';
  ConfPath := ConfDir + '\db.conf';
  if not DirExists(ConfDir) then
    CreateDir(ConfDir);

  // Read user values, falling back to defaults if a field somehow ended up empty.
  HostVal := DbPasswordPage.Values[0];
  if HostVal = '' then HostVal := DEFAULT_DB_HOST;
  PortVal := DbPasswordPage.Values[1];
  if PortVal = '' then PortVal := DEFAULT_DB_PORT;

  Exec(ExpandConstant('{app}\DataViewer.exe'),
       '--encrypt-password=' + DbPasswordPage.Values[2] + ' --encrypted-out=' +
       ConfDir + '\encrypted.tmp',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if not LoadStringFromFile(ConfDir + '\encrypted.tmp', EncPwd) then
    EncPwd := '';
  DeleteFile(ConfDir + '\encrypted.tmp');

  SaveStringToFile(ConfPath,
    '[postgres]' + #13#10 +
    'host = ' + HostVal + #13#10 +
    'port = ' + PortVal + #13#10 +
    'database = dataviewer' + #13#10 +
    'user = dataviewer_app' + #13#10 +
    'password_encrypted = ' + EncPwd + #13#10,
    False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    WriteDbConf;
end;

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Types]
Name: "full";    Description: "Full installation"
Name: "compact"; Description: "Compact installation"
Name: "custom";  Description: "Custom installation"; Flags: iscustom

[Components]
Name: "main";       Description: "DataViewer Enterprise (required)"; Types: full compact custom; Flags: fixed
Name: "translator"; Description: "Document Translator";              Types: full

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
; Main executable
Source: "release\DataViewer.exe"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; Qt DLLs
Source: "release\Qt6Core.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Gui.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Widgets.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Sql.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Network.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\Qt6Svg.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; MinGW runtime
Source: "release\libgcc_s_seh-1.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libstdc++-6.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libwinpthread-1.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; libpq runtime DLLs (PostgreSQL client library + dependencies)
Source: "release\libpq.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libcrypto-3-x64.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libssl-3-x64.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libintl-9.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main
Source: "release\libiconv-2.dll"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; Qt plugins
Source: "release\platforms\*"; DestDir: "{app}\platforms"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\sqldrivers\qsqlite.dll"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion; Components: main
Source: "release\sqldrivers\qsqlpsql.dll"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion; Components: main
Source: "release\imageformats\*"; DestDir: "{app}\imageformats"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\styles\*"; DestDir: "{app}\styles"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\iconengines\*"; DestDir: "{app}\iconengines"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\tls\*"; DestDir: "{app}\tls"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\networkinformation\*"; DestDir: "{app}\networkinformation"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\generic\*"; DestDir: "{app}\generic"; Flags: ignoreversion recursesubdirs; Components: main

; Bundled Python — shipped as a zip to avoid Inno Setup corrupting binary files.
; Extracted at install time by the [Run] section below.
Source: "release\python_bundle.zip"; DestDir: "{app}"; Flags: ignoreversion; Components: main

; Resources
Source: "resources\templates\*"; DestDir: "{app}\resources\templates"; Flags: ignoreversion recursesubdirs; Components: main
Source: "resources\images\*"; DestDir: "{app}\resources\images"; Flags: ignoreversion recursesubdirs; Components: main
Source: "resources\icons\*.svg"; DestDir: "{app}\resources\icons"; Flags: ignoreversion; Components: main
Source: "resources\sops.xlsx"; DestDir: "{app}\resources"; Flags: ignoreversion; Components: main

; Document Translator (optional component)
Source: "dataviewer_translator\dist\DocumentTranslator.exe"; DestDir: "{app}\dataviewer_translator\dist"; Flags: ignoreversion; Components: translator
Source: "dataviewer_translator\resources\*"; DestDir: "{app}\dataviewer_translator\resources"; Flags: ignoreversion recursesubdirs; Components: translator

[Icons]
Name: "{group}\DataViewer Enterprise"; Filename: "{app}\DataViewer.exe"; IconFilename: "{app}\resources\images\ccell_icon.ico"
Name: "{group}\Uninstall DataViewer Enterprise"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DataViewer Enterprise"; Filename: "{app}\DataViewer.exe"; IconFilename: "{app}\resources\images\ccell_icon.ico"; Tasks: desktopicon

[Run]
; Extract bundled Python and remove the zip
Filename: "powershell.exe"; Parameters: "-NoProfile -Command ""Expand-Archive -LiteralPath '{app}\python_bundle.zip' -DestinationPath '{app}\python' -Force"""; Flags: runhidden waituntilterminated; Components: main
Filename: "powershell.exe"; Parameters: "-NoProfile -Command ""Remove-Item -LiteralPath '{app}\python_bundle.zip'"""; Flags: runhidden waituntilterminated; Components: main
Filename: "{app}\DataViewer.exe"; Description: "Launch DataViewer Enterprise"; Flags: nowait postinstall skipifsilent; Components: main
