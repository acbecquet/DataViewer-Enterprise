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

; Qt plugins
Source: "release\platforms\*"; DestDir: "{app}\platforms"; Flags: ignoreversion recursesubdirs; Components: main
Source: "release\sqldrivers\qsqlite.dll"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion; Components: main
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
