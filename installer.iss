[Setup]
AppName=DataViewer Enterprise
AppVersion=0.4.0
AppVerName=DataViewer Enterprise v0.4.0
AppPublisher=SDR
DefaultDirName={autopf}\DataViewer Enterprise
DefaultGroupName=DataViewer Enterprise
OutputDir=dist
OutputBaseFilename=DataViewer-setup
SetupIconFile=resources\images\ccell_icon.ico
Compression=lzma2
SolidCompression=no
ArchitecturesInstallIn64BitMode=x64compatible
WizardStyle=modern
UninstallDisplayName=DataViewer Enterprise
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
; Main executable
Source: "release\DataViewer.exe"; DestDir: "{app}"; Flags: ignoreversion

; Qt DLLs
Source: "release\Qt6Core.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\Qt6Gui.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\Qt6Widgets.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\Qt6Sql.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\Qt6Network.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\Qt6Svg.dll"; DestDir: "{app}"; Flags: ignoreversion

; MinGW runtime
Source: "release\libgcc_s_seh-1.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\libstdc++-6.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\libwinpthread-1.dll"; DestDir: "{app}"; Flags: ignoreversion

; Qt plugins
Source: "release\platforms\*"; DestDir: "{app}\platforms"; Flags: ignoreversion recursesubdirs
Source: "release\sqldrivers\qsqlite.dll"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion
Source: "release\imageformats\*"; DestDir: "{app}\imageformats"; Flags: ignoreversion recursesubdirs
Source: "release\styles\*"; DestDir: "{app}\styles"; Flags: ignoreversion recursesubdirs
Source: "release\iconengines\*"; DestDir: "{app}\iconengines"; Flags: ignoreversion recursesubdirs
Source: "release\tls\*"; DestDir: "{app}\tls"; Flags: ignoreversion recursesubdirs
Source: "release\networkinformation\*"; DestDir: "{app}\networkinformation"; Flags: ignoreversion recursesubdirs
Source: "release\generic\*"; DestDir: "{app}\generic"; Flags: ignoreversion recursesubdirs

; Bundled Python — shipped as a zip to avoid Inno Setup corrupting binary files.
; Extracted at install time by the [Run] section below.
Source: "release\python_bundle.zip"; DestDir: "{app}"; Flags: ignoreversion

; Resources (templates, images/branding)
Source: "resources\templates\*"; DestDir: "{app}\resources\templates"; Flags: ignoreversion recursesubdirs
Source: "resources\images\*"; DestDir: "{app}\resources\images"; Flags: ignoreversion recursesubdirs
Source: "resources\sops.xlsx"; DestDir: "{app}\resources"; Flags: ignoreversion

[Icons]
Name: "{group}\DataViewer Enterprise"; Filename: "{app}\DataViewer.exe"; IconFilename: "{app}\resources\images\ccell_icon.ico"
Name: "{group}\Uninstall DataViewer Enterprise"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DataViewer Enterprise"; Filename: "{app}\DataViewer.exe"; IconFilename: "{app}\resources\images\ccell_icon.ico"; Tasks: desktopicon

[Run]
; Extract bundled Python and remove the zip
Filename: "powershell.exe"; Parameters: "-NoProfile -Command ""Expand-Archive -LiteralPath '{app}\python_bundle.zip' -DestinationPath '{app}\python' -Force"""; Flags: runhidden waituntilterminated
Filename: "powershell.exe"; Parameters: "-NoProfile -Command ""Remove-Item -LiteralPath '{app}\python_bundle.zip'"""; Flags: runhidden waituntilterminated
Filename: "{app}\DataViewer.exe"; Description: "Launch DataViewer Enterprise"; Flags: nowait postinstall skipifsilent
