[Setup]
AppName=DataViewer Enterprise
AppVersion=0.2.0
AppVerName=DataViewer Enterprise v0.2.0
AppPublisher=SDR
DefaultDirName={autopf}\DataViewer Enterprise
DefaultGroupName=DataViewer Enterprise
OutputDir=dist
OutputBaseFilename=DataViewer-setup
SetupIconFile=resources\images\ccell_icon.ico
Compression=lzma2
SolidCompression=yes
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

; Resources (templates, images/branding)
Source: "resources\templates\*"; DestDir: "{app}\resources\templates"; Flags: ignoreversion recursesubdirs
Source: "resources\images\*"; DestDir: "{app}\resources\images"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\DataViewer Enterprise"; Filename: "{app}\DataViewer.exe"; IconFilename: "{app}\resources\images\ccell_icon.ico"
Name: "{group}\Uninstall DataViewer Enterprise"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DataViewer Enterprise"; Filename: "{app}\DataViewer.exe"; IconFilename: "{app}\resources\images\ccell_icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\DataViewer.exe"; Description: "Launch DataViewer Enterprise"; Flags: nowait postinstall skipifsilent
