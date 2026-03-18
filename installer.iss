[Setup]
AppName=DataViewer Enterprise
AppVersion=0.1.0-alpha
AppVerName=DataViewer Enterprise v0.1.0-alpha
AppPublisher=DataViewer
DefaultDirName={autopf}\DataViewer Enterprise
DefaultGroupName=DataViewer Enterprise
OutputDir=dist
OutputBaseFilename=DataViewerEnterprise-v0.1.0-alpha-setup
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
Source: "release\DataViewerEnterprise.exe"; DestDir: "{app}"; Flags: ignoreversion

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

; OpenGL / D3D fallback
Source: "release\D3Dcompiler_47.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "release\opengl32sw.dll"; DestDir: "{app}"; Flags: ignoreversion

; Qt plugins
Source: "release\platforms\*"; DestDir: "{app}\platforms"; Flags: ignoreversion recursesubdirs
Source: "release\sqldrivers\qsqlite.dll"; DestDir: "{app}\sqldrivers"; Flags: ignoreversion
Source: "release\imageformats\*"; DestDir: "{app}\imageformats"; Flags: ignoreversion recursesubdirs
Source: "release\styles\*"; DestDir: "{app}\styles"; Flags: ignoreversion recursesubdirs

; Resources (templates)
Source: "resources\templates\*"; DestDir: "{app}\resources\templates"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\DataViewer Enterprise"; Filename: "{app}\DataViewerEnterprise.exe"
Name: "{group}\Uninstall DataViewer Enterprise"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DataViewer Enterprise"; Filename: "{app}\DataViewerEnterprise.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\DataViewerEnterprise.exe"; Description: "Launch DataViewer Enterprise"; Flags: nowait postinstall skipifsilent
