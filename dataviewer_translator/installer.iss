; Document Translator Installer Script
; Inno Setup configuration file

#define MyAppName "Document Translator"
#define MyAppVersion "2.1.0"
#define MyAppPublisher "SDR"
#define MyAppExeName "DocumentTranslator.exe"

[Setup]
; NOTE: Generate your own GUID using Tools > Generate GUID in Inno Setup
AppId={{B3C4D5E6-F7A8-9012-BCDE-F01234567890}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
OutputDir=installer_output
OutputBaseFilename=DocumentTranslatorSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=resources\ccell_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; 64-bit only
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}";

[Files]
; Main executable
Source: "dist\DocumentTranslator.exe"; DestDir: "{app}"; Flags: ignoreversion
; Python modules
Source: "translate_excel.py";      DestDir: "{app}"; Flags: ignoreversion
Source: "translate_powerpoint.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "translate_image.py";      DestDir: "{app}"; Flags: ignoreversion
; README
Source: "README.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\{#MyAppName}";                          Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}";    Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}";                    Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
