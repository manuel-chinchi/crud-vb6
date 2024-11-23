
[Code]
#define MyAppName "crud vb6 (by Manuel Chinchi)"
#define MyAppVersion "1.0"
#define MyAppPublisher "Manuel Chinchi"
#define MyAppURL "https://github.com/manuel-chinchi/crud-vb6"
#define MyAppExeName "crud_vb6.exe"
#define RootPath SourcePath + "\..\.."

[Setup]
; NOTE: The value of AppId uniquely identifies this application. Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{4A6E1CF4-8BF8-4446-A018-FD22172C42F6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=admin
Compression=lzma
SolidCompression=yes
WizardStyle=classic
SetupIconFile={#RootPath}\Icons\SetupClassicIcon.ico
WizardImageFile={#RootPath}\Images\WizClassicImage.bmp
WizardSmallImageFile={#RootPath}\Images\WizClassicSmallImage.bmp
OutputDir={#RootPath}\InnoSetup_Installer
OutputBaseFilename=setup_crud_vb6

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"
Name: "ukrainian"; MessagesFile: "compiler:Languages\Ukrainian.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "{#RootPath}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#RootPath}\Data\*"; DestDir: "{app}\Data"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RootPath}\Reports\*"; DestDir: "{app}\Reports"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RootPath}\Scripts\*"; DestDir: "{app}\Scripts"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RootPath}\Dependences\*"; DestDir: "{app}\Dependences"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#RootPath}\Dependences\SQLite\sqlite.dll"; DestDir: "{app}"; Flags: ignoreversion
[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent runascurrentuser 
Filename: "{cmd}"; Parameters: "/C ""{app}\Scripts\dependences.bat"""; WorkingDir: "{app}\Scripts"; Flags: runascurrentuser
