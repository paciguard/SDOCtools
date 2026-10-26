#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif
#ifndef PackageDir
  #define PackageDir "..\..\dist\CFL-SubmitTools"
#endif
#ifndef SetupDir
  #define SetupDir "..\..\dist\CFL-Setup"
#endif

[Setup]
AppId={{1A9B83B8-65B6-49F5-BA87-C79DD195710F}
AppName=凯乐提交工具
AppVersion={#AppVersion}
AppPublisher=CFL
AppPublisherURL=https://github.com/paciguard/SDOCtools
DefaultDirName={localappdata}\Programs\CFL\SubmitTools
DefaultGroupName=凯乐提交工具
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
OutputDir={#SetupDir}
OutputBaseFilename=CFL-SubmitTools-Setup-{#AppVersion}-win-x64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile={#PackageDir}\cfl.ico
WizardSmallImageFile={#PackageDir}\cfl-setup.bmp
UninstallDisplayIcon={app}\cfl.ico
DisableProgramGroupPage=yes
DisableWelcomePage=no
CloseApplications=yes
RestartApplications=no
InfoBeforeFile=INSTALL-WINDOWS.txt
VersionInfoDescription=凯乐提交工具安装程序

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："

[Files]
Source: "{#PackageDir}\*"; DestDir: "{app}"; Excludes: "*.pdb,*.log,*.err"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "INSTALL-WINDOWS.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\凯乐提交工具"; Filename: "{app}\ZTools.Standalone.exe"; WorkingDir: "{app}"; IconFilename: "{app}\cfl.ico"
Name: "{autodesktop}\凯乐提交工具"; Filename: "{app}\ZTools.Standalone.exe"; WorkingDir: "{app}"; IconFilename: "{app}\cfl.ico"; Tasks: desktopicon
Name: "{group}\卸载凯乐提交工具"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\ZTools.Standalone.exe"; Description: "启动凯乐提交工具"; Flags: nowait postinstall skipifsilent
