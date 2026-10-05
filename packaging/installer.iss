; 初版 Windows 安装包。只部署已打包程序，不包含输入表格和本地索引。
#ifndef AppVersion
  #define AppVersion "1.1.0"
#endif
#ifndef AppDist
  #define AppDist "..\dist\自动对账工具"
#endif
#ifndef AppExe
  #define AppExe "自动对账工具.exe"
#endif
#ifndef OutputFolder
  #define OutputFolder "..\release"
#endif

[Setup]
AppId={{4093D3C0-6992-45E0-9349-4309A9E96D96}
AppName=自动对账工具
AppVersion={#AppVersion}
AppPublisher=Hzq-0304
AppPublisherURL=https://github.com/Hzq-0304/Auto-Statement-Generator
AppSupportURL=https://github.com/Hzq-0304/Auto-Statement-Generator/issues
DefaultDirName={localappdata}\Programs\AutoStatementGenerator
DefaultGroupName=自动对账工具
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
OutputDir={#OutputFolder}
OutputBaseFilename=AutoStatementGenerator-{#AppVersion}-Setup-x64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#AppExe}
VersionInfoVersion={#AppVersion}
VersionInfoDescription=自动对账工具安装程序
SetupLogging=yes
CloseApplications=no
RestartApplications=no

[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "快捷方式："; Flags: unchecked

[Files]
Source: "{#AppDist}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\自动对账工具"; Filename: "{app}\{#AppExe}"; WorkingDir: "{userdocs}"
Name: "{group}\卸载自动对账工具"; Filename: "{uninstallexe}"
Name: "{autodesktop}\自动对账工具"; Filename: "{app}\{#AppExe}"; WorkingDir: "{userdocs}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "启动自动对账工具"; WorkingDir: "{userdocs}"; Flags: nowait postinstall skipifsilent
