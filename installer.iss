#define MyAppName "PasteMD"
#define MyAppVersion "0.1.4.1"
#define MyAppPublisher "RichQAQ"
#define MyAppExeName "PasteMD.exe"
; AppUserModelID 用于 Win11 通知归属与图标
#define MyAUMID        "RichQAQ.PasteMD"

; 如果是 onefile，发行物在 dist
#define BuildDir       "dist"
; 如果是 onedir，改为发行目录：例如
; #define BuildDir     "dist\\PasteMD"

; ICO 源文件
#define MyIconSrc      "assets\\icons\\logo.ico"

[Setup]
AppId={{4f3f2b18-55a3-4f40-98f6-d01a3e3e0220}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={pf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=PasteMD_pandoc-Setup_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"


[CustomMessages]
english.CreateDesktopIcon=Create a desktop shortcut
english.AdditionalOptions=Additional options:
english.AutoStartup=Start automatically with Windows (current user)
english.RunAfterInstall=Launch {#MyAppName} after installation

chinesesimplified.CreateDesktopIcon=创建桌面快捷方式
chinesesimplified.AdditionalOptions=其他选项：
chinesesimplified.AutoStartup=开机自启（当前用户）
chinesesimplified.RunAfterInstall=安装完成后运行 {#MyAppName}

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalOptions}"
Name: "autorun";     Description: "{cm:AutoStartup}";      GroupDescription: "{cm:AdditionalOptions}"

[Files]
; 主可执行文件
Source: "{#BuildDir}\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Pandoc 打包目录
Source: "third_party\pandoc\*"; DestDir: "{app}\pandoc"; Flags: ignoreversion recursesubdirs createallsubdirs
; 安装时把图标拷到 {app}\icon.ico
Source: "{#MyIconSrc}"; DestDir: "{app}"; DestName: "icon.ico"; Flags: ignoreversion

; 如果你是 onedir，并且有更多运行时文件需要带上，在下方取消注释：
; Source: "{#BuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 为快捷方式写入 AppUserModelID，并指定图标，确保通知归属/图标正确
Name: "{group}\{#MyAppName}";       Filename: "{app}\{#MyAppExeName}"; \
    AppUserModelID: "{#MyAUMID}";   IconFilename: "{app}\icon.ico"

Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; \
    Tasks: desktopicon; AppUserModelID: "{#MyAUMID}"; IconFilename: "{app}\icon.ico"

[Run]
; 安装完成后运行应用
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:RunAfterInstall}"; Flags: nowait postinstall skipifsilent

[Registry]
; 开机自启（当前用户）
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; \
    ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppExeName}"""; \
    Tasks: autorun; Flags: uninsdeletevalue

; 注册 AUMID -> 显示名与图标，供 Win11 通知头部应用区使用
Root: HKCU; Subkey: "Software\Classes\AppUserModelId\{#MyAUMID}"; \
    ValueType: string; ValueName: "DisplayName"; ValueData: "{#MyAppName}"
Root: HKCU; Subkey: "Software\Classes\AppUserModelId\{#MyAUMID}"; \
    ValueType: string; ValueName: "IconUri";    ValueData: "{app}\icon.ico"
; 另外专门加一条“键本身”的项，卸载时直接删整个键（最干净）
Root: HKCU; Subkey: "Software\Classes\AppUserModelId\{#MyAUMID}"; \
    Flags: uninsdeletekey

[UninstallDelete]
; 清理应用数据目录（如有）
Type: filesandordirs; Name: "{userappdata}\{#MyAppName}"

[UninstallRun]
; 卸载前尝试结束进程
Filename: "taskkill"; Parameters: "/f /im {#MyAppExeName}"; Flags: runhidden

