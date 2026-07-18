; ============================================
; AI 备课助手 - Inno Setup 安装脚本
; 编译: ISCC.exe installer.iss
; ============================================
;
; 【如何发布新版本（覆盖升级）】
;   1. 修改下面的 MyAppVersion 为新版本号（必须大于已安装版本）
;      例如：5.0 -> 5.1 -> 5.2 -> 6.0
;   2. 同步更新本目录下 version.txt 中的版本号
;   3. 运行 install.bat 重新编译生成安装包
;   4. 用户直接运行新安装包即可升级，无需先卸载旧版本
;      - 安装程序会自动停止正在运行的服务
;      - 用户数据（app.db、历史记录、uploads）会被保留
;      - 程序文件会被新版本覆盖
;
; ============================================

#define MyAppName "AI 备课助手"
#define MyAppVersion "5.3"
#define MyAppPublisher "LSW"
#define MyAppExeName "启动备课助手.bat"
#define MyAppIcon "app.ico"

[Setup]
; 应用信息（AppId 固定不变，用于识别同一程序的升级）
AppId={{AIBEIKE-ZHUSHOU-5.0-2026-LSW}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/MewoStar/ai-LSW
AppSupportURL=https://github.com/MewoStar/ai-LSW
AppUpdatesURL=https://github.com/MewoStar/ai-LSW
VersionInfoVersion={#MyAppVersion}.0
DefaultDirName={pf}\AI备课助手
DefaultGroupName=AI 备课助手
AllowNoIcons=yes
OutputDir=.
OutputBaseFilename=AI备课助手5.3_安装包
SetupIconFile={#MyAppIcon}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64
UninstallDisplayIcon={app}\app.ico
UninstallDisplayName=AI 备课助手 {#MyAppVersion}
DisableProgramGroupPage=yes
DisableDirPage=no
; 升级相关：记住上次的安装目录和选项
UsePreviousAppDir=yes
UsePreviousTasks=yes
UsePreviousLanguage=yes
; 不弹"关闭应用程序"对话框（pythonw.exe 是后台进程，用户无法手动关闭）
CloseApplications=no
; 欢迎页图标
WizardImageBackColor=clWhite

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
; 简体中文界面覆盖
SetupAppTitle=安装 - {#MyAppName}
SetupWindowTitle={#MyAppName} {#MyAppVersion}
WelcomeLabel1=欢迎安装 {#MyAppName} {#MyAppVersion}
WelcomeLabel2=这将把 {#MyAppName} 安装到您的计算机上。%n%n点击"下一步"继续，或点击"取消"退出安装。
SelectDirLabel3=安装程序将把 {#MyAppName} 安装到以下文件夹。
SelectDirBrowseLabel=点击"下一步"继续。如果您想选择其他文件夹，点击"浏览"。%n%n安装需要至少 300MB 可用磁盘空间。
DiskSpaceMBLabel=至少需要 [mb] MB 可用磁盘空间。
SelectComponentsLabel2=选择需要安装的组件：
PreparingDesc=准备安装中...
InstallingLabel=正在安装 {#MyAppName}，请稍候...
FinishingTitle=安装完成
FinishedHeadingLabel={#MyAppName} 安装完成
FinishedLabelNoIcons={#MyAppName} 已成功安装到您的计算机。
FinishedLabel={#MyAppName} 已成功安装到您的计算机。
ClickFinish=点击"完成"退出安装程序。
FinishedRestartLabel=要完成安装，需要重新启动计算机。%n%n要现在重新启动吗？
FinishedRestartMessage=要完成安装，需要重新启动计算机。%n%n要现在重新启动吗？
ShowReadmeCheck=查看使用说明
ConfirmUninstall=您确定要完全移除 {#MyAppName} 及其所有组件吗？
UninstallDataError=无法删除文件 "%1"。%n%n是否继续？
OnlyAdminCanUninstall=此程序只能由管理员卸载。
UninstalledAll=已成功从您的计算机中移除 {#MyAppName}。
UninstalledMost={#MyAppName} 卸载完成。%n%n部分文件需要手动删除。
StatusClosingApplications=正在关闭应用程序...
StatusCreateDirs=正在创建目录...
StatusExtractFiles=正在解压文件...
StatusCreateIcons=正在创建快捷方式...
StatusCreateIniEntries=正在创建 INI 条目...
StatusCreateRegistryEntries=正在创建注册表条目...
StatusRegisterFiles=正在注册文件...
StatusSavingUninstall=正在保存卸载信息...
StatusRunProgram=正在完成安装...
; 按钮文字
ButtonBack=< 上一步(&B)
ButtonNext=下一步(&N) >
ButtonInstall=安装(&I)
ButtonOK=确定
ButtonCancel=取消
ButtonYes=是(&Y)
ButtonYesToAll=全部是(&A)
ButtonNo=否(&N)
ButtonNoToAll=全部否(&L)
ButtonFinish=完成(&F)
ButtonBrowse=浏览(&R)...
ButtonWizardBrowse=浏览(&R)...
; 向导标题
SelectDirDesc=选择安装位置
SelectComponentsDesc=选择组件
SelectProgramGroupDesc=选择开始菜单文件夹
SelectTasksDesc=选择附加任务
ReadyDesc=准备安装
; 任务
ReadyMemoDir=安装位置：
ReadyMemoGroup=开始菜单文件夹：
ReadyMemoTasks=附加任务：
ReadyMemoComponents=组件：
SelectTasksLabel2=选择安装程序要执行的附加任务：
; 卸载相关
UninstallAppTitle=卸载 - {#MyAppName}
UninstallAppFullTitle=卸载 {#MyAppName}
RemoveAllFiles=您希望完全删除 {#MyAppName} 的所有文件吗？

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式(&D)"; GroupDescription: "附加图标:"; Flags: checkedonce
Name: "quicklaunchicon"; Description: "在快速启动栏创建快捷方式(&Q)"; GroupDescription: "附加图标:"; Flags: checkedonce

[Dirs]
; 授予 Users 组修改权限，确保 app.db / 历史记录 / uploads / output 等可写
; 否则装到 Program Files 下普通用户无法写入数据库，导致服务无法正常工作
Name: "{app}"; Permissions: users-modify

[Files]
; 打包整个便携版目录内容
; 注意：\*.exe 只排除"根目录"的 exe（即安装包自身），不影响 .venv 下的 python.exe
; \*.zip 同理，只排除根目录的 zip 包
Source: "*"; Excludes: "installer.iss,使用说明.txt,app.ico,version.txt,\*.exe,\*.zip"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; 单独确保使用说明、图标、版本文件包含
Source: "使用说明.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.txt"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 开始菜单 - pythonw.exe 静默启动,无黑窗口
Name: "{group}\AI 备课助手"; Filename: "{app}\.venv\pythonw.exe"; Parameters: """{app}\launcher.py"""; IconFilename: "{app}\app.ico"; WorkingDir: "{app}"
Name: "{group}\停止服务"; Filename: "{app}\停止服务.bat"; IconFilename: "{app}\app.ico"; Flags: runminimized
Name: "{group}\使用说明"; Filename: "notepad.exe"; Parameters: """{app}\使用说明.txt"""; IconFilename: "{app}\app.ico"
Name: "{group}\卸载 AI 备课助手"; Filename: "{uninstallexe}"; IconFilename: "{app}\app.ico"
; 桌面
Name: "{commondesktop}\AI 备课助手"; Filename: "{app}\.venv\pythonw.exe"; Parameters: """{app}\launcher.py"""; IconFilename: "{app}\app.ico"; WorkingDir: "{app}"; Tasks: desktopicon
; 快速启动
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\AI 备课助手"; Filename: "{app}\.venv\pythonw.exe"; Parameters: """{app}\launcher.py"""; IconFilename: "{app}\app.ico"; WorkingDir: "{app}"; Tasks: quicklaunchicon

[Run]
; 交互式安装完成后打开使用说明（静默升级不弹 notepad）
Filename: "notepad.exe"; Parameters: """{app}\使用说明.txt"""; Description: "查看使用说明"; Flags: nowait postinstall skipifsilent runasoriginaluser
; 静默升级完成后自动启动应用（用户从软件内触发升级时无感重启）
Filename: "{app}\.venv\pythonw.exe"; Parameters: """{app}\launcher.py"""; WorkingDir: "{app}"; Flags: nowait skipifnotsilent runasoriginaluser

[UninstallDelete]
; 卸载时清理生成的文件
Type: filesandordirs; Name: "{app}\app.db"
Type: filesandordirs; Name: "{app}\launcher_error.log"
Type: filesandordirs; Name: "{app}\output"
Type: filesandordirs; Name: "{app}\历史记录"
Type: filesandordirs; Name: "{app}\uploads"
Type: filesandordirs; Name: "{app}\__pycache__"
Type: filesandordirs; Name: "{app}\exam_module\__pycache__"
Type: filesandordirs; Name: "{app}\visual_module\__pycache__"
Type: dirifempty; Name: "{app}"

[Code]
// ============================================
// 安装前：自动停止正在运行的服务（否则 .py 文件被锁无法覆盖）
// ============================================
function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
  InstallPath: String;
  StopBat: String;
begin
  Result := True;

  // 1. 尝试运行已安装目录里的"停止服务.bat"（最精准）
  InstallPath := '';
  if RegQueryStringValue(HKLM, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{AIBEIKE-ZHUSHOU-5.0-2026-LSW}_is1',
                         'InstallLocation', InstallPath) then
  begin
    if InstallPath <> '' then
    begin
      StopBat := InstallPath + '\停止服务.bat';
      if FileExists(StopBat) then
      begin
        Exec(StopBat, '', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
        Sleep(1500);
      end;
    end;
  end;

  // 2. 兜底：杀掉占用 5000 端口的进程（无论是否找到旧安装）
  Exec(ExpandConstant('{cmd}'), '/C powershell -NoProfile -Command "Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue }"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(1000);
end;

// ============================================
// 卸载时：清理安装目录（仅显式卸载时触发，覆盖升级不会触发）
// ============================================
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    // 先停掉服务，避免文件被占用
    if FileExists(ExpandConstant('{app}\停止服务.bat')) then
    begin
      Exec(ExpandConstant('{app}\停止服务.bat'), '', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      Sleep(1000);
    end;
    // 强制删除 .venv 目录（Inno Setup 默认不删非空目录）
    if DirExists(ExpandConstant('{app}\.venv')) then
      DelTree(ExpandConstant('{app}\.venv'), True, True, True);
    // 强制删除整个安装目录（含残留的 pycache 等）
    if DirExists(ExpandConstant('{app}')) then
      DelTree(ExpandConstant('{app}'), True, True, True);
  end;
end;

