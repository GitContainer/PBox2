unit db.PBoxForm;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI, System.SysUtils, System.StrUtils, System.Classes, System.Types, System.IniFiles, System.Math, System.UITypes, System.ImageList,
  Vcl.Graphics, Vcl.Controls, Vcl.Buttons, Vcl.Forms, Vcl.ExtCtrls, Vcl.Menus, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.ImgList, Vcl.ToolWin, Vcl.Imaging.jpeg,
  db.uCommon, db.uBaseForm, db.uCreateDelphiDll, db.uCreateVCDialogDll, db.AddEXE, db.uCreateEXE;

type
  TfrmPBox = class(TUIBaseForm)
    ilMainMenu: TImageList;
    pnlBottom: TPanel;
    mmMainMenu: TMainMenu;
    clbrPModule: TCoolBar;
    tlbPModule: TToolBar;
    pnlInfo: TPanel;
    lblInfo: TLabel;
    pnlTime: TPanel;
    lblTime: TLabel;
    tmrDateTime: TTimer;
    rzpgcntrlAll: TPageControl;
    tsButton: TTabSheet;
    tsList: TTabSheet;
    tsDll: TTabSheet;
    pnlIP: TPanel;
    lblIP: TLabel;
    bvlIP: TBevel;
    pmTray: TPopupMenu;
    mniTrayShowForm: TMenuItem;
    mniTrayLine01: TMenuItem;
    mniTrayExit: TMenuItem;
    imgDllFormBack: TImage;
    imgButtonBack: TImage;
    imgListBack: TImage;
    ilPModule: TImageList;
    pnlModuleDialog: TPanel;
    pnlModuleDialogTitle: TPanel;
    imgSubModuleClose: TImage;
    bvlModule01: TBevel;
    pnlWeb: TPanel;
    lblWeb: TLabel;
    bvlWeb: TBevel;
    bvlModule02: TBevel;
    pnlLogin: TPanel;
    lblLogin: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tmrDateTimeTimer(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure mniTrayExitClick(Sender: TObject);
    procedure mniTrayShowFormClick(Sender: TObject);
    procedure imgSubModuleCloseClick(Sender: TObject);
    procedure imgSubModuleCloseMouseEnter(Sender: TObject);
    procedure imgSubModuleCloseMouseLeave(Sender: TObject);
  private
    FlstAllDll    : THashedStringList;
    FUIShowStyle  : TShowStyle;
    FUIViewStyle  : TViewStyle;
    FbMaxForm     : Boolean;
    FDelphiDllForm: TForm;
    FintBakRow    : Integer;
    procedure ShowPageTabView(const bShow: Boolean = False);
    procedure ReadConfigUI;
    { 扫描 EXE 文件，读取配置文件 }
    procedure ScanPlugins_EXE;
    { 扫描 Dll 文件，仅限于插件目录(plugins) }
    procedure ScanPlugins_Dll;
    { 扫描插件目录 }
    procedure ScanPlugins;
    { 创建模块菜单 }
    procedure CreateModuleMenu;
    { 点击菜单 }
    procedure OnMenuItemClick(Sender: TObject);
    { 创建新的 Dll 窗体 }
    procedure CreateDllForm;
    { 待实现 }
    function PBoxRun_VC_MFCDll: Boolean;
    function PBoxRun_QT_GUIDll: Boolean;
    { 系统配置 }
    procedure OnSysConfig(Sender: TObject);
    { Delphi Dll Form 窗体关闭事件 }
    procedure OnDelphiDllFormDestory(Sender: TObject);
    { 获取 EXE 文件的图标 }
    function GetExeFileIcon(const strFileName: String): Integer; overload;
    function GetExeFileIcon(const strEXEInfo, strFileName: string): Integer; overload;
    { 获取 Dll 文件的图标 }
    function GetDllFileIcon(const strPModuleName, strSModuleName, strIconFileName: string): Integer;
    function ReadDllIconFromConfig(const strPModule, strSModule: string): Integer;
    procedure ReCreate;
    { 界面显示风格 }
    procedure ChangeUI;
    { 菜单式风格 }
    procedure ChangeUI_Menu;
    { 按钮式风格 }
    procedure ChangeUI_Button;
    { 分栏式风格 }
    procedure ChangeUI_List(const bActivePage: Boolean = True);
    { 参数置空，恢复默认值 }
    procedure FillParamBlank;
    { 点击父模块时，动态创建子模块 }
    procedure OnParentModuleButtonClick(Sender: TObject);
    { 创建所有子模块对话框窗体 }
    procedure CreateSubModulesFormDialog(const strPModuleName: string); overload;
    procedure CreateSubModulesFormDialog(const mmItem: TMenuItem); overload;
    { 列表显示风格，创建子模块 DLL 窗体 }
    procedure OnSubModuleButtonClick(Sender: TObject);
    { 销毁分栏式界面 }
    procedure FreeListViewSubModule;
    { 窗体大小改变时，重新创建分栏式的界面 }
    procedure ResizeListViewSubModule;
    { 分栏式显示时，当鼠标进入 label 时 }
    procedure OnSubModuleMouseEnter(Sender: TObject);
    { 分栏式显示时，当鼠标离开 label 时 }
    procedure OnSubModuleMouseLeave(Sender: TObject);
    { 创建子模块 DLL 模块 }
    procedure OnSubModuleListClick(Sender: TObject);
    { 获取垂直位置间隔 }
    function GetMaxInstance(const intCurrentIndex, intCount: Integer): Integer;
  protected
    { 销毁上一次创建的 Dll 窗体 }
    procedure WMDESTORYPREDLLFORM(var msg: TMessage); message WM_DESTORYPREDLLFORM;
    { 创建新的 DLL 窗体 }
    procedure WMCREATENEWDLLFORM(var msg: TMessage); message WM_CREATENEWDLLFORM;
  end;

var
  frmPBox: TfrmPBox;

implementation

uses db.ConfigForm;

{$R *.dfm}

{ 待实现 }
function TfrmPBox.PBoxRun_VC_MFCDll: Boolean;
begin
  Result := True;
end;

{ 待实现 }
function TfrmPBox.PBoxRun_QT_GUIDll: Boolean;
begin
  Result := True;
end;

{ 系统配置 }
procedure TfrmPBox.OnSysConfig(Sender: TObject);
begin
  if ShowConfigForm(FlstAllDll) then
  begin
    Hide;
    SendMessage(Handle, WM_DESTORYPREDLLFORM, 0, 0);
    ReCreate;
    Show;
  end;
end;

{ Delphi Dll Form 窗体关闭事件 }
procedure TfrmPBox.OnDelphiDllFormDestory(Sender: TObject);
begin
  if FDelphiDllForm <> nil then
  begin
    FreeLibrary(FDelphiDllForm.Tag);
    FDelphiDllForm         := nil;
    g_strCreateDllFileName := '';
    lblInfo.Caption        := '';
    if FUIShowStyle = ssButton then
      rzpgcntrlAll.ActivePageIndex := 0
    else if FUIShowStyle = ssList then
      rzpgcntrlAll.ActivePageIndex := 1;
  end;
end;

{ 创建新的 DLL 窗体 }
procedure TfrmPBox.WMCREATENEWDLLFORM(var msg: TMessage);
var
  hDll                              : HMODULE;
  ShowDllForm                       : Tdb_ShowDllForm_Plugins;
  frm                               : TFormClass;
  strParamModuleName, strModuleName : PAnsiChar;
  strFormClassName, strFormTitleName: PAnsiChar;
  strIconFileName                   : PAnsiChar;
  ft                                : TSPFileType;
  strFileValue                      : String;
begin
  if g_strCreateDllFileName = '' then
    Exit;

  { exe 文件 }
  if (CompareText(ExtractFileExt(g_strCreateDllFileName), '.exe') = 0) or (CompareText(ExtractFileExt(g_strCreateDllFileName), '.msc') = 0) then
  begin
    strFileValue := FlstAllDll.Values[g_strCreateDllFileName];
    PBoxRun_IMAGE_EXE(g_strCreateDllFileName, strFileValue, rzpgcntrlAll, tsDll, lblInfo, FUIShowStyle);
    Exit;
  end;

  { Dll 文件，获取文件类型 }
  hDll := LoadLibrary(PChar(g_strCreateDllFileName));
  try
    ShowDllForm := GetProcAddress(hDll, c_strDllExportName);
    ShowDllForm(frm, ft, strParamModuleName, strModuleName, strFormClassName, strFormTitleName, strIconFileName, False);
  finally
    FreeLibrary(hDll);
  end;

  { 根据 DLL 文件类型的不同，创建 DLL 窗体 }
  case ft of
    ftDelphiDll:
      PBoxRun_DelphiDll(FDelphiDllForm, rzpgcntrlAll, tsDll, OnDelphiDllFormDestory);
    ftVCDialogDll:
      PBoxRun_VC_DLGDll(rzpgcntrlAll, tsDll, lblInfo, FUIShowStyle);
    ftVCMFCDll:
      PBoxRun_VC_MFCDll;
    ftQTDll:
      PBoxRun_QT_GUIDll;
  end;
end;

{ 创建新的 Dll 窗体 }
procedure TfrmPBox.CreateDllForm;
begin
  PostMessage(Handle, WM_CREATENEWDLLFORM, 0, 0);
end;

{ 销毁上一次创建的 Dll 窗体 }
procedure TfrmPBox.WMDESTORYPREDLLFORM(var msg: TMessage);
var
  hDll    : HMODULE;
  hProcess: Cardinal;
begin
  { 是否是 EXE 窗体 }
  if g_hEXEProcessID <> 0 then
  begin
    hProcess := OpenProcess(PROCESS_TERMINATE, False, g_hEXEProcessID);
    TerminateProcess(hProcess, 0);
    g_hEXEProcessID := 0;
    if Visible then
      CreateDllForm;

    Exit;
  end;

  { 是否是 Delphi 创建的 Dll 窗体 }
  if FDelphiDllForm <> nil then
  begin
    hDll := FDelphiDllForm.Tag;
    FDelphiDllForm.Free;
    FDelphiDllForm := nil;
    FreeLibrary(hDll);

    if Visible then
      CreateDllForm;

    Exit;
  end;

  { 是否是 VC 创建的 Dll 窗体 }
  if g_intVCDialogDllFormHandle = 0 then
  begin
    { 创建新的 Dll 窗体 }
    if Visible then
      CreateDllForm;
  end
  else
  begin
    { 销毁 VC Dialog Dll 窗体消息 }
    FreeVCDialogDllForm;
  end;
end;

{ 点击菜单 }
procedure TfrmPBox.OnMenuItemClick(Sender: TObject);
begin
  lblInfo.Caption := TMenuItem(TMenuItem(Sender).Owner).Caption + ' - ' + TMenuItem(Sender).Caption;

  { 如果已经创建了，就不在重复创建了 }
  if (g_strCreateDllFileName <> '') and (g_strCreateDllFileName = FlstAllDll.Names[TMenuItem(Sender).Tag]) then
    Exit;

  g_strCreateDllFileName := FlstAllDll.Names[TMenuItem(Sender).Tag];

  { 销毁上一次创建的 Dll 窗体 }
  PostMessage(Handle, WM_DESTORYPREDLLFORM, 0, 0);
end;

function TfrmPBox.GetExeFileIcon(const strFileName: String): Integer;
var
  IcoExe: TIcon;
begin
  Result := -1;
  if CompareText(ExtractFileExt(strFileName), '.msc') = 0 then
  begin
    IcoExe := TIcon.Create;
    try
      { 从 .msc 文件中获取图标 }
      LoadIconFromMSCFile(strFileName, IcoExe);
      Result := ilMainMenu.AddIcon(IcoExe);
    finally
      IcoExe.Free;
    end;
  end
  else
  begin
    if ExtractIcon(HInstance, PChar(strFileName), $FFFFFFFF) > 0 then
    begin
      IcoExe := TIcon.Create;
      try
        IcoExe.Handle := ExtractIcon(HInstance, PChar(strFileName), 0);
        Result        := ilMainMenu.AddIcon(IcoExe);
      finally
        IcoExe.Free;
      end;
    end;
  end;
end;

{ 获取 EXE 文件的图标 }
function TfrmPBox.GetExeFileIcon(const strEXEInfo, strFileName: string): Integer;
var
  strIconFileName    : String;
  IcoExe             : TIcon;
  strCurrIconFileName: String;
begin
  strIconFileName := strEXEInfo.Split([';'])[4];
  if Trim(strIconFileName) = '' then
  begin
    Result := GetExeFileIcon(strFileName);
  end
  else
  begin
    if FileExists(strIconFileName) then
    begin
      strCurrIconFileName := ExtractFilePath(ParamStr(0)) + 'plugins\Icon\' + ExtractFileName(strIconFileName);

      if not DirectoryExists(ExtractFilePath(strCurrIconFileName)) then
        ForceDirectories(ExtractFilePath(strCurrIconFileName));

      if not FileExists(strCurrIconFileName) then
        CopyFile(PChar(strIconFileName), PChar(strCurrIconFileName), False);

      IcoExe := TIcon.Create;
      try
        IcoExe.LoadFromFile(strCurrIconFileName);
        Result := ilMainMenu.AddIcon(IcoExe);
      finally
        IcoExe.Free;
      end;
    end
    else
    begin
      Result := GetExeFileIcon(strFileName);
    end;
  end;
end;

{ 扫描 EXE 文件，读取配置文件 }
procedure TfrmPBox.ScanPlugins_EXE;
var
  lstEXE      : TStringList;
  I           : Integer;
  strEXEInfo  : String;
  strFileName : String;
  intIconIndex: Integer;
begin
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
  begin
    lstEXE := TStringList.Create;
    try
      ReadSection('EXE', lstEXE);
      for I := 0 to lstEXE.Count - 1 do
      begin
        strFileName  := lstEXE.Strings[I];
        strEXEInfo   := ReadString('EXE', strFileName, '');
        intIconIndex := GetExeFileIcon(strEXEInfo, strFileName);
        if strEXEInfo.CountChar(';') = 4 then
          strEXEInfo := strEXEInfo + ';' + IntToStr(intIconIndex)
        else
          strEXEInfo := strEXEInfo + ';;' + IntToStr(intIconIndex);
        FlstAllDll.Add(Format('%s=%s', [strFileName, strEXEInfo]));
      end;
    finally
      lstEXE.Free;
    end;
    Free;
  end;
end;

function TfrmPBox.ReadDllIconFromConfig(const strPModule, strSModule: string): Integer;
var
  strIconFilePath: String;
  strIconFileName: String;
  IcoExe         : TIcon;
begin
  Result := -1;
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
  begin
    strIconFilePath := ReadString(c_strIniModuleSection, Format('%s_%s_ICON', [strPModule, strSModule]), '');
    strIconFileName := ExtractFilePath(ParamStr(0)) + 'plugins\icon\' + strIconFilePath;
    if FileExists(strIconFileName) then
    begin
      IcoExe := TIcon.Create;
      try
        IcoExe.LoadFromFile(strIconFileName);
        Result := ilMainMenu.AddIcon(IcoExe);
      finally
        IcoExe.Free;
      end;
    end;
    Free;
  end;
end;

{ 获取 Dll 文件的图标 }
function TfrmPBox.GetDllFileIcon(const strPModuleName, strSModuleName, strIconFileName: string): Integer;
var
  strCurrIconFileName: String;
  IcoExe             : TIcon;
begin
  if Trim(strIconFileName) = '' then
  begin
    { 从配置文件中读取图标信息 }
    Result := ReadDllIconFromConfig(strPModuleName, strSModuleName);
  end
  else
  begin
    if FileExists(strIconFileName) then
    begin
      strCurrIconFileName := ExtractFilePath(ParamStr(0)) + 'plugins\icon\' + ExtractFileName(strIconFileName);
      if not DirectoryExists(ExtractFilePath(strCurrIconFileName)) then
        ForceDirectories(ExtractFilePath(strCurrIconFileName));
      if not FileExists(strCurrIconFileName) then
        CopyFile(PChar(strIconFileName), PChar(strCurrIconFileName), False);

      IcoExe := TIcon.Create;
      try
        IcoExe.LoadFromFile(strCurrIconFileName);
        Result := ilMainMenu.AddIcon(IcoExe);
      finally
        IcoExe.Free;
      end;
    end
    else
    begin
      { 从配置文件中读取图标信息 }
      Result := ReadDllIconFromConfig(strPModuleName, strSModuleName);
    end;
  end;
end;

procedure TfrmPBox.ScanPlugins_Dll;
var
  hDll                          : HMODULE;
  ShowDllForm                   : Tdb_ShowDllForm_Plugins;
  frm                           : TFormClass;
  strPModuleName, strSModuleName: PAnsiChar;
  strClassName, strWindowName   : PAnsiChar;
  strIconFileName               : PAnsiChar;
  ft                            : TSPFileType;
  strDllFileName                : String;
  strInfo                       : string;
  I, Count                      : Integer;
  lstTemp                       : TStringList;
  intIconIndex                  : Integer;
begin
  lstTemp := TStringList.Create;
  try
    { 扫描 Dll 文件，仅限于插件目录(plugins) }
    SearchPlugInsDllFile(lstTemp);
    Count := lstTemp.Count;
    if Count <= 0 then
      Exit;

    for I := 0 to Count - 1 do
    begin
      strDllFileName := lstTemp.Strings[I];
      hDll           := LoadLibrary(PChar(strDllFileName));
      if hDll = 0 then
        Continue;

      try
        ShowDllForm := GetProcAddress(hDll, c_strDllExportName);
        if not Assigned(ShowDllForm) then
        begin
          FreeLibrary(hDll);
          Continue;
        end;

        { 获取 Dll 参数 }
        ShowDllForm(frm, ft, strPModuleName, strSModuleName, strClassName, strWindowName, strIconFileName, False);
        intIconIndex := GetDllFileIcon(string(strPModuleName), string(strSModuleName), string(strIconFileName));
        strInfo      := strDllFileName + '=' + string(strPModuleName) + ';' + string(strSModuleName) + ';' + string(strClassName) + ';' + string(strWindowName) + ';' + string(strIconFileName) + ';' + IntToStr(intIconIndex);
        FlstAllDll.Add(strInfo);
      finally
        FreeLibrary(hDll);
      end;
    end;
  finally
    lstTemp.Free;
  end;
end;

{ 创建模块菜单 }
procedure TfrmPBox.CreateModuleMenu;
var
  I             : Integer;
  strInfo       : String;
  strPModuleName: String;
  strSModuleName: String;
  mmPM          : TMenuItem;
  mmSM          : TMenuItem;
  intIconIndex  : Integer;
begin
  for I := 0 to FlstAllDll.Count - 1 do
  begin
    strInfo        := FlstAllDll.ValueFromIndex[I];
    strPModuleName := strInfo.Split([';'])[0];
    strSModuleName := strInfo.Split([';'])[1];
    intIconIndex   := StrToInt(strInfo.Split([';'])[5]);

    { 如果父菜单不存在，创建父菜单 }
    mmPM := mmMainMenu.Items.Find(string(strPModuleName));
    if mmPM = nil then
    begin
      mmPM         := TMenuItem.Create(mmMainMenu);
      mmPM.Caption := string((strPModuleName));
      mmMainMenu.Items.Add(mmPM);
    end;

    { 创建子菜单 }
    mmSM            := TMenuItem.Create(mmPM);
    mmSM.Caption    := string((strSModuleName));
    mmSM.Tag        := I;
    mmSM.ImageIndex := intIconIndex;
    mmSM.OnClick    := OnMenuItemClick;
    mmPM.Add(mmSM);
  end;
end;

{ 扫描插件目录 }
procedure TfrmPBox.ScanPlugins;
var
  strDllModulePath: string;
begin
  strDllModulePath := ExtractFilePath(ParamStr(0)) + 'plugins';
  SetDllDirectory(PChar(strDllModulePath));
  if not DirectoryExists(ExtractFilePath(ParamStr(0)) + 'plugins') then
    Exit;

  { 扫描 Dll 文件，添加到列表；当前插件目录 (plugins) }
  ScanPlugins_Dll;

  { 扫描 EXE 文件，添加到列表；读取配置文件 }
  ScanPlugins_EXE;

  { 排序模块 }
  SortModuleList(FlstAllDll);

  { 创建模块菜单 }
  CreateModuleMenu;
end;

procedure TfrmPBox.tmrDateTimeTimer(Sender: TObject);
const
  WeekDay: array [1 .. 7] of String = ('星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六');
var
  strWebDownSpeed, strWebUpSpeed: String;
begin
  lblTime.Caption := FormatDateTime('YYYY-MM-DD hh:mm:ss', Now) + ' ' + WeekDay[DayOfWeek(Now)];
  GetWebSpeed(strWebDownSpeed, strWebUpSpeed);
  lblWeb.Caption := Format('下载↓：%s  上传↑：%s', [strWebDownSpeed, strWebUpSpeed]);
end;

procedure TfrmPBox.ShowPageTabView(const bShow: Boolean);
var
  I: Integer;
begin
  for I := 0 to rzpgcntrlAll.PageCount - 1 do
  begin
    rzpgcntrlAll.Pages[I].TabVisible := bShow;
  end;
end;

procedure TfrmPBox.ReadConfigUI;
var
  bShowImage  : Boolean;
  strImageBack: String;
begin
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
  begin
    Caption        := ReadString(c_strIniUISection, 'Title', c_strTitle);
    TitleString    := Caption;
    MulScreenPos   := ReadBool(c_strIniUISection, 'MulScreen', False);
    FbMaxForm      := ReadBool(c_strIniUISection, 'MAXSIZE', False);
    FormStyle      := TFormStyle(Integer(ReadBool(c_strIniUISection, 'OnTop', False)) * 3);
    CloseToTray    := ReadBool(c_strIniUISection, 'CloseMini', False);
    pnlWeb.Visible := ReadBool(c_strIniUISection, 'ShowWebSpeed', False);
    bShowImage     := ReadBool(c_strIniUISection, 'showbackimage', False);
    strImageBack   := ReadString(c_strIniUISection, 'filebackimage', '');
    if (bShowImage) and (Trim(strImageBack) <> '') and (FileExists(strImageBack)) then
    begin
      imgDllFormBack.Picture.LoadFromFile(strImageBack);
      imgButtonBack.Picture.LoadFromFile(strImageBack);
      imgListBack.Picture.LoadFromFile(strImageBack);
    end
    else
    begin
      imgDllFormBack.Picture.Assign(nil);
      imgButtonBack.Picture.Assign(nil);
      imgListBack.Picture.Assign(nil);
    end;
    Free;
  end;
end;

{ 参数置空，恢复默认值 }
procedure TfrmPBox.FillParamBlank;
var
  I, J: Integer;
begin
  g_intVCDialogDllFormHandle := 0;
  g_strCreateDllFileName     := '';
  g_bExitProgram             := False;
  g_hEXEProcessID            := 0;
  FUIShowStyle               := GetShowStyle;
  FUIViewStyle               := vsSingle;
  FDelphiDllForm             := nil;
  clbrPModule.Visible        := False;
  pnlWeb.Visible             := False;
  ilMainMenu.Clear;
  ilPModule.Clear;
  lblInfo.Caption   := '';
  tlbPModule.Images := nil;
  tlbPModule.Height := 30;
  tlbPModule.Menu   := nil;
  FintBakRow        := 0;

  for I := tlbPModule.ButtonCount - 1 downto 0 do
  begin
    tlbPModule.Buttons[I].Free;
  end;

  mmMainMenu.AutoMerge := False;
  for I                := mmMainMenu.Items.Count - 1 downto 0 do
  begin
    for J := mmMainMenu.Items.Items[I].Count - 1 downto 0 do
    begin
      mmMainMenu.Items.Items[I].Items[J].Free;
    end;
    mmMainMenu.Items.Items[I].Free;
  end;
  mmMainMenu.Items.Clear;

  FlstAllDll.Clear;
end;

procedure TfrmPBox.ReCreate;
begin
  { 初始化参数 }
  FillParamBlank;

  { 初始化界面 }
  ShowPageTabView(False);
  rzpgcntrlAll.ActivePage := tsDll;
  ReadConfigUI;

  { 扫描插件目录 }
  ScanPlugins;

  { 界面显示风格 }
  ChangeUI;
end;

procedure TfrmPBox.FormCreate(Sender: TObject);
begin
  { 列表显示风格，关闭按钮状态 }
  LoadButtonBmp(imgSubModuleClose, 'Close', 0);
  OnConfig         := OnSysConfig;
  TrayIconPMenu    := pmTray;
  FlstAllDll       := THashedStringList.Create;
  lblLogin.Caption := g_strCurrentLoginName;

  { 显示 时间 }
  tmrDateTime.OnTimer(nil);

  { 显示 IP }
  lblIP.Caption := GetNativeIP;
  lblIP.Left    := (lblIP.Parent.Width - lblIP.Width) div 2;

  ReCreate;
end;

procedure TfrmPBox.FormActivate(Sender: TObject);
begin
  { 最大化窗体 }
  if FbMaxForm then
    pnlDBLClick(nil);
end;

procedure TfrmPBox.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  g_bExitProgram := True;
  CloseDelphiDllForm;
  FreeVCDialogDllForm;
end;

procedure TfrmPBox.FormDestroy(Sender: TObject);
begin
  FlstAllDll.Free;
end;

function EnumChildFunc(hDllForm: THandle; hParentHandle: THandle): Boolean; stdcall;
var
  rctClient: TRect;
begin
  Result := True;

  { 判断是否是 Dll 的窗体句柄 }
  if GetParent(hDllForm) = 0 then
  begin
    { 更改 Dll 窗体大小 }
    GetWindowRect(hParentHandle, rctClient);
    SetWindowPos(hDllForm, hParentHandle, 0, 0, rctClient.Width, rctClient.Height, SWP_NOZORDER + SWP_NOACTIVATE);
  end;
end;

procedure TfrmPBox.FormResize(Sender: TObject);
begin
  { 更改 Dll 窗体大小 }
  EnumChildWindows(Handle, @EnumChildFunc, tsDll.Handle);

  if FUIShowStyle = ssButton then
  begin
    pnlModuleDialog.Top  := clbrPModule.Height + (pnlModuleDialog.Parent.Height - pnlModuleDialog.Height) div 2 - 60;
    pnlModuleDialog.Left := (pnlModuleDialog.Parent.Width - pnlModuleDialog.Width) div 2;
  end;

  if FUIShowStyle = ssList then
  begin
    ResizeListViewSubModule;
  end;
end;

procedure TfrmPBox.mniTrayShowFormClick(Sender: TObject);
begin
  MainTrayIcon.OnDblClick(nil);
end;

procedure TfrmPBox.mniTrayExitClick(Sender: TObject);
begin
  CloseToTray := False;
  Close;
end;

{ 界面显示风格 }
procedure TfrmPBox.ChangeUI;
begin
  case FUIShowStyle of
    ssMenu:
      ChangeUI_Menu;   // 菜单式
    ssButton:          //
      ChangeUI_Button; // 按钮式
    ssList:            //
      ChangeUI_List;   // 分栏式
  end;
end;

{ ------------------------------------------------------------------------------- 菜单式风格 ------------------------------------------------------------------------------- }
procedure TfrmPBox.ChangeUI_Menu;
begin
  tlbPModule.Menu      := mmMainMenu;
  mmMainMenu.AutoMerge := True;
  clbrPModule.Visible  := True;
end;

{ ------------------------------------------------------------------------------- 按钮式风格 ------------------------------------------------------------------------------- }
procedure TfrmPBox.imgSubModuleCloseClick(Sender: TObject);
var
  I: Integer;
begin
  pnlModuleDialog.Visible := False;
  for I                   := 0 to tlbPModule.ButtonCount - 1 do
  begin
    tlbPModule.Buttons[I].Down := False;
  end;
  lblInfo.Caption := '';
end;

procedure TfrmPBox.imgSubModuleCloseMouseEnter(Sender: TObject);
begin
  { 列表显示风格，关闭按钮状态 }
  LoadButtonBmp(imgSubModuleClose, 'Close', 1);
end;

procedure TfrmPBox.imgSubModuleCloseMouseLeave(Sender: TObject);
begin
  { 列表显示风格，关闭按钮状态 }
  LoadButtonBmp(imgSubModuleClose, 'Close', 0);
end;

{ 列表显示风格，创建子模块 DLL 窗体 }
procedure TfrmPBox.OnSubModuleButtonClick(Sender: TObject);
var
  I, J         : Integer;
  mmItem       : TMenuItem;
  strPMouleName: string;
  strSMouleName: string;
begin
  mmItem := nil;

  for I := 0 to tlbPModule.ButtonCount - 1 do
  begin
    if tlbPModule.Components[I] is TToolButton then
    begin
      if TToolButton(tlbPModule.Components[I]).Down then
      begin
        strPMouleName := TToolButton(tlbPModule.Components[I]).Caption;
        Break;
      end;
    end;
  end;
  strSMouleName := TSpeedButton(Sender).Caption;

  for I := 0 to mmMainMenu.Items.Count - 1 do
  begin
    if CompareText(mmMainMenu.Items.Items[I].Caption, strPMouleName) = 0 then
    begin
      for J := 0 to mmMainMenu.Items.Items[I].Count - 1 do
      begin
        if CompareText(mmMainMenu.Items.Items[I].Items[J].Caption, strSMouleName) = 0 then
        begin
          mmItem := mmMainMenu.Items.Items[I].Items[J];
          Break;
        end;
      end;
    end;
  end;

  pnlModuleDialog.Visible := True;
  OnMenuItemClick(mmItem);
end;

{ 创建所有子模块对话框窗体 }
procedure TfrmPBox.CreateSubModulesFormDialog(const mmItem: TMenuItem);
const
  c_intCols         = 5;
  c_intButtonWidth  = 128;
  c_intButtonHeight = 64;
  c_intMiniTop      = 2;
  c_intMiniLeft     = 2;
  c_intHorSpace     = 2;
  c_intVerSpace     = 2;
var
  arrSB   : array of TSpeedButton;
  I, Count: Integer;
begin
  { 释放先前创建的按钮 }
  Count := pnlModuleDialog.ComponentCount;
  if Count > 0 then
  begin
    for I := Count - 1 downto 0 do
    begin
      if pnlModuleDialog.Components[I] is TSpeedButton then
      begin
        TSpeedButton(pnlModuleDialog.Components[I]).Free;
      end;
    end;
  end;

  { 创建新的子模块按钮 }
  SetLength(arrSB, mmItem.Count);
  for I := 0 to mmItem.Count - 1 do
  begin
    arrSB[I]            := TSpeedButton.Create(pnlModuleDialog);
    arrSB[I].Parent     := pnlModuleDialog;
    arrSB[I].Caption    := mmItem.Items[I].Caption;
    arrSB[I].Width      := c_intButtonWidth;
    arrSB[I].Height     := c_intButtonHeight;
    arrSB[I].GroupIndex := 1;
    arrSB[I].Flat       := True;
    arrSB[I].Top        := pnlModuleDialogTitle.Height + c_intMiniTop + (c_intCols + c_intButtonHeight + c_intVerSpace) * (I div c_intCols);
    arrSB[I].Left       := c_intMiniLeft + (c_intButtonWidth + c_intHorSpace) * (I mod c_intCols);
    arrSB[I].Tag        := mmItem.Items[I].Tag;
    arrSB[I].OnClick    := OnSubModuleButtonClick;
    ilMainMenu.GetBitmap(mmItem.Items[I].ImageIndex, arrSB[I].Glyph);
  end;
  pnlModuleDialog.Visible := True;
end;

{ 创建所有子模块对话框窗体 }
procedure TfrmPBox.CreateSubModulesFormDialog(const strPModuleName: string);
var
  I: Integer;
begin
  for I := 0 to mmMainMenu.Items.Count - 1 do
  begin
    if CompareText(mmMainMenu.Items.Items[I].Caption, strPModuleName) = 0 then
    begin
      CreateSubModulesFormDialog(mmMainMenu.Items.Items[I]);
      Break;
    end;
  end;
end;

{ 点击父模块时，动态创建子模块 }
procedure TfrmPBox.OnParentModuleButtonClick(Sender: TObject);
var
  I             : Integer;
  strPMdouleName: string;
begin
  rzpgcntrlAll.ActivePage := tsButton;
  for I                   := 0 to tlbPModule.ButtonCount - 1 do
  begin
    tlbPModule.Buttons[I].Down := False;
  end;
  TToolButton(Sender).Down     := True;
  strPMdouleName               := TToolButton(Sender).Caption;
  pnlModuleDialogTitle.Caption := strPMdouleName;
  CreateSubModulesFormDialog(strPMdouleName);
end;

{ 按钮式风格 }
procedure TfrmPBox.ChangeUI_Button;
var
  tmpTB          : TToolButton;
  I              : Integer;
  strIconFilePath: String;
  strIconFileName: String;
  icoPModule     : TIcon;
begin
  tlbPModule.Images := ilPModule;
  tlbPModule.Height := 58;

  { 获取所有父模块图标 }
  for I := 0 to mmMainMenu.Items.Count - 1 do
  begin
    with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
    begin
      strIconFilePath := mmMainMenu.Items.Items[I].Caption + '_ICON';
      strIconFileName := ExtractFilePath(ParamStr(0)) + 'plugins\icon\' + ReadString(c_strIniModuleSection, strIconFilePath, '');
      Free;
    end;

    if FileExists(strIconFileName) then
    begin
      icoPModule := TIcon.Create;
      try
        icoPModule.LoadFromFile(strIconFileName);
        ilPModule.AddIcon(icoPModule);
      finally
        icoPModule.Free;
      end;
    end;
  end;

  for I := mmMainMenu.Items.Count - 1 downto 0 do
  begin
    tmpTB            := TToolButton.Create(tlbPModule);
    tmpTB.Parent     := tlbPModule;
    tmpTB.Caption    := mmMainMenu.Items.Items[I].Caption;
    tmpTB.ImageIndex := I;
    tmpTB.OnClick    := OnParentModuleButtonClick;
  end;
  clbrPModule.Visible := True;
end;

{ ------------------------------------------------------------------------------- 分栏式风格 ------------------------------------------------------------------------------- }

{ 分栏式显示时，当鼠标进入 label 时 }
procedure TfrmPBox.OnSubModuleMouseEnter(Sender: TObject);
begin
  TLabel(Sender).Font.Color := RGB(0, 0, 255);
  TLabel(Sender).Font.Style := TLabel(Sender).Font.Style + [fsUnderline];
end;

{ 分栏式显示时，当鼠标离开 label 时 }
procedure TfrmPBox.OnSubModuleMouseLeave(Sender: TObject);
begin
  TLabel(Sender).Font.Color := RGB(51, 153, 255);
  TLabel(Sender).Font.Style := TLabel(Sender).Font.Style - [fsUnderline];
end;

{ 创建子模块 DLL 模块 }
procedure TfrmPBox.OnSubModuleListClick(Sender: TObject);
var
  intTag: Integer;
  I, J  : Integer;
  mmItem: TMenuItem;
begin
  mmItem := nil;
  intTag := TLabel(Sender).Tag;
  for I  := 0 to mmMainMenu.Items.Count - 1 do
  begin
    for J := 0 to mmMainMenu.Items.Items[I].Count - 1 do
    begin
      if mmMainMenu.Items.Items[I].Items[J].Tag = intTag then
      begin
        mmItem := mmMainMenu.Items.Items[I].Items[J];
        Break;
      end;
    end;
  end;

  if mmItem <> nil then
    OnMenuItemClick(mmItem);
end;

{ 销毁分栏式界面 }
procedure TfrmPBox.FreeListViewSubModule;
var
  I: Integer;
begin
  for I := tsList.ComponentCount - 1 downto 0 do
  begin
    if tsList.Components[I] is TLabel then
    begin
      TLabel(tsList.Components[I]).Free;
    end
    else if tsList.Components[I] is TImage then
    begin
      if TImage(tsList.Components[I]).Name = '' then
      begin
        TImage(tsList.Components[I]).Free;
      end;
    end;
  end;
end;

{ 窗体大小改变时，重新创建分栏式的界面 }
procedure TfrmPBox.ResizeListViewSubModule;
begin
  if FUIShowStyle = ssList then
  begin
    { 重新创建分栏式显示界面 }
    ChangeUI_List(False);
  end;
end;

{ 获取垂直位置间隔 }
function TfrmPBox.GetMaxInstance(const intCurrentIndex, intCount: Integer): Integer;
var
  intMax: Integer;
  arrInt: array of Integer;
  I     : Integer;
begin
  SetLength(arrInt, mmMainMenu.Items.Count);
  for I := 0 to mmMainMenu.Items.Count - 1 do
  begin
    arrInt[I] := mmMainMenu.Items.Items[I].Count;
  end;
  intMax := MaxIntValue(arrInt);
  if intMax mod 3 = 0 then
    Result := 35 * (0 + intMax div 3)
  else
    Result := 35 * (1 + intMax div 3);
end;

{ 分栏式风格 }
procedure TfrmPBox.ChangeUI_List(const bActivePage: Boolean = True);
var
  I                     : Integer;
  arrParentModuleLabel  : array of TLabel;
  arrParentModuleImage  : array of TImage;
  arrSubModuleLabel     : array of array of TLabel;
  intRow                : Integer;
  strPModuleIconFileName: string;
  strPModuleIconFilePath: string;
  J                     : Integer;
begin
  intRow := IfThen(MaxForm or FullForm, 5, 3);
  if FintBakRow <> intRow then
  begin
    { 销毁分栏式界面 }
    FreeListViewSubModule;
    FintBakRow := intRow;
  end
  else
  begin
    Exit;
  end;

  clbrPModule.Visible := False;
  if bActivePage then
    rzpgcntrlAll.ActivePage := tsList;
  SetLength(arrParentModuleLabel, mmMainMenu.Items.Count);
  SetLength(arrParentModuleImage, mmMainMenu.Items.Count);
  SetLength(arrSubModuleLabel, mmMainMenu.Items.Count);
  for I := 0 to mmMainMenu.Items.Count - 1 do
  begin
    SetLength(arrSubModuleLabel[I], mmMainMenu.Items[I].Count);
  end;

  for I := 0 to mmMainMenu.Items.Count - 1 do
  begin
    { 创建父模块文本 }
    arrParentModuleLabel[I]            := TLabel.Create(tsList);
    arrParentModuleLabel[I].Parent     := tsList;
    arrParentModuleLabel[I].Caption    := mmMainMenu.Items[I].Caption;
    arrParentModuleLabel[I].Font.Name  := '宋体';
    arrParentModuleLabel[I].Font.Size  := 16;
    arrParentModuleLabel[I].Font.Style := [fsBold];
    arrParentModuleLabel[I].Font.Color := RGB(0, 174, 29);
    arrParentModuleLabel[I].Left       := 40 + 400 * (I mod intRow);
    arrParentModuleLabel[I].Top        := 10 + GetMaxInstance(I, intRow) * (I div intRow);

    { 创建父模块图标 }
    arrParentModuleImage[I]         := TImage.Create(tsList);
    arrParentModuleImage[I].Parent  := tsList;
    arrParentModuleImage[I].Height  := 32;
    arrParentModuleImage[I].Width   := 32;
    arrParentModuleImage[I].Stretch := True;
    arrParentModuleImage[I].Left    := arrParentModuleLabel[I].Left - 40;
    arrParentModuleImage[I].Top     := arrParentModuleLabel[I].Top - 5;
    with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
    begin
      strPModuleIconFilePath := ReadString(c_strIniModuleSection, arrParentModuleLabel[I].Caption + '_ICON', '');
      strPModuleIconFileName := ExtractFilePath(ParamStr(0)) + 'plugins\icon\' + strPModuleIconFilePath;
      if FileExists(strPModuleIconFileName) then
        arrParentModuleImage[I].Picture.LoadFromFile(strPModuleIconFileName);
      Free;
    end;

    { 创建子模块文本 }
    for J := 0 to Length(arrSubModuleLabel[I]) - 1 do
    begin
      arrSubModuleLabel[I, J]            := TLabel.Create(tsList);
      arrSubModuleLabel[I, J].Parent     := tsList;
      arrSubModuleLabel[I, J].Caption    := mmMainMenu.Items[I].Items[J].Caption;
      arrSubModuleLabel[I, J].Font.Name  := '宋体';
      arrSubModuleLabel[I, J].Font.Size  := 12;
      arrSubModuleLabel[I, J].Font.Style := [fsBold];
      arrSubModuleLabel[I, J].Font.Color := RGB(51, 153, 255);
      arrSubModuleLabel[I, J].Cursor     := crHandPoint;
      if J mod 3 = 0 then
        arrSubModuleLabel[I, J].Left := arrParentModuleLabel[I].Left + 2
      else
        arrSubModuleLabel[I, J].Left       := arrSubModuleLabel[I, J - 1].Left + arrSubModuleLabel[I, J - 1].Width + 10;
      arrSubModuleLabel[I, J].Top          := arrParentModuleLabel[I].Top + 24 + 20 * (J div 3);
      arrSubModuleLabel[I, J].Tag          := mmMainMenu.Items[I].Items[J].Tag;
      arrSubModuleLabel[I, J].OnMouseEnter := OnSubModuleMouseEnter;
      arrSubModuleLabel[I, J].OnMouseLeave := OnSubModuleMouseLeave;
      arrSubModuleLabel[I, J].OnClick      := OnSubModuleListClick;
    end;
  end;
end;

end.
