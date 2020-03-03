unit uProcessManager;
{$WARN UNIT_PLATFORM OFF}

interface

uses
  Winapi.Windows, System.SysUtils, Winapi.ShlObj, Winapi.TlHelp32, System.Classes, System.IniFiles, System.Math, System.StrUtils, Winapi.ShellAPI, Winapi.ActiveX,
  Vcl.FileCtrl, Vcl.Controls, Vcl.Forms, Vcl.ComCtrls, Vcl.Menus, Vcl.StdCtrls, uProcessAPI, db.uCommon;

type
  TfrmProcessManager = class(TForm)
    lvProcess: TListView;
    lvModule: TListView;
    pmProcess: TPopupMenu;
    mniOpenProcessPath: TMenuItem;
    mniRenameProcessName: TMenuItem;
    mniDeleteProcessFile: TMenuItem;
    mniDllInsertProcess: TMenuItem;
    mniLine01: TMenuItem;
    mniLoadPE: TMenuItem;
    pmModule: TPopupMenu;
    mniOpenModulePath: TMenuItem;
    mniEjectFromProcess: TMenuItem;
    mniDumpToDiskFile: TMenuItem;
    mniLine02: TMenuItem;
    mniCopySelectedModulePath: TMenuItem;
    mniCopySelectedModuleName: TMenuItem;
    mniCopySelectedModuleMemoryAddress: TMenuItem;
    mniLine03: TMenuItem;
    mniCopyFileTo: TMenuItem;
    mniSaveToFile: TMenuItem;
    mniSelectedLineToSaveFile: TMenuItem;
    mniKillProcess: TMenuItem;
    edtParam: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure lvProcessClick(Sender: TObject);
    procedure lvProcessColumnClick(Sender: TObject; Column: TListColumn);
    procedure mniOpenProcessPathClick(Sender: TObject);
    procedure mniRenameProcessNameClick(Sender: TObject);
    procedure mniDeleteProcessFileClick(Sender: TObject);
    procedure mniDllInsertProcessClick(Sender: TObject);
    procedure mniLoadPEClick(Sender: TObject);
    procedure mniOpenModulePathClick(Sender: TObject);
    procedure mniEjectFromProcessClick(Sender: TObject);
    procedure mniDumpToDiskFileClick(Sender: TObject);
    procedure mniCopySelectedModulePathClick(Sender: TObject);
    procedure mniCopySelectedModuleNameClick(Sender: TObject);
    procedure mniCopySelectedModuleMemoryAddressClick(Sender: TObject);
    procedure mniCopyFileToClick(Sender: TObject);
    procedure mniKillProcessClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    procedure EnumProcess(lv: TListView);
    procedure EnumProcessModules(const intPID: Cardinal; lv: TListView);
    function GetFileVersion(const strExeName: string): String;
  public
    { Public declarations }
  end;

type
  { 支持的文件类型 }
  TSPFileType = (ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE);

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;

implementation

{$R *.dfm}

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;
begin
  frm                     := TfrmProcessManager;
  ft                      := ftDelphiDll;
  strParentModuleName     := '程序员工具';
  strModuleName           := 'PM进程管理器';
  strIconFileName         := '';
  strClassName            := '';
  strWindowName           := '';
  Application.Handle      := GetMainFormApplication.Handle;
  Application.Icon.Handle := GetMainFormApplication.Icon.Handle;
end;

{ 获取进程路径 }
function GetProcessName(dwProcessID: LongInt; bFullPath: Bool): String;
var
  pinfo: PROCESS_INFO;
begin
  GetProcessInfo(dwProcessID, pinfo);
  Result := String(pinfo.ImagePathName);
end;

function GetProcessCommandLine(dwProcessID: DWORD): String;
var
  pinfo: PROCESS_INFO;
begin
  GetProcessInfo(dwProcessID, pinfo);
  Result := String(pinfo.CommandLine);
end;

{ 是否是  X64 }
function GetbX64Process(const intPID: Integer): Boolean;
begin
  Result := Is64BitProcess(intPID);
end;

function GetFileDesc(const strExeName: string): String;
var
  n, Len     : DWORD;
  Buf        : PChar;
  Value      : Pointer;
  szName     : array [0 .. 255] of Char;
  Transstring: string;
  Translation: Cardinal;
begin
  Result := '';

  Len := GetFileVersionInfoSize(PChar(strExeName), n);
  if Len > 0 then
  begin
    Buf := AllocMem(Len);
    if GetFileVersionInfo(PChar(strExeName), n, Len, Buf) then
    begin
      VerQueryValue(Buf, '\\VarFileInfo\\Translation', Value, Len);
      if (Len > 0) then
      begin
        Translation := Cardinal(Value^);
        Transstring := Format('%4.4x%4.4x', [(Translation and $0000FFFF), ((Translation shr 16) and $0000FFFF)]);
      end;
      StrPCopy(szName, 'StringFileInfo\' + Transstring + '\FileDescription');
      if VerQueryValue(Buf, szName, Value, Len) then
        Result := StrPas(PChar(Value));
    end;
    FreeMem(Buf, n);
  end;
end;

function GetFileCompany(const strExeName: string): String;
var
  n, Len     : DWORD;
  Buf        : PChar;
  Value      : Pointer;
  szName     : array [0 .. 255] of Char;
  Transstring: string;
  Translation: Cardinal;
begin
  Result := '';

  Len := GetFileVersionInfoSize(PChar(strExeName), n);
  if Len > 0 then
  begin
    Buf := AllocMem(Len);
    if GetFileVersionInfo(PChar(strExeName), n, Len, Buf) then
    begin
      VerQueryValue(Buf, '\\VarFileInfo\\Translation', Value, Len);
      if (Len > 0) then
      begin
        Translation := Cardinal(Value^);
        Transstring := Format('%4.4x%4.4x', [(Translation and $0000FFFF), ((Translation shr 16) and $0000FFFF)]);
      end;
      StrPCopy(szName, 'StringFileInfo\' + Transstring + '\CompanyName');
      if VerQueryValue(Buf, szName, Value, Len) then
        Result := StrPas(PChar(Value));
    end;
    FreeMem(Buf, n);
  end;
end;

procedure TfrmProcessManager.EnumProcess(lv: TListView);
var
  hSnap         : THANDLE;
  pe32          : TProcessEntry32;
  bFind         : Boolean;
  intCount      : Integer;
  strProcessName: String;
begin
  hSnap := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if hSnap = INVALID_HANDLE_VALUE then
    Exit;

  try
    pe32.dwSize := SizeOf(TProcessEntry32);
    bFind       := Process32First(hSnap, pe32);
    if not bFind then
      Exit;

    intCount := 0;

    lv.Items.BeginUpdate;
    while bFind do
    begin
      Inc(intCount);
      with lv.Items.Add do
      begin
        Caption        := Format('%0.3d', [intCount]);
        strProcessName := GetProcessName(pe32.th32ProcessID, True);
        SubItems.Add(pe32.szExeFile);
        SubItems.Add(Format('%4d', [pe32.th32ProcessID]));
        SubItems.Add(IfThen(GetbX64Process(pe32.th32ProcessID), 'x64', 'x86'));
        SubItems.Add(strProcessName);
        SubItems.Add(GetFileDesc(strProcessName));
        SubItems.Add(GetFileCompany(strProcessName));
      end;
      bFind := Process32Next(hSnap, pe32);
    end;
    lv.Items.EndUpdate;
  finally
    CloseHandle(hSnap);
  end;
end;

function TfrmProcessManager.GetFileVersion(const strExeName: string): String;
var
  n, Len     : DWORD;
  Buf        : PChar;
  Value      : Pointer;
  szName     : array [0 .. 255] of Char;
  Transstring: string;
  Translation: Cardinal;
begin
  Result := '';

  Len := GetFileVersionInfoSize(PChar(strExeName), n);
  if Len > 0 then
  begin
    Buf := AllocMem(Len);
    if GetFileVersionInfo(PChar(strExeName), n, Len, Buf) then
    begin
      VerQueryValue(Buf, '\\VarFileInfo\\Translation', Value, Len);
      if (Len > 0) then
      begin
        Translation := Cardinal(Value^);
        Transstring := Format('%4.4x%4.4x', [(Translation and $0000FFFF), ((Translation shr 16) and $0000FFFF)]);
      end;
      StrPCopy(szName, 'StringFileInfo\' + Transstring + '\ProductVersion');
      if VerQueryValue(Buf, szName, Value, Len) then
        Result := StrPas(PChar(Value));
    end;
    FreeMem(Buf, n);
  end;
end;

procedure TfrmProcessManager.EnumProcessModules(const intPID: Cardinal; lv: TListView);
var
  intCount: Integer;
  pinfo   : PROCESS_INFO;
  I       : Integer;
begin
  lv.Clear;

  if intPID = 0 then
    Exit;

  GetProcessInfo(intPID, pinfo);
  intCount := pinfo.ModulesList.Length;
  for I    := 1 to intCount - 1 do
  begin
    with lv.Items.Add do
    begin
      Caption := Format('%0.3d', [I]);
      SubItems.Add(string(pinfo.ModulesList.Modules[I].ModuleName));
      SubItems.Add(string(pinfo.ModulesList.Modules[I].FullPath));
      SubItems.Add(Format('$%0.16x', [pinfo.ModulesList.Modules[I].BaseAddress]));
      SubItems.Add(Format('$%0.16x', [pinfo.ModulesList.Modules[I].EntryAddress]));
      SubItems.Add(Format('$%0.8x', [pinfo.ModulesList.Modules[I].SizeOfImage]));
      SubItems.Add(GetFileVersion(string(pinfo.ModulesList.Modules[I].FullPath)));
      SubItems.Add(GetFileCompany(string(pinfo.ModulesList.Modules[I].FullPath)));
    end;
  end;
end;

procedure TfrmProcessManager.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmProcessManager.FormCreate(Sender: TObject);
begin
  EnumProcess(lvProcess);
end;

procedure TfrmProcessManager.FormResize(Sender: TObject);
begin
  lvProcess.Column[4].Width := Width - 830;
  lvModule.Column[2].Width  := Width - 885;
end;

procedure TfrmProcessManager.lvProcessClick(Sender: TObject);
var
  intPID: Cardinal;
begin
  if lvProcess.Selected = nil then
    Exit;

  intPID := System.SysUtils.StrToInt(lvProcess.Selected.SubItems[1]);
  EnumProcessModules(intPID, lvModule);

  edtParam.Text := GetProcessCommandLine(intPID);
end;

var
  m_bSort: Boolean = False;

function CustomSortProc(Item1, Item2: TListItem; ParamSort: Integer): Integer; stdcall;
var
  txt1, txt2: string;
  intTemp1  : Integer;
  intTemp2  : Integer;
begin
  Result := 0;

  if ParamSort <> 0 then
  begin
    try
      txt1 := Item1.SubItems.Strings[ParamSort - 1];
      txt2 := Item2.SubItems.Strings[ParamSort - 1];
      if TryStrToInt(txt1, intTemp1) and TryStrToInt(txt2, intTemp2) then
      begin
        Result := IfThen(intTemp1 > intTemp2, 1 - Integer(m_bSort), Integer(m_bSort));
      end
      else
      begin
        Result := IfThen(m_bSort, 1, -1) * CompareText(txt1, txt2);
      end;
    except
    end;
  end
  else
  begin
    Result := IfThen(m_bSort, 1, -1) * CompareText(Item1.Caption, Item2.Caption);
  end;
end;

procedure TfrmProcessManager.lvProcessColumnClick(Sender: TObject; Column: TListColumn);
begin
  TListView(Sender).CustomSort(@CustomSortProc, Column.Index);
  m_bSort := not m_bSort;
end;

function SHOpenFolderAndSelectItems(pidlFolder: pItemIDList; cidl: Cardinal; apidl: Pointer; dwFlags: DWORD): HRESULT; stdcall; external shell32;

function OpenFolderAndSelectFile(const strFileName: string; const bEditMode: Boolean = False): Boolean;
var
  IIDL      : pItemIDList;
  pShellLink: IShellLink;
  hr        : Integer;
begin
  Result := False;

  hr := CoCreateInstance(CLSID_ShellLink, nil, CLSCTX_INPROC_SERVER, IID_IShellLink, &pShellLink);
  if hr = S_OK then
  begin
    pShellLink.SetPath(PChar(strFileName));
    pShellLink.GetIDList(&IIDL);
    Result := SHOpenFolderAndSelectItems(IIDL, 0, nil, Cardinal(bEditMode)) = S_OK;
  end;
end;

function FindMainFormHandle: THANDLE;
var
  strTitle      : string;
  strBuffer     : array [0 .. 255] of Char;
  strIniFileName: String;
begin
  GetModuleFileName(0, strBuffer, 256);
  strIniFileName := strBuffer;
  strIniFileName := ChangeFileExt(strIniFileName, '.ini');
  with TIniFile.Create(string(strIniFileName)) do
  begin
    strTitle := ReadString('UI', 'Title', '程序员工具箱 v2.0');
    Free;
  end;
  Result := FindWindow('TfrmMain', PChar(strTitle));
end;

function GetInstanceFromhWnd(const hWnd: Cardinal): TWinControl;
type
  PObjectInstance = ^TObjectInstance;

  TObjectInstance = packed record
    Code: Byte;            { 短跳转 $E8 }
    Offset: Integer;       { CalcJmpOffset(Instance, @Block^.Code); }
    Next: PObjectInstance; { MainWndProc 地址 }
    Self: Pointer;         { 控件对象地址 }
  end;
var
  wc: PObjectInstance;
begin
  Result := nil;
  wc     := Pointer(GetWindowLong(hWnd, GWL_WNDPROC));
  if wc <> nil then
  begin
    Result := wc.Self;
  end;
end;

procedure TfrmProcessManager.mniCopyFileToClick(Sender: TObject);
var
  III        : Integer;
  strFileName: String;
  strSavePath: String;
begin
  if lvModule.SelCount = 0 then
    Exit;

  if not SelectDirectory('选择一个文件夹：', '', strSavePath, [], GetInstanceFromhWnd(FindMainFormHandle)) then
    Exit;

  for III := 0 to lvModule.Items.count - 1 do
  begin
    if lvModule.Items.Item[III].Selected then
    begin
      strFileName := lvModule.Items.Item[III].SubItems[1];
      CopyFile(PChar(strFileName), PChar(strSavePath + '\' + ExtractFileName(strFileName)), True);
    end;
  end;
end;

procedure TfrmProcessManager.mniCopySelectedModuleMemoryAddressClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniCopySelectedModuleNameClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniCopySelectedModulePathClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniDeleteProcessFileClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniDllInsertProcessClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniDumpToDiskFileClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniEjectFromProcessClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniKillProcessClick(Sender: TObject);
var
  hProcess: THANDLE;
  PID     : Cardinal;
begin
  if lvProcess.Selected = nil then
    Exit;

  if Trim(lvProcess.Selected.SubItems[1]) = '' then
    Exit;

  PID      := System.SysUtils.StrToInt(lvProcess.Selected.SubItems[1]);
  hProcess := OpenProcess(PROCESS_TERMINATE, False, PID);
  if TerminateProcess(hProcess, 0) then
  begin
    lvProcess.DeleteSelected;
  end;
end;

procedure TfrmProcessManager.mniLoadPEClick(Sender: TObject);
begin
  //
end;

procedure TfrmProcessManager.mniOpenModulePathClick(Sender: TObject);
begin
  if lvModule.Selected = nil then
    Exit;

  if Trim(lvModule.Selected.SubItems[1]) = '' then
    Exit;

  OpenFolderAndSelectFile(lvModule.Selected.SubItems[1]);
end;

procedure TfrmProcessManager.mniOpenProcessPathClick(Sender: TObject);
begin
  if lvProcess.Selected = nil then
    Exit;

  if Trim(lvProcess.Selected.SubItems[3]) = '' then
    Exit;

  OpenFolderAndSelectFile(lvProcess.Selected.SubItems[3]);
end;

procedure TfrmProcessManager.mniRenameProcessNameClick(Sender: TObject);
begin
  if lvProcess.Selected = nil then
    Exit;

  if Trim(lvProcess.Selected.SubItems[3]) = '' then
    Exit;

  OpenFolderAndSelectFile(lvProcess.Selected.SubItems[3], True);
end;

end.
