unit db.uCommon;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI, Winapi.IpRtrMib, Winapi.ImageHlp, System.SysUtils, System.StrUtils, System.Classes, System.IniFiles, Vcl.Forms, Vcl.Graphics, Vcl.Controls, Data.Win.ADODB, IdIPWatch,
  db.uNetworkManager, FlyUtils.CnXXX.Common, FlyUtils.AES;

type
  TShowMethod = procedure(const str: string) of object;

  { 界面显示方式：菜单、按钮对话框、列表 }
  TShowStyle = (ssMenu, ssButton, ssList);

  { 界面视图方式：单文档、多文档 }
  TViewStyle = (vsSingle, vsMulti);

  { 支持的文件类型 }
  TSPFileType = (ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE);

  {
    Dll 的标准输出函数
    函数名称：db_ShowDllForm_Plugins
    参数说明:
    frm              : 窗体类名    ；TSPFileType 是 ftDelphiDll 时，才可用；Delphi专用，Delphi Dll 主窗体类名；其它语言置空，NULL；
    ft               : 文件类型    ；TSPFileType
    strPModuleName   : 父模块名称  ；
    strSModuleName   : 子模块名称  ；
    strFormClassName : 窗体类名    ；TSPFileType 是 ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE 才会用到， Delphi Dll 时，置空；
    strFormTitleName : 窗体标题名  ；TSPFileType 是 ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE 才会用到， Delphi Dll 时，置空；
    strIconFileName  : 图标文件    ；如果没有指定，从DLL/EXE获取图标；
    bShowForm        : 是否显示窗体；扫描 DLL 时，不用显示 DLL 窗体，只是获取参数；执行时，才显示 DLL 窗体；
  }
  Tdb_ShowDllForm_Plugins = procedure(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;

const
  c_strTitle                                  = '基于 DLL 的模块化开发平台 v2.0';
  c_strUIStyle: array [0 .. 2] of ShortString = ('菜单式', '按钮式', '分栏式');
  c_strIniUISection                           = 'UI';
  c_strIniFormStyleSection                    = 'FormStyle';
  c_strIniModuleSection                       = 'Module';
  c_strIniDBSection                           = 'DB';
  c_strMsgTitle: PChar                        = '系统提示：';
  c_strAESKey                                 = 'dbyoung@sina.com';
  c_strDllExportName                          = 'db_ShowDllForm_Plugins';
  WM_DESTORYPREDLLFORM                        = WM_USER + 1000;
  WM_CREATENEWDLLFORM                         = WM_USER + 1001;

  { 只允许运行一个实例 }
procedure OnlyOneRunInstance;

{ 提升权限 }
function EnableDebugPrivilege(PrivName: string; CanDebug: Boolean): Boolean;

{ 获取本机IP }
function GetNativeIP: String;

{ 排序模块 }
procedure SortModuleList(var lstModuleList: THashedStringList);

{ 搜索插件目录下面的 Dll 文件，添加到列表 }
procedure SearchPlugInsDllFile(var lstDll: TStringList);

function GetShowStyle: TShowStyle;

{ 执行 Sql 脚本，执行成功，是否删除文件 }
function ExeSql(const strFileName: string; ADOCNN: TADOConnection; const bDeleteFileOnSuccess: Boolean = False): Boolean;

{ 升级数据库---执行脚本 }
function UpdateDataBaseScript(const iniFile: TIniFile; const ADOCNN: TADOConnection; const bDeleteFile: Boolean = False): Boolean;

{ 获取数据库库名 }
function GetDBLibraryName(const strLinkDB: string): String;

{ 获取本机当前登录用户名称 }
function GetCurrentLoginUserName: String;

{ 备份数据库，支持远程备份 }
function BackupDataBase(ADOCNN: TADOConnection; const strNativePCLoginName, strNativePCLoginPassword: String; const strSaveFileName: String): Boolean;

{ 恢复数据库，支持远程恢复 }
function RestoreDataBase(ADOCNN: TADOConnection; const strNativePCLoginName, strNativePCLoginPassword: String; const strDBFileName: String; var strErr: String): Boolean;

{ 加密字符串 }
function EncryptString(const strTemp, strKey: string): String;

{ 解密字符串 }
function DecryptString(const strTemp, strKey: string): String;

{ 建立数据库连接 }
function TryLinkDataBase(const strLinkDB: string; var ADOCNN: TADOConnection): Boolean;

{ 从 .msc 文件中获取图标 }
procedure LoadIconFromMSCFile(const strMSCFileName: string; var IcoMSC: TIcon);

{ 从 Dll 中获取导出函数列表 }
function GetPEExport(const strDllFieName: String; var lstFunc: TStringList): Boolean;

{ 数据库密码加密 }
function EncDatabasePassword(const strPassword: string): String;

{ 获取网络下载上传速度 }
procedure GetWebSpeed(var strDnSpeed, strUpSpeed: string);

{ 实时获取 DOS 命令行输出 }
function RunDOS_Real(const cmd: string; CallBackShowRealMessage: TShowMethod = nil): string;

{ 获取 DOS 命令行输出 }
function RunDOS_Result(const CommandLine: string; const bUTF8Output: Boolean = False): string;

{ 获取 PBox 主窗体句柄 <非 DLL 窗体> }
function GetMainFormHandle: hWnd;

{ 根据窗体句柄获取窗体实例 }
function GetInstanceFromhWnd(const hWnd: Cardinal): TWinControl;

{ 获取导出函数全名 C++ 类型函数 }
function GetFullFuncNameCpp(const strFuncName: string): String;

{ 获取 Delphi 窗体的 TApplication }
function GetMainFormApplication: TApplication;

{ 在 DLL 中获取配置文件名称 }
function GetConfigFileName: PAnsiChar;

{ 延时函数 }
procedure DelayTime(const intTime: Integer);

var
  g_intVCDialogDllFormHandle: THandle = 0;
  g_strCreateDllFileName    : string  = '';
  g_bExitProgram            : Boolean = False;
  g_hEXEProcessID           : DWORD   = 0;
  g_strCurrentLoginName     : string  = '';

implementation

{ 只允许运行一个实例 }
procedure OnlyOneRunInstance;
var
  hMainForm       : THandle;
  strTitle        : String;
  bOnlyOneInstance: Boolean;
begin
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
  begin
    strTitle         := ReadString(c_strIniUISection, 'Title', c_strTitle);
    bOnlyOneInstance := ReadBool(c_strIniUISection, 'OnlyOneInstance', False);
    Free;
  end;

  if not bOnlyOneInstance then
    Exit;

  hMainForm := FindWindow('TfrmPBox', PChar(strTitle));
  if hMainForm <> 0 then
  begin
    MessageBox(0, '程序已经运行，无需重复运行', '系统提示：', MB_OK OR MB_ICONERROR);
    if IsIconic(hMainForm) then
      PostMessage(hMainForm, WM_SYSCOMMAND, SC_RESTORE, 0);
    BringWindowToTop(hMainForm);
    SetForegroundWindow(hMainForm);
    Halt;
    Exit;
  end;
end;

{ 提升权限 }
function EnableDebugPrivilege(PrivName: string; CanDebug: Boolean): Boolean;
var
  TP    : Winapi.Windows.TOKEN_PRIVILEGES;
  Dummy : Cardinal;
  hToken: THandle;
begin
  OpenProcessToken(GetCurrentProcess, TOKEN_ADJUST_PRIVILEGES, hToken);
  TP.PrivilegeCount := 1;
  LookupPrivilegeValue(nil, PChar(PrivName), TP.Privileges[0].Luid);
  if CanDebug then
    TP.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED
  else
    TP.Privileges[0].Attributes := 0;
  Result                        := AdjustTokenPrivileges(hToken, False, TP, SizeOf(TP), nil, Dummy);
  hToken                        := 0;
end;

{ 获取本机IP }
function GetNativeIP: String;
var
  IdIPWatch: TIdIPWatch;
begin
  IdIPWatch := TIdIPWatch.Create(nil);
  try
    IdIPWatch.HistoryEnabled := False;
    Result                   := IdIPWatch.LocalIP;
  finally
    IdIPWatch.Free;
  end;
end;

{ 排序父模块 }
procedure SortModuleParent(var lstModuleList: THashedStringList; const strPModuleList: String);
var
  lstTemp           : THashedStringList;
  I, J              : Integer;
  strPModuleName    : String;
  strOrderModuleName: String;
begin
  lstTemp := THashedStringList.Create;
  try
    for J := 0 to Length(strPModuleList.Split([';'])) - 1 do
    begin
      for I := lstModuleList.Count - 1 downto 0 do
      begin
        strPModuleName     := lstModuleList.ValueFromIndex[I].Split([';'])[0];
        strOrderModuleName := strPModuleList.Split([';'])[J];
        if CompareText(strOrderModuleName, strPModuleName) = 0 then
        begin
          lstTemp.Add(lstModuleList.Strings[I]);
          lstModuleList.Delete(I);
        end;
      end;
    end;

    { 有可能会有剩下的；后添加的新模块(父模块)，在未排序之前，是不在排序列表中的 }
    if lstModuleList.Count > 0 then
    begin
      for I := 0 to lstModuleList.Count - 1 do
      begin
        lstTemp.Add(lstModuleList.Strings[I]);
      end;
    end;

    lstModuleList.Clear;
    lstModuleList.Assign(lstTemp);
  finally
    lstTemp.Free;
  end;
end;

{ 交换位置 }
procedure SwapPosHashStringList(var lstModuleList: THashedStringList; const I, J: Integer);
var
  strTemp: String;
begin
  strTemp                  := lstModuleList.Strings[I];
  lstModuleList.Strings[I] := lstModuleList.Strings[J];
  lstModuleList.Strings[J] := strTemp;
end;

{ 查询指定模块的位置 }
function FindSubModuleIndex(const lstModuleList: THashedStringList; const strPModuleName, strSModuleName: String): Integer;
var
  I                  : Integer;
  strParentModuleName: String;
  strSubModuleName   : String;
begin
  Result := -1;
  for I  := 0 to lstModuleList.Count - 1 do
  begin
    strParentModuleName := lstModuleList.ValueFromIndex[I].Split([';'])[0];
    strSubModuleName    := lstModuleList.ValueFromIndex[I].Split([';'])[1];
    if (CompareText(strParentModuleName, strPModuleName) = 0) and (CompareText(strSubModuleName, strSModuleName) = 0) then
    begin
      Result := I;
      Break;
    end;
  end;
end;

{ 查询指定子模块的指定位置的索引号 }
function FindSubModulePos(const lstModuleList: THashedStringList; const strPModuleName: String; const intIndex: Integer): Integer;
var
  I, K               : Integer;
  strParentModuleName: String;
begin
  Result := -1;
  K      := -1;
  for I  := 0 to lstModuleList.Count - 1 do
  begin
    strParentModuleName := lstModuleList.ValueFromIndex[I].Split([';'])[0];
    if CompareText(strParentModuleName, strPModuleName) = 0 then
    begin
      Inc(K);
      if K = intIndex then
      begin
        Result := I;
        Break;
      end;
    end;
  end;
end;

{ 排序子模块 }
procedure SortSubModule_Proc(var lstModuleList: THashedStringList; const strPModuleName: String; const strSModuleOrder: string);
var
  I               : Integer;
  intNewIndex     : Integer;
  intOldIndex     : Integer;
  strSubModuleName: String;
begin
  for I := 0 to Length(strSModuleOrder.Split([';'])) - 1 do
  begin
    strSubModuleName := strSModuleOrder.Split([';'])[I];
    intNewIndex      := FindSubModuleIndex(lstModuleList, strPModuleName, strSubModuleName);
    intOldIndex      := FindSubModulePos(lstModuleList, strPModuleName, I);
    if (intNewIndex <> intOldIndex) and (intNewIndex > -1) and (intOldIndex > -1) then
    begin
      SwapPosHashStringList(lstModuleList, intNewIndex, intOldIndex);
    end;
  end;
end;

{ 排序子模块 }
procedure SortSubModule(var lstModuleList: THashedStringList; const strPModuleOrder: String; const iniModule: TIniFile);
var
  I, Count       : Integer;
  strPModuleName : String;
  strSModuleOrder: String;
begin
  Count := Length(strPModuleOrder.Split([';']));
  for I := 0 to Count - 1 do
  begin
    strPModuleName  := strPModuleOrder.Split([';'])[I];
    strSModuleOrder := iniModule.ReadString(c_strIniModuleSection, strPModuleName, '');
    if Trim(strSModuleOrder) <> '' then
    begin
      SortSubModule_Proc(lstModuleList, strPModuleName, strSModuleOrder);
    end;
  end;
end;

{ 排序模块 }
procedure SortModuleList(var lstModuleList: THashedStringList);
var
  strPModuleOrder: String;
  iniModule      : TIniFile;
begin
  iniModule := TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini'));
  try
    { 排序父模块 }
    strPModuleOrder := iniModule.ReadString(c_strIniModuleSection, 'Order', '');
    if Trim(strPModuleOrder) <> '' then
      SortModuleParent(lstModuleList, strPModuleOrder);

    { 排序子模块 }
    SortSubModule(lstModuleList, strPModuleOrder, iniModule);
  finally
    iniModule.Free;
  end;
end;

{ 搜索插件目录下面的 Dll 文件，添加到列表 }
procedure SearchPlugInsDllFile(var lstDll: TStringList);
var
  strPlugInsPath: String;
  sr            : TSearchRec;
  intFind       : Integer;
begin
  strPlugInsPath := ExtractFilePath(ParamStr(0)) + 'plugins';
  intFind        := FindFirst(strPlugInsPath + '\*.dll', faAnyFile, sr);
  if not DirectoryExists(strPlugInsPath) then
    Exit;

  while intFind = 0 do
  begin
    if (sr.Name <> '.') and (sr.Name <> '..') and (sr.Attr <> faDirectory) then
      lstDll.Add(strPlugInsPath + '\' + sr.Name);
    intFind := FindNext(sr);
  end;
end;

function GetShowStyle: TShowStyle;
begin
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
  begin
    Result := TShowStyle(ReadInteger(c_strIniFormStyleSection, 'index', 0) mod 3);
    Free;
  end;
end;

{ 执行 Sql 脚本，执行成功，是否删除文件 }
function ExeSql(const strFileName: string; ADOCNN: TADOConnection; const bDeleteFileOnSuccess: Boolean = False): Boolean;
var
  strTemp: String;
  I      : Integer;
  qry    : TADOQuery;
begin
  Result := False;
  if not FileExists(strFileName) then
    Exit;

  try
    with TStringList.Create do
    begin
      qry := TADOQuery.Create(nil);
      try
        qry.Connection := ADOCNN;
        LoadFromFile(strFileName);
        strTemp := '';
        for I   := 0 to Count - 1 do
        begin
          if SameText(Trim(Strings[I]), 'GO') then
          begin
            qry.Close;
            qry.SQL.Clear;
            qry.SQL.Text := strTemp;
            qry.ExecSQL;
            strTemp := '';
          end
          else
          begin
            strTemp := strTemp + Strings[I] + #13#10;
          end;
        end;

        if strTemp <> '' then
        begin
          qry.Close;
          qry.SQL.Clear;
          qry.SQL.Text := strTemp;
          qry.ExecSQL;
        end;
      finally
        Result := True;
        if bDeleteFileOnSuccess then
          DeleteFile(strFileName);
        qry.Free;
        Free;
      end;
    end;
  except
    Result := False;
  end;
end;

{ 升级数据库---执行脚本 }
function UpdateDataBaseScript(const iniFile: TIniFile; const ADOCNN: TADOConnection; const bDeleteFile: Boolean = False): Boolean;
var
  strSQLFileName: String;
begin
  Result := False;
  if iniFile.ReadBool(c_strIniDBSection, 'AutoUpdate', False) then
  begin
    strSQLFileName := iniFile.ReadString(c_strIniDBSection, 'AutoUpdateFile', '');
    if (Trim(strSQLFileName) <> '') and (FileExists(strSQLFileName)) then
    begin
      Result := ExeSql(ExtractFilePath(ParamStr(0)) + strSQLFileName, ADOCNN, bDeleteFile);
    end;
  end;
end;

{ 获取数据库库名 }
function GetDBLibraryName(const strLinkDB: string): String;
var
  I, J   : Integer;
  strTemp: String;
begin
  Result := '';
  I      := Pos('initial catalog=', LowerCase(strLinkDB));
  if I > 0 then
  begin
    strTemp := RightStr(strLinkDB, Length(strLinkDB) - I - Length('Initial Catalog=') + 1);
    J       := Pos(';', strTemp);
    Result  := LeftStr(strTemp, J - 1);
  end;
end;

{ 获取本机当前登录用户名称 }
function GetCurrentLoginUserName: String;
var
  Buffer: array [0 .. 255] of Char;
  Count : Cardinal;
begin
  Result := '';
  Count  := 256;
  if GetUserName(Buffer, Count) then
  begin
    Result := Buffer;
  end;
end;

{ 获取本机计算机名称 }
function GetNativePCName: string;
var
  chrPCName: array [0 .. 255] of Char;
  intLen   : Cardinal;
begin
  intLen := 256;
  GetComputerName(@chrPCName[0], intLen);
  Result := chrPCName;
end;

{ 备份数据库，支持远程备份 }
function BackupDataBase(ADOCNN: TADOConnection; const strNativePCLoginName, strNativePCLoginPassword: String; const strSaveFileName: String): Boolean;
const
  c_strbackupDataBase =                                                    //
    ' exec master..xp_cmdshell ''net use z: \\%s\c$ "%s" /user:%s\%s'' ' + //
    ' backup database %s to disk = ''z:\temp.bak''' +                      //
    ' exec master..xp_cmdshell ''net use z: /delete''';
var
  strDBLibraryName: string;
  strNativePCName : string;
begin
  Result := False;
  if not ADOCNN.Connected then
    Exit;

  strDBLibraryName := GetDBLibraryName(ADOCNN.ConnectionString);
  strNativePCName  := GetNativePCName;

  with TADOQuery.Create(nil) do
  begin
    Connection := ADOCNN;
    SQL.Text   := Format(c_strbackupDataBase, [strNativePCName, strNativePCLoginPassword, strNativePCName, strNativePCLoginName, strDBLibraryName]);
    try
      DeleteFile('c:\temp.bak');
      ExecSQL;
      { 移动备份文件到指定位置 }
      Result := MoveFile(PChar('c:\temp.bak'), PChar(strSaveFileName));
    except
      Result := False;
    end;
    Free;
  end;
end;

{ 恢复数据库，支持远程恢复 }
function RestoreDataBase(ADOCNN: TADOConnection; const strNativePCLoginName, strNativePCLoginPassword: String; const strDBFileName: String; var strErr: String): Boolean;
const
  c_strbackupDataBase =                                                    //
    ' exec master..xp_cmdshell ''net use z: \\%s\c$ "%s" /user:%s\%s'' ' + //
    ' restore database %s from disk = ''z:\temp.bak''' +                   //
    ' exec master..xp_cmdshell ''net use z: /delete''';
var
  strDBLibraryName: String;
  strNativePCName : String;
begin
  Result := False;

  if not ADOCNN.Connected then
    Exit;

  { 删除临时文件 }
  DeleteFile('c:\temp.bak');
  strDBLibraryName := GetDBLibraryName(ADOCNN.ConnectionString);
  strNativePCName  := GetNativePCName;

  if Trim(strDBLibraryName) = '' then
    strDBLibraryName := 'RestoreTemp';

  { 复制文件到网络邻居目录中 }
  if CopyFile(PChar(strDBFileName), PChar('c:\temp.bak'), True) then
  begin
    with TADOQuery.Create(nil) do
    begin
      Connection := ADOCNN;
      SQL.Text   := Format(c_strbackupDataBase, [strNativePCName, strNativePCLoginPassword, strNativePCName, strNativePCLoginName, strDBLibraryName]);
      try
        ExecSQL;
        Result := True;
      except
        on E: Exception do
        begin
          strErr := E.Message;
          Result := False;
        end;
      end;
      DeleteFile('c:\temp.bak');
      Free;
    end;
  end;
end;

{ 加密字符串 }
function EncryptString(const strTemp, strKey: string): String;
begin
  Result := AESEncryptStrToHex(strTemp, strKey, TEncoding.Unicode, TEncoding.UTF8, TKeyBit.kb256, '1234567890123456', TPaddingMode.pmPKCS5or7RandomPadding, True, rlCRLF, rlCRLF, nil);
end;

{ 解密字符串 }
function DecryptString(const strTemp, strKey: string): String;
begin
  Result := AESDecryptStrFromHex(strTemp, strKey, TEncoding.Unicode, TEncoding.UTF8, TKeyBit.kb256, '1234567890123456', TPaddingMode.pmPKCS5or7RandomPadding, True, rlCRLF, rlCRLF, nil);
end;

{ 建立数据库连接 }
function TryLinkDataBase(const strLinkDB: string; var ADOCNN: TADOConnection): Boolean;
begin
  Result := False;

  if strLinkDB = '' then
    Exit;

  if not Assigned(ADOCNN) then
    Exit;

  ADOCNN.KeepConnection := True;
  ADOCNN.LoginPrompt    := False;
  ADOCNN.Connected      := False;

  if Pos('.udl', LowerCase(strLinkDB)) > 0 then
  begin
    ADOCNN.ConnectionString := 'FILE NAME=' + strLinkDB;
    ADOCNN.Provider         := strLinkDB;
  end
  else
  begin
    ADOCNN.ConnectionString := strLinkDB;
  end;

  try
    ADOCNN.Connected := True;
    Result           := True;
  except
    Result := False;
  end;
end;

function GetSystemPath: String;
var
  Buffer: array [0 .. 255] of Char;
begin
  GetSystemDirectory(Buffer, 256);
  Result := Buffer;
end;

{ 从 .msc 文件中获取图标 }
procedure LoadIconFromMSCFile(const strMSCFileName: string; var IcoMSC: TIcon);
var
  strLine       : String;
  strSystemPath : String;
  I, J, intIndex: Integer;
  intIconIndex  : Integer;
  strDllFileName: String;
  strTemp       : String;
begin
  with TStringList.Create do
  begin
    intIndex      := -1;
    strSystemPath := GetSystemPath;
    LoadFromFile(strSystemPath + '\' + strMSCFileName, TEncoding.ASCII);

    for I := 0 to Count - 1 do
    begin
      strLine := Strings[I];
      if strLine <> '' then
      begin
        if Pos('<Icon Index="', strLine) > 0 then
        begin
          intIndex := I;
          Break;
        end;
      end;
    end;

    if intIndex <> -1 then
    begin
      strLine      := Strings[intIndex];
      I            := Pos('"', strLine);
      strTemp      := RightStr(strLine, Length(strLine) - I);
      J            := Pos('"', strTemp);
      intIconIndex := StrToIntDef(MidStr(strLine, I + 1, J - 1), 0);

      I              := Pos('File="', strLine);
      strTemp        := RightStr(strLine, Length(strLine) - I - 5);
      J              := Pos('"', strTemp);
      strDllFileName := MidStr(strTemp, 1, J - 1);

      IcoMSC.Handle := ExtractIcon(HInstance, PChar(strDllFileName), intIconIndex);
    end;

    Free;
  end;
end;

{ 获取导出函数全名 C++ 类型函数 }
function GetFullFuncNameCpp(const strFuncName: string): String;
var
  strFuncFullName: array [0 .. 255] of AnsiChar;
begin
  UnDecorateSymbolName(PAnsiChar(AnsiString(strFuncName)), strFuncFullName, 256, 0);
  Result := string(strFuncFullName);
end;

{ 获取导出表在磁盘文件中的位置 }
procedure FindExportTablePos(const sts: array of TImageSectionHeader; const intVA: Cardinal; var intRA: Cardinal);
var
  III, Count: Integer;
begin
  intRA := 0;

  Count   := Length(sts);
  for III := 0 to Count - 2 do
  begin
    if (intVA >= sts[III + 0].VirtualAddress) and (intVA < sts[III + 1].VirtualAddress) then
    begin
      intRA := (intVA - sts[III].VirtualAddress) + sts[III].PointerToRawData;
      Break;
    end;
  end;
end;

{ 从 Dll 中获取导出函数列表 }
function GetPEExport(const strDllFieName: String; var lstFunc: TStringList): Boolean;
var
  idh            : TImageDosHeader;
  bX64           : Boolean;
  hPEFile        : Cardinal;
  intVA          : Cardinal;
  intRA          : Cardinal;
  inhX86         : TImageNtHeaders32;
  inhX64         : TImageNtHeaders64;
  stsArr         : array of TImageSectionHeader;
  eft            : TImageExportDirectory;
  I              : Integer;
  intFuncRA      : Cardinal;
  strFunctionName: array [0 .. 255] of AnsiChar;
begin
  Result  := False;
  hPEFile := FileOpen(strDllFieName, fmOpenRead);
  if hPEFile = INVALID_HANDLE_VALUE then
  begin
    MessageBox(Application.MainForm.Handle, '文件无法打开，检查一下文件是否被占用', '系统提示：', MB_OK or MB_ICONERROR);
    Exit;
  end;

  try
    FileRead(hPEFile, idh, SizeOf(idh));
    if idh.e_magic <> IMAGE_DOS_SIGNATURE then
    begin
      MessageBox(Application.MainForm.Handle, '不是PE文件，请检查文件', '系统提示：', MB_OK or MB_ICONERROR);
      Exit;
    end;

    FileSeek(hPEFile, idh._lfanew, 0);
    FileSeek(hPEFile, idh._lfanew, 0);
    FileRead(hPEFile, inhX86, SizeOf(TImageNtHeaders32));
    bX64 := inhX86.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64;

    if bX64 then
    begin
      FileSeek(hPEFile, idh._lfanew, 0);
      FileRead(hPEFile, inhX64, SizeOf(TImageNtHeaders64));
      intVA := inhX64.OptionalHeader.DataDirectory[0].VirtualAddress;
      SetLength(stsArr, inhX64.FileHeader.NumberOfSections);
      FileRead(hPEFile, stsArr[0], inhX64.FileHeader.NumberOfSections * SizeOf(TImageSectionHeader));
    end
    else
    begin
      intVA := inhX86.OptionalHeader.DataDirectory[0].VirtualAddress;
      SetLength(stsArr, inhX86.FileHeader.NumberOfSections);
      FileRead(hPEFile, stsArr[0], inhX86.FileHeader.NumberOfSections * SizeOf(TImageSectionHeader));
    end;

    { 获取导出表在磁盘文件中的位置 }
    FindExportTablePos(stsArr, intVA, intRA);

    { 获取导出函数列表 }
    FileSeek(hPEFile, intRA, 0);
    FileRead(hPEFile, eft, SizeOf(TImageExportDirectory));
    for I := 0 to eft.NumberOfNames - 1 do
    begin
      FileSeek(hPEFile, (eft.AddressOfNames - intVA) + intRA + DWORD(4 * I), 0);
      FileRead(hPEFile, intFuncRA, 4);
      FileSeek(hPEFile, (intFuncRA - intVA) + intRA, 0);
      FileRead(hPEFile, strFunctionName, 256);
      lstFunc.Add(string(strFunctionName));
    end;

    Result := True;
  finally
    FileClose(hPEFile);
  end;
end;

{ 数据库密码加密 }
function EncDatabasePassword(const strPassword: string): String;
var
  strDllFileName: String;
  strDllFuncName: String;
  hDll          : HMODULE;
  EncFunc       : function(const strPassword: PAnsiChar): PAnsiChar; stdcall;
begin
  Result := strPassword;
  with TIniFile.Create(ChangeFileExt(ParamStr(0), '.ini')) do
  begin
    if ReadBool(c_strIniDBSection, 'PasswordEnc', False) then
    begin
      strDllFileName := ReadString(c_strIniDBSection, 'PasswordEncDllFileName', '');
      strDllFuncName := ReadString(c_strIniDBSection, 'PasswordEncDllFuncName', '');
      if FileExists(strDllFileName) then
      begin
        hDll := LoadLibrary(PChar(strDllFileName));
        if hDll <> INVALID_HANDLE_VALUE then
        begin
          try
            EncFunc := GetProcaddress(hDll, PChar(strDllFuncName));
            if Assigned(EncFunc) = True then
            begin
              Result := string(PChar(EncFunc(PAnsiChar(PChar(strPassword)))));
            end;
          finally
            FreeLibrary(hDll);
          end;
        end;
      end;
    end;
    Free;
  end;
end;

{ 获取网络下载上传速度 }
var
  FintDnBytes: UInt64 = 0;
  FintUpBytes: UInt64 = 0;

procedure GetWebSpeed(var strDnSpeed, strUpSpeed: string);
var
  NetworkManager        : TNetworkManager;
  lsNetworkTraffic      : TList;
  I                     : Integer;
  intDnSpeed, intUpSpeed: Cardinal;
  strNetworkDesc        : String;
  intDnBytes            : UInt64;
  intUpBytes            : UInt64;
begin
  NetworkManager   := TNetworkManager.Create(0);
  lsNetworkTraffic := TList.Create;
  try
    NetworkManager.GetNetworkTraffic(lsNetworkTraffic);
    if lsNetworkTraffic.Count > 0 then
    begin
      intDnBytes := 0;
      intUpBytes := 0;

      for I := 0 to lsNetworkTraffic.Count - 1 do
      begin
        strNetworkDesc := NetworkManager.GetDescrString(PMibIfRow(lsNetworkTraffic.Items[I])^.bDescr);
        if Pos('-0000', strNetworkDesc) = 0 then
        begin
          intDnBytes := intDnBytes + PMibIfRow(lsNetworkTraffic.Items[I])^.dwInOctets;
          intUpBytes := intUpBytes + PMibIfRow(lsNetworkTraffic.Items[I])^.dwOutOctets;
        end;
      end;

      { 第一次 }
      if (FintDnBytes = 0) and (FintUpBytes = 0) then
      begin
        strDnSpeed := Format('%0.2f K/S', [0.0]);
        strUpSpeed := Format('%0.2f K/S', [0.0]);

        FintDnBytes := intDnBytes;
        FintUpBytes := intUpBytes;
        Exit;
      end;

      { 下载速度 }
      intDnSpeed := intDnBytes - FintDnBytes;
      if intDnSpeed > 1024 * 1024 then
        strDnSpeed := Format('%0.2f M/S', [intDnSpeed / 1024 / 1024])
      else
        strDnSpeed := Format('%0.2f K/S', [intDnSpeed / 1024]);

      { 上传速度 }
      intUpSpeed := intUpBytes - FintUpBytes;
      if intUpSpeed > 1024 * 1024 then
        strUpSpeed := Format('%0.2f M/S', [intUpSpeed / 1024 / 1024])
      else
        strUpSpeed := Format('%0.2f K/S', [intUpSpeed / 1024]);

      FintDnBytes := intDnBytes;
      FintUpBytes := intUpBytes;
    end;
  finally
    lsNetworkTraffic.Free;
    NetworkManager.Free;
  end;
end;

{ 实时获取 DOS 命令行输出 }
function RunDOS_Real(const cmd: string; CallBackShowRealMessage: TShowMethod = nil): string;
var
  hReadPipe, hWritePipe: THandle;
  si                   : STARTUPINFO;
  lsa                  : SECURITY_ATTRIBUTES;
  pi                   : PROCESS_INFORMATION;
  cchReadBuffer        : DWORD;
  pOutStr              : PAnsiChar;
  res, strCMD          : string;
begin
  strCMD                   := 'cmd.exe /k ' + cmd;
  pOutStr                  := AllocMem(5000);
  lsa.nLength              := SizeOf(SECURITY_ATTRIBUTES);
  lsa.lpSecurityDescriptor := nil;
  lsa.bInheritHandle       := True;
  if not CreatePipe(hReadPipe, hWritePipe, @lsa, 0) then
    Exit;

  FillChar(si, SizeOf(STARTUPINFO), 0);
  si.cb          := SizeOf(STARTUPINFO);
  si.dwFlags     := (STARTF_USESTDHANDLES or STARTF_USESHOWWINDOW);
  si.wShowWindow := SW_HIDE;
  si.hStdOutput  := hWritePipe;

  if not CreateProcess(nil, PChar(strCMD), nil, nil, True, 0, nil, nil, si, pi) then
    Exit;

  while (True) do
  begin
    if not PeekNamedPipe(hReadPipe, pOutStr, 1, @cchReadBuffer, nil, nil) then
      Break;

    if cchReadBuffer <> 0 then
    begin
      if not ReadFile(hReadPipe, pOutStr^, 4096, cchReadBuffer, nil) then
        Break;

      pOutStr[cchReadBuffer] := chr(0);
      if @CallBackShowRealMessage <> nil then
        CallBackShowRealMessage(string(pOutStr));
      res := res + String(pOutStr);
    end
    else if (WaitForSingleObject(pi.hProcess, 0) = WAIT_OBJECT_0) then
      Break;

    Sleep(10);

    Application.ProcessMessages;
  end;
  pOutStr[cchReadBuffer] := chr(0);

  CloseHandle(hReadPipe);
  CloseHandle(pi.hThread);
  CloseHandle(pi.hProcess);
  CloseHandle(hWritePipe);
  FreeMem(pOutStr);
  Result := res;
end;

{ 获取 DOS 命令行输出 }
function RunDOS_Result(const CommandLine: string; const bUTF8Output: Boolean = False): string;
var
  sa                             : TSecurityAttributes;
  si                             : TStartupInfo;
  pi                             : TProcessInformation;
  StdOutPipeRead, StdOutPipeWrite: THandle;
  bWasOK                         : Boolean;
  Buffer                         : array [0 .. 255] of Byte;
  ReBuffer                       : array of Byte;
  BytesRead                      : Cardinal;
  bCreate                        : Boolean;
begin
  Result                  := '';
  sa.nLength              := SizeOf(sa);
  sa.bInheritHandle       := True;
  sa.lpSecurityDescriptor := nil;
  CreatePipe(StdOutPipeRead, StdOutPipeWrite, @sa, 0);
  try
    FillChar(si, SizeOf(si), 0);
    si.cb          := SizeOf(si);
    si.dwFlags     := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
    si.wShowWindow := SW_HIDE;
    si.hStdInput   := GetStdHandle(STD_INPUT_HANDLE);
    si.hStdOutput  := StdOutPipeWrite;
    si.hStdError   := StdOutPipeWrite;
    bCreate        := CreateProcess(nil, PChar('cmd.exe /C ' + CommandLine), nil, nil, True, 0, nil, nil, si, pi);
    CloseHandle(StdOutPipeWrite);

    if bCreate then
    begin
      try
        repeat
          bWasOK := ReadFile(StdOutPipeRead, Buffer, 255, BytesRead, nil);
          if BytesRead > 0 then
          begin
            SetLength(ReBuffer, BytesRead);
            FillChar(ReBuffer[0], BytesRead, #0);
            Move(Buffer[0], ReBuffer[0], BytesRead);
            if bUTF8Output then
              Result := Result + TEncoding.UTF8.GetString(ReBuffer)
            else
              Result := Result + TEncoding.ANSI.GetString(ReBuffer)
          end;
        until not bWasOK or (BytesRead = 0);
        WaitForSingleObject(pi.hProcess, INFINITE);
      finally
        CloseHandle(pi.hThread);
        CloseHandle(pi.hProcess);
      end;
    end;
  finally
    CloseHandle(StdOutPipeRead);
  end;
end;

function _EnumWindowsProc(P_HWND: Cardinal; LParam: Cardinal): Boolean; stdcall;
var
  PID         : DWORD;
  chrClassName: array [0 .. 255] of Char;
  strClassName: String;
begin
  Result := True;

  GetWindowThreadProcessId(P_HWND, @PID);
  if PCardinal(LParam)^ <> PID then
  begin
    Result := True;
  end
  else
  begin
    GetClassName(P_HWND, chrClassName, 256);
    strClassName := chrClassName;
    if (CompareText(strClassName, 'TApplication') <> 0) and (CompareText(strClassName, 'TPUtilWindow') <> 0) and (CompareText(strClassName, 'IME') <> 0) and (CompareText(strClassName, 'MSCTFIME UI') <> 0) then
    begin
      Result                 := False;
      PCardinal(LParam + 4)^ := P_HWND;
    end;
  end;
end;

{ 获取 PBox 主窗体句柄 <非 DLL 窗体> }
function GetMainFormHandle: hWnd;
var
  Buffer: array [0 .. 1] of Cardinal;
begin
  Result    := 0;
  Buffer[0] := GetCurrentProcessId;
  Buffer[1] := 0;
  EnumWindows(@_EnumWindowsProc, Integer(@Buffer));
  if Buffer[1] > 0 then
    Result := Buffer[1];
end;

{ 根据窗体句柄获取窗体实例 }
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

function _EnumApplicationProc(P_HWND: Cardinal; LParam: Cardinal): Boolean; stdcall;
var
  PID         : DWORD;
  chrClassName: array [0 .. 255] of Char;
  strClassName: String;
begin
  Result := True;

  GetWindowThreadProcessId(P_HWND, @PID);
  if PCardinal(LParam)^ <> PID then
  begin
    Result := True;
  end
  else
  begin
    GetClassName(P_HWND, chrClassName, 256);
    strClassName := chrClassName;
    if CompareText(strClassName, 'TApplication') = 0 then
    begin
      Result                 := False;
      PCardinal(LParam + 4)^ := P_HWND;
    end;
  end;
end;

{ 获取 Delphi 窗体的 TApplication }
function GetMainFormApplication: TApplication;
var
  Buffer   : array [0 .. 1] of Cardinal;
  appHandle: THandle;
begin
  Result    := nil;
  appHandle := 0;
  Buffer[0] := GetCurrentProcessId;
  Buffer[1] := 0;
  EnumWindows(@_EnumApplicationProc, Integer(@Buffer));
  if Buffer[1] > 0 then
    appHandle := Buffer[1];

  if appHandle <> 0 then
  begin
    Result := TApplication(GetInstanceFromhWnd(appHandle));
  end;
end;

{ 在 DLL 中获取配置文件名称 }
function GetConfigFileName: PAnsiChar;
var
  strMainPath   : array [0 .. 255] of Char;
  strIniFileName: String;
begin
  GetModuleFileName(0, strMainPath, 256);
  strIniFileName := strMainPath;
  Result         := PAnsiChar(AnsiString(ChangeFileExt(strIniFileName, '.ini')));
end;

{ 延时函数 }
procedure DelayTime(const intTime: Integer);
var
  intET, intST: Integer;
begin
  intST := GetTickCount;
  while True do
  begin
    Application.ProcessMessages;
    intET := GetTickCount;
    if intET - intST >= intTime then
      Break;
  end;
end;

end.
