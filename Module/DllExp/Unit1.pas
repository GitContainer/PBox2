unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.StrUtils, System.Variants, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.WinXCtrls, Vcl.ComCtrls, db.uCommon;

type
  TfrmDllExport = class(TForm)
    lbl1: TLabel;
    srchbxDllFile: TSearchBox;
    lvDllExport: TListView;
    dlgOpenDllFile: TOpenDialog;
    lbl2: TLabel;
    procedure srchbxDllFileInvokeSearch(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { 从 Dll 中获取导出函数列表 }
    function GetPEExport(const strDllFieName: String): Boolean;
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
  frm                     := TfrmDllExport;
  ft                      := ftDelphiDll;
  strParentModuleName     := '程序员工具';
  strModuleName           := 'Dll函数查看器';
  strIconFileName         := '';
  strClassName            := '';
  strWindowName           := '';
  Application.Handle      := GetMainFormApplication.Handle;
  Application.Icon.Handle := GetMainFormApplication.Icon.Handle;
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

procedure TfrmDllExport.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmDllExport.FormResize(Sender: TObject);
begin
  lvDllExport.Column[1].Width := Width - 370;
end;

{ 从 Dll 中获取导出函数列表 }
function TfrmDllExport.GetPEExport(const strDllFieName: String): Boolean;
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
  intFuncAddress : DWORD;
  intFuncOrder   : WORD;
begin
  Result       := False;
  lbl2.Caption := '';
  lvDllExport.Clear;

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
    bX64         := inhX86.FileHeader.Machine = IMAGE_FILE_MACHINE_AMD64;
    lbl2.Caption := Ifthen(bX64, 'x64', 'x86');

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
      { 获取函数名称 }
      FileSeek(hPEFile, (eft.AddressOfNames - intVA) + intRA + DWORD(4 * I), 0);
      FileRead(hPEFile, intFuncRA, 4);
      FileSeek(hPEFile, (intFuncRA - intVA) + intRA, 0);
      FileRead(hPEFile, strFunctionName, 256);

      { 获取函数序列 }
      FileSeek(hPEFile, (eft.AddressOfNameOrdinals - intVA) + intRA + DWORD(2 * I), 0);
      FileRead(hPEFile, intFuncOrder, 1);

      { 获取函数偏移地址 }
      FileSeek(hPEFile, (eft.AddressOfFunctions - intVA) + intRA + DWORD(4 * I), 0);
      FileRead(hPEFile, intFuncAddress, 4);

      with lvDllExport.Items.Add do
      begin
        Caption := Format('%0.4d(%s)', [intFuncOrder + eft.Base, IntToHex(intFuncOrder + eft.Base, 4)]); // 序列
        SubItems.Add(GetFullFuncNameCpp(string(strFunctionName)));                                       // 函数名称
        SubItems.Add('$' + IntToHex(intFuncAddress));                                                    // 函数内存偏移地址
        SubItems.Add('$' + IntToHex((intFuncRA - intVA) + intRA));                                       // 函数文件偏移地址
      end;
    end;

    Result := True;
  finally
    FileClose(hPEFile);
  end;
end;

procedure TfrmDllExport.srchbxDllFileInvokeSearch(Sender: TObject);
begin
  if not dlgOpenDllFile.Execute(GetMainFormHandle) then
    Exit;

  srchbxDllFile.Text := dlgOpenDllFile.FileName;
  GetPEExport(dlgOpenDllFile.FileName);
end;

end.
