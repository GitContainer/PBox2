unit db.uCreateVCDialogDll;
{
  ���� VC Dialog Dll ����
}

interface

uses Winapi.Windows, Winapi.Messages, System.SysUtils, Vcl.Forms, Vcl.StdCtrls, Vcl.ComCtrls, db.uCommon, HookUtils;

{ ���� VC Dialog Dll ���� }
procedure PBoxRun_VC_DLGDll(Page: TPageControl; const TabDllForm: TTabSheet; lblInfo: TLabel; const uiShowStyle: TShowStyle);

{ ���� VC Dialog Dll ������Ϣ }
procedure FreeVCDialogDllForm;

implementation

var
  FPage                    : TPageControl;
  FTabDllForm              : TTabSheet;
  FUIShowStyle             : TShowStyle;
  FstrVCDialogDllClassName : String  = '';
  FstrVCDialogDllWindowName: String  = '';
  FOldWndProc              : Pointer = nil;
  FOld_CreateWindowExW     : function(dwExStyle: DWORD; lpClassName: LPCWSTR; lpWindowName: LPCWSTR; dwStyle: DWORD; X, Y, nWidth, nHeight: Integer; hWndParent: hWnd; hMenu: hMenu; hins: HINST; lpp: Pointer): hWnd; stdcall;

  { ��� dll �У��� Dll �����ȡ���㣬�������ɷǼ���״̬ }
function NewDllFormProc(hWnd: THandle; msg: UINT; wParam: Cardinal; lParam: Cardinal): Integer; stdcall;
begin
  { ����Ӵ����ȡ����ʱ������������ }
  if msg = WM_ACTIVATE then
  begin
    if Application.MainForm <> nil then
    begin
      SendMessage(Application.MainForm.Handle, WM_NCACTIVATE, Integer(True), 0);
    end;
  end;

  { ��ֹ�����ƶ� }
  if msg = WM_SYSCOMMAND then
  begin
    if wParam = SC_MOVE + 2 then
    begin
      wParam := 0;
    end;
  end;

  { ����ԭ���Ļص����� }
  Result := CallWindowProc(FOldWndProc, hWnd, msg, wParam, lParam);
end;

function _CreateWindowExW(dwExStyle: DWORD; lpClassName: LPCWSTR; lpWindowName: LPCWSTR; dwStyle: DWORD; X, Y, nWidth, nHeight: Integer; hWndParent: hWnd; hMenu: hMenu; hins: HINST; lpp: Pointer): hWnd; stdcall;
begin
  { ��ָ���� VC ���� }
  if (lpClassName <> nil) and (lpWindowName <> nil) and (CompareText(lpClassName, FstrVCDialogDllClassName) = 0) and (CompareText(lpWindowName, FstrVCDialogDllWindowName) = 0) then
  begin
    { ���� VC Dlll ���� }
    FPage.ActivePageIndex     := 2;
    Result                     := FOld_CreateWindowExW($00010101, lpClassName, lpWindowName, $96C80000, 0, 0, 0, 0, hWndParent, hMenu, hins, lpp);
    g_intVCDialogDllFormHandle := Result;                                                                                     // ������ VC Dll ������
    Winapi.Windows.SetParent(Result, FTabDllForm.Handle);                                                                    // ���ø�����Ϊ TabSheet
    RemoveMenu(GetSystemMenu(Result, False), 0, MF_BYPOSITION);                                                               // ɾ���ƶ��˵�
    RemoveMenu(GetSystemMenu(Result, False), 0, MF_BYPOSITION);                                                               // ɾ���ƶ��˵�
    RemoveMenu(GetSystemMenu(Result, False), 0, MF_BYPOSITION);                                                               // ɾ���ƶ��˵�
    RemoveMenu(GetSystemMenu(Result, False), 0, MF_BYPOSITION);                                                               // ɾ���ƶ��˵�
    RemoveMenu(GetSystemMenu(Result, False), 0, MF_BYPOSITION);                                                               // ɾ���ƶ��˵�
    RemoveMenu(GetSystemMenu(Result, False), 0, MF_BYPOSITION);                                                               // ɾ���ƶ��˵�
    SetWindowPos(Result, FTabDllForm.Handle, 0, 0, FTabDllForm.Width, FTabDllForm.Height, SWP_NOZORDER OR SWP_NOACTIVATE); // ��� Dll �Ӵ���
    FOldWndProc := Pointer(GetWindowlong(Result, GWL_WNDPROC));                                                              // ��� DLL �����ȡ����ʱ�������嶪ʧ���������
    SetWindowLong(Result, GWL_WNDPROC, LongInt(@NewDllFormProc));                                                             // ���� DLL ������Ϣ
    PostMessage(Application.MainForm.Handle, WM_NCACTIVATE, 1, 0);                                                            // ����������
    UnHook(@FOld_CreateWindowExW);                                                                                           // UNHOOK
    FOld_CreateWindowExW := nil;                                                                                             // UNHOOK
  end
  else
  begin
    Result := FOld_CreateWindowExW(dwExStyle, lpClassName, lpWindowName, dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hins, lpp);
  end;
end;

{ ���� VC Dialog Dll ���� }
procedure PBoxRun_VC_DLGDll(Page: TPageControl; const TabDllForm: TTabSheet; lblInfo: TLabel; const uiShowStyle: TShowStyle);
var
  hDll                             : HMODULE;
  ShowDllForm                      : Tdb_ShowDllForm_Plugins;
  frm                              : TFormClass;
  ft                               : TSPFileType;
  strParamModuleName, strModuleName: PAnsiChar;
  strClassName, strWindowName      : PAnsiChar;
  strIconFileName                  : PAnsiChar;
  strBakDllFileName                : String;
begin
  FPage        := Page;
  FTabDllForm  := TabDllForm;
  FUIShowStyle := uiShowStyle;

  { ��ȡ���� }
  hDll := LoadLibrary(PChar(g_strCreateDllFileName));
  try
    ShowDllForm := GetProcAddress(hDll, c_strDllExportName);
    ShowDllForm(frm, ft, strParamModuleName, strModuleName, strClassName, strWindowName, strIconFileName, False);
    FstrVCDialogDllClassName  := string(strClassName);
    FstrVCDialogDllWindowName := string(strWindowName);
    @FOld_CreateWindowExW     := HookProcInModule(user32, 'CreateWindowExW', @_CreateWindowExW);
  finally
    FreeLibrary(hDll);
  end;

  { ���� Dll ���� }
  hDll := LoadLibrary(PChar(g_strCreateDllFileName));
  try
    ShowDllForm := GetProcAddress(hDll, c_strDllExportName);
    if not Assigned(ShowDllForm) then
    begin
      MessageBox(Application.MainForm.Handle, PChar(Format('���� %s �ĵ������� %s ���������ļ��Ƿ���ڻ��߱�ռ��', [g_strCreateDllFileName, c_strDllExportName])), c_strMsgTitle, MB_OK or MB_ICONERROR);
      Exit;
    end;

    strBakDllFileName := g_strCreateDllFileName;
    ShowDllForm(frm, ft, strParamModuleName, strModuleName, strClassName, strWindowName, strIconFileName, True);
  finally
    FreeLibrary(hDll);
    g_intVCDialogDllFormHandle := 0;
    FstrVCDialogDllClassName  := '';
    FstrVCDialogDllWindowName := '';

    if g_bExitProgram then
    begin
      Application.MainForm.Close;
    end
    else
    begin
      if CompareText(strBakDllFileName, g_strCreateDllFileName) = 0 then
      begin
        { ������ٵ� Dll��������ǰ���ݵ� Dll����ʾû�� Dll Form ��Ҫ������ }
        g_strCreateDllFileName := '';
        lblInfo.Caption        := '';
        if FUIShowStyle = ssButton then
          FPage.ActivePageIndex := 0
        else if FUIShowStyle = ssList then
          FPage.ActivePageIndex := 1;
      end
      else
      begin
        { ���������ǰ���ݵģ�˵���µ� Dll Form �������� }
        PostMessage(Application.MainForm.Handle, WM_CREATENEWDLLFORM, 0, 0);
      end;
    end;
  end;
end;

{ ���� VC Dialog Dll ������Ϣ }
procedure FreeVCDialogDllForm;
begin
  if g_intVCDialogDllFormHandle = 0 then
    Exit;

  SetWindowLong(g_intVCDialogDllFormHandle, GWL_WNDPROC, LongInt(FOldWndProc));
  PostMessage(g_intVCDialogDllFormHandle, WM_SYSCOMMAND, SC_CLOSE, 0);
end;

end.
