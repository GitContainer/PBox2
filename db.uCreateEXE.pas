unit db.uCreateEXE;
{
  创建 EXE 窗体
}

interface

uses Winapi.Windows, Winapi.ShellAPI, System.SysUtils, Vcl.Forms, Vcl.ComCtrls, Vcl.StdCtrls, Winapi.TlHelp32, db.uCommon;

procedure PBoxRun_IMAGE_EXE(const strEXEFileName, strFileValue: String; pg: TPageControl; ts: TTabSheet; lblInfo: TLabel; const UIShowStyle: TShowStyle);

implementation

var
  g_strFileValue              : string;
  g_strEXEFormClassName       : string = '';
  g_strEXEFormTitleName       : string = '';
  g_PageControl               : TPageControl;
  g_Tabsheet                  : TTabSheet;
  g_lblInfo                   : TLabel;
  g_strCreateDllFileNameBackUp: String;
  g_UIShowStyle               : TShowStyle;

  { 进程是否关闭 }
function CheckProcessExist(const intPID: DWORD): Boolean;
var
  hSnap: THandle;
  vPE  : TProcessEntry32;
begin
  Result     := False;
  hSnap      := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  vPE.dwSize := SizeOf(TProcessEntry32);
  if Process32First(hSnap, vPE) then
  begin
    while Process32Next(hSnap, vPE) do
    begin
      if vPE.th32ProcessID = intPID then
      begin
        Result := True;
        Break;
      end;
    end;
  end;
  CloseHandle(hSnap);
end;

{ 进程关闭后，变量复位 }
procedure EndExeForm(hWnd: hWnd; uMsg, idEvent: UINT; dwTime: DWORD); stdcall;
begin
  if not CheckProcessExist(g_hEXEProcessID) then
  begin
    if g_strCreateDllFileNameBackUp <> g_strCreateDllFileName then
    begin
      g_hEXEProcessID := 0;
    end
    else
    begin
      g_lblInfo.Caption            := '';
      g_hEXEProcessID              := 0;
      g_strCreateDllFileName       := '';
      g_strCreateDllFileNameBackUp := '';
    end;

    if g_UIShowStyle = ssButton then
      g_PageControl.ActivePageIndex := 0
    else if g_UIShowStyle = ssList then
      g_PageControl.ActivePageIndex := 1;

    KillTimer(Application.MainForm.Handle, $2000);
  end;
end;

{ 查找 EXE 的主窗体是否成功创建 }
procedure FindExeForm(hWnd: hWnd; uMsg, idEvent: UINT; dwTime: DWORD); stdcall;
var
  hEXEFormHandle: THandle;
begin
  if Trim(g_strEXEFormClassName) = '' then
    hEXEFormHandle := FindWindow(nil, PChar(g_strEXEFormTitleName))
  else
    hEXEFormHandle := FindWindow(PChar(g_strEXEFormClassName), PChar(g_strEXEFormTitleName));

  if hEXEFormHandle <> 0 then
  begin
    GetWindowThreadProcessId(hEXEFormHandle, g_hEXEProcessID);
    SetWindowPos(hEXEFormHandle, g_Tabsheet.Handle, 0, 0, g_Tabsheet.Width, g_Tabsheet.Height, SWP_NOZORDER OR SWP_NOACTIVATE);                            // 最大化 Dll 子窗体
    Winapi.Windows.SetParent(hEXEFormHandle, g_Tabsheet.Handle);                                                                                           // 设置父窗体为 TabSheet
    RemoveMenu(GetSystemMenu(hEXEFormHandle, False), 0, MF_BYPOSITION);                                                                                    // 删除移动菜单
    RemoveMenu(GetSystemMenu(hEXEFormHandle, False), 0, MF_BYPOSITION);                                                                                    // 删除移动菜单
    RemoveMenu(GetSystemMenu(hEXEFormHandle, False), 0, MF_BYPOSITION);                                                                                    // 删除移动菜单
    RemoveMenu(GetSystemMenu(hEXEFormHandle, False), 0, MF_BYPOSITION);                                                                                    // 删除移动菜单
    RemoveMenu(GetSystemMenu(hEXEFormHandle, False), 0, MF_BYPOSITION);                                                                                    // 删除移动菜单
    RemoveMenu(GetSystemMenu(hEXEFormHandle, False), 0, MF_BYPOSITION);                                                                                    // 删除移动菜单
    SetWindowLong(hEXEFormHandle, GWL_STYLE, Integer(WS_CAPTION OR WS_POPUP OR WS_VISIBLE OR WS_CLIPSIBLINGS OR WS_CLIPCHILDREN OR WS_SYSMENU));           // $96C80000);                                                                        // $96000000
    SetWindowLong(hEXEFormHandle, GWL_EXSTYLE, Integer(WS_EX_LEFT OR WS_EX_LTRREADING OR WS_EX_DLGMODALFRAME OR WS_EX_WINDOWEDGE OR WS_EX_CONTROLPARENT)); // $00010000);                                                                              // $00010101
    ShowWindow(hEXEFormHandle, SW_SHOWNORMAL);                                                                                                             // 显示窗体
    Application.MainForm.Height := Application.MainForm.Height + 1;
    Application.MainForm.Height := Application.MainForm.Height - 1;
    g_PageControl.ActivePage    := g_Tabsheet;
    g_lblInfo.Caption           := g_strFileValue.Split([';'])[0] + ' - ' + g_strFileValue.Split([';'])[1];
    KillTimer(Application.MainForm.Handle, $1000);
    SetTimer(Application.MainForm.Handle, $2000, 100, @EndExeForm);
  end;
end;

procedure PBoxRun_IMAGE_EXE(const strEXEFileName, strFileValue: String; pg: TPageControl; ts: TTabSheet; lblInfo: TLabel; const UIShowStyle: TShowStyle);
begin
  g_PageControl                := pg;
  g_Tabsheet                   := ts;
  g_lblInfo                    := lblInfo;
  g_strFileValue               := strFileValue;
  g_strCreateDllFileNameBackUp := g_strCreateDllFileName;
  g_UIShowStyle                := UIShowStyle;

  g_strEXEFormClassName := strFileValue.Split([';'])[2];
  g_strEXEFormTitleName := strFileValue.Split([';'])[3];
  SetTimer(Application.MainForm.Handle, $1000, 100, @FindExeForm);

  { 创建 EXE 进程，并隐藏窗体 }
  ShellExecute(Application.MainForm.Handle, 'Open', PChar(strEXEFileName), nil, nil, SW_HIDE);
end;

end.
