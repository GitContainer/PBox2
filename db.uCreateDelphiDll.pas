unit db.uCreateDelphiDll;
{
  创建 Delphi DLL窗体
}

interface

uses Vcl.Forms, Winapi.Windows, Winapi.Messages, System.Classes, Vcl.Graphics, Vcl.ComCtrls, Vcl.Controls, Vcl.StdCtrls, db.uCommon;

procedure PBoxRun_DelphiDll(var DllForm: TForm; Page: TPageControl; tsDllForm: TTabSheet; OnDelphiDllFormDestroy: TNotifyEvent);
procedure CloseDelphiDllForm;

implementation

var
  FOnDelphiDllFormDestroy: TNotifyEvent = nil;
  FhDelphiFormDll        : THandle      = 0;
  FbCloseDelphiFormDll   : Boolean      = False;

procedure CloseDelphiDllForm;
begin
  if FhDelphiFormDll = 0 then
    Exit;

  { 关闭 Delphi Dll 窗体 }
  PostMessage(FhDelphiFormDll, WM_SYSCOMMAND, SC_CLOSE, 0);

  { 等待窗体关闭 }
  while True do
  begin
    Application.ProcessMessages;
    if FbCloseDelphiFormDll then
      Break
  end;
end;

{ Delphi Dll 窗体关闭后，变量复位 }
procedure DelphiDllFormDestory(hWnd: hWnd; uMsg, idEvent: UINT; dwTime: DWORD); stdcall;
begin
  if not IsWindowVisible(FhDelphiFormDll) then
  begin
    KillTimer(Application.MainForm.Handle, $3000);
    FOnDelphiDllFormDestroy(nil);
    FbCloseDelphiFormDll := True;
    FhDelphiFormDll      := 0;
  end;
end;

procedure PBoxRun_DelphiDll(var DllForm: TForm; Page: TPageControl; tsDllForm: TTabSheet; OnDelphiDllFormDestroy: TNotifyEvent);
var
  hDll                             : HMODULE;
  ShowDllForm                      : Tdb_ShowDllForm_Plugins;
  frm                              : TFormClass;
  strParamModuleName, strModuleName: PAnsiChar;
  strClassName, strWindowName      : PAnsiChar;
  strIconFileName                  : PAnsiChar;
  ft                               : TSPFileType;
begin
  hDll        := LoadLibrary(PChar(g_strCreateDllFileName));
  ShowDllForm := GetProcAddress(hDll, c_strDllExportName);
  ShowDllForm(frm, ft, strParamModuleName, strModuleName, strClassName, strWindowName, strIconFileName, False);
  DllForm                 := frm.Create(nil);
  DllForm.BorderIcons     := [biSystemMenu];
  DllForm.Position        := poDesigned;
  DllForm.BorderStyle     := bsSingle;
  DllForm.Color           := clWhite;
  DllForm.Anchors         := [akLeft, akTop, akRight, akBottom];
  DllForm.Tag             := hDll;                                                                                         // 将 hDll 放在 DllForm 的 tag 中，卸载时需要用到
  FhDelphiFormDll         := DllForm.Handle;                                                                               //
  FOnDelphiDllFormDestroy := OnDelphiDllFormDestroy;                                                                       //
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // 删除移动菜单
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // 删除大小菜单
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // 删除最小化菜单
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // 删除最大化菜单
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // 删除分割线菜单
  SetWindowPos(DllForm.Handle, tsDllForm.Handle, 0, 0, tsDllForm.Width, tsDllForm.Height, SWP_NOZORDER OR SWP_NOACTIVATE); // 最大化 Dll 子窗体
  Winapi.Windows.SetParent(DllForm.Handle, tsDllForm.Handle);                                                              // 设置父窗体为 TabSheet
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // 删除移动菜单
  DllForm.Show;                                                                                                            // 显示 Dll 子窗体
  Page.ActivePage      := tsDllForm;
  FbCloseDelphiFormDll := False;
  SetTimer(Application.MainForm.Handle, $3000, 100, @DelphiDllFormDestory);
end;

end.
