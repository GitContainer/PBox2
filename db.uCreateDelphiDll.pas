unit db.uCreateDelphiDll;
{
  ���� Delphi DLL����
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

  { �ر� Delphi Dll ���� }
  PostMessage(FhDelphiFormDll, WM_SYSCOMMAND, SC_CLOSE, 0);

  { �ȴ�����ر� }
  while True do
  begin
    Application.ProcessMessages;
    if FbCloseDelphiFormDll then
      Break
  end;
end;

{ Delphi Dll ����رպ󣬱�����λ }
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
  DllForm.Tag             := hDll;                                                                                         // �� hDll ���� DllForm �� tag �У�ж��ʱ��Ҫ�õ�
  FhDelphiFormDll         := DllForm.Handle;                                                                               //
  FOnDelphiDllFormDestroy := OnDelphiDllFormDestroy;                                                                       //
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // ɾ���ƶ��˵�
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // ɾ����С�˵�
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // ɾ����С���˵�
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // ɾ����󻯲˵�
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // ɾ���ָ��߲˵�
  SetWindowPos(DllForm.Handle, tsDllForm.Handle, 0, 0, tsDllForm.Width, tsDllForm.Height, SWP_NOZORDER OR SWP_NOACTIVATE); // ��� Dll �Ӵ���
  Winapi.Windows.SetParent(DllForm.Handle, tsDllForm.Handle);                                                              // ���ø�����Ϊ TabSheet
  RemoveMenu(GetSystemMenu(DllForm.Handle, False), 0, MF_BYPOSITION);                                                      // ɾ���ƶ��˵�
  DllForm.Show;                                                                                                            // ��ʾ Dll �Ӵ���
  Page.ActivePage      := tsDllForm;
  FbCloseDelphiFormDll := False;
  SetTimer(Application.MainForm.Handle, $3000, 100, @DelphiDllFormDestory);
end;

end.
