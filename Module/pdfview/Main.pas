unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Vcl.ExtCtrls,
  PDFium.Frame, db.uCommon;

type
  TMainForm = class(TForm)
    PDFium: TPDFiumFrame;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Open1: TMenuItem;
    N1: TMenuItem;
    Quit1: TMenuItem;
    pnButtons: TPanel;
    btZPlus: TPaintBox;
    btZMinus: TPaintBox;
    btOpen: TPaintBox;
    pbZoom: TPaintBox;
    ppZoom: TPopupMenu;
    N101: TMenuItem;
    N251: TMenuItem;
    N501: TMenuItem;
    N1001: TMenuItem;
    N1002: TMenuItem;
    N1251: TMenuItem;
    N1501: TMenuItem;
    N2001: TMenuItem;
    N4001: TMenuItem;
    N8001: TMenuItem;
    N16001: TMenuItem;
    N24001: TMenuItem;
    N32001: TMenuItem;
    N64001: TMenuItem;
    N2: TMenuItem;
    mnActualSize: TMenuItem;
    mnFitWidth: TMenuItem;
    mnPageLevel: TMenuItem;
    btPageWidth: TPaintBox;
    btFullPage: TPaintBox;
    btActualSize: TPaintBox;
    Close1: TMenuItem;
    procedure Open1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ButtonMouseLeave(Sender: TObject);
    procedure ButtonPaint(Sender: TObject);
    procedure ButtonMouseEnter(Sender: TObject);
    procedure pbZoomPaint(Sender: TObject);
    procedure pbZoomMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure ppZoomPopup(Sender: TObject);
    procedure btZPlusClick(Sender: TObject);
    procedure MenuZoomClick(Sender: TObject);
    procedure mnPageLevelClick(Sender: TObject);
    procedure mnFitWidthClick(Sender: TObject);
    procedure PDFiumResize(Sender: TObject);
    procedure mnActualSizeClick(Sender: TObject);
    procedure Quit1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FButtons: TBitmap;
    FFocus  : TPaintBox;
    procedure CreateButtons;
  end;

var
  MainForm: TMainForm;

type
  { ֧�ֵ��ļ����� }
  TSPFileType = (ftDelphiDll, ftVCDialogDll, ftVCMFCDll, ftQTDll, ftEXE);

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;

implementation

{$R *.dfm}

uses
  DynamicButtons;

resourcestring
  sPDFFiler = 'Adobe PDF files (*.pdf)|*.pdf';
  sPDFPrompt = 'Open';

procedure db_ShowDllForm_Plugins(var frm: TFormClass; var ft: TSPFileType; var strParentModuleName, strModuleName, strClassName, strWindowName, strIconFileName: PAnsiChar; const bShow: Boolean = True); stdcall;
begin
  frm                     := TMainForm;
  ft                      := ftDelphiDll;
  strParentModuleName     := '�ļ�����';
  strModuleName           := 'PDF�鿴��';
  strClassName            := '';
  strWindowName           := '';
  strIconFileName         := '';
  Application.Handle      := GetMainFormApplication.Handle;
  Application.Icon.Handle := GetMainFormApplication.Icon.Handle;
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FButtons.Free;
  Action := caFree;
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  FButtons := TBitmap.Create;
  CreateButtons;
end;

procedure TMainForm.MenuZoomClick(Sender: TObject);
begin
  PDFium.ZoomMode := zmCustom;
  PDFium.Zoom     := TComponent(Sender).Tag;
end;

procedure TMainForm.mnActualSizeClick(Sender: TObject);
begin
  PDFium.ZoomMode := zmCustom;
  PDFium.Zoom     := 100;
end;

procedure TMainForm.mnFitWidthClick(Sender: TObject);
begin
  PDFium.ZoomMode := zmPageWidth;
end;

procedure TMainForm.mnPageLevelClick(Sender: TObject);
begin
  PDFium.ZoomMode := zmPageLevel;
end;

procedure TMainForm.ButtonPaint(Sender: TObject);
var
  D, S: TRect;
  I   : Integer;
begin
  D := TPaintBox(Sender).ClientRect;
  S := TRect.Create(0, 0, 24, 24);
  I := TPaintBox(Sender).Tag;
  if (Sender = FFocus) and (Odd(I) = False) then
    Inc(I);
  S.Offset(24 * I, 0);
  with TPaintBox(Sender).Canvas do
  begin
    CopyRect(D, FButtons.Canvas, S);
  end;
end;

procedure TMainForm.btZPlusClick(Sender: TObject);
var
  Zoom : Integer;
  Z1   : Integer;
  Z2   : Integer;
  Zx   : Integer;
  Index: Integer;
begin
  Z1        := 1000;
  Z2        := 640000;
  Zoom      := Round(PDFium.Zoom * 100);
  for Index := 0 to ppZoom.Items.count - 1 do
  begin
    Zx := 100 * ppZoom.Items[Index].Tag;
    if Zx <> 0 then
    begin
      if (Zx < Zoom) and (Zx > Z1) then
        Z1 := Zx
      else if (Zx > Zoom) and (Zx < Z2) then
        Z2 := Zx;
    end;
  end;
  if Sender = btZPlus then
    Z1            := Z2;
  PDFium.ZoomMode := zmCustom;
  PDFium.Zoom     := Z1 / 100;
  pbZoom.Invalidate;
end;

procedure TMainForm.ButtonMouseEnter(Sender: TObject);
begin
  if Sender = FFocus then
    Exit;
  if FFocus <> nil then
    FFocus.Invalidate;
  FFocus := TPaintBox(Sender);
  if FFocus <> nil then
    FFocus.Invalidate;
end;

procedure TMainForm.ButtonMouseLeave(Sender: TObject);
begin
  if FFocus <> nil then
  begin
    FFocus.Invalidate;
    FFocus := nil;
  end;
end;

procedure TMainForm.Close1Click(Sender: TObject);
begin
  PDFium.CloseDocument;
end;

procedure TMainForm.CreateButtons;
begin
  FButtons.PixelFormat := pf32Bit;
  FButtons.SetSize(2 * 2 * 7 * 24, 2 * 24);
  with FButtons.Canvas do
  begin
    Brush.Color := clBtnFace;
    FillRect(Rect(0, 0, FButtons.Width, FButtons.Height));

    Pen.Width := 4;
    DrawButtons(0, $6F6F6F);
    DrawButtons(48, $C87521);
  end;
  // FButtons.SaveToFile('BUTTONS48.BMP');
  AntiAliaze(FButtons);
  with FButtons.Canvas do
  begin
    Pen.Width := 1;
    FixButtons(0, $6F6F6F);
    FixButtons(24, $C87521);
  end;
  // FButtons.SaveToFile('BUTTONS24.BMP');
end;

procedure TMainForm.Open1Click(Sender: TObject);
var
  Str: string;
begin
  Str := '';
  if PromptForFileName(Str, sPDFFiler, sPDFPrompt) then
  begin
    PDFium.LoadFromFile(Str);
    PDFium.SetFocus();
    Caption := ExtractFileName(Str) + ' - ' + Application.Title;
  end;
end;

procedure TMainForm.pbZoomMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  p: TPoint;
begin
  if Button = mbLeft then
  begin
    p.X := X;
    p.Y := Y;
    p   := pbZoom.ClientToScreen(p);
    ppZoom.Popup(p.X, p.Y);
  end;
end;

procedure TMainForm.pbZoomPaint(Sender: TObject);
var
  Str: string;
  R  : TRect;
begin
  R := pbZoom.ClientRect;
  with pbZoom.Canvas do
  begin
    Pen.Color   := $CBCBCB;
    Brush.Color := clWhite;
    Rectangle(R);
    if FFocus = Sender then
    begin
      Pen.Color   := $D28E4A;
      Brush.Color := $C87521;
    end
    else
    begin
      Pen.Color   := $8A8A8A;
      Brush.Color := $6F6F6F;
    end;
    Arrow(60, 12);
    Dec(R.Right, 24);
    Str         := FloatToStrF(PDFium.Zoom, TFloatFormat.ffFixed, 18, 0) + '%';
    Brush.Color := clWhite;
    Font.Color  := $4D4D4D;
    TextRect(R, Str, [TTextFormats.tfSingleLine, TTextFormats.tfCenter, TTextFormats.tfVerticalCenter]);
  end;
end;

procedure TMainForm.PDFiumResize(Sender: TObject);
begin
  btPageWidth.Tag  := 6 + Ord(PDFium.ZoomMode = zmPageWidth);
  btFullPage.Tag   := 8 + Ord(PDFium.ZoomMode = zmPageLevel);
  btActualSize.Tag := 10 + Ord((PDFium.ZoomMode = zmCustom) and (Round(PDFium.Zoom * 100) = 10000));
  pnButtons.Invalidate;
end;

procedure TMainForm.ppZoomPopup(Sender: TObject);
var
  Zoom : Integer;
  Index: Integer;
begin
  Zoom      := Round(PDFium.Zoom * 100);
  for Index := 0 to ppZoom.Items.count - 1 do
  begin
    ppZoom.Items[Index].Checked := (PDFium.ZoomMode = zmCustom) and (ppZoom.Items[Index].Tag * 100 = Zoom);
  end;
  mnActualSize.Checked := (PDFium.ZoomMode = zmCustom) and (Zoom = 10000);
  mnPageLevel.Checked  := PDFium.ZoomMode = zmPageLevel;
  mnFitWidth.Checked   := PDFium.ZoomMode = zmPageWidth;
end;

procedure TMainForm.Quit1Click(Sender: TObject);
begin
  Close();
end;

end.
