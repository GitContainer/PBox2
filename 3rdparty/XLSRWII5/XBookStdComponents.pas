unit XBookStdComponents;

{-
********************************************************************************
******* XLSSpreadSheet V3.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}

interface

uses Classes, SysUtils, Windows, Messages, vcl.Controls, vcl.Graphics, vcl.Forms, UxTheme,
     vcl.StdCtrls,
     Xc12Utils5, Xc12DataStyleSheet5,
     XLSUtils5,
     XBookTypes2, XBookPaintGDI2,
     ExcelColorPicker;

const XSS_COMBO_BTN_WIDTH = 18;

type TXBookComboState = (xbcsDisabled,xbcsNormal,xbcsHoover,xbcsClicked);

type TXBookDropDownWin = class(TCustomControl)
protected
     FCombo : TCustomControl;
     FCancel: boolean;

     procedure WMKillFocus(var Message: TWMKillFocus); message WM_KILLFOCUS;
     procedure WMKeyDown(var Message: TWMKeyDown); message WM_KEYDOWN;
public
     procedure Paint; override;

     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     end;

type TXBookColorPicker = class(TXBookDropDownWin)
private
     procedure SetColor(const Value: TXc12Color);
     function  GetCompact: boolean;
     procedure SetCompact(const Value: boolean);
protected
     FCPTheme    : TExcelColorPicker;
     FCPStandard : TExcelColorPicker;
     FTitleHeight: integer;
     FSelColor   : TXc12Color;

     procedure ThemeColorClicked(ASender: TObject);
     procedure StandardColorClicked(ASender: TObject);
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure Paint; override;

     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

     property Color: TXc12Color write SetColor;
     property SelectedColor: TXc12Color read FSelColor;
published
     property Compact: boolean read GetCompact write SetCompact;
     end;

type TXBookWindowComboBox = class(TCustomControl)
protected
     FHTheme  : HTHEME;

     FState   : TXBookComboState;

     FDropDown: TXBookDropDownWin;
     FCanceled: boolean;

     FIsDestroyingDropdown: boolean;

     procedure CMMouseEnter(var Message: TMessage); message CM_MOUSEENTER;
     procedure CMMouseLeave(var Message: TMessage); message CM_MOUSELEAVE;
     procedure WMSetFocus(var Message: TWMSetFocus); message WM_SETFOCUS;
     procedure WMKillFocus(var Message: TWMKillFocus); message WM_KILLFOCUS;
     procedure SetEnabled(Value: Boolean); override;
     procedure CreateDropDown; virtual;
     procedure BeforeCloseUp; virtual;
     procedure CloseUp; virtual;
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure Paint; override;

     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

     property Canceled: boolean read FCanceled;

     property TabStop;
     end;

type TXBookColorComboBox = class(TXBookWindowComboBox)
private
     procedure SetColor(const Value: TXc12Color);
protected
     FColor      : TXc12Color;
     FSelColor   : TXc12Color;
     FSelectEvent: TNotifyEvent;

     procedure CreateDropDown; override;
     procedure BeforeCloseUp; override;
     procedure CloseUp; override;
public

     procedure Paint; override;
published
     property Color: TXc12Color read FColor write SetColor;

     property OnSelect: TNotifyEvent read FSelectEvent write FSelectEvent;
     end;

type TXBookCellBorderPickerStyle = (xcbpsCell,xcbpsColumn,xcbpsRow,xcbpsArea);
type TXBookCellBorderPickerSide = (xcbpsLeft,xcbpsTop,xcbpsBottom,xcbpsRight,xcbpsInsideHoriz,xcbpsInsideVert,xcbpsDiagonalUp,xcbpsDiagonalDown);

type TXBookCellBorderLine = record
     Pt1     : TXYPoint;
     Pt2     : TXYPoint;
     Click   : TXYRect;
     CanClick: boolean; // For diag lines. Click is not implemented on these.
     Style   : TXc12CellBorderStyle;
     Color   : TXc12Color;
     Changed : boolean;
     end;

type TXBookCellBorderPicker = class(TCustomControl)
private
     procedure SetStyle(const Value: TXBookCellBorderPickerStyle);
     function  GetButton(const Index: Integer): TButton;
     procedure SetButton(const Index: Integer; const Value: TButton);
     function  GetButtonPreset(const Index: Integer): TButton;
     procedure SetButtonPreset(const Index: Integer; const Value: TButton);
     function  GetSides(Index: TXBookCellBorderPickerSide): TXBookCellBorderLine;
protected
     FMargin1         : integer;
     FMargin2         : integer;
     FStyle           : TXBookCellBorderPickerStyle;
     // +---1---+---1---+
     // 0       5       3
     // +---4---+---4---+
     // 0       5       3
     // +---2---+---2---+
     FLines           : array of TXBookCellBorderLine;
     FButtons         : array[0..7] of TButton;
     FBtnPreset       : array[0..2] of TButton; // 0 = None, 1 = Outline, 2 = Inside
     FBorderStyle     : TXc12CellBorderStyle;
     FBorderColor     : TXc12Color;
     FSampleText      : AxUCString;

     FChanged         : boolean;

     FSelectEvent: TNotifyEvent;

     procedure CalcLines;
     function  FindLine(const AX,AY: integer): integer;
     procedure SetLine(const AIndex: integer);
     procedure UpdateLine(const AIndex: integer);
     procedure ButtonClick(ASender: TObject);
     procedure ButtonPresetClick(ASender: TObject);
public
     constructor Create(AOwner: TComponent); override;

     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure Paint; override;

     property Style              : TXBookCellBorderPickerStyle read FStyle write SetStyle;

     property BorderStyle        : TXc12CellBorderStyle read FBorderStyle write FBorderStyle;
     property BorderColor        : TXc12Color read FBorderColor write FBorderColor;

     property Sides[Index: TXBookCellBorderPickerSide]: TXBookCellBorderLine read GetSides;
published
     property ButtonLeft         : TButton index 0 read GetButton write SetButton;
     property ButtonTop          : TButton index 1 read GetButton write SetButton;
     property ButtonBottom       : TButton index 2 read GetButton write SetButton;
     property ButtonRight        : TButton index 3 read GetButton write SetButton;
     property ButtonInsideHoriz  : TButton index 4 read GetButton write SetButton;
     property ButtonInsideVert   : TButton index 5 read GetButton write SetButton;
     property ButtonDiagUp       : TButton index 6 read GetButton write SetButton;
     property ButtonDiagDown     : TButton index 7 read GetButton write SetButton;
     property ButtonPresetNone   : TButton index 0 read GetButtonPreset write SetButtonPreset;
     property ButtonPresetOutline: TButton index 1 read GetButtonPreset write SetButtonPreset;
     property ButtonPresetInside : TButton index 2 read GetButtonPreset write SetButtonPreset;

     // Symbol Changed exists in base class.
     property IsChanged          : boolean read FChanged;

     property OnSelect: TNotifyEvent read FSelectEvent write FSelectEvent;
     end;


type TXBookCellBorderStyle = record
     Style: TXc12CellBorderStyle;
     Click: TXYRect;
     end;

type TXBookCellBorderStylePicker = class(TCustomControl)
private
     function  GetLineStyle: TXc12CellBorderStyle;
     procedure SetLineStyle(const Value: TXc12CellBorderStyle);
     procedure SetLineColor(const Value: TXc12Color);
protected
     FStyles: array[0..13] of TXBookCellBorderStyle;
     FSelected: integer;

     FLineColor : TXc12Color;

     FMargin: integer;
     FLineCellHeight: integer;

     FSelectEvent: TNotifyEvent;

     procedure CalcStyles;

     procedure CMWantSpecialKey(var Msg: TCMWantSpecialKey); message CM_WANTSPECIALKEY;
     procedure WMKeyDown(var Message: TWMKeyDown); message WM_KEYDOWN;

     procedure SetSelectedIndex(const AIndex: integer);
public
     constructor Create(AOwner: TComponent); override;

     procedure Paint; override;
     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;

     property LineColor: TXc12Color read FLineColor write SetLineColor;
     property LineStyle: TXc12CellBorderStyle read GetLineStyle write SetLineStyle;
published
     property OnSelect: TNotifyEvent read FSelectEvent write FSelectEvent;
     end;

implementation

{ TXBookWindowComboBox }

procedure TXBookWindowComboBox.BeforeCloseUp;
begin

end;

procedure TXBookWindowComboBox.CloseUp;
begin

end;

procedure TXBookWindowComboBox.CMMouseEnter(var Message: TMessage);
begin
  if FState = xbcsNormal then
    FState := xbcsHoover;
  Paint;
end;

procedure TXBookWindowComboBox.CMMouseLeave(var Message: TMessage);
begin
  if FState = xbcsHoover then
    FState := xbcsNormal;
  Paint;
end;

constructor TXBookWindowComboBox.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  Height := 21;
  Width := 110;

  FState := xbcsNormal;
end;

procedure TXBookWindowComboBox.CreateDropDown;
begin
  FDropDown := TXBookDropDownWin.Create(Owner);
end;

destructor TXBookWindowComboBox.Destroy;
begin
  CloseThemeData(FHTheme);

  inherited;
end;

procedure TXBookWindowComboBox.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;

  if FDropDown <> Nil then begin
    FIsDestroyingDropdown := True;
    FDropDown.Free;
    FIsDestroyingDropdown := False;
    FDropDown := Nil;
    FState := xbcsHoover;
  end
  else begin
    CreateDropDown;
    FDropDown.Parent := Parent;
    FDropDown.FCombo := Self;
    FDropDown.Left := Left;
    FDropDown.Top := Top + Height + 1;
    if FDropDown.Width <= 0 then
      FDropDown.Width := 150;
    if FDropDown.Height <= 0 then
      FDropDown.Height := 150;
    FDropDown.SetFocus;
    FState := xbcsClicked;
  end;
  Paint;
end;

procedure TXBookWindowComboBox.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  inherited;

end;

procedure TXBookWindowComboBox.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
//
//  if ClientRect.Contains(Point(X,Y)) then
//    FState := xbcsHoover
//  else
//    FState := xbcsNormal;
//  Paint;
end;

procedure TXBookWindowComboBox.Paint;
{$ifndef DELPHI_2007_OR_LATER}
const
  CP_BORDER     = 4;
  CBB_NORMAL    = 1;
  CBB_HOT       = 2;
  CBB_FOCUSED   = 3;
  CBB_DISABLED  = 4;
{$endif}
var
  ButtonRect: TRect;
begin
  inherited;

  if FHTheme = 0 then
    FHTheme := OpenThemeData(Handle,'COMBOBOX');

  Canvas.Pen.Color := clSilver;
  Canvas.Brush.Color := clWhite;
  Canvas.Rectangle(0,0,Width,Height);

  if Focused then
    Canvas.DrawFocusRect(Rect(2,2,Width - XSS_COMBO_BTN_WIDTH - 2,Height - 2));

  ButtonRect.Left := Width - XSS_COMBO_BTN_WIDTH;
  ButtonRect.Top := 1;
  ButtonRect.Right := Width - 1;
  ButtonRect.Bottom := Height - 1;

  case FState of
    xbcsDisabled : begin
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_BORDER,CBB_DISABLED,ButtonRect,Nil);
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_DROPDOWNBUTTON,CBXS_DISABLED,ButtonRect,Nil);
    end;
    xbcsNormal   : begin
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_BORDER,CBB_NORMAL,ButtonRect,Nil);
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_DROPDOWNBUTTON,CBXS_NORMAL,ButtonRect,Nil);
    end;
    xbcsHoover   : begin
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_BORDER,CBB_HOT,ButtonRect,Nil);
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_DROPDOWNBUTTON,CBXS_HOT,ButtonRect,Nil);
    end;
    xbcsClicked  : begin
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_BORDER,CBB_FOCUSED,ButtonRect,Nil);
      DrawThemeBackground(FHtheme,Canvas.Handle,CP_DROPDOWNBUTTON,CBXS_PRESSED,ButtonRect,Nil);
    end;
  end;
end;

procedure TXBookWindowComboBox.SetEnabled(Value: Boolean);
begin
  inherited;

  FState := xbcsDisabled;
  Paint;
end;

procedure TXBookWindowComboBox.WMKillFocus(var Message: TWMKillFocus);
begin
  FState := xbcsNormal;
  Paint;
end;

procedure TXBookWindowComboBox.WMSetFocus(var Message: TWMSetFocus);
begin
  if (FDropDown <> Nil) and not FIsDestroyingDropdown then begin
    FCanceled := FDropDown.FCancel;
    BeforeCloseUp;
    FDropDown.Free;
    FDropDown := Nil;
    CloseUp;
  end;
  FState := xbcsHoover;
  Paint;
end;

{ TXBookDropDownWin }

procedure TXBookDropDownWin.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FCombo <> Nil then
    FCombo.SetFocus;
end;

procedure TXBookDropDownWin.Paint;
begin
  inherited;

  Canvas.Pen.Color := clBlack;
  Canvas.Brush.Color := clWhite;
  Canvas.Rectangle(0,0,Width,Height);
end;

procedure TXBookDropDownWin.WMKeyDown(var Message: TWMKeyDown);
begin
  if Message.CharCode = VK_ESCAPE then begin
    FCancel := True;
    if FCombo <> Nil then
      FCombo.SetFocus;
  end;
end;

procedure TXBookDropDownWin.WMKillFocus(var Message: TWMKillFocus);
begin
  if FCombo <> Nil then
    FCombo.SetFocus;
end;

{ TXBookColorComboBox }

procedure TXBookColorComboBox.BeforeCloseUp;
begin
  FSelColor := TXBookColorPicker(FDropDown).SelectedColor;
end;

procedure TXBookColorComboBox.CloseUp;
begin
  if FSelColor.ColorType <> exctUnassigned then begin
    FColor := FSelColor;
    if Assigned(FSelectEvent) then
      FSelectEvent(Self);
  end;
end;

procedure TXBookColorComboBox.CreateDropDown;
begin
  FDropDown := TXBookColorPicker.Create(Owner);
  FDropDown.Parent := Parent.Parent;
end;

procedure TXBookColorComboBox.Paint;
begin
  inherited;

  if FColor.ColorType <> exctUnassigned then begin
    if FColor.ColorType = exctAuto then begin
      Canvas.Pen.Color := clWhite;
      Canvas.Brush.Color := clWhite;
      Canvas.Rectangle(4,4,Width - XSS_COMBO_BTN_WIDTH - 4,Height - 4);
      Canvas.TextOut(8,4,'Automatic');
    end
    else begin
      Canvas.Pen.Color := RevRGB(FColor.ARGB);
      Canvas.Brush.Color := RevRGB(FColor.ARGB);
      Canvas.Rectangle(4,4,Width - XSS_COMBO_BTN_WIDTH - 4,Height - 4);
    end;
  end;
end;

procedure TXBookColorComboBox.SetColor(const Value: TXc12Color);
begin
  FColor := Value;
  Repaint;
end;

{ TXBookColorDropDownWin }

constructor TXBookColorPicker.Create(AOwner: TComponent);
begin
  inherited;

  FSelColor.ColorType := exctUnassigned;

  FTitleHeight := 20;

  FCPTheme := TExcelColorPicker.Create(Owner);
  FCPTheme.Parent := Self;
  FCPTheme.Left := 1;
  FCPTheme.Top := FTitleHeight * 2;
  FCPTheme.ColorMode := ecpmExcel2007ThemeCompact;
  FCPTheme.OnClick := ThemeColorClicked;

  FCPStandard := TExcelColorPicker.Create(Owner);
  FCPStandard.Parent := Self;
  FCPStandard.Left := 1;
  FCPStandard.Top := FCPTheme.Top + FCPTheme.Height + FTitleHeight;
  FCPStandard.ColorMode := ecpmExcel2007Standard;
  FCPStandard.LinkedPicker := FCPTheme;
  FCPStandard.OnClick := StandardColorClicked;

  Width := FCPStandard.Width + 2;
  Height := FCPStandard.Top + FCPStandard.Height + 2;
end;

destructor TXBookColorPicker.Destroy;
begin
  FCPTheme.Free;
  FCPStandard.Free;
  inherited;
end;

function TXBookColorPicker.GetCompact: boolean;
begin
  Result := FCPTheme.ColorMode = ecpmExcel2007ThemeCompact;
end;

procedure TXBookColorPicker.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if Y < FTitleHeight then begin
    FSelColor := MakeXc12ColorAuto;
    if FCombo <> Nil then
      FCombo.SetFocus;
  end;
end;

procedure TXBookColorPicker.Paint;
begin
  inherited;

  Canvas.Pen.Color := RGB(219,229,241);
  Canvas.Brush.Style := bsSolid;
  Canvas.Brush.Color := RGB(219,229,241);
  Canvas.Font.Color := clNavy;
  Canvas.Font.Style := [];

  Canvas.Rectangle(1,FTitleHeight,Width - 2,FTitleHeight * 2);
  Canvas.Rectangle(1,FCPStandard.Top - FTitleHeight,Width - 2,FCPStandard.Top - 2);

  Canvas.Brush.Color := clBlack;
  Canvas.Rectangle(4,4,FTitleHeight - 4,FTitleHeight - 4);
  if FCPTheme.ExcelColor.Auto then begin
    Canvas.Brush.Style := bsClear;
    Canvas.Pen.Color := SWATCH_MOUSECOLOR1;
    Canvas.Pen.Width := 2;
    Canvas.Rectangle(3,3,FTitleHeight - 2,FTitleHeight - 2);
    Canvas.Pen.Width := 1;
  end;

  Canvas.Brush.Style := bsClear;
  Canvas.TextOut(24,4,'Automatic');

  Canvas.Font.Style := [fsBold];

  Canvas.TextOut(8,FTitleHeight + 2,'Theme colors');
  Canvas.TextOut(8,FCPStandard.Top - FTitleHeight + 2,'Standard colors');
end;

procedure TXBookColorPicker.SetColor(const Value: TXc12Color);
begin
  FCPTheme.FindAndSelect(Value);
end;

procedure TXBookColorPicker.SetCompact(const Value: boolean);
begin
  if Value <> GetCompact then begin
    if Value then
      FCPTheme.ColorMode := ecpmExcel2007ThemeCompact
    else
      FCPTheme.ColorMode := ecpmExcel2007Theme;
  end;
end;

procedure TXBookColorPicker.StandardColorClicked(ASender: TObject);
begin
  FSelColor := FCPStandard.ExcelColor;
  if FCombo <> Nil then
    FCombo.SetFocus;
end;

procedure TXBookColorPicker.ThemeColorClicked(ASender: TObject);
begin
  FSelColor := FCPTheme.ExcelColor;
  if FCombo <> Nil then
    FCombo.SetFocus;
end;

{ TXBookCellBorderPicker }

procedure TXBookCellBorderPicker.ButtonClick(ASender: TObject);
begin
  if TButton(ASender).Tag > 5 then begin
    case FStyle of
      xcbpsCell  : UpdateLine(TButton(ASender).Tag);
      xcbpsColumn,
      xcbpsRow   : begin
        UpdateLine(TButton(ASender).Tag);
        UpdateLine(TButton(ASender).Tag + 2);
      end;
      xcbpsArea  : begin
        if TButton(ASender).Tag = 6 then begin
          UpdateLine(6);
          UpdateLine(7);
          UpdateLine(8);
        end
        else if TButton(ASender).Tag = 7 then begin
          UpdateLine(9);
          UpdateLine(10);
          UpdateLine(11);
        end;
      end;
    end;
  end
  else
    UpdateLine(TButton(ASender).Tag);
  Paint;

  if Assigned(FSelectEvent) then
    FSelectEvent(Self);
end;

procedure TXBookCellBorderPicker.ButtonPresetClick(ASender: TObject);
var
  i: integer;
begin
  case TButton(ASender).Tag of
    0: begin
      for i := 0 to High(FLines) do begin
        FLines[i].Style := cbsNone;
        FLines[i].Changed := True;
      end;
      FChanged := True;
    end;
    1: begin
      for i := 0 to 3 do
        SetLine(i);
    end;
    2: begin
      case FStyle of
        xcbpsColumn: SetLine(4);
        xcbpsRow   : SetLine(5);
        xcbpsArea  : begin
          SetLine(4);
          SetLine(5);
        end;
      end;
    end;
  end;
  Paint;

  if Assigned(FSelectEvent) then
    FSelectEvent(Self);
end;

procedure TXBookCellBorderPicker.CalcLines;
var
  i: integer;
begin
  case FStyle of
    xcbpsCell  : SetLength(FLines,8);
    xcbpsColumn: SetLength(FLines,10);
    xcbpsRow   : SetLength(FLines,10);
    xcbpsArea  : SetLength(FLines,12);
  end;

  FLines[0].Pt1.X := FMargin2 - 1;
  FLines[0].Pt1.Y := FMargin2;
  FLines[0].Pt2.X := FMargin2 - 1;
  FLines[0].Pt2.Y := Height - FMargin2;

  FLines[1].Pt1.X := FMargin2 - 1;
  FLines[1].Pt1.Y := FMargin2;
  FLines[1].Pt2.X := Width - FMargin2;
  FLines[1].Pt2.Y := FMargin2;

  FLines[2].Pt1.X := FMargin2 - 1;
  FLines[2].Pt1.Y := Height - FMargin2;
  FLines[2].Pt2.X := Width - FMargin2;
  FLines[2].Pt2.Y := Height - FMargin2;

  FLines[3].Pt1.X := Width - FMargin2;
  FLines[3].Pt1.Y := FMargin2;
  FLines[3].Pt2.X := Width - FMargin2;
  FLines[3].Pt2.Y := Height - FMargin2;

  FLines[4].Pt1.X := FMargin2 - 1;
  FLines[4].Pt1.Y := Height div 2;
  FLines[4].Pt2.X := Width - FMargin2;
  FLines[4].Pt2.Y := Height div 2;

  FLines[5].Pt1.X := Width div 2;
  FLines[5].Pt1.Y := FMargin2;
  FLines[5].Pt2.X := Width div 2;
  FLines[5].Pt2.Y := Height - FMargin2;

  FLines[0].CanClick := True;
  FLines[1].CanClick := True;
  FLines[2].CanClick := True;
  FLines[3].CanClick := True;

  case FStyle of
    xcbpsCell  : begin
      FLines[4].CanClick := False;
      FLines[5].CanClick := False;
      if FButtons[4] <> Nil then FButtons[4].Enabled := False;
      if FButtons[5] <> Nil then FButtons[5].Enabled := False;
      if FBtnPreset[2] <> Nil then FBtnPreset[2].Enabled := False;

      FLines[6].Pt1 := FLines[2].Pt1;
      FLines[6].Pt2 := FLines[1].Pt2;
      FLines[7].Pt1 := FLines[1].Pt1;
      FLines[7].Pt2 := FLines[2].Pt2;
    end;
    xcbpsColumn: begin
      FLines[4].CanClick := True;
      FLines[5].CanClick := False;
      if FButtons[5] <> Nil then FButtons[5].Enabled := False;

      FLines[6].Pt1 := FLines[4].Pt1;
      FLines[6].Pt2 := FLines[1].Pt2;
      FLines[7].Pt1 := FLines[1].Pt1;
      FLines[7].Pt2 := FLines[4].Pt2;

      FLines[8].Pt1 := FLines[2].Pt1;
      FLines[8].Pt2 := FLines[4].Pt2;
      FLines[9].Pt1 := FLines[4].Pt1;
      FLines[9].Pt2 := FLines[2].Pt2;
    end;
     // +---1---+---1---+
     // 0       5       3
     // +---4---+---4---+
     // 0       5       3
     // +---2---+---2---+
    xcbpsRow   : begin
      FLines[4].CanClick := False;
      FLines[5].CanClick := True;
      if FButtons[4] <> Nil then FButtons[4].Enabled := False;

      FLines[6].Pt1 := FLines[2].Pt1;
      FLines[6].Pt2 := FLines[5].Pt1;
      FLines[7].Pt1 := FLines[1].Pt1;
      FLines[7].Pt2 := FLines[5].Pt2;

      FLines[8].Pt1 := FLines[5].Pt2;
      FLines[8].Pt2 := FLines[3].Pt1;
      FLines[9].Pt1 := FLines[5].Pt1;
      FLines[9].Pt2 := FLines[2].Pt2;
    end;
    xcbpsArea  : begin
      FLines[4].CanClick := True;
      FLines[5].CanClick := True;

      FLines[6].Pt1 := FLines[4].Pt1;
      FLines[6].Pt2 := FLines[5].Pt1;
      FLines[7].Pt1 := FLines[2].Pt1;
      FLines[7].Pt2 := FLines[1].Pt2;
      FLines[8].Pt1 := FLines[5].Pt2;
      FLines[8].Pt2 := FLines[4].Pt2;

      FLines[9].Pt1 := FLines[4].Pt1;
      FLines[9].Pt2 := FLines[5].Pt2;
      FLines[10].Pt1 := FLines[1].Pt1;
      FLines[10].Pt2 := FLines[2].Pt2;
      FLines[11].Pt1 := FLines[5].Pt1;
      FLines[11].Pt2 := FLines[4].Pt2;
    end;
  end;

  for i := 0 to 5 do begin
    FLines[i].Color := RGBColorToXc12($000000);
    FLines[i].Style := cbsNone;

    FLines[i].Click.X1 := FLines[i].Pt1.X - FMargin2;
    FLines[i].Click.Y1 := FLines[i].Pt1.Y - FMargin2;
    FLines[i].Click.X2 := FLines[i].Pt2.X + FMargin2;
    FLines[i].Click.Y2 := FLines[i].Pt2.Y + FMargin2;
  end;
  for i := 6 to High(Flines) do
    FLines[i].CanClick := False;
end;

constructor TXBookCellBorderPicker.Create(AOwner: TComponent);
begin
  inherited;

  FSampleText := 'Text';

  Width := 150;
  Height := 100;

  FMargin1 := 5;
  FMargin2 := 10;

  FBorderStyle := cbsThin;
  FBorderColor := RGBColorToXc12($000000);

  SetStyle(xcbpsArea);
end;

function TXBookCellBorderPicker.FindLine(const AX, AY: integer): integer;
begin
  for Result := 0 to High(FLines) do begin
    if FLines[Result].CanClick and HitXYRect(AX,AY,FLines[Result].Click) then
      Exit;
  end;
  Result := -1;
end;

function TXBookCellBorderPicker.GetButton(const Index: Integer): TButton;
begin
  Result := FButtons[Index];
end;

function TXBookCellBorderPicker.GetButtonPreset(const Index: Integer): TButton;
begin
  Result := FBtnPreset[Index];
end;

function TXBookCellBorderPicker.GetSides(Index: TXBookCellBorderPickerSide): TXBookCellBorderLine;
begin
  Result := FLines[Integer(Index)];
end;

procedure TXBookCellBorderPicker.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  inherited;

  i := FindLine(X,Y);
  if i >= 0 then begin
    UpdateLine(i);
    Paint;

    if Assigned(FSelectEvent) then
      FSelectEvent(Self);
  end;
end;

procedure TXBookCellBorderPicker.Paint;
var
  i: integer;
  TH: integer;
  X2,Y2: integer;
  X4,Y4: integer;
  X4_3,Y4_3: integer;
begin
  inherited;

  Canvas.Pen.Color := clBlack;
  Canvas.Brush.Color := clWhite;
  Canvas.Rectangle(0,0,Width,Height);

  TH := Canvas.TextHeight(FSampleText) div 2;
  X2 := (Width - FMargin2 * 2) div 2 + FMargin2;
  Y2 := (Height - FMargin2 * 2) div 2 + FMargin2 - TH;
  X4 := (Width - FMargin2 * 2) div 4 + FMargin2;
  Y4 := (Height - FMargin2 * 2) div 4 + FMargin2 - TH;
  X4_3 := ((Width - FMargin2 * 2) div 4) * 3 + FMargin2;
  Y4_3 := ((Height - FMargin2 * 2) div 4) * 3 + FMargin2 - TH;
  Windows.SetTextAlign(Canvas.Handle,TA_CENTER + TA_TOP);
  case FStyle of
    xcbpsCell  : Canvas.TextOut(X2,Y2,FSampleText);
    xcbpsColumn: begin
      Canvas.TextOut(X2,Y4,FSampleText);
      Canvas.TextOut(X2,Y4_3,FSampleText);
    end;
    xcbpsRow   : begin
      Canvas.TextOut(X4,Y2,FSampleText);
      Canvas.TextOut(X4_3,Y2,FSampleText);
    end;
    xcbpsArea  : begin
      Canvas.TextOut(X4,Y4,FSampleText);
      Canvas.TextOut(X4_3,Y4,FSampleText);
      Canvas.TextOut(X4,Y4_3,FSampleText);
      Canvas.TextOut(X4_3,Y4_3,FSampleText);
    end;
  end;
  Canvas.Pen.Width := 1;

  Canvas.MoveTo(FMargin1 - 1,FMargin2);
  Canvas.LineTo(FMargin2 - 1,FMargin2);
  Canvas.LineTo(FMargin2 - 1,FMargin1);

  Canvas.MoveTo(Width - FMargin1,FMargin2);
  Canvas.LineTo(Width - FMargin2,FMargin2);
  Canvas.LineTo(Width - FMargin2,FMargin1);

  Canvas.MoveTo(FMargin1 - 1,Height - FMargin2);
  Canvas.LineTo(FMargin2 - 1,Height - FMargin2);
  Canvas.LineTo(FMargin2 - 1,Height - FMargin1);

  Canvas.MoveTo(Width - FMargin1,Height - FMargin2);
  Canvas.LineTo(Width - FMargin2,Height - FMargin2);
  Canvas.LineTo(Width - FMargin2,Height - FMargin1);

  if FStyle in [xcbpsColumn,xcbpsArea] then begin
    Canvas.MoveTo(Width div 2,FMargin1);
    Canvas.LineTo(Width div 2,FMargin2);
    Canvas.MoveTo(Width div 2 - FMargin1,FMargin2);
    Canvas.LineTo(Width div 2 + FMargin1,FMargin2);

    Canvas.MoveTo(Width div 2,Height - FMargin1);
    Canvas.LineTo(Width div 2,Height - FMargin2);
    Canvas.MoveTo(Width div 2 - FMargin1,Height - FMargin2);
    Canvas.LineTo(Width div 2 + FMargin1,Height - FMargin2);
  end;

  if FStyle in [xcbpsRow,xcbpsArea] then begin
    Canvas.MoveTo(FMargin1,Height div 2);
    Canvas.LineTo(FMargin2,Height div 2);
    Canvas.MoveTo(FMargin2 - 1,Height div 2 - FMargin1);
    Canvas.LineTo(FMargin2 - 1,Height div 2 + FMargin1);

    Canvas.MoveTo(Width - FMargin1,Height div 2);
    Canvas.LineTo(Width - FMargin2,Height div 2);
    Canvas.MoveTo(Width - FMargin2,Height div 2 - FMargin1);
    Canvas.LineTo(Width - FMargin2,Height div 2 + FMargin1);
  end;

  for i := 0 to High(FLines) do begin
    if FLines[i].Style > cbsNone then begin
      DrawExcelLine(Canvas.Handle,FLines[i].Style,FLines[i].Color.ARGB,FLines[i].Pt1.X,FLines[i].Pt1.Y,FLines[i].Pt2.X,FLines[i].Pt2.Y);
//      Canvas.Pen.Width := 1;
//      Canvas.Pen.Color := FLines[i].Color.ARGB;
//      Canvas.MoveTo(FLines[i].Pt1.X,FLines[i].Pt1.Y);
//      Canvas.LineTo(FLines[i].Pt2.X,FLines[i].Pt2.Y);
    end;
  end;
end;

procedure TXBookCellBorderPicker.SetButton(const Index: Integer; const Value: TButton);
begin
  FButtons[Index] := Value;
  if FButtons[Index] <> Nil then begin
    FButtons[Index].Tag := Index;
    FButtons[Index].OnClick := ButtonClick;
  end;
end;

procedure TXBookCellBorderPicker.SetButtonPreset(const Index: Integer; const Value: TButton);
begin
  FBtnPreset[Index] := Value;
  if FBtnPreset[Index] <> Nil then begin
    FBtnPreset[Index].Tag := Index;
    FBtnPreset[Index].OnClick := ButtonPresetClick;
  end;
end;

procedure TXBookCellBorderPicker.SetLine(const AIndex: integer);
begin
  FLines[AIndex].Style := FBorderStyle;
  FLines[AIndex].Color := FBorderColor;
  FLines[AIndex].Changed := True;

  FChanged := True;
end;

procedure TXBookCellBorderPicker.SetStyle(const Value: TXBookCellBorderPickerStyle);
begin
  FStyle := Value;
  CalcLines;
end;

procedure TXBookCellBorderPicker.UpdateLine(const AIndex: integer);
begin
  if (FLines[AIndex].Style = FBorderStyle) and Xc12ColorEqual(FLines[AIndex].Color,FBorderColor) then
    FLines[AIndex].Style := cbsNone
  else
    SetLine(AIndex);
end;

{ TXBookCellStylePicker }

procedure TXBookCellBorderStylePicker.CalcStyles;
var
  X,Y: integer;
  W,H: integer;
begin
  X := FMargin;
  Y := FMargin;

  W := (Width div 2) - FMargin - FMargin;
  H := FLineCellHeight;

  FStyles[0].Style := cbsNone;
  FStyles[0].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[1].Style := cbsHair;
  FStyles[1].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[2].Style := cbsDotted;
  FStyles[2].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[3].Style := cbsDashDotDot;
  FStyles[3].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[4].Style := cbsDashDot;
  FStyles[4].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[5].Style := cbsDashed;
  FStyles[5].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[6].Style := cbsThin;
  FStyles[6].Click := SetXYRect(X,Y,X + W,Y + H);

  X := Width - FMargin - W;
  Y := FMargin;

  FStyles[7].Style := cbsMediumDashDotDot;
  FStyles[7].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[8].Style := cbsSlantedDashDot;
  FStyles[8].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[9].Style := cbsMediumDashDot;
  FStyles[9].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[10].Style := cbsMediumDashed;
  FStyles[10].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[11].Style := cbsMedium;
  FStyles[11].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[12].Style := cbsThick;
  FStyles[12].Click := SetXYRect(X,Y,X + W,Y + H);
  Inc(Y,H);

  FStyles[13].Style := cbsDouble;
  FStyles[13].Click := SetXYRect(X,Y,X + W,Y + H);
end;

procedure TXBookCellBorderStylePicker.CMWantSpecialKey(var Msg: TCMWantSpecialKey);
begin
  if Msg.CharCode in [VK_LEFT,VK_RIGHT,VK_UP,VK_DOWN] then
    Msg.Result := 1;
end;

constructor TXBookCellBorderStylePicker.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FMargin := 5;
  FLineCellHeight := 16;

  Width := 110;
  Height := (FMargin * 2) + ((Length(FStyles) div 2) * FLineCellHeight);

  FLineColor := RGBColorToXc12($000000);

  CalcStyles;
  FSelected := 0;

  TabStop := True;
end;

function TXBookCellBorderStylePicker.GetLineStyle: TXc12CellBorderStyle;
begin
  Result := FStyles[FSelected].Style;
end;

procedure TXBookCellBorderStylePicker.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  i: integer;
begin
  inherited;

  SetFocus;

  for i := 0 to High(FStyles) do begin
    if HitXYRect(X,Y,FStyles[i].Click) then begin
      SetLineStyle(FStyles[i].Style);
      Paint;
      Exit;
    end;
  end;
end;

procedure TXBookCellBorderStylePicker.Paint;
var
  i: integer;
  Y: integer;
begin
  inherited;

  Canvas.Pen.Style := psSolid;
  Canvas.Pen.Color := clBlack;
  Canvas.Brush.Color := clWhite;
  Canvas.Rectangle(0,0,Width,Height);

  for i := 0 to High(FStyles) do begin
    Y := FStyles[i].Click.Y1 + ((FStyles[i].Click.Y2 - FStyles[i].Click.Y1) div 2);

    if FStyles[i].Style = cbsNone then begin
      Canvas.TextOut(FStyles[i].Click.X1 + 4,Y - Canvas.TextHeight('None') div 2,'None');
    end
    else
      DrawExcelLine(Canvas.Handle,FStyles[i].Style,FLineColor.ARGB,FStyles[i].Click.X1,Y,FStyles[i].Click.X2,Y);

    if  i = FSelected then
      Canvas.DrawFocusRect(Rect(FStyles[i].Click.X1 - 2,FStyles[i].Click.Y1,FStyles[i].Click.X2 + 2,FStyles[i].Click.Y2));
  end;
end;

procedure TXBookCellBorderStylePicker.SetLineColor(const Value: TXc12Color);
begin
  FLineColor := Value;
  Paint;
end;

procedure TXBookCellBorderStylePicker.SetLineStyle(const Value: TXc12CellBorderStyle);
var
  i: integer;
begin
  for i := 0 to High(FStyles) do begin
    if FStyles[i].Style = Value then begin
      if i <> FSelected then begin
        SetSelectedIndex(i);
        if Assigned(FSelectEvent) then
          FSelectEvent(Self);
      end;
    end;
  end;
end;

procedure TXBookCellBorderStylePicker.SetSelectedIndex(const AIndex: integer);
begin
  if AIndex <> FSelected then begin
    FSelected := AIndex;
    Paint;
  end;
end;

procedure TXBookCellBorderStylePicker.WMKeyDown(var Message: TWMKeyDown);
begin
  if Message.CharCode in [VK_RIGHT,VK_DOWN] then begin
    Inc(FSelected);
    if FSelected > High(FStyles) then
      FSelected := 0;
    Paint;
  end
  else if Message.CharCode in [VK_LEFT,VK_UP] then begin
    Dec(FSelected);
    if FSelected < 0  then
      FSelected := High(FStyles);
    Paint;
  end;
end;

end.
