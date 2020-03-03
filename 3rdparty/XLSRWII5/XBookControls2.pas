unit XBookControls2;

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
{$I AxCompilers.inc}

interface

uses { Delphi  } Classes, SysUtils, vcl.Controls, Math, vcl.StdCtrls, vcl.Forms,
     { Windows } Windows, UxTheme,
     { XLSRWII } XLSUtils5,
     { XBook   } XBookUtils2, XBookWindows2, XBookSkin2, XBookPaintGDI2;

// TXScrollState and TXScrollHitTest must match
type TXScrollState   = (xssIdle,xssUpClicked,xssDownClicked,xssLargeUpClicked,xssLargeDownClicked,xssTrack,xssReleaseTrack);
type TXScrollHitTest = (xshtOutside,xshtUp,xshtDown,xshtLargeUp,xshtLargeDown,xshtThumb);

type TXLSScrollEvent   = procedure(Sender: TObject; Value: TXScrollState) of object;

const MIN_THUMBSIZE = 8;
const XSCROLL_TIMERTIME = 50;
const XSCROLL_STARTTIMERTIME = 200;

type TXLSShapeAnchor = record
     Col1: integer;
     Col1Offset: integer;
     Row1: integer;
     Row1Offset: integer;
     Col2: integer;
     Col2Offset: integer;
     Row2: integer;
     Row2Offset: integer;
     end;

type TXLSShapeClientAnchor = class(TObject)
protected
     FAnchor: TXLSShapeAnchor;
public

     property Anchor: TXLSShapeAnchor read FAnchor;
     end;

// Base class for controls.
type TXBookClientShape = class(TXSSClientWindow)
protected
     FShape: TXLSShapeClientAnchor;
//     FParams: TEscherSettings;
public
     constructor Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
     procedure Paint; override;

     property Shape: TXLSShapeClientAnchor read FShape;
//     property EscherSettings: TEscherSettings read FParams write FParams;
     end;

type TXLSControl = class(TXBookClientShape)
protected
     FSetCtrlFontEvent: TXLSProcEvent;
public
     procedure SetCtrlFont; virtual;
     property OnSetCtrlFont: TXLSProcEvent read FSetCtrlFontEvent write FSetCtrlFontEvent;
     end;

type TXLSScrollBar = class(TXLSControl)
private
     FHTheme: longword;
     FKindHoriz: boolean;
     FBtnSize: integer;
     FMax: integer;
     FMin: integer;
     FPosition: integer;
     FTotalValues: integer;
     FState: TXScrollState;
     FHit,FLastHit: TXScrollHitTest;
     FThumbPos: integer;
     FThumbSize: integer;
     FThumbVisible: boolean;
     FSmallChange: integer;
     FLargeChange: integer;
     FClickDPos: integer;
     FLastX,FLastY: integer;
     FTimerActive: boolean;
     FChangeEvent: TNotifyEvent;
     FScrollEvent: TXLSScrollEvent;
     FEnabled: boolean;
     FExcelStyle: boolean;

     function PosXY(X,Y: integer): integer; {$ifdef D2006PLUS} inline; {$endif}
     function PosC1: integer; {$ifdef D2006PLUS} inline; {$endif}
     function PosC2: integer; {$ifdef D2006PLUS} inline; {$endif}

     procedure SetMax(const Value: integer);
     procedure SetMin(const Value: integer);
     procedure SetPosition(const Value: integer);
     function  HitTest(X,Y: integer): TXScrollHitTest;
     procedure SetKindHoriz(const Value: boolean);
     procedure SetEnabled(const Value: boolean);
protected
     procedure CalcThumbPos;
     procedure DoPosChange;
     procedure ScrollChanged; virtual;

     procedure PaintNormal;
     procedure PaintThemed;
public
     constructor Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
     destructor Destroy; override;
     procedure Paint; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;

     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseEnter(X, Y: Integer); override;
     procedure MouseLeave; override;
     procedure HandleTimer(ASender: TObject; AData: Pointer);

     // Width or height of the scroll bar as defined by the system.
     function  SysSize: integer;

     property MinVal: integer read FMin write SetMin;
     property MaxVal: integer read FMax write SetMax;
     property Position: integer read FPosition write SetPosition;
     property TotalValues: integer read FTotalValues write FTotalValues;
     property SmallChange: integer read FSmallChange write FSmallChange;
     property LargeChange: integer read FLargeChange write FLargeChange;
     property KindHoriz: boolean read FKindHoriz write SetKindHoriz;
     property Enabled: boolean read FEnabled write SetEnabled;
     property ExcelStyle: boolean read FExcelStyle write FExcelStyle;

     property OnChange: TNotifyEvent read FChangeEvent write FChangeEvent;
     property OnScroll: TXLSScrollEvent read FScrollEvent write FScrollEvent;
     end;

const XLSLISTBOX_TOPMARG = 4;

type TXLSListBox = class(TXLSControl)
protected
     FScroll: TXLSScrollBar;
//     FValues: TXLSWideStringList;
     FSelectedIndex: integer;
     FTopIndex: integer;
     FChangeEvent: TNotifyEvent;

     function  VisibleItems: integer;
     procedure ValuesChanged(Sender: TObject);
     procedure ScrollChanged(Sender: TObject);
     procedure SetSelected(Y: integer);
     procedure PaintLine(Index: integer);
     procedure SetTopIndex(Value: integer);
     procedure SetSelectedIndex(Value: integer);
     procedure ListChanged; virtual;
public
     constructor Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
     destructor Destroy; override;
     procedure Paint; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

//     property Values: TXLSWideStringList read FValues;

     property OnChange: TNotifyEvent read FChangeEvent write FChangeEvent;
     property SelectedIndex: integer read FSelectedIndex write FSelectedIndex;
     end;

// It's not possible to change the focused item in the list box by moving the
// mouse. In order to enable this, the control needs to keep track of both
// selected and focused item. Now, these are the same.
type TXLSComboBox = class(TXLSControl)
private
//     function  GetValues: TXLSWideStringList;
     function  GetSelectedIndex: integer;
     procedure SetSelectedIndex(const Value: integer);
protected
     FList: TXLSListBox;
     FValue: AxUCString;
     FY2CbUp: integer;
     FShowOnlyButton: boolean;
     FHighlight: boolean;

     procedure ListChanged(Sender: TObject); virtual;
     procedure HideList; virtual;
     procedure ShowList; virtual;
public
     constructor Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
     destructor Destroy; override;
     procedure Paint; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

     property Value: AxUCString read FValue write FValue;
//     property Values: TXLSWideStringList read GetValues;
     property SelectedIndex: integer read GetSelectedIndex write SetSelectedIndex;
     property Highlight: boolean read FHighlight write FHighlight;
     end;

const XLSRADIOBUTTONSIZE = 13;

type TXLSRadioButton = class(TXLSControl)
protected
     FChecked: boolean;
     FText: AxUCString;
     FChangeEvent: TNotifyEvent;
public
     constructor Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
     destructor Destroy; override;
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

     property Checked: boolean read FChecked write FChecked;
     property Text: AxUCString read FText write FText;
     end;

type TXLSCheckBoxState = (xcbsUnChecked,xcbsChecked,xcbsMixed);

type TXLSCheckBox = class(TXLSControl)
protected
     FState: TXLSCheckBoxState;
     FText: AxUCString;
public
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

     property Checked: TXLSCheckBoxState read FState write FState;
     property Text: AxUCString read FText write FText;
     end;

type TXLSButton = class(TXLSControl)
protected
     FText: AxUCString;
public
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

     property Text: AxUCString read FText write FText;
     end;

type TXLSLabel = class(TXLSControl)
protected
     FText: AxUCString;
public
     procedure Paint; override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;

     property Text: AxUCString read FText write FText;
     end;

implementation

{$ifndef DELPHI_2009_OR_LATER}
const ABS_UPNORMAL      = 1;
const ABS_UPHOT         = 2;
const ABS_UPPRESSED     = 3;
const ABS_UPDISABLED    = 4;
const ABS_DOWNNORMAL    = 5;
const ABS_DOWNHOT       = 6;
const ABS_DOWNPRESSED   = 7;
const ABS_DOWNDISABLED  = 8;
const ABS_LEFTNORMAL    = 9;
const ABS_LEFTHOT       = 10;
const ABS_LEFTPRESSED   = 11;
const ABS_LEFTDISABLED  = 12;
const ABS_RIGHTNORMAL   = 13;
const ABS_RIGHTHOT      = 14;
const ABS_RIGHTPRESSED  = 15;
const ABS_RIGHTDISABLED = 16;
const ABS_UPHOVER       = 17;
const ABS_DOWNHOVER     = 18;
const ABS_LEFTHOVER     = 19;
const ABS_RIGHTHOVER    = 20;
{$endif}

{ TXLSScrollBar }

{
The following code sets the thumb tab of a scrollbar so that it represents the
proportion of the scrolling range that is visible:

var

  TrackHeight: Integer; // the size of the scroll bar track
  MinHeight: Integer; // the default size of the thumb tab
begin
  MinHeight := GetSystemMetrics(SM_CYVTHUMB); // save default size
  with ScrollBar1 do
  begin
    TrackHeight := ClientHeight - 2 * GetSystemMetrics(SM_CYVSCROLL);
    PageSize := TrackHeight div (Max - Min + 1);
    if PageSize < MinHeight then PageSize := MinHeight;
  end;

end;
}
procedure TXLSScrollBar.CalcThumbPos;
var
  Sz: integer;
begin
  FThumbVisible := ((PosC2 - PosC1) > MIN_THUMBSIZE) and (FMax > FMin);
  if FThumbVisible then begin
    if FTotalValues > 0 then
      FThumbSize := Max(Round((PosC2 - PosC1) * ((FTotalValues - FMax) / FTotalValues)),MIN_THUMBSIZE)
    else
      FThumbSize := Windows.GetSystemMetrics(SM_CXVSCROLL);
    Sz := PosC2 - PosC1 - FThumbSize;
    FThumbPos := Round((Sz / (FMax - FMin)) * FPosition);
  end;
end;

constructor TXLSScrollBar.Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
begin
  inherited Create(AParent,AShape);

  FEnabled := True;
  FMin := 0;
  FMax := 100;
  FTotalValues := 0;
  FPosition := 0;
  FThumbPos := 0;
  FThumbSize := Windows.GetSystemMetrics(SM_CXVSCROLL);
  FSmallChange := 1;
  FLargeChange := 10;
  FCursor := xctHandPoint;
  SetSize(0,0,0,0);
  SetKindHoriz(False);

  if (Skin.Style >= xssExcel2007) and (FSkin.HandleHWND <> 0) then
    FHTheme := OpenThemeData(FSkin.HandleHWND,'SCROLLBAR');
end;

destructor TXLSScrollBar.Destroy;
begin
  if FHTheme <> 0 then
    CloseThemeData(FHTheme);
  inherited;
end;

procedure TXLSScrollBar.DoPosChange;
var
  P: integer;
begin
  if not FExcelStyle then begin
    P := FPosition;
    case FState of
      xssIdle: ;
      xssUpClicked:        FPosition := Max(FPosition - FSmallChange,FMin);
      xssDownClicked:      FPosition := Min(FPosition + FSmallChange,FMax);
      xssLargeUpClicked:   FPosition := Max(FPosition - FLargeChange,FMin);
      xssLargeDownClicked: FPosition := Min(FPosition + FLargeChange,FMax);
    end;
    if P <> FPosition then begin
      CalcThumbPos;
      Paint;
      if Assigned(FChangeEvent) then
        FChangeEvent(Self);
      ScrollChanged;
    end;
  end;
end;

function TXLSScrollBar.HitTest(X, Y: integer): TXScrollHitTest;
begin
  Result := xshtOutside;
  if Hit(X,Y) then begin
    if ClientHit(X,Y) then begin
      if PosXY(X,Y) < (PosC1 + FThumbPos) then
        Result := xshtLargeUp
      else if PosXY(X,Y) > (PosC1 + FThumbPos + FThumbSize) then
        Result := xshtLargeDown
      else
        Result := xshtThumb;
    end
    else begin
      if PosXY(X,Y) < PosC1 then
        Result := xshtUp
      else
        Result := xshtDown;
    end;
  end;
end;

// TODO Different for themed or not.
procedure TXLSScrollBar.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  // Result := inherited MouseDown(Button,Shift,X,Y);
  FSkin.GDI.SetCursor(FCursor);
  if not FEnabled then
    Exit;
  if FKindHoriz then
    FClickDPos := X - FCX1 - FThumbPos
  else
    FClickDPos := Y - FCY1 - FThumbPos;
  FHit := HitTest(X,Y);
  FState := TXScrollState(FHit);
  if Assigned(FScrollEvent) and (FState in [xssUpClicked,xssDownClicked,xssLargeUpClicked,xssLargeDownClicked]) then
    FScrollEvent(Self,FState);
  if FState > xssIdle then
    FSystem.StartTimer(Self,HandleTimer,XSCROLL_STARTTIMERTIME,Nil);
  Paint;
end;

procedure TXLSScrollBar.MouseEnter(X, Y: Integer);
begin
  inherited;
  FHit := HitTest(X,Y);
  Paint;
end;

procedure TXLSScrollBar.MouseLeave;
begin
  inherited;
  FHit := xshtOutside;
  FState := xssIdle;
  Paint;
end;

procedure TXLSScrollBar.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
var
  Sz,P: integer;
begin
  // Result := inherited MouseMove(Shift,X,Y);
  FSkin.GDI.SetCursor(FCursor);
  if not FEnabled then
    Exit;
  FLastX := X;
  FLastY := Y;
  FHit := HitTest(X,Y);
  if FHit <> FLastHit then
    Paint;
  FLastHit := FHit;
  P := FPosition;
  if FKindHoriz then begin
    if (FState = xssTrack) then begin
      Dec(X,FClickDPos);
      Sz := FCX2 - FCX1 - FThumbSize;
      FPosition := Round(((FMax - FMin) / Sz) * (X - FCX1));
      if FPosition < FMin then
        FPosition := FMin
      else if FPosition > FMax then
        FPosition := FMax;
    end;
  end
  else begin
    if (FState = xssTrack) then begin
      Dec(Y,FClickDPos);
      Sz := FCY2 - FCY1 - FThumbSize;
      FPosition := Round(((FMax - FMin) / Sz) * (Y - FCY1));
      if FPosition < FMin then
        FPosition := FMin
      else if FPosition > FMax then
        FPosition := FMax;
    end;
  end;
  if P <> FPosition then begin
    CalcThumbPos;
    Paint;
    if Assigned(FChangeEvent) then
      FChangeEvent(Self);
    ScrollChanged;
    if Assigned(FScrollEvent) and (FState = xssTrack) then
      FScrollEvent(Self,FState);
  end;
end;

procedure TXLSScrollBar.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  FSkin.GDI.SetCursor(FCursor);
  if not FEnabled then
    Exit;
  FSystem.StopTimer(Self);
  if not FTimerActive then
    DoPosChange;
  if (FState = xssTrack) and Assigned(FScrollEvent) then
     FScrollEvent(Self,xssReleaseTrack);
  FState := xssIdle;
  FTimerActive := False;
  Paint;
end;

procedure TXLSScrollBar.Paint;
begin
//  BeginPaint;
  if FHTheme <> 0 then
    PaintThemed
  else
    PaintNormal;
//  EndPaint;
end;

procedure TXLSScrollBar.PaintNormal;

procedure PaintBackground(x1,y1,x2,y2: integer; Clicked: boolean);
begin
  if Clicked then
    FSkin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_HIGHLIGHT))
  else
    FSkin.GDI.SelectPatternBrush(spbLightGray);
  FSkin.GDI.PenSolid := False;
  FSkin.GDI.Rectangle(x1 + 1,y1,x2,y2);
  FSkin.GDI.PenSolid := True;

  FSkin.GDI.PenColor := RevRGB(GetSysColor(COLOR_BTNSHADOW));
  if FKindHoriz then begin
    FSkin.GDI.Line(x1,y1,x2,y1);
    FSkin.GDI.Line(x1,y2,x2,y2);
  end
  else begin
    FSkin.GDI.Line(x1,y1,x1,y2);
    FSkin.GDI.Line(x2,y1,x2,y2);
  end;

  FSkin.GDI.SelectPatternBrush(spbNone);
end;

procedure PaintThumb(x1,y1,x2,y2: integer);
begin
  FSkin.GDI.DrawEdge(x1,y1,x2,y2,EDGE_RAISED,BF_TOPRIGHT + BF_BOTTOMLEFT);
  FSkin.GDI.PenSolid := False;
  FSkin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_BTNFACE));
  FSkin.GDI.Rectangle(x1 + 2,y1 + 2,x2 - 2,y2 - 2);
  FSkin.GDI.PenSolid := True;
end;

begin
  if FKindHoriz then begin
    if not FEnabled then begin
      FSkin.GDI.DrawControl(gctScrollButtonLeftDisabled,FX1, FY1, FX1 + FBtnSize, FY2);

      FSkin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_WINDOW));
      FSkin.GDI.PenColor := RevRGB(GetSysColor(COLOR_BTNSHADOW));
      FSkin.GDI.Rectangle(FX1 + FBtnSize,FY1,FX2 - FBtnSize - 1,FY2 - 1);

      FSkin.GDI.DrawControl(gctScrollButtonRightDisabled,FX2 - FBtnSize, FY1, FX2, FY2);
    end
    else begin
      if not FTimerActive then
        FSkin.GDI.DrawControl(gctScrollButtonLeft,FX1, FY1, FX1 + FBtnSize, FY2);

      if FThumbVisible then begin
        if FThumbPos > 0 then
          PaintBackground(FX1 + FBtnSize,FY1,FX1 + FBtnSize + FThumbPos,FY2 - 1,FState = xssLargeUpClicked);

        PaintThumb(FX1 + FBtnSize + FThumbPos,FY1,FX1 + FBtnSize + FThumbPos + FThumbSize,FY2);

        PaintBackground(FX1 + FBtnSize + FThumbPos + FThumbSize,FY1,FX2 - FBtnSize,FY2 - 1,FState = xssLargeDownClicked);
      end
      else
        PaintBackground(FCX1,FY1,FCX2,FY2,FState in [xssLargeUpClicked,xssLargeDownClicked]);


      if not FTimerActive then
        FSkin.GDI.DrawControl(gctScrollButtonRight,FX2 - FBtnSize, FY1, FX2, FY2);
    end;
  end
  else begin
    if not FEnabled then begin
      FSkin.GDI.DrawControl(gctScrollButtonUpDisabled,FX1, FY1, FX2, FY1 + FBtnSize);

      FSkin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_WINDOW));
      FSkin.GDI.PenColor := RevRGB(GetSysColor(COLOR_BTNSHADOW));
      FSkin.GDI.Rectangle(FX1,FY1 + FBtnSize,FX2 - 1,FY2 - FBtnSize - 1);

      FSkin.GDI.DrawControl(gctScrollButtonDownDisabled,FX1, FY2 - FBtnSize, FX2, FY2);
    end
    else begin
      if not FTimerActive then
        FSkin.GDI.DrawControl(gctScrollButtonUp,FX1, FY1, FX2, FY1 + FBtnSize);

      if FThumbVisible then begin
        if FThumbPos > 0 then
          PaintBackground(FX1,FY1 + FBtnSize,FX2 - 1,FY1 + FBtnSize + FThumbPos,FState = xssLargeUpClicked);

        PaintThumb(FX1,FY1 + FBtnSize + FThumbPos,FX2,FY1 + FBtnSize + FThumbPos + FThumbSize);

        PaintBackground(FX1,FY1 + FBtnSize + FThumbPos + FThumbSize,FX2 - 1,FY2 - FBtnSize,FState = xssLargeDownClicked);
      end
      else
        PaintBackground(FX1,FCY1,FX2 - 1,FCY2,FState in [xssLargeUpClicked,xssLargeDownClicked]);

      if not FTimerActive then
        FSkin.GDI.DrawControl(gctScrollButtonDown,FX1, FY2 - FBtnSize, FX2, FY2);
    end;
  end;
end;

procedure TXLSScrollBar.PaintThemed;

procedure PaintBackground(AClicked: boolean; x1,y1,x2,y2: integer);
begin
  if AClicked then
    FSkin.GDI.PaintColor := RevRGB(GetSysColor(COLOR_BTNSHADOW))
  else
    FSkin.GDI.PaintColor := RevRGB(GetSysColor(COLOR_BTNFACE));

  FSkin.GDI.Rectangle(x1,y1,x2,y2);

  FSkin.GDI.PenColor := RevRGB(GetSysColor(COLOR_MENU));
  if FKindHoriz then begin
//    FSkin.GDI.Line(x1,y1 + 1,x2,y1 + 1);
    FSkin.GDI.Line(x1,y2,x2,y2);
  end
  else begin
    FSkin.GDI.Line(x1,y1,x1,y2);
    FSkin.GDI.Line(x2,y1,x2,y2);
  end;

  FSkin.GDI.SelectPatternBrush(spbNone);
end;

procedure PaintThumb(AState: longword; x1,y1,x2,y2: integer; Horizontal: boolean);
begin
  if Horizontal then begin
    DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_THUMBBTNHORZ, AState, Rect(x1,y1,x2,y2), Nil);
    if (x2 - x1) > 15 then
      DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_GRIPPERHORZ, AState, Rect(x1,y1,x2,y2), Nil);
  end
  else begin
    DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_THUMBBTNVERT, AState, Rect(x1,y1,x2,y2), Nil);
    if (y2 - y1) > 15 then
      DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_GRIPPERVERT, AState, Rect(x1,y1,x2,y2), Nil);
  end;
end;

procedure PaintArrow(AState: longword; x1,y1,x2,y2: integer);
begin
  DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_ARROWBTN, AState, Rect(x1,y1,x2,y2), Nil);
end;

begin
  FSkin.GDI.InvalidateRect(FX1,FY1,FX2,FY2);
  if FKindHoriz then begin
    if not FEnabled then begin
      DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_ARROWBTN, ABS_LEFTDISABLED, Rect(FX1, FY1, FX1 + FBtnSize, FY2), Nil);
      DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_ARROWBTN, ABS_RIGHTDISABLED, Rect(FX2 - FBtnSize, FY1, FX2, FY2), Nil);
    end
    else begin
      PaintBackground(False,FX1 + FBtnSize,FY1,FX2 - FBtnSize,FY2);

//      FSkin.GDI.Rectangle(FX1, FY1, FX1 + FBtnSize, FY2);
      PaintArrow(ABS_LEFTNORMAL, FX1, FY1, FX1 + FBtnSize, FY2);
      PaintThumb(SCRBS_NORMAL,FX1 + FBtnSize + FThumbPos,FY1,FX1 + FBtnSize + FThumbPos + FThumbSize,FY2,True);
      PaintArrow(ABS_RIGHTNORMAL, FX2 - FBtnSize, FY1, FX2, FY2);
      case FHit of
        xshtOutside  : ;
        xshtUp       : begin
          if FState = xssUpClicked then
            PaintArrow(ABS_LEFTPRESSED, FX1, FY1, FX1 + FBtnSize, FY2)
          else
            PaintArrow(ABS_LEFTHOT, FX1, FY1, FX1 + FBtnSize, FY2);
        end;
        xshtDown     : begin
          if FState = xssDownClicked then
            PaintArrow(ABS_RIGHTPRESSED, FX2 - FBtnSize, FY1, FX2, FY2)
          else
            PaintArrow(ABS_RIGHTHOT, FX2 - FBtnSize, FY1, FX2, FY2);
        end;
        xshtLargeUp  : begin
          PaintArrow(ABS_LEFTHOVER, FX1, FY1, FX1 + FBtnSize, FY2);
          if FState = xssLargeUpClicked then
            PaintBackground(True,FX1 + FBtnSize,FY1,FX1 + FBtnSize + FThumbPos,FY2 - 1);
          PaintArrow(ABS_RIGHTHOVER, FX2 - FBtnSize, FY1, FX2, FY2);
        end;
        xshtLargeDown: begin
          PaintArrow(ABS_LEFTHOVER, FX1, FY1, FX1 + FBtnSize, FY2);
          if FState = xssLargeDownClicked then
            PaintBackground(True,FX1 + FBtnSize + FThumbPos + FThumbSize,FY1,FX2 - FBtnSize,FY2 - 1);
          PaintArrow(ABS_RIGHTHOVER, FX2 - FBtnSize, FY1, FX2, FY2);
        end;
        xshtThumb    : begin
          PaintArrow(ABS_LEFTHOVER, FX1, FY1, FX1 + FBtnSize, FY2);
          if FState = xssTrack then
            PaintThumb(SCRBS_PRESSED,FX1 + FBtnSize + FThumbPos,FY1,FX1 + FBtnSize + FThumbPos + FThumbSize,FY2,True)
          else
            PaintThumb(SCRBS_HOT,FX1 + FBtnSize + FThumbPos,FY1,FX1 + FBtnSize + FThumbPos + FThumbSize,FY2,True);
          PaintArrow(ABS_RIGHTHOVER, FX2 - FBtnSize, FY1, FX2, FY2);
        end;
      end;
    end;
  end
  else begin
    if not FEnabled then begin
      DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_ARROWBTN, ABS_LEFTDISABLED, Rect(FX1, FY1, FX1 + FBtnSize, FY2), Nil);
      DrawThemeBackground(FHTheme,FSkin.GDI.DC, SBP_ARROWBTN, ABS_RIGHTDISABLED, Rect(FX2 - FBtnSize, FY1, FX2, FY2), Nil);
    end
    else begin
      PaintBackground(False,FX1 + FBtnSize,FY1,FX2 - FBtnSize,FY2 - 1);

      PaintArrow(ABS_UPNORMAL, FX1, FY1, FX2, FY1 + FBtnSize);
      PaintThumb(SCRBS_NORMAL,FX1,FY1 + FBtnSize + FThumbPos,FX2,FY1 + FBtnSize + FThumbPos + FThumbSize,False);
      PaintArrow(ABS_DOWNNORMAL, FX1, FY2 - FBtnSize, FX2, FY2);
      case FHit of
        xshtOutside  : ;
        xshtUp       : begin
          if FState = xssUpClicked then
            PaintArrow(ABS_UPPRESSED, FX1, FY1, FX2, FY1 + FBtnSize)
          else
            PaintArrow(ABS_UPHOT, FX1, FY1, FX2, FY1 + FBtnSize);
        end;
        xshtDown     : begin
          if FState = xssDownClicked then
            PaintArrow(ABS_DOWNPRESSED, FX1, FY2 - FBtnSize, FX2, FY2)
          else
            PaintArrow(ABS_DOWNHOT, FX1, FY2 - FBtnSize, FX2, FY2);
        end;
        xshtLargeUp  : begin
          PaintArrow(ABS_UPHOVER, FX1, FY1, FX2, FY1 + FBtnSize);
          if FState = xssLargeUpClicked then
            PaintBackground(True,FX1,FY1 + FBtnSize,FX2,FY1 + FBtnSize + FThumbPos);
          PaintArrow(ABS_DOWNHOVER, FX1, FY2 - FBtnSize, FX2, FY2);
        end;
        xshtLargeDown: begin
          PaintArrow(ABS_UPHOVER, FX1, FY1, FX2, FY1 + FBtnSize);
          if FState = xssLargeDownClicked then
            PaintBackground(True,FX1,FY1 + FBtnSize + FThumbPos + FThumbSize,FX2,FY2 - FBtnSize);
          PaintArrow(ABS_DOWNHOVER, FX1, FY2 - FBtnSize, FX2, FY2);
        end;
        xshtThumb    : begin
          PaintArrow(ABS_UPHOVER, FX1, FY1, FX2, FY1 + FBtnSize);
          if FState = xssTrack then
            PaintThumb(SCRBS_PRESSED,FX1,FY1 + FBtnSize + FThumbPos,FX2,FY1 + FBtnSize + FThumbPos + FThumbSize,False)
          else
            PaintThumb(SCRBS_HOT,FX1,FY1 + FBtnSize + FThumbPos,FX2,FY1 + FBtnSize + FThumbPos + FThumbSize,False);
          PaintArrow(ABS_DOWNHOVER, FX1, FY2 - FBtnSize, FX2, FY2);
        end;
      end;
    end;
  end;
end;

function TXLSScrollBar.PosC1: integer;
begin
  if FKindHoriz then
    Result := FCX1
  else
    Result := FCY1;
end;

function TXLSScrollBar.PosC2: integer;
begin
  if FKindHoriz then
    Result := FCX2
  else
    Result := FCY2;
end;

function TXLSScrollBar.PosXY(X, Y: integer): integer;
begin
  if FKindHoriz then
    Result := X
  else
    Result := Y;
end;

procedure TXLSScrollBar.ScrollChanged;
begin

end;

procedure TXLSScrollBar.SetEnabled(const Value: boolean);
begin
  FEnabled := Value;
end;

procedure TXLSScrollBar.SetKindHoriz(const Value: boolean);
begin
  FKindHoriz := Value;
  if FKindHoriz then begin
    FThumbSize := Windows.GetSystemMetrics(SM_CYHSCROLL);
    SetSize(FX1,FY1,FX2,FY1 + FThumbSize);
  end
  else begin
    FThumbSize := Windows.GetSystemMetrics(SM_CXVSCROLL);
    SetSize(FX1,FY1,FX1 + FThumbSize,FY2);
  end;
end;

procedure TXLSScrollBar.SetMax(const Value: integer);
begin
  FMax := Value;
  FEnabled := FMax > FMin;
  CalcThumbPos;
end;

procedure TXLSScrollBar.SetMin(const Value: integer);
begin
  FMin := Value;
  FEnabled := FMax > FMin;
  CalcThumbPos;
end;

procedure TXLSScrollBar.SetPosition(const Value: integer);
begin
  FPosition := Value;
  CalcThumbPos;
end;

procedure TXLSScrollBar.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited SetSize(pX1,pY1,pX2,pY2);

  if FKindHoriz then begin
    FCY2 := FCY1 + Windows.GetSystemMetrics(SM_CYHSCROLL);
    FBtnSize := FY2 - FY1;
    SetClientSize(FX1 + FBtnSize,FY1,FX2 - FBtnSize,FY2);
  end
  else begin
    FCX2 := FCX1 + Windows.GetSystemMetrics(SM_CXVSCROLL);
    FBtnSize := FX2 - FX1;
    SetClientSize(FX1,FY1 + FBtnSize,FX2,FY2 - FBtnSize);
  end;
  CalcThumbPos;
end;

function TXLSScrollBar.SysSize: integer;
begin
  if FKindHoriz then
    Result := Windows.GetSystemMetrics(SM_CYHSCROLL)
  else
    Result := Windows.GetSystemMetrics(SM_CXVSCROLL);
end;

procedure TXLSScrollBar.HandleTimer(ASender: TObject; AData: Pointer);
begin
  FSystem.StopTimer(Self);
  if Assigned(FScrollEvent) and (FState in [xssUpClicked,xssDownClicked,xssLargeUpClicked,xssLargeDownClicked]) then begin
    FScrollEvent(Self,FState);
    Paint;
    FSystem.StartTimer(Self,HandleTimer,XSCROLL_TIMERTIME,Nil);
  end;
end;

{ TXLSListBox }

constructor TXLSListBox.Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
begin
  inherited Create(AParent,AShape);
//  FValues := TXLSWideStringList.Create;
//  FValues.OnChange := ValuesChanged;
  FScroll := TXLSScrollBar.Create(Self,AShape);
  FScroll.OnChange := ScrollChanged;
  FScroll.OnSetCtrlFont := SetCtrlFont;
  FScroll.FCursor := xctHandPoint;
//   FScroll.Visible := False;
  Add(FScroll);
end;

destructor TXLSListBox.Destroy;
begin
//  FValues.OnChange := Nil;
//  FValues.Free;
  inherited;
end;

procedure TXLSListBox.ListChanged;
begin

end;

procedure TXLSListBox.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
//  inherited;
  FSkin.GDI.SetCursor(FCursor);
  SetSelected(Y);
end;

procedure TXLSListBox.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
  if xssLeft in Shift then
    SetSelected(Y);
end;

procedure TXLSListBox.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
//  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSListBox.Paint;
//var
//  i: integer;
//  y,TH: integer;
begin
//  inherited;
  Skin.GDI.PenColor := $000000;
  Skin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_WINDOW));;
  Skin.GDI.PenSolid := False;
  Skin.GDI.Rectangle(FCX1,FCY1,FCX2 + 1,FCY2);
  Skin.GDI.PenSolid := True;
  Skin.GDI.BrushSolid := False;
  Skin.GDI.Rectangle(FX1,FY1,FX2,FY2);
  Skin.GDI.BrushSolid := True;

//  Skin.SetTabFont(False);
//  FSkin.GDI.GetTextMetricSys;
  SetCtrlFont;
  Skin.GDI.SetTextAlign(xhtaLeft,xvtaTop);
//  TH := FSkin.GDI.TM_CELL.tmHeight;
//
//  y := FCY1 + XLSLISTBOX_TOPMARG;
//  for i := FTopIndex to FValues.Count - 1 do begin
//    if i = FSelectedIndex then begin
//      FSkin.GDI.FontColor := RevRGB(GetSysColor(COLOR_WINDOW));
//      FSkin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_HIGHLIGHT));
//      FSkin.GDI.Rectangle(FCX1,y,FCX2,y + TH);
//      FSkin.GDI.TextOut(FCX1 + XLSLISTBOX_TOPMARG,y,Rect(FCX1,y,FCX2,y + TH),FValues[i]);
//      FSkin.GDI.FontColor := RevRGB(GetSysColor(COLOR_WINDOWTEXT));
//    end
//    else
//      FSkin.GDI.TextOut(FCX1 + XLSLISTBOX_TOPMARG,y,Rect(FCX1,y,FCX2,y + TH),FValues[i]);
//    Inc(y,TH);
//    if (y + TH) > FCY2 then
//      Break;
//  end;
end;

procedure TXLSListBox.PaintLine(Index: integer);
//var
//  y,TH: integer;
begin
  if (Index < FTopIndex) or (Index >= (FTopIndex + VisibleItems)) then
    Exit;
//  SetCtrlFont;
//  TH := FSkin.GDI.TM_CELL.tmHeight;
//  y := FCY1 + XLSLISTBOX_TOPMARG + (TH * (Index - FTopIndex));
//  if Index = FSelectedIndex then begin
//    FSkin.GDI.FontColor := RevRGB(GetSysColor(COLOR_WINDOW));
//    FSkin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_HIGHLIGHT));
//    FSkin.GDI.Rectangle(FCX1,y,FCX2,y + TH);
//    FSkin.GDI.TextOut(FCX1 + XLSLISTBOX_TOPMARG,y,Rect(FCX1,y,FCX2,y + TH),FValues[Index]);
//    FSkin.GDI.FontColor := RevRGB(GetSysColor(COLOR_WINDOWTEXT));
//  end
//  else
//    FSkin.GDI.TextOut(FCX1 + XLSLISTBOX_TOPMARG,y,Rect(FCX1,y,FCX2,FCY1 + y + TH),FValues[Index]);
end;

procedure TXLSListBox.ScrollChanged(Sender: TObject);
begin
  SetTopIndex(FScroll.Position);
  Paint;
end;

procedure TXLSListBox.SetSelected(Y: integer);
//var
//  i: integer;
//  j: integer;
begin
//  SetCtrlFont;
//  i := (Y - FY1 - XLSLISTBOX_TOPMARG) div FSkin.GDI.TM_CELL.tmHeight;
//  Inc(i,FTopIndex);
//  if i < 0 then
//    i := 0;
//  if (i <= (FValues.Count - 1)) and (i <> FSelectedIndex) and (i >= FTopIndex) and (i < FTopIndex + VisibleItems) then begin
//    j := FSelectedIndex;
//    FSelectedIndex := i;
//    PaintLine(j);
//    PaintLine(i);
//    Paint;
//    SetSelectedIndex(i);
//  end;
end;

procedure TXLSListBox.SetSelectedIndex(Value: integer);
begin
  FSelectedIndex := Value;
  if Assigned(FChangeEvent) then
    FChangeEvent(Self);
  ListChanged;
end;

procedure TXLSListBox.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited;
  FScroll.SetSize(-1,pY1 + 1,pX2,pY2);
{
  FScroll.FTotalValues := 0;
  FScroll.MaxVal := 0;
}
  SetClientSize(pX1,pY1,FScroll.FX1 - 1,pY2);
  ValuesChanged(Self);
  FScroll.SetSize(-1,pY1 + 1,pX2,pY2);
end;

procedure TXLSListBox.SetTopIndex(Value: integer);
begin
  FTopIndex := Value;
end;

procedure TXLSListBox.ValuesChanged(Sender: TObject);
begin
//  FScroll.FTotalValues := FValues.Count;
//  FScroll.MaxVal := FValues.Count - VisibleItems;
//  FScroll.SetSize(-1,FY1 + 1,FX2,FY2);
end;

function TXLSListBox.VisibleItems: integer;
begin
//  Skin.SetTabFont(False);
//  FSkin.GDI.GetTextMetricSys;
  SetCtrlFont;
  Result := (FY2 - FY1 - XLSLISTBOX_TOPMARG) div FSkin.GDI.TM.Height;
end;

{ TXLSComboBox }

constructor TXLSComboBox.Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
begin
  inherited Create(AParent,AShape);
  FList := TXLSListBox.Create(Self,Nil);
  FList.OnChange := ListChanged;
  FList.Visible := False;
  FList.OnSetCtrlFont := SetCtrlFont;
  FList.FCursor := xctHandPoint;
  Add(FList);
end;

destructor TXLSComboBox.Destroy;
begin

  inherited;
end;

function TXLSComboBox.GetSelectedIndex: integer;
begin
  Result := FList.FSelectedIndex;
end;

//function TXLSComboBox.GetValues: TXLSWideStringList;
//begin
//  Result := FList.FValues;
//end;

procedure TXLSComboBox.HideList;
begin

end;

procedure TXLSComboBox.ListChanged(Sender: TObject);
begin
//  if FList.FSelectedIndex >= 0 then
//    FValue := FList.FValues[FList.FSelectedIndex]
//  else
//    FValue := '';
  FList.Visible := False;
  Paint;
  HideList;
end;

procedure TXLSComboBox.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
  FList.Visible := not FList.Visible;
  if FList.Visible then
    ShowList;
  FList.Repaint;
  Paint;
  if not FList.Visible then
    HideList;
end;

procedure TXLSComboBox.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSComboBox.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSComboBox.Paint;
var
  H: integer;
  TH: integer;
  CBPaintStyle: TGDICtrlType;
begin
//  inherited;

  H := FCY2 - FCY1 - 2;
  if not FShowOnlyButton then begin
    Skin.GDI.PenColor := $000000;
    Skin.GDI.BrushColor := RevRGB(GetSysColor(COLOR_WINDOW));;
    Skin.GDI.Rectangle(FCX1,FCY1,FCX2,FCY2);

    SetCtrlFont;
    FSkin.GDI.SetTextAlign(xhtaLeft,xvtaTop);
    TH := FSkin.GDI.TM.Height;
    FSkin.GDI.TextOut(FCX1 + XLSLISTBOX_TOPMARG,FCY1 + ((FCY2 - FCY1) div 2) - (TH div 2),Rect(FCX1 + 1,FCY1 + 1,FCX2 - H,FCY2 - 1),FValue);
  end;

  if FHighlight then begin
    if FList.Visible then
      CBPaintStyle := gctComboHighlightClicked
    else
      CBPaintStyle := gctComboHighlight;
  end
  else begin
    if FList.Visible then
      CBPaintStyle := gctComboClicked
    else
      CBPaintStyle := gctCombo;
  end;
  Skin.GDI.DrawControl(CBPaintStyle,FCX2 - H,FCY1 + 1,FCX2,FCY2);
end;

procedure TXLSComboBox.SetSelectedIndex(const Value: integer);
begin
  Flist.FSelectedIndex := Value;
end;

procedure TXLSComboBox.SetSize(const pX1, pY1, pX2, pY2: integer);
var
  ppX1,ppY1,ppX2,ppY2,W: integer;
begin
  ppX1 := pX1;
  ppY1 := pY1;
  ppX2 := pX2;
  ppY2 := pY2;

  if pY2 = -1 then
    ppY2 := pY1 + Windows.GetSystemMetrics(SM_CYVSCROLL) + 4;

  if FShowOnlyButton then begin
    ppX1 := pX2 - (Windows.GetSystemMetrics(SM_CYVSCROLL) + 2);
    ppY1 := PY2 - (Windows.GetSystemMetrics(SM_CYVSCROLL) + 2);
  end;

  inherited  SetSize(ppX1,ppY1,ppX2,ppY2);
  SetCtrlFont;
  W := FSkin.GDI.StdCharWidth * 13 + Windows.GetSystemMetrics(SM_CXVSCROLL);
  if (pX2 - pX1) < W then
    FList.SetSize(pX2 - W,pY2,pX2 - 1,pY2 + 100)
  else
    FList.SetSize(pX1,pY2,pX2 - 1,pY2 + 100);
  FY2CbUp := FY2;
end;

procedure TXLSComboBox.ShowList;
begin

end;

{ TXLSRadioButton }

constructor TXLSRadioButton.Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
begin
  inherited Create(AParent,AShape);
end;

destructor TXLSRadioButton.Destroy;
begin

  inherited;
end;

procedure TXLSRadioButton.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSRadioButton.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSRadioButton.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
  FChecked := True;
  Paint;
  if Assigned(FChangeEvent) then
    FChangeEvent(Self);
end;

procedure TXLSRadioButton.Paint;
var
  y: integer;
begin
//  inherited;
  y := FY1 + ((FY2 - FY1) div 2) - (XLSRADIOBUTTONSIZE div 2);
  if FChecked then
    FSkin.GDI.DrawControl(gctRadioButtonChecked,FX1, y, FX1 + XLSRADIOBUTTONSIZE, y + XLSRADIOBUTTONSIZE)
  else
    FSkin.GDI.DrawControl(gctRadioButton,FX1, y, FX1 + XLSRADIOBUTTONSIZE , y + XLSRADIOBUTTONSIZE);
  FSkin.GDI.CenterTextVert(FX1 + XLSRADIOBUTTONSIZE + 4,FY1,FX2,FY2,FText);
end;

{ TXLSControl }

procedure TXLSControl.SetCtrlFont;
begin
  if Assigned(FSetCtrlFontEvent) then
    FSetCtrlFontEvent;
end;

{ TXLSCheckBox }

procedure TXLSCheckBox.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSCheckBox.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSCheckBox.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
  if FState in [xcbsChecked,xcbsMixed] then
    FState := xcbsUnChecked
  else
    FState := xcbsChecked;
  Paint;
end;

procedure TXLSCheckBox.Paint;
begin
//  inherited;

end;

{ TXLSButton }

procedure TXLSButton.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSButton.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSButton.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSButton.Paint;
begin
//  inherited;
  SetCtrlFont;
  FSkin.GDI.DrawControl(gctButton,FX1,FY1,FX2,FY2);
  FSkin.GDI.BrushSolid := False;
  FSkin.GDI.CenterTextRect(FX1 + 2,FY1 + 2,FX2 - 2,FY2 - 2,FText);
  FSkin.GDI.BrushSolid := True;
end;

{ TXLSLabel }

procedure TXLSLabel.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSLabel.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSLabel.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited;
  FSkin.GDI.SetCursor(FCursor);
end;

procedure TXLSLabel.Paint;
begin
//  inherited;
  SetCtrlFont;
  FSkin.GDI.BrushSolid := False;
  FSkin.GDI.SetTextAlign(xhtaLeft,xvtaTop);
  FSkin.GDI.TextOut(FX1,FY1,Rect(FX1,FY1,FX2,FY2),FText);
  FSkin.GDI.BrushSolid := True;
end;

{ TXBookClientShape }

constructor TXBookClientShape.Create(AParent: TXSSWindow; AShape: TXLSShapeClientAnchor);
begin
  inherited Create(AParent);
  FShape := AShape;
end;

procedure TXBookClientShape.Paint;
begin
//  FParams.SetEscherOptions(FSkin.GDI,FShape);
end;

end.
