unit XBookSkin2;

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

uses { Delphi  } Classes, SysUtils, Windows, Types, vcl.Controls, Contnrs, vcl.Graphics,
     { XLSRWII } Xc12Utils5, Xc12DataStyleSheet5,
                 XLSUtils5,
     { XBook   } XBookUtils2, XBookPaint2, XBookOptions2, XBookHintWindow2,
                 XBookPaintGDI2;

const BOOK_BORDERWIDTH        = 1;
const GRIDLINE_SIZE           = 1;
const SPLITTER_SIZE           = 8;
const TABSLANT                = 7;

const GUTTER_SIZE             = 15;

const SEL_CELL_MARGIN         = 1;

const MARG_CELLCURSORHIT      = 3;
const MARG_CELLCURSORSIZEHIT  = 6;

type PXSSSkinButtonStateColors = ^TXSSSkinButtonStateColors;
     TXSSSkinButtonStateColors = record
     Cl1   : longword;
     Cl2   : longword;
     Sel1  : longword;
     Sel2  : longword;
     Focus1: longword;
     Focus2: longword;
     Font  : longword;
     Bold  : boolean;
     end;

type PXSSSkinColors = ^TXSSSkinColors;
     TXSSSkinColors = record
     WindowBkg: longword;
     SheetBkg: longword;
     SheetSel: longword;

     Header: TXSSSkinButtonStateColors;
     Row: TXSSSkinButtonStateColors;
     Tab: TXSSSkinButtonStateColors;

     HasHoover: boolean;
     HeaderHoover: TXSSSkinButtonStateColors;
     RowHoover: TXSSSkinButtonStateColors;
     TabHoover: TXSSSkinButtonStateColors;

     HeaderLine: longword;
     HeaderSelLine: longword;
     HeaderFocusLine: longword;
     HeaderFrozenLine: longword;
     HeaderAutofiltFnt: longword;
     HeaderAutofiltFntSel: longword;
     HeaderLineWidth: integer;
     SheetGridline: longword;

     TabColor: longword;
     TabColorSel: longword;
     TabFontColor: longword;
     TabFontColorSel: longword;

     GutterColor: longword;
     GutterLineColor: longword;
     end;

type TXBookSkinStyle = (xssExcelXP,xssExcel2007,xssExcel2013,xssExcelNone);
type TXBookSkinHeaderStyle = (xshsNormal,xshsRow,xshsGutter);

// **** SHAPES ****
const NOTE_SHADOWSIZE = 2;

const COLOR_SHAPESHADOW = $7F7F7F;

type THeaderSelState = (hssNormal,hssFocused,hssSelected);
type TTabState       = (tsNormal,tsFocused,tsSelected);
type TCursorSizePos  = (cspCell,cspCellReadOnly,cspColumn,cspRow,cspNone);

type TStraightLine = record
     A1,A2,B: integer;
     Horiz: boolean;
     end;

type TStraightLineArray = array of TStraightLine;

type TXLSBookSkin = class(TXLSBookPaint)
private
     FStyle     : TXBookSkinStyle;
     FColors    : PXSSSkinColors;
     FColorsXP  : TXSSSkinColors;
     FColors2007: TXSSSkinColors;
     FColors2013: TXSSSkinColors;
     FHandleHWND: longword;
     FCX1,FCY1,FCX2,FCY2: integer;

     function  GetHeaderMargin: integer;
     procedure SetupColors;
     procedure SetStyle(const Value: TXBookSkinStyle);
protected
public
     // HandleHWND required by themed controls.
     constructor Create(HandleHWND: longword; AGDI: TAXWGDI);
     destructor Destroy; override;

     class procedure LoadResources;
     class procedure FreeResources;

     procedure PaintHeader(x1, y1, x2, y2: integer; State: THeaderSelState; HeaderStyle: TXBookSkinHeaderStyle = xshsNormal; Hoover: boolean = False; Text: AxUCString = ''; Filtered: boolean = False);
     procedure Paint3dFrame(x1, y1, x2, y2,Width: integer; Raised: boolean = False);
     procedure PaintSheetBkg(x1, y1, x2, y2: integer);
     procedure PaintSplitterHandle(x1, y1, x2, y2: integer; Vertical,FSmall: boolean);
     procedure PaintTab(x1,x2,y: integer; Text: AxUCString; State: TTabState; Color: longword);
     procedure PaintCursor(x1, y1, x2, y2: integer; SizePos: TCursorSizePos);
     procedure PaintFocusCell(x1, y1, x2, y2: integer);
     procedure PaintMoveRect(Lines: TStraightLineArray);
     procedure PaintDeleteRect(x1,y1,x2,y2: integer);

     procedure GetBmpBtnSize(ButtonType: TBmpButtonType; var H, W: integer);
     procedure PaintBmpButton(ButtonType: TBmpButtonType; ButtonState: TBmpButtonState; Index, X, Y: integer);
     procedure PaintStockBMP(StockBMP: TStockXBookBitmap; X, Y: integer);
     procedure PaintTabMoveBmp(X, Y: integer);

     procedure SetClientSize(x1, y1, x2, y2: integer);

     procedure SetFont(AFont: TXc12Font);

     procedure Test;

     property HeaderMargin: integer read GetHeaderMargin;

     property HandleHWND: longword read FHandleHWND;

     property CX1: integer read FCX1;
     property CY1: integer read FCY1;
     property CX2: integer read FCX2;
     property CY2: integer read FCY2;

     property Style: TXBookSkinStyle read FStyle write SetStyle;

     property Colors: PXSSSkinColors read FColors;
     end;

implementation

var
  bmpBtnTabSet : vcl.Graphics.TBitmap;
  bmpBtnOutline: vcl.Graphics.TBitmap;
  bmpBtnNumbers: vcl.Graphics.TBitmap;

  bmpTabMove   : vcl.Graphics.TBitmap;
  bmpNote      : vcl.Graphics.TBitmap;
  bmpCheckBox  : vcl.Graphics.TBitmap;
  bmpHSplit    : vcl.Graphics.TBitmap;
  bmpVSplit    : vcl.Graphics.TBitmap;

{ TXLSBookSkin }

constructor TXLSBookSkin.Create(HandleHWND: longword; AGDI: TAXWGDI);
begin
  inherited Create(AGDI);
  FHandleHWND := HandleHWND;

  SetupColors;
  SetStyle(xssExcel2013);
end;

destructor TXLSBookSkin.Destroy;
begin
  inherited;
end;

class procedure TXLSBookSkin.FreeResources;
begin
  FreeAndNil(bmpBtnTabSet);
  FreeAndNil(bmpBtnOutline);
  FreeAndNil(bmpBtnNumbers);
  FreeAndNil(bmpTabMove);
  FreeAndNil(bmpNote);
  FreeAndNil(bmpCheckBox);
  FreeAndNil(bmpHSplit);
  FreeAndNil(bmpVSplit);
end;

procedure TXLSBookSkin.GetBmpBtnSize(ButtonType: TBmpButtonType; var H, W: integer);
begin
  W := 0;
  H := 0;
  case ButtonType of
    bbtTabset: begin
      W := bmpBtnTabSet.Width div 4;
      H := bmpBtnTabSet.Height div 3;
    end;
    bbtOutline: begin
      W := bmpBtnOutline.Width div 2;
      H := bmpBtnOutline.Height div 3;
    end;
    bbtNumbers: begin
      W := bmpBtnNumbers.Width div 8;
      H := bmpBtnNumbers.Height div 3;
    end;
  end;
end;

function TXLSBookSkin.GetHeaderMargin: integer;
begin
  Result := GDI.TM_SYS.Descent;
end;

class procedure TXLSBookSkin.LoadResources;
begin
  bmpBtnTabSet := vcl.Graphics.TBitmap.Create;
  bmpBtnTabSet.LoadFromResourceName(HInstance,'TABBUTTONS');

  bmpBtnOutline := vcl.Graphics.TBitmap.Create;
  bmpBtnOutline.LoadFromResourceName(HInstance,'OUTLINEBTNS');

  bmpBtnNumbers := vcl.Graphics.TBitmap.Create;
  bmpBtnNumbers.LoadFromResourceName(HInstance,'NUMBUTTONS');

  bmpTabMove := vcl.Graphics.TBitmap.Create;
  bmpTabMove.LoadFromResourceName(HInstance,'TABMOVE2');

  bmpNote := vcl.Graphics.TBitmap.Create;
  bmpNote.LoadFromResourceName(HInstance,'BMPNOTE');

  bmpCheckBox := vcl.Graphics.TBitmap.Create;
  bmpCheckBox.LoadFromResourceName(HInstance,'BMPCHECKBOX');

  bmpHSplit := vcl.Graphics.TBitmap.Create;
  bmpHSplit.LoadFromResourceName(HInstance,'BMPHSPLIT');

  bmpVSplit := vcl.Graphics.TBitmap.Create;
  bmpVSplit.LoadFromResourceName(HInstance,'BMPVSPLIT');
end;

procedure TXLSBookSkin.Paint3dFrame(x1, y1, x2, y2, Width: integer; Raised: boolean);
begin
  FGDI.BrushColor := $000080;
  FGDI.PenColor := $000080;
  Dec(Width);
  FGDI.Rectangle(x1,y1,x2,x1 + Width);
  FGDI.Rectangle(x1,y1,x1 + Width,y2);
  FGDI.Rectangle(x2 - Width,y1,x2,y2);
  FGDI.Rectangle(x1,y2 - Width,x2,y2);
end;

procedure TXLSBookSkin.PaintBmpButton(ButtonType: TBmpButtonType; ButtonState: TBmpButtonState; Index, X, Y: integer);
var
  W,H: integer;
  rectSrc,rectDest: TRect;
begin
  GetBmpBtnSize(ButtonType,H,W);
  rectSrc.Left    := Index * W;
  rectSrc.Top     := Ord(ButtonState) * H;
  rectSrc.Right   := rectSrc.Left + W;
  rectSrc.Bottom  := rectSrc.Top + H;

  rectDest.Left   := 0;
  rectDest.Top    := 0;
  rectDest.Right  := W;
  rectDest.Bottom := H;
  OffsetRect(rectDest,X,Y);
  case ButtonType of
    bbtTabset : FGDI.BmpBlt(bmpBtnTabset.Canvas.Handle,rectDest.Left,rectDest.Top,W,H,Index * W,Ord(ButtonState) * H,SRCCOPY);
    bbtOutline: FGDI.BmpBlt(bmpBtnOutline.Canvas.Handle,rectDest.Left,rectDest.Top,W,H,Index * W,Ord(ButtonState) * H,SRCCOPY);
    bbtNumbers: FGDI.BmpBlt(bmpBtnNumbers.Canvas.Handle,rectDest.Left,rectDest.Top,W,H,Index * W,Ord(ButtonState) * H,SRCCOPY);
  end;
end;

procedure TXLSBookSkin.PaintHeader(x1, y1, x2, y2: integer; State: THeaderSelState; HeaderStyle: TXBookSkinHeaderStyle = xshsNormal; Hoover: boolean = False; Text: AxUCString = ''; Filtered: boolean = False);
var
  StatColors: PXSSSkinButtonStateColors;
  Cl1,Cl2: longword;
begin
  case HeaderStyle of
    xshsNormal: begin
      if FColors.HasHoover and Hoover then
        StatColors := @FColors.HeaderHoover
      else
        StatColors := @FColors.Header;
    end;
    xshsRow   : begin
      if FColors.HasHoover and Hoover then
        StatColors := @FColors.RowHoover
      else
        StatColors := @FColors.Row;
    end;
    xshsGutter   : begin
      FGDI.BrushColor := FColors.GutterColor;
      FGDI.PenColor := FColors.GutterLineColor;
      FGDI.Rectangle(x1, y1, x2 - GRIDLINE_SIZE, y2 - GRIDLINE_SIZE);
      Exit;
    end
    else        StatColors := @FColors.Header;
  end;

  case State of
    hssNormal:   begin
      if Filtered then
        FGDI.FontColor := FColors.HeaderAutofiltFnt;
      Cl1 := StatColors.Cl1;
      Cl2 := StatColors.Cl2;
    end;
    hssFocused:  begin
      if Filtered then
        FGDI.FontColor := FColors.HeaderAutofiltFnt;
      Cl1 := StatColors.Focus1;
      Cl2 := StatColors.Focus2;
    end;
    hssSelected: begin
      if Filtered then
        FGDI.FontColor := FColors.HeaderAutofiltFntSel;
      Cl1 := StatColors.Sel1;
      Cl2 := StatColors.Sel2;
    end;
    else begin
      Cl1 := StatColors.Cl1;
      Cl2 := StatColors.Cl2;
    end;
  end;

  if Cl1 = Cl2 then begin
    FGDI.PaintColor := Cl1;
    FGDI.Rectangle(x1, y1, x2 - GRIDLINE_SIZE, y2 - GRIDLINE_SIZE);
  end
  else
    FGDI.GradientFillRect(x1, y1, x2 - GRIDLINE_SIZE, y2 - GRIDLINE_SIZE,Cl1,Cl2,False);

  if State > hssNormal then begin
    FGDI.FontColor := StatColors.Font;
    FGDI.SetFontStyle(StatColors.Bold,False,False);
  end
  else begin
    FGDI.FontColor := $000000;
    FGDI.SetFontStyle(False,False,False);
  end;

  if Text <> '' then
    FGDI.CenterTextRect(x1,y1,x2,y2,y2 - GetHeaderMargin - GRIDLINE_SIZE + 1,Text);

  case State of
    hssNormal:   FGDI.PenColor := FColors.HeaderLine;
    hssFocused:  begin
      FGDI.PenColor := FColors.HeaderFocusLine;
      FGDI.PenWidth := FColors.HeaderLineWidth;
    end;
    hssSelected: begin
      FGDI.PenColor := FColors.HeaderSelLine;
      FGDI.PenWidth := FColors.HeaderLineWidth;
    end;
  end;
  if GRIDLINE_SIZE = 1 then begin
    if HeaderStyle = xshsNormal then begin
      FGDI.Line(x1,y2,x2 + 1,y2);
      FGDI.PenWidth := 1;
      FGDI.Line(x2,y1,x2,y2);
    end
    else begin
      FGDI.Line(x2,y1,x2,y2);
      FGDI.PenWidth := 1;
      FGDI.Line(x1,y2,x2 + 1,y2);
    end;
  end
  else begin
    FGDI.BrushColor := FColors.HeaderLine;
    FGDI.Rectangle(x1,y2 - GRIDLINE_SIZE + 1,x2,y2);
    FGDI.Rectangle(x2 - GRIDLINE_SIZE + 1,y1,x2,y2);
  end;

  FGDI.PenWidth := 1;
end;

procedure TXLSBookSkin.PaintCursor(x1, y1, x2, y2: integer; SizePos: TCursorSizePos);
begin
  FGDI.XorMode := True;
  FGDI.PaintColor := 0;

  if x1 >= x2 then
    FGDI.Rectangle(x1 - 2,y1 + 1,x1,y2 + 2)
  else if y1 >= y2 then
    FGDI.Rectangle(x1 - 2,y1 - 2,x2 + 2,y1)
  else begin
    case SizePos of
      cspCell: begin
        FGDI.Rectangle(x1 - 2,y1 - 2,x2,y1);
        FGDI.Rectangle(x1 - 2,y2 - 2,x2 - 5,y2);
        FGDI.Rectangle(x1 - 2,y1 + 1,x1,y2 - 3);
        FGDI.Rectangle(x2 - 2,y1 + 1,x2,y2 - 5);

        FGDI.Rectangle(x2 - 3,y2 - 3,x2 + 1,y2 + 1);
      end;
      cspCellReadOnly: begin
        FGDI.Rectangle(x1 - 2,y1 - 2,x2,y1);
        FGDI.Rectangle(x1 - 2,y2 - 2,x2 - 0,y2);
        FGDI.Rectangle(x1 - 2,y1 + 1,x1,y2 - 3);
        FGDI.Rectangle(x2 - 2,y1 + 1,x2,y2 - 3);
      end;
      cspColumn: begin
        FGDI.Rectangle(x1 - 2,y1 - 2,x2 - 4,y1);
        FGDI.Rectangle(x1 - 2,y2 - 2,x2,y2);
        FGDI.Rectangle(x1 - 2,y1 + 1,x1,y2 - 3);
        FGDI.Rectangle(x2 - 2,y1 + 7,x2,y2 - 3);

        FGDI.Rectangle(x2 - 3,y1 + 1,x2 + 1,y1 + 5);
      end;
      cspRow: begin
        FGDI.Rectangle(x1 - 2,y1 - 2,x2,y1);
        FGDI.Rectangle(x1 + 7,y2 - 2,x2,y2);
        FGDI.Rectangle(x1 - 2,y1 + 1,x1,y2 - 4);
        FGDI.Rectangle(x2 - 2,y1 + 1,x2,y2 - 3);

        FGDI.Rectangle(x1 + 1,y2 - 3,x1 + 5,y2 + 1);
      end;
    end;
  end;

  FGDI.XorMode := False;
end;

procedure TXLSBookSkin.PaintSheetBkg(x1, y1, x2, y2: integer);
begin
  FGDI.PaintColor := FColors.SheetBkg;
  FGDI.Rectangle(x1, y1, x2, y2);
end;

procedure TXLSBookSkin.PaintSplitterHandle(x1, y1, x2, y2: integer; Vertical,FSmall: boolean);
begin
  if FStyle >= xssExcel2007 then begin
    FGDI.PaintColor := FColors.WindowBkg;
    FGDI.Rectangle(x1, y1, x2, y2);
    if Vertical then
      PaintStockBMP(sxbVSplit,x1 + 2,y1 + 1)
    else
      PaintStockBMP(sxbHSplit,x1,y1 + 2);
  end
  else begin
    FGDI.PaintColor := FColors.Header.Cl1;
    FGDI.Rectangle(x1, y1, x2, y2);
    FGDI.PenColor := $FFFFFF;
    FGDI.Line(x1 + 1,y2 - 1,x1 + 1,y1);
    FGDI.Line(x1 + 2,y1 + 1,x2 - 1,y1 + 1);
    FGDI.PenColor := $808080;
    FGDI.Line(x2 - 1,y1 + 1,x2 - 1,y2 - 1);
    FGDI.Line(x2 - 1,y2 - 1,x1,y2 - 1);
    if FSmall then begin
       FGDI.PenColor := 0;
      if Vertical then
        FGDI.Line(x1,y2,x2,y2)
      else
        FGDI.Line(x2,y1,x2,y2);
    end;
  end;
end;

procedure TXLSBookSkin.PaintStockBMP(StockBMP: TStockXBookBitmap; X, Y: integer);
begin
  case StockBMP of
    sxbCommentMark: FGDI.BmpBlt(bmpNote.Canvas.Handle,X - bmpNote.Width + 1,Y,bmpNote.Width,bmpNote.Height,0,0,SRCCOPY);
    sxbCheckBox   : FGDI.BmpBlt(bmpCheckbox.Canvas.Handle,X,Y,bmpCheckBox.Width,bmpCheckBox.Height,0,0,SRCCOPY);
    sxbHSplit     : FGDI.BmpBlt(bmpHSplit.Canvas.Handle,X,Y,bmpHSplit.Width,bmpHSplit.Height,0,0,SRCCOPY);
    sxbVSplit     : FGDI.BmpBlt(bmpVSplit.Canvas.Handle,X,Y,bmpVSplit.Width,bmpVSplit.Height,0,0,SRCCOPY);
  end;
end;

procedure TXLSBookSkin.PaintTab(x1, x2, y: integer; Text: AxUCString; State: TTabState; Color: longword);
var
  i: integer;
  Pts: array[0..5] of TPoint;
  BgColor: longword;

function VisibleContrast(BackGroundColor : longword): longword;
const
 cHalfBrightness = ((0.3 * 255.0) + (0.59 * 255.0) + (0.11 * 255.0)) / 2.0;
var
 Brightness : double;
begin
 with TRGBQuad(BackGroundColor) do
   BrightNess := (0.3 * rgbRed) + (0.59 * rgbGreen) + (0.11 * rgbBlue);
 if Brightness > cHalfBrightNess then
   Result := $000000
 else
   Result := $FFFFFF;
end;

begin
  SetTabFont(False);

  if FStyle >= xssExcel2007 then
    Dec(y,2);

  Pts[0].X := x1;
  Pts[0].Y := y;

  Pts[1].X := x2;
  Pts[1].Y := y;

  case FStyle of
    xssExcelNone,
    xssExcelXP  : begin
      Pts[2].X := x2 - TABSLANT;
      Pts[2].Y := y + GDI.TM_SYS.Height;
      Pts[3].X := x1 + TABSLANT;
      Pts[3].Y := y + GDI.TM_SYS.Height;
      Pts[4] := Pts[3];
    end;
    xssExcel2007: begin
      Pts[2].X := x2 - GDI.TM_SYS.Height;
      Pts[2].Y := y + GDI.TM_SYS.Height;

      Pts[3].X := x1;
      Pts[3].Y := y + GDI.TM_SYS.Height;
      Pts[4].X := x1;
      Pts[4].Y := y + GDI.TM_SYS.Height - 1;
    end;
    xssExcel2013: begin
      Pts[2].X := x2;
      Pts[2].Y := y + GDI.TM_SYS.Height;

      Pts[3].X := x1;
      Pts[3].Y := y + GDI.TM_SYS.Height;
      Pts[4].X := x1;
      Pts[4].Y := y;
    end;
  end;

  Pts[5] := Pts[0];

  if State <> tsNormal then
    BgColor := FColors.TabColorSel
  else begin
    if Color = XLSCOLOR_AUTO then
      BgColor := FColors.TabColor
    else
      BgColor := Color;
  end;

  FGDI.PaintColor := BgColor;

  FGDI.Polygon(Pts);

  if State <> tsNormal then begin
    FGDI.PenColor := FColors.SheetBkg;
    FGDI.Line(Pts[0].X,Pts[0].Y,Pts[3].X,Pts[3].Y + 1);
    if FStyle >= xssExcel2007 then
      FGDI.PenColor := FColors.HeaderLine
    else
      FGDI.PenColor := 0;
    FGDI.MoveTo(Pts[1].X,Pts[1].Y);
    for i := 2 to High(Pts) do
      FGDI.LineTo(Pts[i].X,Pts[i].Y);

    if State = tsSelected then
      SetTabFont(True);
    if (FStyle < xssExcel2007) and (Color <> FColors.Header.Cl1) then begin
      FGDI.PenColor := Color;
      FGDI.Line(Pts[2].X,Pts[2].Y - 1,Pts[3].X - 1,Pts[3].Y - 1);
      FGDI.Line(Pts[2].X,Pts[2].Y - 2,Pts[3].X - 1,Pts[3].Y - 2);
    end;
  end
  else begin
    FGDI.PenColor := FColors.SheetBkg;
    FGDI.Line(Pts[0].X,Pts[0].Y,Pts[3].X,Pts[3].Y + 1);
    if FStyle >= xssExcel2007 then
      FGDI.PenColor := FColors.HeaderLine
    else
      FGDI.PenColor := 0;
    FGDI.MoveTo(Pts[1].X,Pts[1].Y);
    for i := 2 to High(Pts) do
      FGDI.LineTo(Pts[i].X,Pts[i].Y);
  end;
  if State < tsSelected then begin
    FGDI.PenColor := FColors.HeaderLine;
    FGDI.Line(Pts[0].X,Pts[0].Y,Pts[1].X,Pts[1].Y);
  end;

  if State = tsSelected then
    FGDI.FontColor := RevRGB(FColors.TabFontColorSel)
  else
    FGDI.FontColor := VisibleContrast(BgColor);

  case FStyle of
    xssExcelNone,
    xssExcelXP  : begin
      TextAlign(ptaCenter);
      FGDI.CenterText(x1,x2,y + GDI.TM_SYS.Ascent,Text);
    end;
    xssExcel2007: begin
      FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
      FGDI.TextOut(x1 + GDI.TM_SYS.Ascent div 2,y + GDI.TM_SYS.Ascent,Text);
      FGDI.PenColor := FColors.HeaderLine;
      FGDI.Pixel(Pts[3].X + 1,Pts[3].Y - 1);
    end;
    xssExcel2013: begin
      TextAlign(ptaCenter);
      FGDI.CenterText(x1,x2,y + GDI.TM_SYS.Ascent,Text);

      if State = tsSelected then begin
        FGDI.PenColor := FGDI.FontColor;
        FGDI.Line(x1,y + GDI.TM_SYS.Height - 1,x2,y + GDI.TM_SYS.Height - 1);
        FGDI.Line(x1,y + GDI.TM_SYS.Height - 2,x2,y + GDI.TM_SYS.Height - 2);
      end;
    end;
  end;
end;

procedure TXLSBookSkin.PaintTabMoveBmp(X, Y: integer);
var
  w,h: integer;
begin
  w := bmpTabMove.Width;
  h := bmpTabMove.Height;
  FGDI.BmpBlt(bmpTabMove.Canvas.Handle,X - w div 2,y - h,w,h,0,0,SRCINVERT);
end;

procedure TXLSBookSkin.Test;
begin
  inherited Test;
end;

procedure TXLSBookSkin.PaintFocusCell(x1, y1, x2, y2: integer);
begin
  FGDI.XorMode := True;
  FGDI.PenColor := 0;
  FGDI.Line(x1, y1, x2, y1);
  FGDI.Line(x2, y1, x2, y2 + 1);
  FGDI.Line(x1, y2, x2, y2);
  FGDI.Line(x1, y1, x1, y2);
  FGDI.XorMode := False;
end;

procedure TXLSBookSkin.PaintMoveRect(Lines: TStraightLineArray);
var
  i: integer;
begin
  FGDI.XorMode := True;
  FGDI.SelectPatternBrush(spbGray);

  for i := 0 to High(Lines) do begin
    if Lines[i].Horiz then
      PaintHMoveLine(Lines[i].A1,Lines[i].B,Lines[i].A2)
    else
      PaintVMoveLine(Lines[i].B,Lines[i].A1,Lines[i].A2);
  end;

  FGDI.SelectPatternBrush(spbNone);
  FGDI.XorMode := False;
end;

procedure TXLSBookSkin.PaintDeleteRect(x1,y1,x2,y2: integer);
begin
  FGDI.SelectPatternBrush(spbWarning);
  FGDI.XorMode := True;
//  SetBmpBrushMode(True);
  FGDI.Rectangle(x1,y1,x2,y2);
//  SetBmpBrushMode(False);
  FGDI.SelectPatternBrush(spbNone);
  FGDI.XorMode := False;
end;

procedure TXLSBookSkin.SetClientSize(x1, y1, x2, y2: integer);
begin
  FCX1 := x1;
  FCY1 := y1;
  FCX2 := x2;
  FCY2 := y2;
end;

procedure TXLSBookSkin.SetFont(AFont: TXc12Font);
var
  H: HFONT;
begin
  H := AFont.GetHandle(FGDI.PixelsPerInchY);
  FGDI.SelectFont(H);
  FGDI.FontColor := RevRGB(Xc12ColorToRGB(AFont.Color));
end;

procedure TXLSBookSkin.SetStyle(const Value: TXBookSkinStyle);
begin
  FStyle := Value;
  bmpBtnTabSet.Free;
  bmpBtnOutline.Free;
  bmpBtnNumbers.Free;
  case FStyle of
    xssExcelXP  : begin
      FColors := @FColorsXP;

      bmpBtnTabSet := TBitmap.Create;
      bmpBtnTabSet.LoadFromResourceName(HInstance,'TABBUTTONS');

      bmpBtnOutline := TBitmap.Create;
      bmpBtnOutline.LoadFromResourceName(HInstance,'OUTLINEBTNS');

      bmpBtnNumbers := TBitmap.Create;
      bmpBtnNumbers.LoadFromResourceName(HInstance,'NUMBUTTONS');
    end;
    xssExcel2007: begin
      FColors := @FColors2007;

      bmpBtnTabSet := TBitmap.Create;
      bmpBtnTabSet.LoadFromResourceName(HInstance,'TABBUTTONS2007');

      bmpBtnOutline := TBitmap.Create;
      bmpBtnOutline.LoadFromResourceName(HInstance,'OUTLINEBTNS');

      bmpBtnNumbers := TBitmap.Create;
      bmpBtnNumbers.LoadFromResourceName(HInstance,'NUMBUTTONS');
    end;
    xssExcel2013: begin
      FColors := @FColors2013;

      bmpBtnTabSet := TBitmap.Create;
      bmpBtnTabSet.LoadFromResourceName(HInstance,'TABBUTTONS2013');

      bmpBtnOutline := TBitmap.Create;
      bmpBtnOutline.LoadFromResourceName(HInstance,'OUTLINEBTNS');

      bmpBtnNumbers := TBitmap.Create;
      bmpBtnNumbers.LoadFromResourceName(HInstance,'NUMBUTTONS');
    end;
  end;
end;

procedure TXLSBookSkin.SetupColors;
begin
  FColorsXP.WindowBkg := $DBD8D1;
  FColorsXP.SheetBkg := $FFFFFF;
  FColorsXP.SheetSel := $A9B2CA;

  FColorsXP.Header.Cl1 := $DBD8D1;
  FColorsXP.Header.Cl2 := $DBD8D1;
  FColorsXP.Header.Sel1 := $49422D;
  FColorsXP.Header.Sel2 := $49422D;
  FColorsXP.Header.Focus1 := $B6BDD2;
  FColorsXP.Header.Focus2 := $B6BDD2;
  FColorsXP.Header.Font := $000000;
  FColorsXP.Header.Bold := False;

  FColorsXP.Row.Cl1 := $DBD8D1;
  FColorsXP.Row.Cl2 := $DBD8D1;
  FColorsXP.Row.Sel1 := $49422D;
  FColorsXP.Row.Sel2 := $49422D;
  FColorsXP.Row.Focus1 := $B6BDD2;
  FColorsXP.Row.Focus2 := $B6BDD2;
  FColorsXP.Row.Font := $000000;
  FColorsXP.Row.Bold := False;

  FColorsXP.HasHoover := False;
  FColorsXP.HeaderHoover.Cl1 := $FFFFFF;
  FColorsXP.HeaderHoover.Cl2 := $FFFFFF;
  FColorsXP.HeaderHoover.Sel1 := $FFFFFF;
  FColorsXP.HeaderHoover.Sel2 := $FFFFFF;
  FColorsXP.HeaderHoover.Focus1 := $FFFFFF;
  FColorsXP.HeaderHoover.Focus2 := $FFFFFF;
  FColorsXP.HeaderHoover.Font := $000000;
  FColorsXP.HeaderHoover.Bold := False;

  FColorsXP.HeaderLine := $808080;
  FColorsXP.HeaderSelLine := $F5DB95;
  FColorsXP.HeaderFocusLine := $0A246A;
  FColorsXP.HeaderFrozenLine := $000000;
  FColorsXP.HeaderAutofiltFnt := $FF0000;
  FColorsXP.HeaderAutofiltFntSel := $00FFFF;
  FColorsXP.SheetGridline := $B0B0B0;
  FColorsXP.HeaderLineWidth := 1;

  FColorsXP.TabColor := $7F7F7F;
  FColorsXP.TabColorSel := clWhite;
  FColorsXP.TabFontColor := clWhite;
  FColorsXP.TabFontColorSel := clBlack;

  FColorsXP.Tab.Font := $000000;
  FColorsXP.Tab.Bold := False;

  FColorsXP.GutterColor := $DBD8D1;
  FColorsXP.GutterLineColor := $A0A0A0;

  // Excel 2007

  FColors2007.WindowBkg := $9EB8DC;
  FColors2007.SheetBkg := $FFFFFF;
  FColors2007.SheetSel := $A9B2CA;

  FColors2007.Header.CL1 := $F9FCFD;
  FColors2007.Header.Cl2 := $D3DBE9;
  FColors2007.Header.Sel1 := $9DC1E4;
  FColors2007.Header.Sel2 := $9DC1E4;
  FColors2007.Header.Focus1 := $F9D99F;
  FColors2007.Header.Focus2 := $F1C15F;
  FColors2007.Header.Font := $000000;
  FColors2007.Header.Bold := False;

  FColors2007.Row.CL1 := $E4ECF7;
  FColors2007.Row.Cl2 := $E4ECF7;
  FColors2007.Row.Sel1 := $9DC1E4;
  FColors2007.Row.Sel2 := $9DC1E4;
  FColors2007.Row.Focus1 := $FFD58D;
  FColors2007.Row.Focus2 := $FFD58D;
  FColors2007.Row.Font := $000000;
  FColors2007.Row.Bold := False;

  FColors2007.Tab.CL1 := $E4ECF7;
  FColors2007.Tab.Cl2 := $E4ECF7;
  FColors2007.Tab.Sel1 := $9DC1E4;
  FColors2007.Tab.Sel2 := $9DC1E4;
  FColors2007.Tab.Focus1 := $FFD58D;
  FColors2007.Tab.Focus2 := $FFD58D;
  FColors2007.Tab.Font := $000000;
  FColors2007.Tab.Bold := False;

  FColors2007.HasHoover := True;

  FColors2007.HeaderHoover.Cl1 := $DFE2E4;
  FColors2007.HeaderHoover.Cl2 := $BCC5D2;
  FColors2007.HeaderHoover.Sel1 := $DFE2E4;
  FColors2007.HeaderHoover.Sel2 := $BCC5D2;
  FColors2007.HeaderHoover.Focus1 := $FFD58D;
  FColors2007.HeaderHoover.Focus2 := $F2923A;
  FColors2007.HeaderHoover.Font := $000000;
  FColors2007.HeaderHoover.Bold := False;

  FColors2007.RowHoover.CL1 := $BCC5D2;
  FColors2007.RowHoover.Cl2 := $BCC5D2;
  FColors2007.RowHoover.Sel1 := $BBC4D1;
  FColors2007.RowHoover.Sel2 := $BBC4D1;
  FColors2007.RowHoover.Focus1 := $F1C05C;
  FColors2007.RowHoover.Focus2 := $F1C05C;
  FColors2007.RowHoover.Font := $000000;
  FColors2007.RowHoover.Bold := False;

  FColors2007.HeaderLine := $9EB6CE;
  FColors2007.HeaderSelLine := $FFFFFF;
  FColors2007.HeaderFocusLine := $F29536;
  FColors2007.HeaderFrozenLine := $000000;
  FColors2007.HeaderAutofiltFnt := $FF0000;
  FColors2007.HeaderAutofiltFntSel := $00FFFF;
  FColors2007.SheetGridline := $D0D7E5;
  FColors2007.HeaderLineWidth := 1;

  FColors2007.TabColor := $C8DDF7;
  FColors2007.TabColorSel := clWhite;
  FColors2007.TabFontColor := $15429F;
  FColors2007.TabFontColorSel := $15429F;

  FColors2007.GutterColor := $D7E6F7;
  FColors2007.GutterLineColor := $A0A0A0;

  // Excel 2013

  FColors2013.WindowBkg := $F6F6F6;
  FColors2013.SheetBkg := $FFFFFF;
  FColors2013.SheetSel := $A6A6A6;

  FColors2013.Header.CL1 := $FFFFFF;
  FColors2013.Header.Cl2 := $FFFFFF;
  FColors2013.Header.Sel1 := $D3F0E0;
  FColors2013.Header.Sel2 := $D3F0E0;
  FColors2013.Header.Focus1 := $E1E1E1;
  FColors2013.Header.Focus2 := $E1E1E1;
  FColors2013.Header.Font := $217346;
  FColors2013.Header.Bold := True;

  FColors2013.Row.CL1 := $FFFFFF;
  FColors2013.Row.Cl2 := $FFFFFF;
  FColors2013.Row.Sel1 := $E1E1E1;
  FColors2013.Row.Sel2 := $E1E1E1;
  FColors2013.Row.Focus1 := $E1E1E1;
  FColors2013.Row.Focus2 := $E1E1E1;
  FColors2013.Row.Font := $217346;
  FColors2013.Row.Bold := True;

  FColors2013.Tab.CL1 := $FFFFFF;
  FColors2013.Tab.Cl2 := $FFFFFF;
  FColors2013.Tab.Sel1 := $FFFFFF;
  FColors2013.Tab.Sel2 := $FFFFFF;
  FColors2013.Tab.Focus1 := $FFFFFF;
  FColors2013.Tab.Focus2 := $FFFFFF;
  FColors2013.Tab.Font := $217346;
  FColors2013.Tab.Bold := True;

  FColors2013.HasHoover := True;

  FColors2013.HeaderHoover.Cl1 := $9FD5B7;
  FColors2013.HeaderHoover.Cl2 := $9FD5B7;
  FColors2013.HeaderHoover.Sel1 := $9FD5B7;
  FColors2013.HeaderHoover.Sel2 := $9FD5B7;
  FColors2013.HeaderHoover.Focus1 := $9FD5B7;
  FColors2013.HeaderHoover.Focus2 := $9FD5B7;
  FColors2013.HeaderHoover.Font := $217346;
  FColors2013.HeaderHoover.Bold := False;

  FColors2013.RowHoover.CL1 := $9FD5B7;
  FColors2013.RowHoover.Cl2 := $9FD5B7;
  FColors2013.RowHoover.Sel1 := $9FD5B7;
  FColors2013.RowHoover.Sel2 := $9FD5B7;
  FColors2013.RowHoover.Focus1 := $9FD5B7;
  FColors2013.RowHoover.Focus2 := $9FD5B7;
  FColors2013.RowHoover.Font := $217346;
  FColors2013.RowHoover.Bold := False;

  FColors2013.HeaderLine := $9EB6CE;
  FColors2013.HeaderSelLine := $217346;
  FColors2013.HeaderFocusLine := $217346;
  FColors2013.HeaderFrozenLine := $000000;
  FColors2013.HeaderAutofiltFnt := $FF0000;
  FColors2013.HeaderAutofiltFntSel := $00FFFF;
  FColors2013.SheetGridline := $D6D6D6;
  FColors2013.HeaderLineWidth := 2;

  FColors2013.TabColor := $F6F6F6;
  FColors2013.TabColorSel := clWhite;
  FColors2013.TabFontColor := $000000;
  FColors2013.TabFontColorSel := $217346;

  FColors2013.GutterColor := $FF0000;
  FColors2013.GutterLineColor := $A0A0A0;
end;

initialization
  TXLSBookSkin.LoadResources;

finalization
  TXLSBookSkin.FreeResources;

end.

