unit XBookPaintGDI2;

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

uses {* Delphi  *} Classes, SysUtils, Windows, Math, Contnrs, vcl.Graphics,
     {* XLSRWII *} Xc12Utils5, Xc12DataStyleSheet5, Xc12DataAutofilter5,
                   XLSUtils5,

     {* XLSBook *} XBookTypes2, XBookUtils2, XBookSysVar2;

const GDIMETRIC_CHECKBOXSIZE = 13;

type TXGDICanvas = (xgdiBackground,xgdiBitmap);

type TPaintLineStyle = (plsNone,plsThin,plsMedium,plsDashed,plsDotted,plsThick,
                        plsDouble,plsHair,plsMediumDashed,plsDashDot,
                        plsMediumDashDot,plsDashDotDot,plsMediumDashDotDot,
                        plsSlantedDashDot,pls_Clear_,pls_MediumDotted_);

type TXGDIHorizTextAlign = (xhtaDefault,xhtaLeft,xhtaCenter,xhtaRight,xhtaLeftNoUpdateCP);
type TXGDIVertTextAlign = (xvtaDefault,xvtaNone,xvtaTop,xvtaBaseline,xvtaBottom);

type TXLSCursorType     = (xctCurrent,xctArrow,xctHandPoint,xctCell,xctColSelect,xctFill,
                        xctHorizSplit,xctHoriz2Split,xctRowSelect,xctVertSplit,
                        xctVert2Split,xctCenter,xctMoveCells,xctSizeCells,xctNeSwArrow,
                        xctNwSeArrow);

type TBmpButtonState = (bbsNormal,bbsFocused,bbsClicked);
type TBmpButtonType  = (bbtTabset,bbtOutline,bbtNumbers);

type TStockPatternBrush = (spbNone,spbGray,spbLightGray,spbWarning);
type TStockXBookBitmap = (sxbCommentMark,sxbCheckBox,sxbHSplit,sxbVSplit);

type TGDICtrlType = (gctButton,gctCheckBox,gctCheckBoxChecked,gctCheckBoxMixed,
                     gctScrollButtonUp,
                     gctScrollButtonUpClicked,
                     gctScrollButtonUpDisabled,
                     gctScrollButtonDown,
                     gctScrollButtonDownClicked,
                     gctScrollButtonDownDisabled,

                     gctScrollButtonLeft,
                     gctScrollButtonLeftClicked,
                     gctScrollButtonLeftDisabled,
                     gctScrollButtonRight,
                     gctScrollButtonRightClicked,
                     gctScrollButtonRightDisabled,

                     gctCombo,gctComboClicked,gctComboHighlight,gctComboHighlightClicked,
                     gctRadioButton,gctRadioButtonChecked);

type TLineEndcapStyle = (lesCircle,lesSquare,lesPoint);
type TLineColorStyle = (lcsSolid,lcsGradient,lcsRainbow);

type TDrawTextOption = (dtoHCenter,dtoLeft,dtoRight,dtoTop,dtoVCenter,dtoBaseline,dtoBottom,dtoVDistributed,dtoWordBreak);
     TDrawTextOptions = set of TDrawTextOption;

type TFloatMatrix = array[0..2, 0..2] of double;

const DefaultMatrix: TFloatMatrix = ((1, 0, 0),(0, 1, 0),(0, 0, 1));

type TAXWTextmetrics = record
     UnderscoreSize    : integer;
     UnderscorePosition: integer;
     LineGap           : integer;
     Height            : integer;
     Ascent            : integer;
     InternalLeading   : integer;
     Descent           : integer;
     AveCharWidth      : integer;
     BreakChar         : AxUCChar;
     end;

type TXLSCurveType = (xctLineTo,xctPolyline,xctBezier);

type TXLSCurvePoints = record
     Points: PPointArray;
     Count: integer;
     CurveType: TXLSCurveType;
     end;

type TXLSCurve = class(TObject)
private
     FStart: TPoint;
     FCount: integer;
     FCurves: array of TXLSCurvePoints;

     function GetItems(Index: integer): TXLSCurvePoints;
public
     constructor Create;
     destructor Destroy; override;
     procedure Add(TTCurve: PTTPOLYCURVE; Height,GlyphDescent: integer);
     procedure CloseCurve;
     procedure Adjust(X,Y: integer);

     property Start: TPoint read FStart write FStart;
     property Count: integer read FCount;
     property Items[Index: integer]: TXLSCurvePoints read GetItems; default;
     end;

type TXLSFigure = class(TObjectList)
private
     FWidth,FHeight: integer;
     function GetItems(Index: integer): TXLSCurve;
public
     procedure Add(TTCurves: PTTPOLYGONHEADER; Size,W,H,GlyphDescent: integer);
     procedure Adjust(X,Y: integer);

     property Items[Index: integer]: TXLSCurve read GetItems; default;
     property Width: integer read FWidth;
     property Height: integer read FHeight;
     end;

type TTransformation = class(TObject)
protected
     FMatrix: TFloatMatrix;

     function Mult(M1,M2: TFloatMatrix): TFloatMatrix;
public
     constructor Create;
     procedure Clear;

     procedure Transform(var X,Y: integer); overload;
     procedure Transform(var Arr: array of TPoint); overload;
     procedure Rotate(X,Y,Alpha: double);
     procedure Skew(X,Y: double);
     procedure Scale(X,Y: double);
     procedure Translate(X,Y: double);
     end;

type TAXWGDI = class(TObject)
private
     function  GetXorMode: boolean;
     procedure SetXorMode(const Value: boolean);
     function  GetPenColor: longword;
     procedure SetPenColor(const Value: longword);
     function  GetBrushColor: longword;
     procedure SetBrushColor(const Value: longword);
     function  GetBrushSolid: boolean;
     procedure SetBrushSolid(const Value: boolean);
     function  GetPenSolid: boolean;
     procedure SetPenSolid(const Value: boolean);
     procedure SetPaintColor(const Value: longword);
     function  GetFontColor: longword;
     procedure SetFontColor(const Value: longword);
     function  GetTheDC: longword;
     function  GetPenWidth: longword;
     procedure SetPenWidth(const Value: longword);
     function  GetTransparentMode: boolean;
     procedure SetTransparentMode(const Value: boolean);
     function  GetPenWidthF: double;
     procedure SetPenWidthF(const Value: double);
protected
     FHWND: HWND;
     FDC: HDC;
     FCurrentCursor: integer;
     FBrush: LOGBRUSH;
     FHBrush: HBRUSH;
     FPen,FSavedPen: LOGPEN;
     FHPen: HPEN;
     FFontHandle: HFONT;
     FMinX,FMinY,FMaxX,FMaxY: integer;
     FFontRotation: integer;
     FTM_SYS: TAXWTextmetrics;
     FTM: TAXWTextmetrics;
     // Delphi specific, remove this...
     FFillBmp: TBitmap;
     FDirty: boolean;
     FLineEndcapStyle: TLineEndcapStyle;
     FLineColorStyle: TLineColorStyle;
     FTransform: TTransformation;
     FZoom: double;

     procedure SetDirtyArea(x,y: integer); overload; {$ifdef D2006PLUS} inline; {$endif}
     procedure SetDirtyArea(Points: array of TPoint); overload;
     procedure SetDefaultData;
     procedure ResetLimits;
     procedure CheckHandle; virtual;
     procedure SetPixel(x,y: integer; Cl: longword); virtual;
     procedure SetPixelAA(x,y: integer; Cl: longword; A: byte); virtual;

     procedure LineWide(x1,y1,x2,y2: integer);
     procedure LineWide2(x1,y1,x2,y2: integer);
     procedure LineAADraw(x1,y1,x2,y2: integer);

     procedure StrToFont32Array(Source: AxUCString; Dest: PWideChar);
public
     constructor Create(WinHandle: longword);
     destructor Destroy; override;

     class procedure LoadResources;
     class procedure FreeResources;

     procedure ReleaseHandle; virtual;
     procedure Render(Dest: TAXWGDI); virtual;
     procedure RenderDirty(Dest: TAXWGDI); virtual;
     procedure Clear; virtual;
     procedure SetDirtyArea(x1,y1,x2,y2: integer); overload; {$ifdef D2006PLUS} inline; {$endif}

     // Basic gdi functions
     procedure Line(x1,y1,x2,y2: integer);
     procedure LineF(x1,y1,x2,y2: double);
     procedure LineStyled(x1,y1,x2,y2: integer; PenColor,BackColor: longword; Style: TPaintLineStyle; Horiz: boolean; Diag: boolean = False);
     procedure MoveTo(x,y: integer);
     procedure MoveToF(x,y: double);
     procedure LineTo(x,y: integer);
     procedure LineToF(x,y: double);
     procedure Rectangle(x1,y1,x2,y2: integer); overload;
     procedure Rectangle(Rect: TXYRect); overload;
     procedure RectangleF(x1,y1,x2,y2: double); overload;
     procedure FillRect(Rect: TXYRect);
     procedure Ellipse(x1,y1,x2,y2: integer);
     procedure Circle(x,y,r: integer);
     procedure CircleF(x,y,r: double);
     procedure PizzaSlice(const x,y,r: integer; const AStartAng,AEndAng: double);
     procedure PizzaSliceF(const x,y,r: double; const AStartAng,AEndAng: double);
     procedure Arc(x1,y1,x2,y2,StartX,StartY,EndX,EndY: integer);
     procedure ArcTo(x1,y1,x2,y2,StartX,StartY,EndX,EndY: integer);
     procedure RoundRect(x1,y1,x2,y2: integer; Adjust: double);
     procedure Polygon(Points: array of TPoint);
     procedure Polyline(Points: array of TPoint);
     procedure PolylineTo(Points: array of TPoint);
     procedure PolyBezier(Points: array of TPoint);
     procedure PolyBezierTo(Points: array of TPoint);
     procedure Pixel(x,y: integer);

     procedure BmpBlt(Handle: longword; X, Y, Width, Height, XSrc, YSrc: integer; ROP: longword = SRCCOPY);
     procedure GradientFillRect(x1,y1,x2,y2: integer; StartColor,EndColor: longword; HorizontalFill: boolean);
     procedure Draw(const AX,AY: integer; AGraphic: TObject);
     procedure DrawF(const AX,AY: double; AGraphic: TObject); overload;
     procedure DrawF(const AX,AY,AWidth,AHeight: double; AGraphic: TObject); overload;

     procedure SetArcDirection(CCW: boolean);

     procedure SavePen;
     procedure RecallPen;

     procedure AdjustBrush(Percent: double);

     // Antialiasing
     procedure LineAA(x1,y1,x2,y2: integer);

     procedure DrawControl(ControlType: TGDICtrlType; x1,y1,x2,y2: integer);
     procedure DrawEdge(x1, y1, x2, y2: integer; EdgeType,EdgeFlags: longword);
     procedure WidthRectangle(x1,y1,x2,y2,W: integer);

     procedure CreatePen(Color: longword; Style: longword; Width: longword);

     procedure Invalidate; virtual;
     procedure InvalidateRect(x1,y1,x2,y2: integer); virtual;

     // Functions needed by XLSBook
     procedure PaintSelectBMP(x1,y1,x2,y2: integer);
     procedure PaintTabMoveBmp(X, Y: integer);
     procedure SizeLine(x1, y1, x2, y2: integer);
     procedure CenterText(x1,x2,y: integer; Text: AxUCString);
     procedure CenterTextRect(x1, y1, x2, y2: integer; Text: AxUCString); overload;
     procedure CenterTextRect(x1, y1, x2, y2, y: integer; Text: AxUCString); overload;
     procedure CenterTextVert(x1, y1, x2, y2: integer; Text: AxUCString);
     procedure _GetTextMetricSys;
     procedure GetTextMetric;
     procedure Scroll(X1,Y1,X2,Y2,dX,dY: integer);
     procedure TextOut(X,Y: integer; Rect: TRect; Text: AxUCString); overload;
     procedure TextOut(X,Y: integer; Text: AxUCString); overload;
     procedure TextOutF(X,Y: double; Text: AxUCString); overload;
     procedure TextOutF(X,Y: double; Text: AxUCString; AOptions: TDrawTextOptions); overload;
     procedure DrawText(X1,Y1,X2,Y2: double; Text: AxUCString; Options: TDrawTextOptions); overload;
     procedure DrawText(Rect: TRect; Text: AxUCString; Options: TDrawTextOptions); overload;
     procedure PaintBMPImage(Bitmap: TBitmap; X1,Y1,X2,Y2: integer); overload;
     procedure PaintEMFImage(Buf: Pointer; BufSize, X1,Y1,X2,Y2: integer);
     procedure PaintWMFImage(Buf: Pointer; BufSize, X1,Y1,X2,Y2: integer);
     procedure SelectPatternBrush(PatternBrush: TStockPatternBrush);
     procedure PaintStockBMP(StockBMP: TStockXBookBitmap; X,Y: integer);
     procedure PaintCondFmtIcon(const AStyle: TXc12IconSetType; const AIndex, AX ,AY: integer);
     procedure CheckMinFontSize;

     // Clip functions
     function  CreateClipRect(x1,y1,x2,y2: integer): longword;
     function  AppendClipRect(x1,y1,x2,y2: integer): longword;
     procedure SelectClipRect(ClipHandle: longword);
     procedure DeleteClipRect(ClipHandle: longword);

     // Font functions
     procedure SetFont(AFont: TXc12Font);
     function  CreateFont(FaceName: AxUCString; Size: double; Bold,Italic,Underline: boolean; CharSet: integer): longword;
     function  CreateFontByStr(FaceName,FontStyle: AxUCString; Size: integer): integer;
     procedure SelectFont(FontHandle: longword);
     procedure DeleteFont(FontHandle: longword);
     function  GetFont: longword;
     procedure CopyFontToDC(DestDC: longword);

     function  PixelsPerInchX: integer; virtual;
     function  PixelsPerInchY: integer; virtual;
     function  PtToPixX(const APoints: double): integer;
     function  PtToPixY(const APoints: double): integer;
     function  PixToPtX(const APixels: integer): double;
     function  PixToPtY(const APixels: integer): double;
     function  PixelsPerCm: integer;
     function  TextHeight(const Text: AxUCString): integer;
     function  TextHeightF(const Text: AxUCString): double;
     function  TextWidth(const Text: AxUCString): integer;
     function  TextWidthF(const Text: AxUCString): double;
     function  TextWidthFloat(const Text: AxUCString): double;
     function  CharWidth(C: WideChar): integer;
     function  StdCharWidth: integer;
     function  GetFontRotation(Handle: longword): double;
     function  SetFontRotation(Handle: longword; const Value: integer): longword;
     function  SetFontStyle(Bold,Italic,Underline: boolean): longword;
     function  SetFontStrikeOut(StrikeOut: boolean): longword;
     procedure GetFontSuperscriptData(var FontOffset,FontSize: TPoint);
     procedure GetFontSubscriptData(var FontOffset,FontSize: TPoint);
     function  SetFontSize(Size: integer): longword;
     function  SetFontSizePix(Size: integer): longword;
     function  SetFontName(FaceName: AxUCString): longword; overload;
     function  SetFontName(FaceName: AxUCString; ACharset: byte): longword; overload;
     function  GetFontName: AxUCString;
     function  SetTextJustification(const ABreakExtra,ABreakCount: integer): integer;
     function  SetTextJustificationF(const ABreakExtra,ABreakCount: double): integer;
     // Right To Left
     function  CharIsRTL(C: WideChar): boolean;
     function  GetFontSize: integer;
     function  FntPixHeightToPt(PixHeight: integer): integer;
//     function  GetGlyphCurves(C: WideChar; var Buf: Pointer): integer;
     procedure GetGlyphCurves(C: WideChar; Curves: TXLSFigure; GlyphHeight,GlyphDescent: integer; XScale,YScale: double);
     procedure DrawGlyphCurves(Buf: Pointer; Sz: integer);
     procedure DrawXLSCurves(Curves: TXLSFigure);

     // Bitmap functions
     procedure StretchBlt(DestDC: longword; DestX,DestY,DestWidth,DestHeight,ScrX,SrcY,SrcWidth,SrcHeight: integer);

     procedure SetCursor(Cursor: TXLSCursorType);
     procedure SetTextAlign(HPos: TXGDIHorizTextAlign; VPos: TXGDIVertTextAlign);

     function  EMUToPixels(const EMU: integer): integer;
     function  PixelsToEMU(const Pixels: integer): integer;

     procedure DebugDumpBMP(N: integer);

     property BrushColor: longword read GetBrushColor write SetBrushColor;
     property BrushSolid: boolean read GetBrushSolid write SetBrushSolid;
     property PenColor: longword read GetPenColor write SetPenColor;
     property PenWidth: longword read GetPenWidth write SetPenWidth;
     property PenWidthF: double read GetPenWidthF write SetPenWidthF;
     property PenSolid: boolean read GetPenSolid write SetPenSolid;
     property PaintColor: longword write SetPaintColor;
     property FontColor: longword read GetFontColor write SetFontColor;
     property XorMode: boolean read GetXorMode write SetXorMode;
     property TransparentMode: boolean read GetTransparentMode write SetTransparentMode;
     property FontRotation: integer read FFontRotation;

     property TM_SYS: TAXWTextmetrics read FTM_SYS;
//     property TM_CELL: TEXTMETRIC read FTM_CELL;
     property TM: TAXWTextmetrics read FTM;

     property DC: longword read GetTheDC;
     property Dirty: boolean read FDirty;

     property Trans: TTransformation read FTransform;

     property Zoom: double read FZoom write FZoom;
     end;

type TAXWGDIMetafile = class(TAXWGDI)
private
     FMetafile: TMetafile;
     FMCanvas : TMetafileCanvas;
     FXScale  : double;
     FYScale  : double;
public
     constructor Create;
     destructor Destroy; override;

     function  PixelsPerInchX: integer; override;
     function  PixelsPerInchY: integer; override;

     function  DevicePixelsPerInchX: integer;
     function  DevicePixelsPerInchY: integer;

     procedure BeginMetafile(const AReferenceDC: longword; const AWidthCm,AHeightCm: double);
     procedure EndMetafile;
     procedure PlayMetafile(ACanvas: TCanvas; const AX1,AY1,AX2,AY2: integer);

     procedure CheckHandle; override;
     procedure ReleaseHandle; override;

     property XScale: double read FXScale;
     property YScale: double read FYScale;

     property Metafile: TMetafile read FMetafile;
     end;

type TAXWGDIBMP = class(TAXWGDI)
private
     FHBMP: HBITMAP;
     FBmpBits: Pointer;
     FBMPWidth,FBMPHeight: integer;
     FTransparentColor: integer;
protected
     procedure CheckHandle; override;
     procedure SetPixel(x,y: integer; Cl: longword); override;
public
     constructor Create;
     destructor Destroy; override;
     procedure ReleaseHandle; override;

     procedure Invalidate; override;
     procedure InvalidateRect(x1,y1,x2,y2: integer); override;
//     procedure LineStyled(x1,y1,x2,y2: integer; PenColor,BackColor: longword; Style: TPaintLineStyle; Horiz: boolean); override;
     end;


type TAXWGDIBMPTrans = class(TAXWGDIBMP)
private
public
     constructor Create;
     procedure Clear; override;
     procedure Render(Dest: TAXWGDI); override;
     procedure RenderDirty(Dest: TAXWGDI); override;
     end;

function  BrightColorChannel(const Channel: Byte; const Pct: Single): Byte;
procedure DrawExcelLine(const ADC: HDC; AStyle: TXc12CellBorderStyle; AColor: longword; AX1,AY1,AX2,AY2: integer);

implementation

type TRGBARec = packed record
     R,G,B,A: byte;
     end;

type PPoints = ^TPoints;
     TPoints = array[0..0] of TPoint;

var
  crCell,
  crColSelect,
  crFill,
  crHorizSplit,
  crHoriz2Split,
  crRowSelect,
  crVertSplit,
  crVert2Split,
  crCenter,
  crMoveCells,
  crSizeCells,
  crNeSwArrow,
  crNwSeArrow: longword;

  bmpGray        : TBitmap;
  bmpLightGray   : TBitmap;
  bmpWarning     : TBitmap;
  bmpComment     : TBitmap;
  bmpCheckBox    : TBitmap;
  bmpCondFmtIcons: TBitmap;

const ExcelLineStyleDot : array[0..1] of dword = (2,2);
const ExcelLineStyleDash: array[0..1] of dword = (3,1);
const ExcelLineStyleDashDot: array[0..3] of dword = (9,3,3,3);
const ExcelLineStyleDashDotDot: array[0..5] of dword = (9,3,3,3,3,3);
const ExcelLineStyleMediumDash: array[0..1] of dword = (9,3);
const ExcelLineStyleMediumDashDot: array[0..3] of dword = (9,3,3,3);
const ExcelLineStyleMediumDashDotDot: array[0..5] of dword = (9,3,3,3,3,3);
const ExcelLineSlantedDashDot1: array[0..3] of dword = (11,1,5,1);
const ExcelLineSlantedDashDot2: array[0..3] of dword = (10,2,4,2);

// cbsDouble and cbsSlantedDashDot are only drawn correct for horizontal lines.
procedure DrawExcelLine(const ADC: HDC; AStyle: TXc12CellBorderStyle; AColor: longword; AX1,AY1,AX2,AY2: integer);
var
  H,HOld: HPEN;
  LB: TLogBrush;
begin
  AColor := RevRGB(AColor);

  LB.lbStyle := BS_SOLID;
  LB.lbColor := AColor;
  case AStyle of
    cbsNone            : Exit;
    cbsThin            : H := Windows.CreatePen(PS_SOLID,1,AColor);
    cbsMedium          : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_SOLID or PS_ENDCAP_FLAT,2,LB,0,Nil);
    cbsDashed          : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,1,LB,2,@ExcelLineStyleDash);
    cbsDotted          : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,1,LB,2,@ExcelLineStyleDot);
    cbsThick           : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_SOLID or PS_ENDCAP_FLAT,3,LB,0,Nil);
    cbsDouble          : H := Windows.CreatePen(PS_SOLID,1,AColor);
    cbsHair            : H := Windows.ExtCreatePen(PS_COSMETIC or PS_ALTERNATE,1,LB,0,Nil);
    cbsMediumDashed    : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,2,LB,2,@ExcelLineStyleMediumDash);
    cbsDashDot         : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,1,LB,4,@ExcelLineStyleDashDot);
    cbsMediumDashDot   : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,2,LB,4,@ExcelLineStyleMediumDashDot);
    cbsDashDotDot      : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,1,LB,6,@ExcelLineStyleDashDotDot);
    cbsMediumDashDotDot: H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,2,LB,6,@ExcelLineStyleMediumDashDotDot);
    cbsSlantedDashDot  : H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,1,LB,4,@ExcelLineSlantedDashDot1);
    else Exit;
  end;

  HOld := Windows.SelectObject(ADC,H);

  case AStyle of
    cbsDouble        : begin
      if AX1 = AX2 then begin
        Windows.MoveToEx(ADC,AX1 - 1,AY1,Nil);
        Windows.LineTo  (ADC,AX2 - 1,AY2);
        Windows.MoveToEx(ADC,AX1 + 1,AY1,Nil);
        Windows.LineTo  (ADC,AX2 + 1,AY2);
      end
      else begin
        Windows.MoveToEx(ADC,AX1,AY1 - 1,Nil);
        Windows.LineTo  (ADC,AX2,AY2 - 1);
        Windows.MoveToEx(ADC,AX1,AY1 + 1,Nil);
        Windows.LineTo  (ADC,AX2,AY2 + 1);
      end;
    end;
    cbsSlantedDashDot: begin
      Windows.MoveToEx(ADC,AX1,AY1,Nil);
      Windows.LineTo(ADC,AX2,AY2);

      Windows.DeleteObject(H);

      H := Windows.ExtCreatePen(PS_GEOMETRIC or PS_USERSTYLE or PS_ENDCAP_FLAT,1,LB,4,@ExcelLineSlantedDashDot2);
      Windows.SelectObject(ADC,H);

      if AX1 = AX2 then begin
        Windows.MoveToEx(ADC,AX1 + 1,AY1,Nil);
        Windows.LineTo  (ADC,AX2 + 1,AY2);
      end
      else begin
        Windows.MoveToEx(ADC,AX1,AY1 + 1,Nil);
        Windows.LineTo  (ADC,AX2,AY2 + 1);
      end;
    end
    else begin
      Windows.MoveToEx(ADC,AX1,AY1,Nil);
      Windows.LineTo(ADC,AX2,AY2);
    end;
  end;
  Windows.SelectObject(ADC,HOld);
  Windows.DeleteObject(H);
end;

const MaxBytePercent = High(Byte) * 0.01;

function DarkColorChannel(const Channel: Byte; const Pct: Single): Byte;
var
  Temp: Integer;
begin
  if Pct < 0 then
    Result := BrightColorChannel(Channel, -Pct)
  else
  begin
    Temp := Round(Channel - Pct * MaxBytePercent);
    if Temp < Low(Result) then
      Result := Low(Result)
    else
      Result := Temp;
  end;
end;

function BrightColorChannel(const Channel: Byte; const Pct: Single): Byte;
var
  Temp: Integer;
begin
  if Pct < 0 then
    Result := DarkColorChannel(Channel, -Pct)
  else
  begin
    Temp := Round(Channel + Pct * MaxBytePercent);
    if Temp > High(Result) then
      Result := High(Result)
    else
      Result := Temp;
  end;
end;

function BrightColor(const Color: longword; const Pct: Single): longword;
type TRGBRec = packed record
    case Integer of
      0: (Value: Longint);
      1: (Red, Green, Blue: Byte);
      2: (R, G, B, Flag: Byte);
      3: (Index: Word);
  end;

var
  Temp: TRGBRec;
begin
  Temp.Value := ColorToRGB(Color);
  Temp.R := BrightColorChannel(Temp.R, Pct);
  Temp.G := BrightColorChannel(Temp.G, Pct);
  Temp.B := BrightColorChannel(Temp.B, Pct);
  Result := Temp.Value;
end;

function PtFxToPT(ptfx: POINTFX; Height,GlyphDescent: integer): TPoint;
begin
  Result.x := ptfx.x.value;
  if ptfx.x.fract >= $8000 then
    Inc(Result.x);
  Result.y := ptfx.y.value;
  if ptfx.y.fract >= $8000 then
    Inc(Result.Y);
  Result.Y := (Height - Result.Y) - GlyphDescent;
end;

{ TAXWGDI }

procedure TAXWGDI.AdjustBrush(Percent: double);
var
  Cl: longword;
begin
  CheckHandle;
  Cl := GetBrushColor;
  Cl := BrightColor(Cl,Percent * 100);
  SetBrushColor(Cl);
end;

function TAXWGDI.AppendClipRect(x1, y1, x2, y2: integer): longword;
begin
  CheckHandle;
  Result := Windows.CreateRectRgn(x1,y1,x2,y2);
  Windows.ExtSelectClipRgn(FDC,Result,RGN_OR);
end;

procedure TAXWGDI.Arc(x1, y1, x2, y2, StartX, StartY, EndX, EndY: integer);
begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2 + 1,y2 + 1);
  Windows.Arc(FDC,x1,y1,x2,y2,StartX,StartY,EndX,EndY);
end;

procedure TAXWGDI.ArcTo(x1, y1, x2, y2, StartX, StartY, EndX, EndY: integer);
begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2 + 1,y2 + 1);
  Windows.ArcTo(FDC,x1,y1,x2,y2,StartX,StartY,EndX,EndY);
end;

procedure TAXWGDI.BmpBlt(Handle: longword; X, Y, Width, Height, XSrc, YSrc: integer; ROP: longword);
begin
  CheckHandle;
  SetDirtyArea(X,Y,X + Width,Y + Height);
  Windows.BitBlt(FDC,X,Y,Width,Height,Handle,XSrc,YSrc,ROP);
end;

procedure TAXWGDI.CenterText(x1, x2, y: integer; Text: AxUCString);
begin
  CheckHandle;
  Windows.SetTextAlign(FDC,TA_CENTER + TA_BASELINE);
  Windows.TextOutW(FDC,x1 + ((x2 - x1) div 2),y,PWideChar(Text),Length(Text));
end;

procedure TAXWGDI.CenterTextRect(x1, y1, x2, y2, y: integer; Text: AxUCString);
var
  R: TRect;
begin
  CheckHandle;
  Windows.SetTextAlign(FDC,TA_CENTER + TA_BASELINE);
  R := Rect(x1, y1, x2, y2);
  Windows.ExtTextOutW(FDC, x1 + ((x2 - x1) div 2), y, ETO_CLIPPED, @R, PWideChar(Text), Length(Text), nil);
end;

procedure TAXWGDI.CenterTextVert(x1, y1, x2, y2: integer; Text: AxUCString);
var
  R: TRect;
begin
  CheckHandle;
  Windows.SetTextAlign(FDC,TA_LEFT + TA_TOP);
  R := Rect(x1, y1, x2, y2);
  GetTextMetric;
  Windows.ExtTextOutW(FDC, x1, y1 + ((y2 - y1) div 2) - (FTM.Height div 2), ETO_CLIPPED, @R, PWideChar(Text), Length(Text), nil);
end;

procedure TAXWGDI.CenterTextRect(x1, y1, x2, y2: integer; Text: AxUCString);
var
  R: TRect;
begin
  CheckHandle;
  Windows.SetTextAlign(FDC,TA_CENTER + TA_TOP);
  R := Rect(x1, y1, x2, y2);
  GetTextMetric;
  Windows.ExtTextOutW(FDC, x1 + ((x2 - x1) div 2), y1 + ((y2 - y1) div 2) - (FTM.Height div 2), ETO_CLIPPED, @R, PWideChar(Text), Length(Text), nil);
end;

function TAXWGDI.CharIsRTL(C: WideChar): boolean;
begin
  // Hebrew and Arabic
  Result := ((Word(C) >= $0591) and (Word(C) <= $06FE)) or ((Word(C) >= $FB1D) and (Word(C) <= $FDFB));
end;

function TAXWGDI.CharWidth(C: WideChar): integer;
var
  Sz: SIZE;
  S: AxUCString;
  ABC: array [0..0] of TABC;
begin
  CheckHandle;
  if Windows.GetCharABCWidths(FDC,longword(C),longword(C),ABC) then
    Result := ABC[0].abcA + Integer(ABC[0].abcB) + ABC[0].abcC
  else begin
    S := C;
    GetTextExtentPointW(FDC,PWideChar(S),1,Sz);
    Result := Sz.cx;
  end;
end;

procedure TAXWGDI.CheckHandle;
begin
  if FDC = 0 then begin
    FDC := GetDC(FHWND);
    SetDefaultData;
  end;
end;

// Optimize the selection of small font...
procedure TAXWGDI.CheckMinFontSize;
var
  W: double;
  H: longword;
  S: AxUCString;
begin
  S := GetFontName;
  if S <> FONT_SMALL then begin
    W := TextWidth('0');
    if (W > 0) and (W <= 4) then begin
      H := SetFontName(FONT_SMALL);
      SelectObject(FDC,H);
    end;
  end;
end;

procedure TAXWGDI.Circle(x, y, r: integer);
begin
  CheckHandle;
  SetDirtyArea(x - r,y - r,x + r,y + r);
  Windows.Ellipse(FDC,x - r,y - r,x + r,y + r);
end;

procedure TAXWGDI.CircleF(x, y, r: double);
begin
  Circle(Round(x),Round(y),Round(r));
end;

procedure TAXWGDI.Clear;
begin

end;

procedure TAXWGDI.CopyFontToDC(DestDC: longword);
var
  LF: LOGFONT;
  Pt: double;
begin
  CheckHandle;
  Windows.GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LOGFONT),@LF);
  Pt := (LF.lfHeight * 72) / PixelsPerInchY;
  LF.lfHeight := MulDiv(Round(Pt), PixelsPerInchY, 72);
  Windows.SelectObject(DestDC,CreateFontIndirect(LF));
end;

constructor TAXWGDI.Create(WinHandle: longword);
begin
  FHWND := WinHandle;
  // Default
  FCurrentCursor := 0;

  FZoom := 1.00;

  FTransform := TTransformation.Create;

  FFillBmp := TBitmap.Create;
  FFillBmp.Width := 2;
  FFillBmp.Height := 2;
  FFillBmp.Canvas.Pen.Color := $CAB2A9;
  FFillBmp.Canvas.Brush.Color := clRed;
  FFillBmp.Canvas.Rectangle(0,0,2,2);
end;

function TAXWGDI.CreateClipRect(x1, y1, x2, y2: integer): longword;
//var
//  E: integer;
begin
  CheckHandle;

  if not ((x1 >= x2) or (y1 >= y2)) then begin
    Result := Windows.CreateRectRgn(x1,y1,x2,y2);
//    E := Windows.SelectClipRgn(FDC,Result);
    Windows.SelectClipRgn(FDC,Result);
//    if E <> SIMPLEREGION then
//      raise XLSRWException.CreateFmt('Clip error %d',[E]);
  end
  else
    Result := 0;
end;

procedure TAXWGDI.CreatePen(Color, Style, Width: longword);
begin
  CheckHandle;
  FPen.lopnStyle := Style;
  FPen.lopnColor := RevRGB(Color);
  FPen.lopnWidth.X := Width;
  FHPen := CreatePenIndirect(FPen);
  DeleteObject(SelectObject(FDC,FHPen));
end;

function TAXWGDI.CreateFont(FaceName: AxUCString; Size: double; Bold, Italic, Underline: boolean; CharSet: integer): longword;
var
  LF: LOGFONTW;
begin
  CheckHandle;

//  GetObject(GetStockObject(ANSI_VAR_FONT),SizeOf(LF),@LF);
  FillChar(LF,SizeOf(LOGFONT),#0);
  if Bold then
    LF.lfWeight := 700
  else
    LF.lfWeight := 400;
  if Italic then
    LF.lfItalic := 1
  else
    LF.lfItalic := 0;
  if Underline then
    LF.lfUnderline := 1
  else
    LF.lfUnderline := 0;
  if Size < 0 then
    LF.lfHeight := Round(Size)
  else
    LF.lfHeight := -MulDiv(Round(Size), PixelsPerInchY, 72);
  LF.lfCharSet := CharSet;
  LF.lfOutPrecision := OUT_TT_ONLY_PRECIS;
  StrToFont32Array(FaceName,LF.lfFaceName);
  Result := CreateFontIndirectW(LF);
end;

function XLSEnumFontFamilies(const EnumLogFontEx: TEnumLogFontExW; const NewTextMetricEx: TNewTextMetricExW; FontType: Integer; LF: PLOGFONTW): Integer; stdcall;
begin
  Result := integer(not CompareMem(@EnumLogFontEx.elfStyle,@LF.lfFaceName,LF_FACESIZE * 2));
  if Result = 0 then
    Move(EnumLogFontEx.elfLogFont,LF^,SizeOf(LOGFONTW));
end;

function TAXWGDI.CreateFontByStr(FaceName, FontStyle: AxUCString; Size: integer): integer;
var
  LF: LOGFONTW;
  LFFound: LOGFONTW;
begin
  FaceName := Copy(FaceName,1,LF_FACESIZE);
  FontStyle := Copy(FontStyle,1,LF_FACESIZE);

  FillChar(LF,SizeOf(LOGFONTW),0);
  FillChar(LFFound,SizeOf(LOGFONTW),0);
  Move(FaceName[1], LF.lfFaceName, Length(FaceName) * 2);
  Move(FontStyle[1], LFFound.lfFaceName, Length(FontStyle) * 2);
  LF.lfCharSet := DEFAULT_CHARSET;
  lf.lfOutPrecision := OUT_TT_ONLY_PRECIS;
  EnumFontFamiliesExW(FDC,LF,@XLSEnumFontFamilies,Integer(@LFFound),0);
  if LFFound.lfWeight > 0 then begin
    LFFound.lfHeight := -MulDiv(Round(Size), PixelsPerInchY, 72);
    LFFound.lfWidth := 0;
    Result := CreateFontIndirectW(LFFound);
  end
  else
    Result := 0;
end;

procedure TAXWGDI.DeleteFont(FontHandle: longword);
begin
  DeleteObject(FontHandle);
end;

function TAXWGDI.GetFont: longword;
begin
  CheckHandle;
  Result := Windows.GetCurrentObject(FDC,OBJ_FONT);
end;

procedure TAXWGDI.DebugDumpBMP(N: integer);
var
  B: TBitmap;
begin
  if True then begin
    B := TBitmap.Create;
    B.Width := 1600;
    B.Height := 1200;
    Windows.BitBlt(B.Canvas.Handle,0,0,1600,1200,FDC,0,0,SRCCOPY);
    B.SaveToFile('d:\temp\temp' + IntToStr(N) + '.bmp');
    B.Free;
  end;
end;

procedure TAXWGDI.DeleteClipRect(ClipHandle: longword);
begin
  CheckHandle;
  if ClipHandle <> 0 then begin
    Windows.SelectClipRgn(FDC,0);
    Windows.DeleteObject(ClipHandle);
  end;
end;

destructor TAXWGDI.Destroy;
begin
  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);
  DeleteObject(FHBrush);
  FFillBmp.Free;
  FTransform.Free;

  inherited;
end;

procedure TAXWGDI.Draw(const AX, AY: integer; AGraphic: TObject);
begin
  raise XLSRWException.Create('TODO');
end;

procedure TAXWGDI.DrawControl(ControlType: TGDICtrlType; x1, y1, x2, y2: integer);
type
  TCtrlArrowDir = (cadLeft,cadUp,cadRight,cadDown);
  TCtrlArrowState = (casUp,casDown,casDisabled);
var
  x,y: integer;
  C: longword;

procedure MakeHighlight;
begin
  x := x1 + ((x2 - x1) div 2);
  y := y1 + ((y2 - y1) div 2);
  C := Windows.GetPixel(FDC,x,y);
  BrushColor := $0000FF;
  Windows.ExtFloodFill(FDC,x,y,C,FLOODFILLSURFACE)
end;

procedure DrawButton(Clicked: boolean);
var
  R: TRect;
begin
  if Clicked then begin
    BrushColor := RevRGB(GetSysColor(COLOR_BTNFACE));
    PenColor := RevRGB(GetSysColor(COLOR_BTNSHADOW));
    Windows.Rectangle(FDC,x1,y1,x2,y2);
  end
  else begin
    R := Rect(x1,y1,x2,y2);
    Windows.DrawEdge(FDC,R,EDGE_RAISED,BF_TOPRIGHT + BF_BOTTOMLEFT);
    PenSolid := False;
    BrushColor := RevRGB(GetSysColor(COLOR_BTNFACE));
    Windows.Rectangle(FDC,x1 + 2,y1 + 2,x2 - 1,y2 - 1);
  end;
end;

procedure DrawArrowButton(Dir: TCtrlArrowDir; State: TCtrlArrowState; Color: longword = $000000);
var
  Pts: array[0..3] of TPoint;
  ArrowHeight: integer;
  cX,cX2,cY,cY2: integer;
begin
  DrawButton(State = casDown);

  ArrowHeight := (x2 - x1) div 4;
  cX := ((x2 - x1) div 2) + x1;
  cX2 := cX - (ArrowHeight div 2);
  cY := ((y2 - y1) div 2) + y1;
  cY2 := cY - (ArrowHeight div 2);

  if State = casDown then begin
    Inc(cX);
    Inc(cX2);
    Inc(cY);
    Inc(cY2);
  end;

  case Dir of
    cadLeft: begin
      Pts[0].X := cX2;
      Pts[0].Y := cY;

      Pts[1].X := cX2 + ArrowHeight;
      Pts[1].Y := cY - ArrowHeight;

      Pts[2].X := cX2 + ArrowHeight;
      Pts[2].Y := cY + ArrowHeight;

      Pts[3].X := Pts[0].X;
      Pts[3].Y := Pts[0].Y;
    end;
    cadUp: begin
      Pts[0].X := cX;
      Pts[0].Y := cY2;

      Pts[1].X := cX + ArrowHeight;
      Pts[1].Y := cY2 + ArrowHeight;

      Pts[2].X := cX - ArrowHeight;
      Pts[2].Y := cY2 + ArrowHeight;

      Pts[3].X := Pts[0].X;
      Pts[3].Y := Pts[0].Y;
    end;
    cadRight: begin
      Pts[0].X := cX2 + ArrowHeight;
      Pts[0].Y := cY;

      Pts[1].X := cX2;
      Pts[1].Y := cY - ArrowHeight;

      Pts[2].X := cX2;
      Pts[2].Y := cY + ArrowHeight;

      Pts[3].X := Pts[0].X;
      Pts[3].Y := Pts[0].Y;
    end;
    cadDown: begin
      Pts[0].X := cX;
      Pts[0].Y := cY2 + ArrowHeight;

      Pts[1].X := cX + ArrowHeight;
      Pts[1].Y := cY2;

      Pts[2].X := cX - ArrowHeight;
      Pts[2].Y := cY2;

      Pts[3].X := Pts[0].X;
      Pts[3].Y := Pts[0].Y;
    end;
  end;
  PenSolid := True;
  if State = casDisabled then
    PaintColor := RevRGB(GetSysColor(COLOR_BTNSHADOW))
  else
    PaintColor := Color;
  Windows.Polygon(FDC,PPoints(@Pts)^,Length(Pts));
end;

begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2 + 1,y2 + 1);  //!!!
  case ControlType of
    gctButton:          DrawButton(False);
    gctCheckBox,gctCheckBoxChecked,gctCheckBoxMixed: begin
      SetPenColor($000000);
      SetPenWidth(1);
      BrushSolid := True;
      if ControlType = gctCheckBoxMixed then
        SelectPatternBrush(spbGray)
      else
        SetBrushColor($FFFFFF);
      y := y1 + ((y2 - y1) div 2);
      Windows.Rectangle(FDC,x1,y - (GDIMETRIC_CHECKBOXSIZE div 2),x1 + GDIMETRIC_CHECKBOXSIZE,y + (GDIMETRIC_CHECKBOXSIZE div 2) + 1);
      case ControlType of
        gctCheckBoxChecked: PaintStockBMP(sxbCheckBox,x1 + 3,y - (GDIMETRIC_CHECKBOXSIZE div 4));
        gctCheckBoxMixed:   SelectPatternBrush(spbNone);
      end;
    end;
    gctScrollButtonUp             : DrawArrowButton(cadUp,casUp);
    gctScrollButtonUpClicked      : DrawArrowButton(cadUp,casDown);
    gctScrollButtonUpDisabled     : DrawArrowButton(cadUp,casDisabled);
    gctScrollButtonDown           : DrawArrowButton(cadDown,casUp);
    gctScrollButtonDownClicked    : DrawArrowButton(cadDown,casDown);
    gctScrollButtonDownDisabled   : DrawArrowButton(cadDown,casDisabled);

    gctScrollButtonLeft           : DrawArrowButton(cadLeft,casUp);
    gctScrollButtonLeftClicked    : DrawArrowButton(cadLeft,casDown);
    gctScrollButtonLeftDisabled   : DrawArrowButton(cadLeft,casDisabled);
    gctScrollButtonRight          : DrawArrowButton(cadRight,casUp);
    gctScrollButtonRightClicked   : DrawArrowButton(cadRight,casDown);
    gctScrollButtonRightDisabled  : DrawArrowButton(cadRight,casDisabled);

    gctCombo                      : DrawArrowButton(cadDown,casUp);
    gctComboClicked               : DrawArrowButton(cadDown,casDown);
    gctComboHighlight             : DrawArrowButton(cadDown,casUp,$0000FF);
    gctComboHighlightClicked      : DrawArrowButton(cadDown,casDown,$0000FF);
    gctRadioButton                : begin
      SetPenColor($000000);
      SetBrushColor($FFFFFF);
      Windows.Ellipse(FDC,x1, y1, x2, y2);
    end;
    gctRadioButtonChecked         : begin
      SetPenColor($000000);
      SetBrushColor($FFFFFF);
      Windows.Ellipse(FDC,x1, y1, x2, y2);
      SetBrushColor($000000);
      Windows.Ellipse(FDC,x1 + 3, y1 + 3, x2 - 3, y2 - 3)
    end;
  end;
end;

procedure TAXWGDI.DrawEdge(x1, y1, x2, y2: integer; EdgeType, EdgeFlags: longword);
var
  R: TRect;
begin
  CheckHandle;
  R := Rect(x1,y1,x2,y2);
  Windows.DrawEdge(FDC,R,EdgeType, EdgeFlags);
end;

procedure TAXWGDI.DrawF(const AX, AY, AWidth, AHeight: double; AGraphic: TObject);
begin
  raise XLSRWException.Create('TODO');
end;

procedure TAXWGDI.DrawF(const AX, AY: double; AGraphic: TObject);
begin
  raise XLSRWException.Create('TODO');
end;

procedure TAXWGDI.DrawText(Rect: TRect; Text: AxUCString; Options: TDrawTextOptions);
var
  H,D: integer;
  R: TRect;
  Opt: longword;
begin
  CheckHandle;
  Opt := 0;
  if dtoHCenter      in Options then Opt := Opt or DT_CENTER;
  if dtoLeft         in Options then Opt := Opt or DT_LEFT;
  if dtoRight        in Options then Opt := Opt or DT_RIGHT;
  if dtoTop          in Options then Opt := Opt or DT_TOP;
  if dtoWordBreak    in Options then Opt := Opt or DT_WORDBREAK;
  if dtoBottom in Options then begin
    R := Rect;
    H := Windows.DrawTextExW(FDC,PWideChar(Text),Length(Text),R,Opt or DT_CALCRECT,Nil);
    D := (Rect.Bottom - Rect.Top) - H;
    if D > 0 then begin
      Inc(Rect.Top,D);
      Inc(Rect.Bottom,D);
    end
  end
  else if (dtoVCenter in Options) or (dtoVDistributed in Options) then begin
    R := Rect;
    H := Windows.DrawTextExW(FDC,PWideChar(Text),Length(Text),R,Opt or DT_CALCRECT,Nil);
    D := (Rect.Bottom - Rect.Top) - H;
    if D > 0 then
      Inc(Rect.Top,D div 2);
  end;
  Windows.DrawTextExW(FDC,PWideChar(Text),Length(Text),Rect,Opt,Nil);
end;

procedure TAXWGDI.DrawXLSCurves(Curves: TXLSFigure);
var
  i,j: integer;
  CPoints: TXLSCurvePoints;
begin
  CheckHandle;
  for i := 0 to Curves.Count - 1 do begin
//    SetDirtyArea(Curves[i].Start.X,Curves[i].Start.Y,Curves[i].Start.X + FWidth,Curves[i].Start.Y - FHeight);
    Windows.MoveToEx(FDC,Curves[i].Start.X,Curves[i].Start.Y,Nil);
    for j := 0 to Curves[i].Count - 1 do begin
      CPoints := Curves[i][j];
      case CPoints.CurveType of
        xctLineTo:   Windows.LineTo(FDC,Curves[i].Start.X,Curves[i].Start.Y);
        xctPolyline: Windows.PolyLineTo(FDC,CPoints.Points^,CPoints.Count);
        xctBezier:   Windows.PolyBezierTo(FDC,CPoints.Points^,CPoints.Count);
      end;
    end;
  end;
end;

procedure TAXWGDI.Ellipse(x1,y1,x2,y2: integer);
begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2 + 1,y2 + 1);
  Windows.Ellipse(FDC,x1, y1, x2 + 1, y2 + 1);
end;

function TAXWGDI.EMUToPixels(const EMU: integer): integer;
begin
  Result := Round(PixelsPerInchY * (EMU / 914400));
end;

procedure TAXWGDI.FillRect(Rect: TXYRect);
begin
  CheckHandle;
  SetDirtyArea(Rect.X1,Rect.Y1,Rect.X2,Rect.Y2);
  Windows.FillRect(FDC,TRect(Rect),FHBrush);
end;

function TAXWGDI.FntPixHeightToPt(PixHeight: integer): integer;
begin
  CheckHandle;
  Result := Round((PixHeight * 72) / PixelsPerInchY);
end;

class procedure TAXWGDI.FreeResources;
begin
  FreeAndNil(bmpGray);
  FreeAndNil(bmpLightGray);
  FreeAndNil(bmpWarning);
  FreeAndNil(bmpComment);
  FreeAndNil(bmpCheckBox);
  FreeAndNil(bmpCondFmtIcons);
end;

function TAXWGDI.GetBrushColor: longword;
begin
  Result := RevRGB(FBrush.lbColor);
end;

function TAXWGDI.GetBrushSolid: boolean;
begin
  Result := FBrush.lbStyle = BS_SOLID;
end;

{
function TAXWGDI.GetGlyphCurves(C: WideChar; var Buf: Pointer): integer;
const
  // Missing in Windows.
  GGO_BEZIER = 3;
var
  GM: GLYPHMETRICS;
  Identity: MAT2;
begin
  CheckHandle;
  Identity.eM11.fract := 0;
  Identity.eM11.value := 1;
  Identity.eM12.fract := 0;
  Identity.eM12.value := 0;
  Identity.eM21.fract := 0;
  Identity.eM21.value := 0;
  Identity.eM22.fract := 0;
  Identity.eM22.value := 1;
  Result := Windows.GetGlyphOutlineW(FDC,Longword(C),GGO_BEZIER,GM,0,Nil,Identity);
  GetMem(Buf,Result);
  Windows.GetGlyphOutlineW(FDC,Longword(C),GGO_BEZIER,GM,Result,Buf,Identity);
end;
}
procedure TAXWGDI.DrawGlyphCurves(Buf: Pointer; Sz: integer);
const
  // Missing in Windows.
  TT_PRIM_CSPLINE = 3;
var
  i: integer;
  P: Pointer;
  pEnd: integer;
  pHeader,pNext: ^TTPOLYGONHEADER;
  pCurve: ^TTPOLYCURVE;
  Start,Pt: TPoint;
  Pts: array of TPoint;

procedure PATH_PTFXtoPT(ptfx: POINTFX);
begin
  Pt.x := 500 + ptfx.x.value;
  if ptfx.x.fract >= $8000 then
    Inc(Pt.x);
  Pt.y := 500 - ptfx.y.value;
  if ptfx.y.fract >= $8000 then
    Dec(Pt.Y);
end;

begin
  CheckHandle;
//  SetDirtyArea(500,500,600,600);

  P := Buf;
  pEnd := NativeInt(Buf) + Sz;

  // For each contour
  while NativeInt(P) < pEnd do begin
    pHeader := P;
    pNext := Pointer(Longword(P) + pHeader.cb);
    PATH_PTFXtoPT(pHeader.pfxStart);
    Start := Pt;
    Windows.MoveToEx(FDC,Pt.x,Pt.y,Nil);
    P := Pointer(NativeInt(P) + SizeOf(TTPOLYGONHEADER));
    // For each curve
    while NativeInt(P) < NativeInt(pNext) do begin
      pCurve := p;

      SetLength(Pts,pCurve.cpfx);
       // For each point
      for i := 0 to pCurve.cpfx - 1 do begin
        PATH_PTFXtoPT(pCurve.apfx[i]);
        Pts[i] := Pt;
      end;
      case pCurve.wType of
        TT_PRIM_LINE   : Windows.PolyLineTo(FDC,Pts[0],pCurve.cpfx);
        TT_PRIM_CSPLINE: Windows.PolyBezierTo(FDC,Pts[0],pCurve.cpfx);
        else raise XLSRWException.Create('Illegal curve type in glyph');
      end;
      P := Pointer(NativeInt(P) + SizeOf(TTPOLYCURVE) + (SizeOf(POINTFX) * (pCurve.cpfx - 1)));

    end;
    if (pt.x <> start.x) or (pt.y <> start.y) then
      Windows.LineTo(FDC,Start.X,Start.Y);
  end;
end;

procedure TAXWGDI.DrawText(X1, Y1, X2, Y2: double; Text: AxUCString; Options: TDrawTextOptions);
var
  R: TRect;
begin
  R.Left := Round(X1);
  R.Top := Round(Y1);
  R.Right := Round(X2);
  R.Bottom := Round(Y2);

  Drawtext(R,Text,Options);
end;

function TAXWGDI.GetTheDC: longword;
begin
  CheckHandle;
  Result := FDC;
end;

function TAXWGDI.GetTransparentMode: boolean;
begin
  CheckHandle;
  Result := Windows.GetBkMode(FDC) = TRANSPARENT;
end;

function TAXWGDI.GetFontColor: longword;
begin
  CheckHandle;
  Result := Windows.GetTextColor(FDC);
end;

function TAXWGDI.GetFontName: AxUCString;
begin
  CheckHandle;
  SetLength(Result,255);
  SetLength(Result,Windows.GetTextFaceW(FDC,255,PWideChar(Result)) - 1);
end;

function TAXWGDI.GetFontRotation(Handle: longword): double;
var
  LogRec: TLogFont;
begin
  GetObject(Handle, SizeOf(LogRec), Addr(LogRec));
  Result := LogRec.lfEscapement / 10;
end;

function TAXWGDI.GetFontSize: integer;
var
  LF: LOGFONT;
begin
  CheckHandle;
  GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LF),@LF);
  Result := -Round((LF.lfHeight * 72) / PixelsPerInchY);
end;

// Values adjusted for use with headers/footers
procedure TAXWGDI.GetFontSubscriptData(var FontOffset, FontSize: TPoint);
var
  OTM: OUTLINETEXTMETRICW;
begin
  CheckHandle;
  Windows.GetOutlineTextMetrics(FDC,SizeOf(OUTLINETEXTMETRICW),@OTM);
  FontOffset := OTM.otmptSubscriptOffset;
  FontOffset.X := 0;
  FontOffset.Y := OTM.otmTextMetrics.tmAscent - FontOffset.Y;
  FontSize := OTM.otmptSubscriptOffset;
end;

// Values adjusted for use with headers/footers
procedure TAXWGDI.GetFontSuperscriptData(var FontOffset, FontSize: TPoint);
var
  t: integer;
  OTM: OUTLINETEXTMETRICW;
begin
  CheckHandle;
  Windows.GetOutlineTextMetrics(FDC,SizeOf(OUTLINETEXTMETRICW),@OTM);
  FontOffset := OTM.otmptSuperscriptOffset;
  t := FontOffset.X;
  FontOffset.X := FontOffset.Y;
  FontOffset.Y := -t;
  FontSize := OTM.otmptSuperscriptSize;
end;

procedure TAXWGDI.GetGlyphCurves(C: WideChar; Curves: TXLSFigure; GlyphHeight,GlyphDescent: integer; XScale,YScale: double);
const
  // Missing in Windows.
  GGO_BEZIER = 3;
var
  GM: GLYPHMETRICS;
  Identity: MAT2;
  Sz: integer;
  Buf: Pointer;
begin
  CheckHandle;
  Identity.eM11.fract := Round(Frac(XScale) * $FFFF);
  Identity.eM11.value := Trunc(Int(XScale));
  Identity.eM12.fract := 0;
  Identity.eM12.value := 0;
  Identity.eM21.fract := 0;
  Identity.eM21.value := 0;
  Identity.eM22.fract := Round(Frac(YScale) * $FFFF);
  Identity.eM22.value := Trunc(Int(YScale));
  Sz := Windows.GetGlyphOutlineW(FDC,Longword(C),GGO_BEZIER,GM,0,Nil,Identity);
  GetMem(Buf,Sz);
  try
    Windows.GetGlyphOutlineW(FDC,Longword(C),GGO_BEZIER,GM,Sz,Buf,Identity);
    Curves.Add(Buf,Sz,GM.gmCellIncX + GM.gmptGlyphOrigin.X,GlyphHeight,GlyphDescent);
  finally
    FreeMem(Buf);
  end;
end;

function TAXWGDI.GetPenColor: longword;
begin
  Result := RevRGB(FPen.lopnColor);
end;

function TAXWGDI.GetPenSolid: boolean;
begin
  Result := FPen.lopnStyle = PS_SOLID;
end;

function TAXWGDI.GetPenWidth: longword;
begin
  Result := FPen.lopnWidth.X;
end;

function TAXWGDI.GetPenWidthF: double;
begin
  Result := FPen.lopnWidth.X;
end;

procedure TAXWGDI.GetTextMetric;
var
  OTM: TOUTLINETEXTMETRIC;
  TM: TTEXTMETRIC;
begin
  CheckHandle;

  // Don't works on True Type fonts
  if GetOutlineTextMetricsW(FDC,SizeOf(TOUTLINETEXTMETRIC),@OTM) <> 0 then begin
    FTM.Height := OTM.otmTextMetrics.tmHeight;
    FTM.Ascent := OTM.otmTextMetrics.tmAscent;
    FTM.Descent := OTM.otmTextMetrics.tmDescent;
    FTM.InternalLeading := OTM.otmTextMetrics.tmInternalLeading;
    FTM.AveCharWidth := OTM.otmTextMetrics.tmAveCharWidth;
    FTM.BreakChar := AxUCChar(OTM.otmTextMetrics.tmBreakChar);

    FTM.UnderscoreSize := OTM.otmsUnderscoreSize;
    FTM.UnderscorePosition := OTM.otmsUnderscorePosition;
    FTM.LineGap := OTM.otmLineGap;
  end
  else begin
    GetTextMetrics(FDC,TM);
    FTM.Height := TM.tmHeight;
    FTM.Ascent := TM.tmAscent;
    FTM.Descent := TM.tmDescent;
    FTM.InternalLeading := TM.tmInternalLeading;
    FTM.AveCharWidth := TM.tmAveCharWidth;
    FTM.BreakChar := AxUCChar(TM.tmBreakChar);

    FTM.UnderscoreSize := 1;
    FTM.UnderscorePosition := -Round(TM.tmHeight * 0.12);
    FTM.LineGap := Round(TM.tmHeight * 0.67);
  end;
end;

procedure TAXWGDI._GetTextMetricSys;
var
  OTM: TOUTLINETEXTMETRIC;
  TM: TTEXTMETRIC;
begin
  CheckHandle;

  // Don't works on True Type fonts
  if GetOutlineTextMetrics(FDC,SizeOf(TOUTLINETEXTMETRIC),@OTM) <> 0 then begin
    FTM_SYS.Height := OTM.otmTextMetrics.tmHeight;
    FTM_SYS.Ascent := OTM.otmTextMetrics.tmAscent;
    FTM_SYS.Descent := OTM.otmTextMetrics.tmDescent;
    FTM_SYS.InternalLeading := OTM.otmTextMetrics.tmInternalLeading;
    FTM_SYS.AveCharWidth := OTM.otmTextMetrics.tmAveCharWidth;
    FTM_SYS.BreakChar := AxUCChar(OTM.otmTextMetrics.tmBreakChar);

    FTM_SYS.UnderscoreSize := OTM.otmsUnderscoreSize;
    FTM_SYS.UnderscorePosition := OTM.otmsUnderscorePosition;
    FTM_SYS.LineGap := OTM.otmLineGap;
  end
  else begin
    GetTextMetrics(FDC,TM);
    FTM_SYS.Height := TM.tmHeight;
    FTM_SYS.Ascent := TM.tmAscent;
    FTM_SYS.Descent := TM.tmDescent;
    FTM_SYS.InternalLeading := TM.tmInternalLeading;
    FTM_SYS.AveCharWidth := TM.tmAveCharWidth;
    FTM_SYS.BreakChar := AxUCChar(TM.tmBreakChar);

    FTM_SYS.UnderscoreSize := 1;
    FTM_SYS.UnderscorePosition := -Round(TM.tmHeight * 0.12);
    FTM_SYS.LineGap := Round(TM.tmHeight * 0.67);
  end;
end;

function TAXWGDI.GetXorMode: boolean;
begin
  CheckHandle;
  Result := GetROP2(FDC) = R2_NOTXORPEN;
end;

procedure TAXWGDI.GradientFillRect(x1, y1, x2, y2: integer; StartColor, EndColor: longword; HorizontalFill: boolean);
var
  Vertex: array[0..1] of TRIVERTEX;
  GRect: GRADIENT_RECT;
begin
  CheckHandle;

  Vertex[0].x := x1;
  Vertex[0].y := y1;
  Vertex[0].Red :=  (StartColor and $00FF0000) shr 8;
  Vertex[0].Green := StartColor and $0000FF00;
  Vertex[0].Blue := (StartColor and $000000FF) shl 8;
  Vertex[0].Alpha := 0;
  Vertex[1].x := x2 + 1;
  Vertex[1].y := y2 + 1;
  Vertex[1].Red :=  (EndColor and $00FF0000) shr 8;
  Vertex[1].Green := EndColor and $0000FF00;
  Vertex[1].Blue := (EndColor and $000000FF) shl 8;
  Vertex[1].Alpha := 0;

  GRect.UpperLeft := 0;
  GRect.LowerRight := 1;

  // Declaration of GradientFill in Delphi 7 (Windows.pas) is wrong.
{$ifdef DELPHI_2006_OR_LATER}
  if HorizontalFill then
    Windows.GradientFill(FDC,@Vertex,2,@GRect,1,GRADIENT_FILL_RECT_H)
  else
    Windows.GradientFill(FDC,@Vertex,2,@GRect,1,GRADIENT_FILL_RECT_V);
{$else}
  BrushColor := EndColor;
  Rectangle(x1, y1, x2 + 1, y2 + 1);
{$endif}
end;

procedure TAXWGDI.Invalidate;
begin

end;

procedure TAXWGDI.InvalidateRect(x1,y1,x2,y2: integer);
begin
  SetDirtyArea(x1,y1,x2,y2);
end;

procedure TAXWGDI.Line(x1, y1, x2, y2: integer);
begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2,y2);
  Windows.MoveToEx(FDC,x1,y1,Nil);
  Windows.LineTo(FDC,x2,y2);
end;

procedure TAXWGDI.LineAA(x1, y1, x2, y2: integer);
begin
  Checkhandle;
  if FPen.lopnWidth.X = 1 then
    LineAADraw(x1,y1,x2,y2)
  else
    LineWide(x1,y1,x2,y2);
end;

procedure TAXWGDI.LineAADraw(x1, y1, x2, y2: integer);
var
   IntensityShift, ErrorAdj, ErrorAcc: word;
   ErrorAccTemp, Weighting, WeightingComplementMask: word;
   DeltaX, DeltaY, Temp, XDir,BaseColor,IntensityBits,NumLevels: integer;
// Antialiasing: Wu Algorithm - The Code Project - Fonts & GDI
begin
   SetDirtyArea(x1,y1,x2,y2);
   BaseColor := $000000;
   IntensityBits := 8;
   NumLevels := 256;
   // Make sure the line runs top to bottom */
   if y1 > y2 then begin
      Temp := y1;
      y1 := y2;
      y2 := Temp;
      Temp := x1;
      x1 := x2;
      x2 := Temp;
   end;
   // Draw the initial pixel, which is always exactly intersected by
   // the line and so needs no weighting
   SetPixelAA(x1, y1, BaseColor,0);

   DeltaX := x2 - x1;
   if DeltaX >= 0 then
     XDir := 1
   else begin
     XDir := -1;
    DeltaX := -DeltaX; // make DeltaX positive */
   end;
   // Special-case horizontal, vertical, and diagonal lines, which
   // require no weighting because they go right through the center of
   // every pixel
   DeltaY := y2 - y1;
   if DeltaY = 0 then begin
      //* Horizontal line */
      while DeltaX <> 0 do begin
         Inc(x1,XDir);
         SetPixelAA(x1, y1, BaseColor,0);
         Dec(DeltaX);
      end;
      Exit;
   end;
   if DeltaX = 0 then begin
     // Vertical line
     repeat
       Inc(y1);
       SetPixelAA(x1, y1, BaseColor,0);
       Dec(DeltaY);
     until (DeltaY <> 0);
      Exit;
   end;
   if DeltaX = DeltaY then begin
     // Diagonal line
     repeat
       Inc(x1,XDir);
       Inc(y1);
       SetPixelAA(x1, y1, BaseColor,0);
       Dec(DeltaY);
     until (DeltaY <> 0);
     Exit;
   end;
   // Line is not horizontal, diagonal, or vertical
   ErrorAcc := 0;  // initialize the line error accumulator to 0
   // # of bits by which to shift ErrorAcc to get intensity level
   IntensityShift := 16 - IntensityBits;
   // Mask used to flip all bits in an intensity weighting, producing the
   // result (1 - intensity weighting)
   WeightingComplementMask := NumLevels - 1;
   // Is this an X-major or Y-major line?
   if DeltaY > DeltaX then begin
      // Y-major line; calculate 16-bit fixed-point fractional part of a
      // pixel that X advances each time Y advances 1 pixel, truncating the
      // result so that we won't overrun the endpoint along the X axis
      ErrorAdj := (DeltaX shl 16) div DeltaY;
      // Draw all pixels other than the first and last
      while DeltaY > 0 do begin
        ErrorAccTemp := ErrorAcc;   // remember currrent accumulated error
        Inc(ErrorAcc, ErrorAdj);    // calculate error for next pixel
        if ErrorAcc <= ErrorAccTemp then
           // The error accumulator turned over, so advance the X coord
           Inc(x1,XDir);
        Inc(y1); // Y-major, so always advance Y
        // The IntensityBits most significant bits of ErrorAcc give us the
        // intensity weighting for this pixel, and the complement of the
        // weighting for the paired pixel */
        Weighting := ErrorAcc shr IntensityShift;
        SetPixelAA(x1, y1, BaseColor + Weighting,0);

//         DrawPixel(Canvas.Handle,x1 + XDir, y1, BaseColor + (Weighting ^ WeightingComplementMask));
        SetPixelAA(x1 + XDir, y1, BaseColor + (Weighting xor WeightingComplementMask),0);
        Dec(DeltaY);
      end;
      // Draw the final pixel, which is always exactly intersected by the line
      // and so needs no weighting
      SetPixelAA(x2, y2, BaseColor,0);
      Exit;
   end;
   // It's an X-major line; calculate 16-bit fixed-point fractional part of a
   // pixel that Y advances each time X advances 1 pixel, truncating the
   // result to avoid overrunning the endpoint along the X axis
   ErrorAdj := (DeltaY shl 16) div DeltaX;
   // Draw all pixels other than the first and last */
   while DeltaX > 0 do begin
     ErrorAccTemp := ErrorAcc;  // remember currrent accumulated error */
     Inc(ErrorAcc,ErrorAdj);      // calculate error for next pixel */
     if ErrorAcc <= ErrorAccTemp then
        // The error accumulator turned over, so advance the Y coord */
       Inc(y1);
     Inc(x1,XDir); // X-major, so always advance X */
     // The IntensityBits most significant bits of ErrorAcc give us the
     // intensity weighting for this pixel, and the complement of the
     // weighting for the paired pixel */
     Weighting := ErrorAcc shr IntensityShift;
     SetPixelAA(x1, y1, BaseColor + Weighting,0);
//      DrawPixel(Canvas.Handle,x1, y1 + 1,BaseColor + (Weighting ^ WeightingComplementMask));
     SetPixelAA(x1, y1 + 1,BaseColor + (Weighting xor WeightingComplementMask),0);
     Dec(DeltaX);
   end;
   // Draw the final pixel, which is always exactly intersected by the line
   //   and so needs no weighting */
   SetPixelAA(x2, y2, BaseColor,0);
end;

procedure TAXWGDI.LineF(x1, y1, x2, y2: double);
begin
  Line(Round(x1),Round(y1),Round(x2),Round(y2));
end;

procedure TAXWGDI.LineStyled(x1, y1, x2, y2: integer; PenColor,BackColor: longword; Style: TPaintLineStyle; Horiz: boolean; Diag: boolean = False);

procedure DrawStyled(LineStyle: array of boolean);
var
  i: integer;
  x,y: integer;
begin
  SetDirtyArea(x1,y1,x2,y2);
  i := 0;
  if Horiz then begin
    for x := x1 to x2 do begin
      if LineStyle[i] then
        Windows.SetPixelV(FDC,x,y1,PenColor)
      else
        Windows.SetPixelV(FDC,x,y1,BackColor);
      Inc(i);
      if i > High(LineStyle) then
        i := 0;
    end;
  end
  else begin
    for y := y1 to y2 do begin
      if LineStyle[i] then
        Windows.SetPixelV(FDC,x1,y,PenColor)
      else
        Windows.SetPixelV(FDC,x1,y,BackColor);
      Inc(i);
      if i > High(LineStyle) then
        i := 0;
    end;
  end;
end;

begin
  CheckHandle;

  SetPenColor(PenColor);

  if Diag then begin
    SetBrushColor(BackColor);

    case Style of
      plsNone: ;
      plsThin            : ;
      plsMedium          : PenWidth := 2;
      plsDashed          : CreatePen(PenColor,PS_DASH,1);
      plsDotted          : CreatePen(PenColor,PS_DOT,1);
      plsThick           : PenWidth := 3;
      plsDouble          : begin
        Line(x1 - 1,y1,x2 - 1,y2);
        Line(x1 + 1,y1,x2 + 1,y2);
        Exit;
      end;
      plsHair            : ;
      plsMediumDashed    : CreatePen(PenColor,PS_DASH,2);
      plsDashDot         : CreatePen(PenColor,PS_DASHDOT,1);
      plsMediumDashDot   : CreatePen(PenColor,PS_DASHDOT,2);
      plsDashDotDot      : CreatePen(PenColor,PS_DASHDOTDOT,1);
      plsMediumDashDotDot: CreatePen(PenColor,PS_DASHDOTDOT,2);
      plsSlantedDashDot  : PenWidth := 2;
      pls_Clear_: ;
      pls_MediumDotted_: ;
    end;

    Line(x1,y1,x2,y2);

    PenWidth := 1;
    SetPenSolid(True);

    Exit;
  end;

  case Style of
    pls_Clear_: Exit;
    plsNone,plsThin: begin
      Line(x1,y1,x2,y2);
    end;
    plsMedium: begin
      if Horiz then begin
        Line(x1,y1,x2,y1);
        Line(x1,y2,x2,y2);
      end
      else begin
        Line(x1,y1,x1,y2);
        Line(x2,y1,x2,y2);
      end;
    end;
    plsDashed: begin
      SetBrushColor(BackColor);
      CreatePen(PenColor,PS_DASH,1);
      Line(x1,y1,x2,y2);
      SetPenSolid(True);
//      DrawStyled([True,True,True,False]);
    end;
    plsDotted: begin
      SetBrushColor(BackColor);
      CreatePen(PenColor,PS_DOT,1);
      Line(x1,y1,x2,y2);
      SetPenSolid(True);
//      DrawStyled([True,True,False,False]);
    end;
    plsThick: begin
      if Horiz then begin
        Line(x1,y1,x2,y1);
        Line(x1,y1 + 1,x2,y1 + 1);
        Line(x1,y1 + 2,x2,y1 + 2);
      end
      else begin
        Line(x1,y1,x1,y2);
        Line(x1 + 1,y1,x1 + 1,y2);
        Line(x1 + 2,y1,x1 + 2,y2);
      end;
    end;
    plsDouble: begin
      if Horiz then begin
        Line(x1,y1,x1,y2);
        PenColor := BackColor;
        Line(x1 + 1,y1,x1 + 1,y2);
        PenColor := PenColor;
        Line(x1 + 2,y1,x1 + 2,y2);
      end
      else begin
        Line(x1,y1,x1,y2);
        PenColor := BackColor;
        Line(x1 + 1,y1,x1 + 1,y2);
        PenColor := PenColor;
        Line(x1 + 2,y1,x1 + 2,y2);
      end;
    end;
    plsHair:
      DrawStyled([True,False]);
    pls_MediumDotted_: begin
      SetXorMode(True);

      CreatePen(PenColor,PS_DOT,1);

      if Horiz then begin
        Line(x1,y1,x2,y1);
        Line(x1,y2,x2,y2);
      end
      else begin
        Line(x1,y1,x1,y2);
        Line(x2,y1,x2,y2);
      end;

      SetXorMode(False);

      SetPenSolid(True);
    end;
    plsMediumDashed: begin
//      SetBrushColor(BackColor);
//
//      DrawStyled([True,True,True,False]);
//      if Horiz then
//        Inc(y1)
//      else
//        Inc(x1);
//
//      Line(x1,y1,x2,y2);
//      DrawStyled([True,True,True,False]);

      CreatePen(PenColor,PS_DASH,1);

      if Horiz then begin
        Line(x1,y1,x2,y1);
        Line(x1,y2,x2,y2);
      end
      else begin
        Line(x1,y1,x1,y2);
        Line(x2,y1,x2,y2);
      end;

      SetPenSolid(True);
    end;
    plsDashDot: begin
      SetBrushColor(BackColor);
      CreatePen(PenColor,PS_DASHDOT,1);
      Line(x1,y1,x2,y2);
      SetPenSolid(True);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False]);
    end;
    plsMediumDashDot: begin
      SetBrushColor(BackColor);
      CreatePen(PenColor,PS_DASHDOT,1);

      Line(x1,y1,x2,y2);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False]);
      if Horiz then
        Inc(y1)
      else
        Inc(x1);

      Line(x1,y1,x2,y2);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False]);
      SetPenSolid(True);
    end;
    plsDashDotDot: begin
      SetBrushColor(BackColor);
      CreatePen(PenColor,PS_DASHDOTDOT,1);
      Line(x1,y1,x2,y2);
      SetPenSolid(True);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False,True,False,False,False]);
    end;
    plsMediumDashDotDot: begin
      SetBrushColor(BackColor);
      CreatePen(PenColor,PS_DASHDOTDOT,1);

      Line(x1,y1,x2,y2);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False,True,False,False,False]);
      if Horiz then
        Inc(y1)
      else
        Inc(x1);

      Line(x1,y1,x2,y2);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False,True,False,False,False]);
      SetPenSolid(True);
    end;
    plsSlantedDashDot: begin
      DrawStyled([False,True,True,True,True,True,True,True,True,True,True,True,False,True,True,True,True,True]);
      if Horiz then
        Inc(y1)
      else
        Inc(x1);
      DrawStyled([True,True,True,True,True,True,True,True,True,True,False,False,True,True,True,True,False,False]);
    end;
  end;
end;

procedure TAXWGDI.LineTo(x, y: integer);
begin
  CheckHandle;
  SetDirtyArea(x,y);
  Windows.LineTo(FDC,x,y);
end;

procedure TAXWGDI.LineToF(x, y: double);
begin
  CheckHandle;
  SetDirtyArea(Round(x),Round(y));
  Windows.LineTo(FDC,Round(x),Round(y));
end;

procedure TAXWGDI.LineWide(x1, y1, x2, y2: integer);
var
  x,y,W:integer;
  d0,d1:integer;  {difference terms d0=perpendicular to line, d1=along line}
  dd:integer;     {distance along line}
  u,v:integer;    {u=x2-x1;  v=y2-y1;}
  p:integer;    {pel counters, p=along line, q=perpendicular to line}
  ku,kv,kd,{ks,}kt:integer;  {loop constants}
  XInc,YInc: integer;
  tk:single;               {thickness threshold}

procedure parallelline(px,py,d1: integer);
begin  {implements Figure 5B}
  p := 0;
  d1 := -d1;
  while (p <= u) do begin {test for end of parallel line}
    SetPixel(px,py,FPen.lopnColor);
    if d1 <= kt then begin {square move}
      Inc(px,XInc);
      d1 := d1 + kv;
    end
    else begin {diagonal move}
      Inc(px,XInc);
      Inc(py,YInc);
      d1 := d1 + kd;
   end;
   Inc(p);
  end;
end; {perpendicular}

function normconstant(u,v:integer): single;
var choice:integer;
begin
    u:=abs(u); v:=abs(v);
    choice:=3;   {change this to suit taste/compute power etc.}
    case choice of
      1: Result:=u+v/4;    {12% thickness error - uses add and shift only}
      2: begin             {2.7% thickness error, uses compare, add and shift only}
             if (v+v+v)>u then Result:=u-(u/8)+v/2
                          else Result:=u+v/8;
         end;
      3: Result:=sqrt(u*u+v*v);  {ideal}
      else
         Result := 0;
    end; {case}
end;

begin  {implements figure 5A }
   W := 12; // FPen.lopnWidth.X;
   SetDirtyArea(x1 - W,y1 - W,x2 + W,y2 + W);
   {Initialisation}
   u := x2 - x1;   {delta x}
   v := y2 - y1;   {delta y}
   if u < 0 then begin
     u := -u;
     XInc := -1;
   end
   else
     XInc := 1;
   if v < 0 then begin
     v := -v;
     YInc := -1;
   end
   else
     YInc := 1;

   ku := u + u;     {change in l for square shift}
   kv := v + v;     {change in d for square shift}
   kd := kv - ku;   {change in d for diagonal shift}
   kt := u - kv;    {diag/square decision threshold}
   tk := 2 * W * normconstant(u,v);  {used here for constant thickness line}
   d0 := 0;
   d1 := 0;
   dd := 0;
   x := x1;
   y := y1;
   while (dd < tk) do begin {outer loop, stepping perpendicular to line}
     parallelline(x,y,d1);
     if d0 < kt then {square move}
       Inc(y,YInc)
     else begin {diagonal move}
      dd := dd + kv;
      d0 := d0 - ku;
      if d1 < kt then begin  {normal diagonal}
        Dec(x,XInc);
        Inc(y,YInc);
        d1 := d1 + kv;
      end
      else begin {double square move, extra parallel line}
        Dec(x,XInc);
        d1 := d1 + kd;
        if dd > tk then
          Exit;    {breakout on the extra line}
        parallelline(x,y,d1);
        Inc(y,YInc);
      end;
     end;
     dd := dd + ku;
     d0 := d0 + kv;
   end;
end;

procedure TAXWGDI.LineWide2(x1, y1, x2, y2: integer);
var
  LB: TLOGBRUSH;
  W: integer;
begin
  CheckHandle;
  W := FPen.lopnWidth.X;
  SetDirtyArea(x1 - W,y1 - W,x2 + W,y2 + W);

  LB.lbStyle := BS_SOLID;
  LB.lbColor := FPen.lopnColor;
  LB.lbHatch := 0;
  FHPen := ExtCreatePen(PS_GEOMETRIC or PS_SOLID or PS_ENDCAP_FLAT,FPen.lopnWidth.X,LB,0,Nil);
  DeleteObject(SelectObject(FDC,FHPen));

  Windows.MoveToEx(FDC,x1,y1,Nil);
  Windows.LineTo(FDC,x2,y2);
end;

class procedure TAXWGDI.LoadResources;
begin
  crCell        := Windows.LoadCursor(HInstance,'CELLCURSOR');
  crColSelect   := Windows.LoadCursor(HInstance,'COLSELCURSOR');
  crFill        := Windows.LoadCursor(HInstance,'FILLCURSOR');
  crHorizSplit  := Windows.LoadCursor(HInstance,'HSPLITCURSOR');
  crHoriz2Split := Windows.LoadCursor(HInstance,'HSPLIT2CURSOR');
  crRowSelect   := Windows.LoadCursor(HInstance,'ROWSELCURSOR');
  crVertSplit   := Windows.LoadCursor(HInstance,'VSPLITCURSOR');
  crVert2Split  := Windows.LoadCursor(HInstance,'VSPLIT2CURSOR');
  crCenter      := Windows.LoadCursor(HInstance,'CENTERCURSOR');
  crMoveCells   := Windows.LoadCursor(HInstance,'MOVECELLSCURSOR');
  crSizeCells   := Windows.LoadCursor(HInstance,'SIZECELLSCURSOR');
  crNeSwArrow   := Windows.LoadCursor(HInstance,'NESWCURSOR');
  crNwSeArrow   := Windows.LoadCursor(HInstance,'NWSECURSOR');

  bmpGray := TBitmap.Create;
  bmpGray.LoadFromResourceName(HInstance,'BMPGRAY');

  bmpLightGray := TBitmap.Create;
  bmpLightGray.LoadFromResourceName(HInstance,'BMPLIGHTGRAY');

  bmpWarning := TBitmap.Create;
  bmpWarning.LoadFromResourceName(HInstance,'BMPWARNING');

  bmpComment := TBitmap.Create;
  bmpComment.LoadFromResourceName(HInstance,'BMPNOTE');

  bmpCheckBox := TBitmap.Create;
  bmpCheckBox.LoadFromResourceName(HInstance,'BMPCHECKBOX');

  bmpCondFmtIcons := TBitmap.Create;
  bmpCondFmtIcons.LoadFromResourceName(HInstance,'BMPCONDFMTICONS_2');
  bmpCondFmtIcons.Transparent := True;
  bmpCondFmtIcons.TransparentColor := clFuchsia;
end;

procedure TAXWGDI.MoveTo(x, y: integer);
begin
  CheckHandle;
  SetDirtyArea(x,y);
  Windows.MoveToEx(FDC,x,y,Nil);
end;

procedure TAXWGDI.MoveToF(x, y: double);
begin
  CheckHandle;
  SetDirtyArea(Round(x),Round(y));
  Windows.MoveToEx(FDC,Round(x),Round(y),Nil);
end;

procedure TAXWGDI.PaintBMPImage(Bitmap: TBitmap; X1, Y1, X2, Y2: integer);
var
  HDCBMP: HDC;
{$ifdef DELPHI_2009_OR_LATER}
  Blendfunc: TBLENDFUNCTION;
{$endif}
begin
  CheckHandle;
  SetDirtyArea(X1,Y1,X2,Y2);

  HDCBMP := Windows.CreateCompatibleDC(FDC);
  Windows.SelectObject(HDCBMP,Bitmap.Handle);

  Windows.SetStretchBltMode(FDC,HALFTONE);
{$ifdef DELPHI_2009_OR_LATER}
  if Bitmap.AlphaFormat > afIgnored then begin
    Blendfunc.BlendOp := AC_SRC_OVER;
    Blendfunc.BlendFlags := 0;
    Blendfunc.SourceConstantAlpha := $FF;
    BlendFunc.AlphaFormat := AC_SRC_ALPHA;

    Windows.SetBrushOrgEx(FDC,0,0,Nil);
    Windows.AlphaBlend(FDC,X1,Y1,X2 - X1,Y2 - Y1,HDCBMP,0,0,Bitmap.Width,Bitmap.Height,Blendfunc);
  end
  else
{$endif}
    Windows.StretchBlt(FDC,X1,Y1,X2 - X1,Y2 - Y1,HDCBMP,0,0,Bitmap.Width,Bitmap.Height,SRCCOPY);
  Windows.DeleteDC(HDCBMP);
end;

procedure TAXWGDI.PaintCondFmtIcon(const AStyle: TXc12IconSetType; const AIndex, AX, AY: integer);
var
  MaskDC: HDC;
  Save: THandle;
  XOffs,YOffs: integer;
  XSpace,YSpace: integer;
{$ifdef _DELPHI_2006_OR_LATER}
  Blendfunc: TBLENDFUNCTION;
{$endif}
begin
  CheckHandle;
  SetDirtyArea(AX,AY,AX + 16,AY + 16);

  XSpace := 8;
  YSpace := 10;
  XOffs := 7;
  YOffs :=  5;
  case AStyle of
    x12ist3Arrows        : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,0 * (16 + YSpace));
    end;
    x12ist3ArrowsGray    : begin
      Inc(XOffs,125 + AIndex * (16 + XSpace));
      Inc(YOffs,0 * (16 + YSpace));
    end;
    x12ist3Flags         : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,8 * (16 + YSpace));
    end;
    x12ist3TrafficLights1: begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,4 * (16 + YSpace));
    end;
    x12ist3TrafficLights2: begin
      Inc(XOffs,125 + AIndex * (16 + XSpace));
      Inc(YOffs,4 * (16 + YSpace));
    end;
    x12ist3Signs         : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,2 * (16 + YSpace));
    end;
    x12ist3Symbols       : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,7 * (16 + YSpace));
    end;
    x12ist3Symbols2      : begin
      Inc(XOffs,125 + AIndex * (16 + XSpace));
      Inc(YOffs,7 * (16 + YSpace));
    end;
    x12ist4Arrows        : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,2 * (16 + YSpace));
    end;
    x12ist4ArrowsGray    : begin
      Inc(XOffs,125 + AIndex * (16 + XSpace));
      Inc(YOffs,1 * (16 + YSpace));
    end;
    x12ist4RedToBlack    : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,6 * (16 + YSpace));
    end;
    x12ist4Rating        : begin
      Inc(XOffs,126 + AIndex * (16 + XSpace));
      Inc(YOffs,9 * (16 + YSpace));
    end;
    x12ist4TrafficLights : begin
      Inc(XOffs,126 + AIndex * (16 + XSpace));
      Inc(YOffs,-1 + 5 * (16 + YSpace));
    end;
    x12ist5Arrows        : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,3 * (16 + YSpace));
    end;
    x12ist5ArrowsGray    : begin
      Inc(XOffs,125 + AIndex * (16 + XSpace));
      Inc(YOffs,2 * (16 + YSpace));
    end;
    x12ist5Rating        : begin
      Inc(XOffs,126 + AIndex * (16 + XSpace));
      Inc(YOffs,10 * (16 + YSpace));
    end;
    x12ist5Quarters      : begin
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,10 * (16 + YSpace));
    end;
    else begin // Icon set new in Excel 2010. "4 squares".
      Inc(XOffs,AIndex * (16 + XSpace));
      Inc(YOffs,11 * (16 + YSpace));
    end;
  end;

{$ifdef _DELPHI_2006_OR_LATER}
    HDCBMP := Windows.CreateCompatibleDC(FDC);
    Windows.SelectObject(HDCBMP,bmpCondFmtIcons.Handle);

    Windows.SetStretchBltMode(FDC,HALFTONE);

    Blendfunc.BlendOp := AC_SRC_OVER;
    Blendfunc.BlendFlags := 0;
    Blendfunc.SourceConstantAlpha := $7F;
    BlendFunc.AlphaFormat := 0;//AC_SRC_ALPHA;

    Windows.SetBrushOrgEx(FDC,0,0,Nil);
    Windows.AlphaBlend(FDC,AX,AY,16,16,HDCBMP,XOffs,YOffs,16,16,Blendfunc);

    Windows.DeleteDC(HDCBMP);
{$else}
    MaskDC := Windows.CreateCompatibleDC(0);
    Save := SelectObject(MaskDC, bmpCondFmtIcons.MaskHandle);

    TransparentStretchBlt(FDC,AX,AY,16,16,bmpCondFmtIcons.Canvas.Handle,XOffs,YOffs,16,16,MaskDC,XOffs,YOffs);

    Windows.SelectObject(MaskDC, Save);
    Windows.DeleteDC(MaskDC);
//    Windows.BitBlt(FDC,AX,AY,16,16,bmpCondFmtIcons.Canvas.Handle,XOffs,YOffs,SRCCOPY);
{$endif}
end;

procedure TAXWGDI.PaintEMFImage(Buf: Pointer; BufSize, X1, Y1, X2, Y2: integer);
var
  HM: HENHMETAFILE;
begin
  CheckHandle;
  SetDirtyArea(X1,Y1,X2,Y2);
  HM := SetEnhMetaFileBits(BufSize,Buf);
  if HM <> 0 then begin
    PlayEnhMetaFile(FDC,HM,Rect(X1,Y1,X2,Y2));
    DeleteEnhMetaFile(HM);
  end;
end;

procedure TAXWGDI.PaintWMFImage(Buf: Pointer; BufSize, X1, Y1, X2, Y2: integer);
var
  HM: HENHMETAFILE;
  MP: METAFILEPICT;
begin
  CheckHandle;
  SetDirtyArea(X1,Y1,X2,Y2);
  MP.mm := MM_ANISOTROPIC;
  MP.xExt := 100;
  MP.yExt := 100;
  HM := SetWinMetaFileBits(BufSize,Buf,0,MP);
  if HM <> 0 then begin
    PlayEnhMetaFile(FDC,HM,Rect(X1,Y1,X2,Y2));
    DeleteEnhMetaFile(HM);
  end;
end;

procedure TAXWGDI.PizzaSlice(const x, y, r: integer; const AStartAng, AEndAng: double);
var
  x1,y1,x2,y2: integer;
  x3,y3,x4,y4: integer;
begin
  x1 := x - r;
  y1 := y - r;
  x2 := x + r;
  y2 := y + r;

  X3 := x + Round(r * Cos(Pi * 2 - AStartAng));
  Y3 := y - Round(r * Sin(Pi * 2 - AStartAng));

  X4 := x + Round(r * Cos(Pi * 2 - AEndAng));
  Y4 := y - Round(r * Sin(Pi * 2 - AEndAng));

  CheckHandle;
  SetDirtyArea(x1,y1,x2,y2);
  Windows.Pie(FDC,x1,y1,x2,y2,x3,y3,x4,y4);
end;

procedure TAXWGDI.PizzaSliceF(const x, y, r: double; const AStartAng, AEndAng: double);
begin
  PizzaSlice(Round(x),Round(y),Round(r),AStartAng, AEndAng);
end;

procedure TAXWGDI.Pixel(x, y: integer);
begin
  CheckHandle;
  SetDirtyArea(x,y);
  Windows.SetPixel(FDC,x,y,FPen.lopnColor)
end;

function TAXWGDI.PixelsPerCm: integer;
begin
  CheckHandle;
  Result := Round((PixelsPerInchY) / 2.54);
end;

function TAXWGDI.PixelsPerInchX: integer;
begin
  CheckHandle;
  Result := Round(GetDeviceCaps(FDC, LOGPIXELSX) * FZoom);
end;

function TAXWGDI.PixelsPerInchY: integer;
begin
  CheckHandle;
  Result := Round(GetDeviceCaps(FDC, LOGPIXELSY) * FZoom);
end;

//  Result := PixelsPerInch * (EMU / 914400);
function TAXWGDI.PixelsToEMU(const Pixels: integer): integer;
begin
  Result := Round((Pixels / PixelsPerInchY) * 914400);
end;

function TAXWGDI.PixToPtX(const APixels: integer): double;
begin
  Result := APixels / (PixelsPerInchX / 72);
end;

function TAXWGDI.PixToPtY(const APixels: integer): double;
begin
  Result := APixels / (PixelsPerInchY / 72);
end;

function TAXWGDI.PtToPixX(const APoints: double): integer;
begin
  Result := Round(APoints * (PixelsPerInchX / 72));
end;

function TAXWGDI.PtToPixY(const APoints: double): integer;
begin
  Result := Round(APoints * (PixelsPerInchY / 72));
end;

procedure TAXWGDI.PolyBezier(Points: array of TPoint);
begin
  CheckHandle;
  SetDirtyArea(Points);
  Windows.PolyBezier(FDC,PPoints(@Points)^,Length(Points));
end;

procedure TAXWGDI.PolyBezierTo(Points: array of TPoint);
begin
  CheckHandle;
  SetDirtyArea(Points);
  Windows.PolyBezierTo(FDC,PPoints(@Points)^,Length(Points));
end;

procedure TAXWGDI.Polygon(Points: array of TPoint);
begin
  CheckHandle;
  SetDirtyArea(Points);
  Windows.Polygon(FDC,PPoints(@Points)^,Length(Points));
end;

procedure TAXWGDI.Polyline(Points: array of TPoint);
begin
  CheckHandle;
  SetDirtyArea(Points);
  Windows.Polyline(FDC,PPoints(@Points)^,Length(Points));
end;

procedure TAXWGDI.PolylineTo(Points: array of TPoint);
begin
  CheckHandle;
  SetDirtyArea(Points);
  Windows.PolylineTo(FDC,PPoints(@Points)^,Length(Points));
end;

procedure TAXWGDI.Rectangle(x1, y1, x2, y2: integer);
begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2 + 1,y2 + 1);
  Windows.Rectangle(FDC,x1, y1, x2 + 1, y2 + 1);
end;

procedure TAXWGDI.RecallPen;
begin
  Move(FSavedPen,FPen,SizeOf(LOGPEN));
  FHPen := CreatePenIndirect(FPen);
  DeleteObject(SelectObject(FDC,FHPen));
end;

procedure TAXWGDI.Rectangle(Rect: TXYRect);
begin
  CheckHandle;
  Rectangle(Rect.x1,Rect.y1,Rect.x2,Rect.y2);
end;

procedure TAXWGDI.RectangleF(x1, y1, x2, y2: double);
begin
  CheckHandle;
  SetDirtyArea(Round(x1),Round(y1),Round(x2 + 1),Round(y2 + 1));
  Windows.Rectangle(FDC,Round(x1),Round(y1),Round(x2 + 1),Round(y2 + 1));
end;

procedure TAXWGDI.ReleaseHandle;
begin
  ReleaseDC(FHWND,FDC);
  FDC := 0;
  ResetLimits;
end;

procedure TAXWGDI.PaintSelectBMP(x1, y1, x2, y2: integer);
begin
  CheckHandle;
  Windows.StretchBlt(FDC,x1,y1,x2 - x1 + 1,y2 - y1 + 1,FFillBmp.Canvas.Handle,0,0,FFillBmp.Width,FFillBmp.Height,SRCAND);
end;

procedure TAXWGDI.PaintStockBMP(StockBMP: TStockXBookBitmap; X, Y: integer);
begin
  CheckHandle;
  case StockBMP of
    sxbCommentMark: begin
      SetDirtyArea(X - bmpComment.Width + 1,Y,X - bmpComment.Width + 1 + bmpComment.Width,Y + bmpComment.Height);
      BitBlt(FDC,X - bmpComment.Width + 1,Y,bmpComment.Width,bmpComment.Height,bmpComment.Canvas.Handle,0,0,SRCCOPY);
    end;
    sxbCheckBox: begin
      SetDirtyArea(X,Y,X + bmpCheckBox.Width,Y + bmpCheckBox.Height);
      BitBlt(FDC,X,Y,bmpCheckBox.Width,bmpCheckBox.Height,bmpCheckBox.Canvas.Handle,0,0,SRCCOPY);
    end;
  end;
end;

procedure TAXWGDI.PaintTabMoveBmp(X, Y: integer);
begin

end;

procedure TAXWGDI.Render(Dest: TAXWGDI);
var
  x1,y1,x2,y2: integer;
begin
  CheckHandle;
  x1 := FMinX;
  y1 := FMinY;
  x2 := FMaxX + 1;
  y2 := FMaxY + 1;
  Windows.BitBlt(Dest.DC,x1,y1,x2 - x1,y2 - y1,FDC,x1,y1,SRCCOPY);
end;

procedure TAXWGDI.RenderDirty(Dest: TAXWGDI);
var
  x1,y1,x2,y2: integer;
begin
  if FDirty then begin
    CheckHandle;
    x1 := FMinX;
    y1 := FMinY;
    x2 := FMaxX + 1;
    y2 := FMaxY + 1;
    Windows.BitBlt(Dest.DC,x1,y1,x2 - x1,y2 - y1,FDC,x1,y1,SRCCOPY);
    FDirty := False;
  end;
//  DebugDumpBMP(2);
end;

procedure TAXWGDI.ResetLimits;
begin
  FMinX := MAXINT;
  FMinY := MAXINT;
  FMaxX := -(MAXINT - 1);
  FMaxY := -(MAXINT - 1);
  FDirty := False;
end;

procedure TAXWGDI.RoundRect(x1, y1, x2, y2: integer; Adjust: double);
var
  Roundiness: integer;
begin
  CheckHandle;
  SetDirtyArea(x1,y1,x2 + 1,y2 + 1);
  Roundiness := Round(Min(x2 - x1,y2 - y1) * Adjust);
  Windows.RoundRect(FDC,x1, y1, x2 + 1, y2 + 1,Roundiness,Roundiness);
end;

procedure TAXWGDI.SavePen;
begin
  Move(FPen,FSavedPen,SizeOf(LOGPEN));
end;

procedure TAXWGDI.Scroll(X1, Y1, X2, Y2, dX, dY: integer);
begin
  CheckHandle;
  Windows.BitBlt(FDC,X1 + dX,Y1 + dY,X2 - X1,Y2 - Y1,FDC,X1,Y1,SRCCOPY);
end;

procedure TAXWGDI.SelectClipRect(ClipHandle: longword);
begin
  CheckHandle;
  Windows.SelectClipRgn(FDC,ClipHandle);
end;

procedure TAXWGDI.SelectPatternBrush(PatternBrush: TStockPatternBrush);
begin
  CheckHandle;
  case PatternBrush of
    spbNone     : FHBrush := CreateBrushIndirect(FBrush);
    spbGray     : FHBrush := CreatePatternBrush(bmpGray.Handle);
    spbLightGray: FHBrush := CreatePatternBrush(bmpLightGray.Handle);
    spbWarning  : FHBrush := CreatePatternBrush(bmpWarning.Handle);
  end;
  DeleteObject(SelectObject(FDC,FHBrush));
end;

procedure TAXWGDI.SelectFont(FontHandle: longword);
var
  Res: longword;
begin
  CheckHandle;
  if FFontHandle <> 0 then begin
    DeleteObject(FFontHandle);
    FFontHandle := 0;
  end;

  Res := SelectObject(FDC,FontHandle);
  Assert(Res <> 0);
end;

procedure TAXWGDI.SetArcDirection(CCW: boolean);
begin
  CheckHandle;
  if CCW then
    Windows.SetArcDirection(FDC,AD_COUNTERCLOCKWISE)
  else
    Windows.SetArcDirection(FDC,AD_CLOCKWISE);
end;

procedure TAXWGDI.SetBrushColor(const Value: longword);
begin
  CheckHandle;
  FBrush.lbColor := RevRGB(Value);
  FHBrush := CreateBrushIndirect(FBrush);
  DeleteObject(SelectObject(FDC,FHBrush));
end;

procedure TAXWGDI.SetBrushSolid(const Value: boolean);
begin
  CheckHandle;
  if Value then
    FBrush.lbStyle := BS_SOLID
  else
    FBrush.lbStyle := BS_NULL;
  FHBrush := CreateBrushIndirect(FBrush);
  DeleteObject(SelectObject(FDC,FHBrush));
end;

procedure TAXWGDI.SetCursor(Cursor: TXLSCursorType);
begin
  case Cursor of
    xctArrow      : FCurrentCursor := Windows.LoadCursor(0,IDC_ARROW);
    xctHandPoint  : FCurrentCursor := Windows.LoadCursor(0,IDC_HAND);
    xctCell       : FCurrentCursor := crCell;
    xctColSelect  : FCurrentCursor := crColSelect;
    xctFill       : FCurrentCursor := crFill;
    xctHorizSplit : FCurrentCursor := crHorizSplit;
    xctHoriz2Split: FCurrentCursor := crHoriz2Split;
    xctRowSelect  : FCurrentCursor := crRowSelect;
    xctVertSplit  : FCurrentCursor := crVertSplit;
    xctVert2Split : FCurrentCursor := crVert2Split;
    xctCenter     : FCurrentCursor := crCenter;
    xctMoveCells  : FCurrentCursor := crMoveCells;
    xctSizeCells  : FCurrentCursor := crSizeCells;
    xctNeSwArrow  : FCurrentCursor := crNeSwArrow;
    xctNwSeArrow  : FCurrentCursor := crNwSeArrow;
  end;
  Windows.SetCursor(FCurrentCursor);
end;

procedure TAXWGDI.SetDefaultData;
begin
  // Do not check validity of DC here, it must be valid.
  SetBkMode(FDC,TRANSPARENT);
end;

procedure TAXWGDI.SetFont(AFont: TXc12Font);
begin
  CheckHandle;

  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);

  FFontHandle := CreateFont(AFont.Name,AFont.Size,xfsBold in AFont.Style,xfsItalic in AFont.Style,AFont.Underline <> xulNone,0);

  SelectObject(FDC,FFontHandle);

  SetFontColor(RevRGB(Xc12ColorToRGB(AFont.Color)));
end;

procedure TAXWGDI.SetFontColor(const Value: longword);
begin
  CheckHandle;
  Windows.SetTextColor(FDC,Value);
end;

function TAXWGDI.SetFontName(FaceName: AxUCString; ACharset: byte): longword;
begin
  Result := SetFontName(FaceName);
end;

function TAXWGDI.SetFontName(FaceName: AxUCString): longword;
var
  LF: LOGFONTW;
  i: integer;
begin
  CheckHandle;
  GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LF),@LF);

  FaceName := Copy(FaceName,1,31);
  for i := 1 to Length(FaceName) do
    LF.lfFaceName[i - 1] := FaceName[i];
  LF.lfFaceName[Length(FaceName)] := #0;

  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);
  FFontHandle := CreateFontIndirectW(LF);
  SelectObject(FDC,FFontHandle);
  Result := FFontHandle;
end;

function TAXWGDI.SetFontRotation(Handle: longword; const Value: integer): longword;
var
  LogRec: TLogFont;
begin
  CheckHandle;
  FFontRotation := Value;
  GetObject(Handle, SizeOf(LogRec), Addr(LogRec));
  LogRec.lfEscapement := Round(FFontRotation * 10);
//  LogRec.lfOrientation := Value * 10;
  Result:= CreateFontIndirect(LogRec);
  SelectObject(FDC, Result);
end;

function TAXWGDI.SetFontSize(Size: integer): longword;
var
  LF: LOGFONT;
  H: integer;
begin
  CheckHandle;
  Result := GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LF),@LF);

  H := -MulDiv(Round(Size), PixelsPerInchY, 72);
  if H = LF.lfHeight then
    Exit;
  LF.lfHeight := H;

  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);
  FFontHandle := CreateFontIndirect(LF);
  SelectObject(FDC,FFontHandle);
  Result := FFontHandle;
end;

function TAXWGDI.SetFontSizePix(Size: integer): longword;
var
  LF: LOGFONT;
begin
  CheckHandle;
  Result := GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LF),@LF);

  Size := Size;
  if Size = LF.lfHeight then
    Exit;
  LF.lfHeight := Size;

  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);
  FFontHandle := CreateFontIndirect(LF);
  SelectObject(FDC,FFontHandle);
  Result := FFontHandle;
end;

function TAXWGDI.SetFontStrikeOut(StrikeOut: boolean): longword;
var
  LF: LOGFONT;
begin
  CheckHandle;
  Result := GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LF),@LF);

  if Integer(StrikeOut) = LF.lfStrikeOut then
    Exit;

  if StrikeOut then
    LF.lfStrikeOut := 1
  else
    LF.lfStrikeOut := 0;

  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);
  FFontHandle := CreateFontIndirect(LF);
  SelectObject(FDC,FFontHandle);
  Result := FFontHandle;
end;

function TAXWGDI.SetFontStyle(Bold, Italic, Underline: boolean): longword;
var
  LF: LOGFONT;
begin
  CheckHandle;
  Result := GetObject(GetCurrentObject(FDC,OBJ_FONT),SizeOf(LF),@LF);

  if (Bold = (LF.lfWeight = 700)) and (Integer(Italic) = LF.lfItalic) and (Integer(Underline) = LF.lfUnderline) then
    Exit;

  if Bold then
    LF.lfWeight := 700
  else
    LF.lfWeight := 400;

  LF.lfItalic := Integer(Italic);
  LF.lfUnderline := Integer(Underline);
  if FFontHandle <> 0 then
    DeleteObject(FFontHandle);
  FFontHandle := CreateFontIndirect(LF);
  SelectObject(FDC,FFontHandle);
  Result := FFontHandle;
end;

procedure TAXWGDI.SetDirtyArea(x, y: integer);
begin
       if x < FMinX then FMinX := x
  else if x > FMaxX then FMaxX := x;
       if y < FMinY then FMinY := y
  else if y > FMaxY then FMaxY := y;
  FDirty := True;
end;

procedure TAXWGDI.SetDirtyArea(x1, y1, x2, y2: integer);
begin
       if x1 < FMinX then FMinX := x1
  else if x1 > FMaxX then FMaxX := x1;
       if y1 < FMinY then FMinY := y1
  else if y1 > FMaxY then FMaxY := y1;
       if x2 < FMinX then FMinX := x2
  else if x2 > FMaxX then FMaxX := x2;
       if y2 < FMinY then FMinY := y2
  else if y2 > FMaxY then FMaxY := y2;
  FDirty := True;
end;

procedure TAXWGDI.SetPenColor(const Value: longword);
begin
  CheckHandle;
  FPen.lopnColor := RevRGB(Value);
  FHPen := CreatePenIndirect(FPen);
  DeleteObject(SelectObject(FDC,FHpen));
end;

procedure TAXWGDI.SetPenSolid(const Value: boolean);
begin
  CheckHandle;
  if Value then
    FPen.lopnStyle := PS_SOLID
  else
    FPen.lopnStyle := PS_NULL;
  FHPen := CreatePenIndirect(FPen);
  DeleteObject(SelectObject(FDC,FHPen));
end;

procedure TAXWGDI.SetPenWidth(const Value: longword);
var
  LB: LOGBRUSH;
begin
  CheckHandle;
  FPen.lopnWidth.X := Value;
//  FHPen := CreatePenIndirect(FPen);
  LB.lbStyle := 0;
  LB.lbColor := FPen.lopnColor;
  LB.lbHatch := 0;
  FHPen := ExtCreatePen(66048,Value,LB,0,Nil);
  DeleteObject(SelectObject(FDC,FHPen));
end;

procedure TAXWGDI.SetPenWidthF(const Value: double);
begin
  SetPenWidth(Round(Value));
end;

procedure TAXWGDI.SetPixel(x, y: integer; Cl: longword);
begin
  // **** Main routine must call Checkhandle! ***
  Windows.SetPixel(FDC,x,y,Cl);
end;

procedure TAXWGDI.SetPixelAA(x, y: integer; Cl: longword; A: byte);
begin
  // **** Main routine must call Checkhandle! ***
  Windows.SetPixel(FDC,x,y,RGB(Cl,Cl,Cl));
end;

procedure TAXWGDI.SetTextAlign(HPos: TXGDIHorizTextAlign; VPos: TXGDIVertTextAlign);
var
  V,H: longword;
begin
  CheckHandle;
  V := 0;
  H := 0;
  case HPos of
    xhtaLeft:   V := TA_LEFT;
    xhtaCenter: V := TA_CENTER;
    xhtaRight:  V := TA_RIGHT;
    xhtaLeftNoUpdateCP: V := TA_LEFT or TA_NOUPDATECP;
  end;
  case VPos of
    xvtaTop:      H := TA_TOP;
    xvtaBaseline: H := TA_BASELINE;
    xvtaBottom:   H := TA_BOTTOM;
  end;
  Windows.SetTextAlign(FDC,V or H);
end;

function TAXWGDI.SetTextJustification(const ABreakExtra, ABreakCount: integer): integer;
begin
  CheckHandle;
  Result := Windows.SetTextJustification(FDC,ABreakExtra,ABreakCount);
end;

function TAXWGDI.SetTextJustificationF(const ABreakExtra, ABreakCount: double): integer;
begin
  CheckHandle;
  Result := Windows.SetTextJustification(FDC,Round(ABreakExtra),Round(ABreakCount));
end;

procedure TAXWGDI.SetTransparentMode(const Value: boolean);
begin
  CheckHandle;
  if Value then
    Windows.SetBkMode(FDC,TRANSPARENT)
  else
    Windows.SetBkMode(FDC,OPAQUE);
end;

procedure TAXWGDI.SetXorMode(const Value: boolean);
begin
  CheckHandle;
  if Value then
    SetROP2(FDC,R2_NOTXORPEN)
  else
    SetROP2(FDC,R2_COPYPEN);
end;

procedure TAXWGDI.SizeLine(x1, y1, x2, y2: integer);
begin
  CheckHandle;
  PenSolid := False;
  XorMode := True;
  FHBrush := CreatePatternBrush(bmpGray.Handle);
  DeleteObject(SelectObject(FDC,FHBrush));
  SetDirtyArea(x1,y1,x2,y2);
  Windows.Rectangle(FDC,x1, y1, x2, y2);
  XorMode := False;
  PenSolid := True;
end;

function TAXWGDI.StdCharWidth: integer;
var
  Sz: SIZE;
  S: AxUCString;
begin
  CheckHandle;
  S := '0';
  GetTextExtentPointW(FDC,PWideChar(S),Length(S),Sz);
  Result := Sz.cx;
end;

procedure TAXWGDI.StrToFont32Array(Source: AxUCString; Dest: PWideChar);
var
  i: integer;
begin
  Source := Copy(Source,1,32);
  for i := 1 to Length(Source) do
    Dest[i - 1] := Source[i];
  if Length(Source) < 32 then
    Dest[Length(Source) + 1] := #0;
end;

function TAXWGDI.TextHeight(const Text: AxUCString): integer;
var
  Sz: SIZE;
begin
  CheckHandle;
  GetTextExtentPointW(FDC,PWideChar(Text),Length(Text),Sz);
  Result := Sz.cy;
end;

function TAXWGDI.TextHeightF(const Text: AxUCString): double;
var
  Sz: SIZE;
begin
  CheckHandle;
  GetTextExtentPointW(FDC,PWideChar(Text),Length(Text),Sz);
  Result := Sz.cy;
end;

procedure TAXWGDI.TextOut(X, Y: integer; Text: AxUCString);
begin
  CheckHandle;
  Windows.TextOutW(FDC,X,Y,PWideChar(Text),Length(Text));
end;

// dtoHCenter,dtoLeft,dtoRight,dtoTop,dtoVCenter,dtoBottom,dtoVDistributed,dtoWordBreak
procedure TAXWGDI.TextOutF(X, Y: double; Text: AxUCString; AOptions: TDrawTextOptions);
var
  TA: longword;
begin
  CheckHandle;

  TA := 0;
  if dtoRight in AOptions then
    TA := TA + TA_RIGHT
  else if dtoLeft in AOptions then
    TA := TA + TA_LEFT
  else if dtoHCenter in AOptions then begin
    TA := TA + TA_LEFT;
    X := X - (TextWidthF(Text) / 2);
  end;

  if dtoBaseline in AOptions then
    TA := TA + TA_BASELINE
  else if dtoVCenter in AOptions then begin
    TA := TA + TA_TOP;
    Y := Y - (TextHeightF(Text) / 2);
  end;

  Windows.SetTextAlign(FDC,TA);
  Windows.TextOutW(FDC,Round(X),Round(Y),PWideChar(Text),Length(Text));
end;

procedure TAXWGDI.TextOutF(X, Y: double; Text: AxUCString);
begin
  CheckHandle;
  Windows.TextOutW(FDC,Round(X),Round(Y),PWideChar(Text),Length(Text));
end;

procedure TAXWGDI.TextOut(X, Y: integer; Rect: TRect; Text: AxUCString);
begin
  CheckHandle;
  SetDirtyArea(Rect.Left,Rect.Top,Rect.Right + 1,Rect.Bottom + 1);
  Windows.ExtTextOutW(FDC, X, Y, ETO_CLIPPED, @Rect, PWideChar(Text),Length(Text), nil);
end;

function TAXWGDI.TextWidth(const Text: AxUCString): integer;
var
  Sz: SIZE;
begin
  CheckHandle;

  Sz.cx := 0;
  Sz.cy := 0;
  Windows.GetTextExtentPoint32W(FDC, PWideChar(Text), Length(Text), Sz);
  Result := Sz.cx;
end;

function TAXWGDI.TextWidthF(const Text: AxUCString): double;
var
  Sz: SIZE;
begin
  CheckHandle;

  Sz.cx := 0;
  Sz.cy := 0;
  Windows.GetTextExtentPoint32W(FDC, PWideChar(Text), Length(Text), Sz);
  Result := Sz.cx;
end;

function TAXWGDI.TextWidthFloat(const Text: AxUCString): double;
begin
  Result := TextWidth(Text);
end;

procedure TAXWGDI.WidthRectangle(x1, y1, x2, y2, W: integer);
var
  C: longword;
begin
  CheckHandle;
  SetDirtyArea(x1 - W,y1 - W,x2 + W,y2 + W);
  C := GetBrushColor;
  SetBrushColor(GetPenColor);
  Windows.Rectangle(FDC,x1,y1,x2 + 1,y1 + W + 1);
  Windows.Rectangle(FDC,x1,y1 + W + 1,x1 + W + 1,y2 - W - 1);
  Windows.Rectangle(FDC,x2 - W - 1,y1 + W + 1,x2 + 1,y2 - W - 1);
  Windows.Rectangle(FDC,x1,y2 - W - 1,x2 + 1,y2 + 1);
  SetBrushColor(C);
  PenSolid := False;
  Windows.Rectangle(FDC,x1 + W + 1,y1 + W + 1,x2 - W,y2 - W);
  PenSolid := True;
end;

procedure TAXWGDI.SetPaintColor(const Value: longword);
begin
  CheckHandle;
  FPen.lopnColor := RevRGB(Value);
  FHPen := CreatePenIndirect(FPen);
  DeleteObject(SelectObject(FDC,FHpen));

  FBrush.lbColor := RevRGB(Value);
  FHBrush := CreateBrushIndirect(FBrush);
  DeleteObject(SelectObject(FDC,FHBrush));
end;

procedure TAXWGDI.SetDirtyArea(Points: array of TPoint);
var
  i: integer;
begin
  for i := 0 to High(Points) do begin
         if Points[i].x < FMinX then FMinX := Points[i].x
    else if Points[i].x > FMaxX then FMaxX := Points[i].x;
         if Points[i].y < FMinY then FMinY := Points[i].y
    else if Points[i].y > FMaxY then FMaxY := Points[i].y;
  end;
  FDirty := True;
end;

procedure TAXWGDI.StretchBlt(DestDC: longword; DestX, DestY, DestWidth, DestHeight, ScrX, SrcY, SrcWidth, SrcHeight: integer);
begin
  CheckHandle;
  Windows.SetStretchBltMode(DestDC,HALFTONE);

  Windows.StretchBlt(DestDC,
                     DestX, DestY, DestWidth, DestHeight,
                     FDC,
                     ScrX, SrcY, SrcWidth, SrcHeight,SRCCOPY);
end;

{ TAXWGDIBMP }

procedure TAXWGDIBMP.CheckHandle;
begin
//  SetDefaultData;
  // Do nothing.
end;

constructor TAXWGDIBMP.Create;
var
  BMP: BITMAP;
  BI: BITMAPINFO;
begin
//  FTransparentColor := $D1D1D1;
  FTransparentColor := $123456;

  FDC := CreateCompatibleDC(0);
  FillChar(BI, SizeOf(BITMAPINFO), 0);
  with BI.bmiHeader do begin
    biSize := SizeOf(TBitmapInfoHeader);
    biPlanes := 1;
    biBitCount := 32;
    biCompression := BI_RGB;

    biWidth := GetDeviceCaps(FDC,HORZRES);
    biHeight := -GetDeviceCaps(FDC,VERTRES);
    FBMPWidth := GetDeviceCaps(FDC,HORZRES);
    FBMPHeight := GetDeviceCaps(FDC,VERTRES);

    FHBMP := CreateDIBSection(0, BI, DIB_RGB_COLORS, Pointer(FBmpBits), 0, 0);
    SelectObject(FDC,FHBMP);


    if GetObject(FHBMP, SizeOf(BMP), @BMP) = SizeOf(BMP) then
      FBmpBits := BMP.bmBits;

    inherited Create(0);

    SetDefaultData;

    Clear;
  end;
end;

destructor TAXWGDIBMP.Destroy;
begin
  DeleteObject(FHBMP);
  DeleteDC(FDC);
  inherited;
end;

procedure TAXWGDIBMP.Invalidate;
begin
  CheckHandle;
  SetBrushSolid(True);
  SetPenColor(RevRGB(FTransparentColor));
  SetBrushColor(RevRGB(FTransparentColor));
  Windows.Rectangle(FDC,0,0,FBMPWidth - 1,FBMPHeight - 1);
end;

procedure TAXWGDIBMP.InvalidateRect(x1, y1, x2, y2: integer);
begin
  inherited InvalidateRect(x1,y1,x2 + 1,y2 + 1);
  CheckHandle;
  SetPenColor(RevRGB(FTransparentColor));
  SetBrushColor(RevRGB(FTransparentColor));
  Windows.Rectangle(FDC,x1,y1,x2 + 1,y2 + 1);
end;

//procedure TAXWGDIBMP.LineStyled(x1, y1, x2, y2: integer; PenColor,BackColor: longword; Style: TPaintLineStyle; Horiz: boolean);
//var
//  i: integer;
//
//procedure DrawStyled(LineStyle: array of boolean);
//var
//  i,j: integer;
//begin
//  j := 0;
//  if Horiz then
//    for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do begin
//      if LineStyle[j] then
//        PLongWordArray(FBmpBits)[i] := PenColor
//      else
//        PLongWordArray(FBmpBits)[i] := BackColor;
//      Inc(j);
//      if j > High(LineStyle) then
//        j := 0;
//    end
//  else begin
//    for i := y1 to y2 do begin
//      if LineStyle[j] then
//        PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor
//      else
//        PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := BackColor;
//      Inc(j);
//      if j > High(LineStyle) then
//        j := 0;
//    end
//  end;
//end;
//
//begin
//  case Style of
//    plsNone,plsThin: begin
//      if Horiz then begin
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//      end
//      else begin
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//      end;
//    end;
//    plsMedium: begin
//      if Horiz then begin
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//        for i := FBMPWidth * y2 + x1 to FBMPWidth * y2 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//      end
//      else begin
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x2] := PenColor;
//      end;
//    end;
//    plsDashed:
//      DrawStyled([True,True,True,False]);
//    plsDotted:
//      DrawStyled([True,True,False,False]);
//    plsThick: begin
//      if Horiz then begin
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//        Inc(y1);
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//        Inc(y1);
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//      end
//      else begin
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//        Inc(x1);
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//        Inc(x1);
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//      end;
//    end;
//    plsDouble: begin
//      if Horiz then begin
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//        Inc(y1);
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := BackColor;
//        Inc(y1);
//        for i := FBMPWidth * y1 + x1 to FBMPWidth * y1 + x2 do
//          PLongWordArray(FBmpBits)[i] := PenColor;
//      end
//      else begin
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//        Inc(x1);
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := BackColor;
//        Inc(x1);
//        for i := y1 to y2 do
//          PLongWordArray(FBmpBits)[i * FBMPWidth + x1] := PenColor;
//      end;
//    end;
//    plsHair:
//      DrawStyled([True,False]);
//    plsMediumDashed: begin
//      DrawStyled([True,True,True,False]);
//      if Horiz then
//        Inc(y1)
//      else
//        Inc(x1);
//      DrawStyled([True,True,True,False]);
//    end;
//    plsDashDot:
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False]);
//    plsMediumDashDot: begin
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False]);
//      if Horiz then
//        Inc(y1)
//      else
//        Inc(x1);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False]);
//    end;
//    plsDashDotDot:
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False,True,False,False,False]);
//    plsMediumDashDotDot: begin
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False,True,False,False,False]);
//      if Horiz then
//        Inc(y1)
//      else
//        Inc(x1);
//      DrawStyled([True,True,True,True,True,True,True,True,True,False,False,False,True,True,True,False,False,False,True,False,False,False]);
//    end;
//    plsSlantedDashDot: begin
//      DrawStyled([False,True,True,True,True,True,True,True,True,True,True,True,False,True,True,True,True,True]);
//      if Horiz then
//        Inc(y1)
//      else
//        Inc(x1);
//      DrawStyled([True,True,True,True,True,True,True,True,True,True,False,False,True,True,True,True,False,False]);
//    end;
//  end;
//end;

procedure TAXWGDIBMP.ReleaseHandle;
begin
  ResetLimits;
  // Do nothing.
end;

procedure TAXWGDIBMP.SetPixel(x, y: integer; Cl: longword);
begin
  PLongWordArray(FBmpBits)[FBMPWidth * y + x] := Cl;
end;

{ TAXWGDIBMPTrans }

procedure TAXWGDIBMPTrans.Clear;
begin
  inherited;
  CheckHandle;
  SetBrushColor(RevRGB(FTransparentColor));
  SetPenColor(RevRGB(FTransparentColor));
  Windows.Rectangle(FDC,0,0,MAXINT,MAXINT);
end;

constructor TAXWGDIBMPTrans.Create;
begin
  inherited Create;
end;

procedure TAXWGDIBMPTrans.Render(Dest: TAXWGDI);
{
var
  Blend: BLENDFUNCTION;
}
begin
{
  Blend.BlendOp := AC_SRC_OVER;
  Blend.BlendFlags := 0;
  Blend.SourceConstantAlpha := 128;
  Blend.AlphaFormat := 0; // AC_SRC_ALPHA;
  SetBkColor(hSrc,$FFFFFF);
  Windows.AlphaBlend(hDest,X1,Y1,X2 - X1,Y2 - Y1,hSrc,X1,Y1,X2 - X1,Y2 - Y1,Blend);

  x1 := Min(0,Dest.FMinX);
  y1 := Min(0,Dest.FMinY);
  x2 := Max(FWidth,Dest.FMaxX);
  y2 := Max(FHeight,Dest.FMaxY);
}
  Windows.TransparentBlt(Dest.FDC,0,0,FBMPWidth,FBMPHeight,FDC,0,0,FBMPWidth,FBMPHeight,FTransparentColor);
end;

procedure TAXWGDIBMPTrans.RenderDirty(Dest: TAXWGDI);
var
  x1,y1,x2,y2: integer;
begin
  // TransparentBlt can't handle negative coordinates!!!!!! (BitBlt can).
  x1 := Max(FMinX,0);
  y1 := Max(FMinY,0);
  x2 := Min(FMaxX + 1,FBMPWidth);
  y2 := Min(FMaxY + 1,FBMPHeight);
  if FDirty then begin
    if (FMinX <> MAXINT) or (FMinY <> MAXINT) then begin
//      Dest.SetDirtyArea(FMinX,FMinY,FMaxX,FMaxY);
      Windows.TransparentBlt(Dest.DC,x1,y1,x2 - x1,y2 - y1,FDC,x1,y1,x2 - x1,y2 - y1,FTransparentColor);
    end;
    FDirty := False;
  end
//  else
//    Windows.TransparentBlt(Dest.DC,101,70,100,100,FDC,101,70,100,100,FTransparentColor);
  else
    Windows.TransparentBlt(Dest.DC,x1,y1,x2 - x1,y2 - y1,FDC,x1,y1,x2 - x1,y2 - y1,FTransparentColor);
end;

{ TTransformation }

procedure TTransformation.Clear;
begin
  FMatrix := DefaultMatrix;
end;

constructor TTransformation.Create;
begin
  Clear;
end;

function TTransformation.Mult(M1,M2: TFloatMatrix): TFloatMatrix;
var
  i, j: Integer;
begin
  for i := 0 to 2 do begin
    for j := 0 to 2 do
      Result[i, j] := M1[0, j] * M2[i, 0] + M1[1, j] * M2[i, 1] + M1[2, j] * M2[i, 2];
  end;
end;

procedure TTransformation.Rotate(X,Y,Alpha: double);
var
  S, C: double;
  M: TFloatMatrix;
begin
  if (X <> 0) or (Y <> 0) then
    Translate(-X, -Y);
  Alpha := DegToRad(Alpha);
  S := Sin(Alpha);
  C := Cos(Alpha);
  M := DefaultMatrix;
  M[0,0] := C;   M[1,0] := S;
  M[0,1] := -S;  M[1,1] := C;
  FMatrix := Mult(M, FMatrix);
  if (X <> 0) or (Y <> 0) then
    Translate(X, Y);
end;

procedure TTransformation.Scale(X,Y: double);
var
  M: TFloatMatrix;
begin
  M := DefaultMatrix;
  M[0,0] := X;
  M[1,1] := Y;
  FMatrix := Mult(M, FMatrix);
end;

procedure TTransformation.Skew(X, Y: double);
var
  M: TFloatMatrix;
begin
  M := DefaultMatrix;
  M[1,0] := X;
  M[0,1] := Y;
  FMatrix := Mult(M, FMatrix);
end;

procedure TTransformation.Transform(var X,Y: integer);
var
  fX,fY: double;
begin
  fX := X;
  fY := Y;
  X := Round(fX * FMatrix[0,0] + fY * FMatrix[1,0] + FMatrix[2,0]);
  Y := Round(fX * FMatrix[0,1] + fY * FMatrix[1,1] + FMatrix[2,1]);
end;

procedure TTransformation.Transform(var Arr: array of TPoint);
var
  i: integer;
  fX,fY: double;
begin
  for i := 0 to High(Arr) do begin
    fX := Arr[i].X;
    fY := Arr[i].Y;
    Arr[i].X := Round(fX * FMatrix[0,0] + fY * FMatrix[1,0] + FMatrix[2,0]);
    Arr[i].Y := Round(fX * FMatrix[0,1] + fY * FMatrix[1,1] + FMatrix[2,1]);
  end;
end;

procedure TTransformation.Translate(X,Y: double);
var
  M: TFloatMatrix;
begin
  M := DefaultMatrix;
  M[2,0] := X;
  M[2,1] := Y;
  FMatrix := Mult(M, FMatrix);
end;

{ TAXWGDIPrint }

procedure TAXWGDIMetafile.BeginMetafile(const AReferenceDC: longword; const AWidthCm, AHeightCm: double);
begin
{$ifdef DELPHI_2006_OR_LATER}
  FMetafile.SetSize(Round(AWidthCm * 1000),Round(AHeightCm * 1000));
{$else}
  FMetafile.Width := Round(AWidthCm * 1000);
  FMetafile.Height := Round(AHeightCm * 1000);
{$endif}
  FMCanvas := TMetafileCanvas.Create(FMetafile,AReferenceDC);
  FDC := FMCanvas.Handle;

//  FXScale := (GetDeviceCaps(AReferenceDC,HORZSIZE) * 100) / GetDeviceCaps(AReferenceDC,HORZRES);
  FXScale := (AWidthCm * 1000) / ((AWidthCm / 2.54) * DevicePixelsPerInchX);
  if FXScale <= 0 then
    FXScale := 1;

//  FYScale := (GetDeviceCaps(AReferenceDC,VERTSIZE) * 100) / GetDeviceCaps(AReferenceDC,VERTRES);
  FYScale := (AHeightCm * 1000) / ((AHeightCm / 2.54) * DevicePixelsPerInchY);

  if FYScale <= 0 then
    FYScale := 1;
end;

procedure TAXWGDIMetafile.CheckHandle;
begin
  // Do nothing
end;

constructor TAXWGDIMetafile.Create;
begin
  inherited Create(0);

  FMetafile := TMetafile.Create;
end;

destructor TAXWGDIMetafile.Destroy;
begin
  if FMCanvas <> Nil then
    FMCanvas.Free;
  FMetafile.Free;

  inherited;
end;

function TAXWGDIMetafile.DevicePixelsPerInchX: integer;
begin
  Result := GetDeviceCaps(FDC, LOGPIXELSX);
end;

function TAXWGDIMetafile.DevicePixelsPerInchY: integer;
begin
  Result := GetDeviceCaps(FDC, LOGPIXELSY);
end;

procedure TAXWGDIMetafile.EndMetafile;
begin
  FMCanvas.Free;
  FMCanvas := Nil;
  FDC := 0;
end;

function TAXWGDIMetafile.PixelsPerInchX: integer;
begin
  Result := Round(GetDeviceCaps(FDC, LOGPIXELSX) * FZoom * FXScale);
end;

function TAXWGDIMetafile.PixelsPerInchY: integer;
begin
  Result := Round(GetDeviceCaps(FDC, LOGPIXELSY) * FZoom * FYScale);
end;

procedure TAXWGDIMetafile.PlayMetafile(ACanvas: TCanvas; const AX1,AY1,AX2,AY2: integer);
begin
{$ifdef DELPHI_2006_OR_LATER}
  FMetafile.SetSize(AX2 - AX1,AY2 - AY1);
{$else}
  FMetafile.Width := AX2 - AX1;
  FMetafile.Height := AY2 - AY1;
{$endif}
  ACanvas.Draw(AX1,AY1,FMetafile);
end;

procedure TAXWGDIMetafile.ReleaseHandle;
begin
  // Do nothing
end;

{ TXLSCurve }

procedure TXLSCurve.Add(TTCurve: PTTPOLYCURVE; Height,GlyphDescent: integer);
const
  // Missing in Windows.
  TT_PRIM_CSPLINE = 3;
var
  i: integer;
begin
  if FCount >= Length(FCurves) then
    SetLength(FCurves,Length(FCurves) + 16);
  FCurves[FCount].Count := TTCurve.cpfx;
  GetMem(FCurves[FCount].Points,FCurves[FCount].Count * SizeOf(TPoint));
   // For each point
  for i := 0 to TTCurve.cpfx - 1 do
    FCurves[FCount].Points[i] := PtFxToPT(TTCurve.apfx[i],Height,GlyphDescent);
  case TTCurve.wType of
    TT_PRIM_LINE   : FCurves[FCount].CurveType := xctPolyline;
    TT_PRIM_CSPLINE: FCurves[FCount].CurveType := xctBezier;
    else raise XLSRWException.Create('Illegal curve type in glyph');
  end;
  Inc(FCount);
end;

procedure TXLSCurve.Adjust(X, Y: integer);
var
  i,j: integer;
begin
  Inc(FStart.X,X);
  Inc(FStart.Y,Y);
//  FStart.Y := Y - FStart.Y;
  for i := 0 to FCount - 1 do begin
    for j := 0 to FCurves[i].Count - 1 do begin
      Inc(FCurves[i].Points[j].X,X);
      Inc(FCurves[i].Points[j].Y,Y);
//      FCurves[i].Points[j].Y := Y - FCurves[i].Points[j].Y;
    end;
  end;
end;

procedure TXLSCurve.CloseCurve;
begin
  if (FCount > 0) and (FCurves[0].Count > 0) and ((FCurves[0].Points[0].X <> FStart.X) or (FCurves[0].Points[0].Y <> FStart.Y)) then begin
    if FCount >= Length(FCurves) then
      SetLength(FCurves,Length(FCurves) + 1);
    FCurves[FCount].Count := 0;
    FCurves[FCount].Points := Nil;
    FCurves[FCount].CurveType := xctLineTo;
    Inc(FCount);
  end;
end;

constructor TXLSCurve.Create;
begin
  SetLength(FCurves,16);
end;

destructor TXLSCurve.Destroy;
var
  i: integer;
begin
  for i := 0 to FCount - 1 do
    FreeMem(Items[i].Points);
  inherited;
end;

function TXLSCurve.GetItems(Index: integer): TXLSCurvePoints;
begin
  Result := FCurves[Index];
end;

{ TXLSCurves }

procedure TXLSFigure.Add(TTCurves: PTTPOLYGONHEADER; Size,W,H,GlyphDescent: integer);
const
  // Missing in Windows.
  TT_PRIM_CSPLINE = 3;
var
  P: Pointer;
  pEnd: NativeInt;
  pHeader,pNext: ^TTPOLYGONHEADER;
  pCurve: PTTPOLYCURVE;
  Start: TPoint;
  Curve: TXLSCurve;
begin
  FWidth := W;
  FHeight := H;

  P := TTCurves;
  pEnd := NativeInt(TTCurves) + Size;

  // For each contour
  while NativeInt(P) < pEnd do begin
    Curve := TXLSCurve.Create;
    pHeader := P;
    pNext := Pointer(Longword(P) + pHeader.cb);
    Start := PtFxToPT(pHeader.pfxStart,FHeight,GlyphDescent);
    Curve.Start := Start;
    P := Pointer(NativeInt(P) + SizeOf(TTPOLYGONHEADER));
    // For each curve
    while NativeInt(P) < NativeInt(pNext) do begin
      pCurve := P;
      Curve.Add(pCurve,FHeight,GlyphDescent);
      P := Pointer(NativeInt(P) + SizeOf(TTPOLYCURVE) + (SizeOf(POINTFX) * (pCurve.cpfx - 1)));
    end;
    Curve.CloseCurve;
    inherited Add(Curve);
  end;
end;

procedure TXLSFigure.Adjust(X, Y: integer);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Adjust(X,Y);
end;

function TXLSFigure.GetItems(Index: integer): TXLSCurve;
begin
  Result := TXLSCurve(inherited Items[Index]);
end;

initialization
  TAXWGDI.LoadResources;

finalization
  TAXWGDI.FreeResources;

end.
