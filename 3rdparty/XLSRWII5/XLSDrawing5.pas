unit XLSDrawing5;

{-
********************************************************************************
******* XLSReadWriteII V6.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSReadWriteII component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSReadWriteII is supplied as is. The author disclaims all warranties,     **
** expressedor implied, including, without limitation, the warranties of      **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSReadWriteII.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}
{$I XLSRWII.inc}

interface

uses Classes, SysUtils, Contnrs, Math,
{$ifdef BABOON}
  {$ifdef DELPHI_XE5_OR_LATER}
     FMX.Graphics,
  {$else}
     FMX.Types,
  {$endif}
{$else}
     vcl.Graphics,
{$endif}
     xpgParseDrawingCommon, xpgParseDrawing, xpgParseChart, xpgPSimpleDOM,
     Xc12Utils5, Xc12Graphics, Xc12DataStyleSheet5, Xc12Manager5,
{$ifdef XLS_BIFF}
     BIFF_DrawingObj5, BIFF_ControlObj5,
{$endif}
     XLSUtils5, XLSColumn5, XLSRow5, XLSRelCells5;

type TAddImageEvent97 = function(AAnchor: TCT_TwoCellAnchor): boolean of object;

type TDrwOptImagePositioning = (doipMoveButDoNotSize,doipMoveAndSize,doipDoNotMoveOrSize);

type TXLSDrwFillType = (xdftNone,xdftSolid,xdftGradient,xdftPicture,xdftAutomatic);
type TXLSDrwColorType = (xdctRGB,xdctHSL,xdctSystem,xdctScheme,xdctPreset);
type TXLSDrwXMLColorType = (xdxctRGB,xdxctCRGB,xdxctHSL,xdxctSystem,xdxctScheme,xdxctPreset);
type TXLSDrwGradientFillType = (xdgftLinear,xdgftRadial,xdgftRectangular,xdgftPath);
type TXLSDrwGradientFillDirection = (xdgfdNone,xdgfdCentre,xdgfdTopLeft,xdgfdTopRight,xdgfdBottomLeft,xdgfdBottomRight);
type TXLSDrwMirrorType = (xdmtNone,xdmtHorizontal,xdmtVertical,xdmtBoth);
type TXLSDrwLineJoinType = (xdljtNone,xdljtRound,xdljtBevel,xdljtMiter);
type TXLSDrwTextAlign = (xdtaLeft,xdtaCentre,xdtaRight,xdtaJustify,xdtaJustifyLow,xdtaDistributed);

type TXLSDrwShapeProp = class(TObject)
end;

type TXLSDrwColor = class(TXLSDrwShapeProp)
private
     function  GetAsARGB: longword;
     function  GetAsHSL: longword;
     function  GetAsPreset: TST_PresetColorVal;
     function  GetAsRGB: longword;
     function  GetAsScheme: TST_SchemeColorVal;
     function  GetAsSystem: TST_SystemColorVal;
     function  GetTransparency: double;
     procedure SetAsARGB(const Value: longword);
     procedure SetAsHSL(const Value: longword);
     procedure SetAsPreset(const Value: TST_PresetColorVal);
     procedure SetAsRGB(const Value: longword);
     procedure SetAsScheme(const Value: TST_SchemeColorVal);
     procedure SetAsSystem(const Value: TST_SystemColorVal);
     procedure SetTransparency(const Value: double);
     function  GetColorType: TXLSDrwColorType;
     function  GetAsTColor: TColor;
     procedure SetAsTColor(const Value: TColor);
protected
     FColorChoice: TEG_ColorChoice;
     FColorTransform: TEG_ColorTransform;

     procedure AssignColorTransform;
     function  XlsColorType: TXLSDrwXMLColorType;
public
     constructor Create(AColorChoice: TEG_ColorChoice);

     property ColorType: TXLSDrwColorType read GetColorType;
     property Transparency: double read GetTransparency write SetTransparency;

     property AsTColor: TColor read GetAsTColor write SetAsTColor;
     property AsRGB: longword read GetAsRGB write SetAsRGB;
     property AsARGB: longword read GetAsARGB write SetAsARGB;
     property AsHSL: longword read GetAsHSL write SetAsHSL;
     property AsSystem: TST_SystemColorVal read GetAsSystem write SetAsSystem;
     property AsScheme: TST_SchemeColorVal read GetAsScheme write SetAsScheme;
     property AsPreset: TST_PresetColorVal read GetAsPreset write SetAsPreset;
     end;

type TXLSDrwFill = class(TXLSDrwShapeProp)
protected
     FFillType: TXLSDrwFillType;

     function GetFillType: TXLSDrwFillType; virtual; abstract;
public

     property FillType: TXLSDrwFillType read GetFillType;
     end;

type TXLSDrwFillNone = class(TXLSDrwFill)
protected
     FFill: TCT_NoFillProperties;

     function GetFillType: TXLSDrwFillType; override;
public
     constructor Create(AFill: TCT_NoFillProperties);
     end;

type TXLSDrwFillSolid = class(TXLSDrwFill)
protected
     FColor: TXLSDrwColor;
     FFill: TCT_SolidColorFillProperties;

     function GetFillType: TXLSDrwFillType; override;
public
     constructor Create(AFill: TCT_SolidColorFillProperties);
     destructor Destroy; override;

     property Color: TXLSDrwColor read FColor;
     end;

type TXLSDrwFillGradientStop = class(TXLSDrwShapeProp)
private
     function  GetStopPosition: double;
     function  GetTransparency: double;
     procedure SetStopPosition(const Value: double);
     procedure SetTransparency(const Value: double);
protected
     FStop:  TCT_GradientStop;
     FColor: TXLSDrwColor;
public
     constructor Create(AStop:  TCT_GradientStop);
     destructor Destroy; override;

     property StopPosition: double read GetStopPosition write SetStopPosition;
     property Color: TXLSDrwColor read FColor;
     property Transparency: double read GetTransparency write SetTransparency;
     end;

type TXLSDrwFillGradientStops = class(TXLSDrwShapeProp)
private
     function GetItems(const Index: integer): TXLSDrwFillGradientStop;
protected
     FItems: TObjectList;
     FGSList: TCT_GradientStopList;

     procedure SetDefault;
public
     constructor Create(AGSList: TCT_GradientStopList);
     destructor Destroy; override;

     procedure Clear;
     function  Add: TXLSDrwFillGradientStop;
     procedure Remove(const AIndex: integer);
     function  Count: integer;

     property Items[const Index: integer]: TXLSDrwFillGradientStop read GetItems; default;
     end;

type TXLSDrwFillGradient = class(TXLSDrwFill)
private
     function  GetDirection: TXLSDrwGradientFillDirection;
     function  GetGradientType: TXLSDrwGradientFillType;
     function  GetRotateWithShape: boolean;
     procedure SetDirection(const Value: TXLSDrwGradientFillDirection);
     procedure SetGradientType(const Value: TXLSDrwGradientFillType);
     procedure SetRotateWithShape(const Value: boolean);
     function  GetAngle: double;
     procedure SetAngle(const Value: double);
protected
     FFill : TCT_GradientFillProperties;
     FStops: TXLSDrwFillGradientStops;

     function GetFillType: TXLSDrwFillType; override;
public
     constructor Create(AFill: TCT_GradientFillProperties);
     destructor Destroy; override;

     property Stops: TXLSDrwFillGradientStops read FStops;
     property GradientType: TXLSDrwGradientFillType read GetGradientType write SetGradientType;
     property Direction: TXLSDrwGradientFillDirection read GetDirection write SetDirection;
     property RotateWithShape: boolean read GetRotateWithShape write SetRotateWithShape;
     // Only when GradientType = Linear.
     property Angle: double read GetAngle write SetAngle;
     end;

type TXLSDrwFillPictStretchOpts = class(TXLSDrwShapeProp)
private
     function  GetOffsetBottom: double;
     function  GetOffsetLeft: double;
     function  GetOffsetRight: double;
     function  GetOffsetTop: double;
     procedure SetOffsetBottom(const Value: double);
     procedure SetOffsetLeft(const Value: double);
     procedure SetOffsetRight(const Value: double);
     procedure SetOffsetTop(const Value: double);
protected
     FStretch: TCT_StretchInfoProperties;
public
     constructor Create(AStretch: TCT_StretchInfoProperties);

     property OffsetLeft: double read GetOffsetLeft write SetOffsetLeft;
     property OffsetRight: double read GetOffsetRight write SetOffsetRight;
     property OffsetTop: double read GetOffsetTop write SetOffsetTop;
     property OffsetBottom: double read GetOffsetBottom write SetOffsetBottom;
     end;

type TXLSDrwFillPictTileOpts = class(TXLSDrwShapeProp)
private
     function  GetAlignment: TST_RectAlignment;
     function  GetMirrorType: TXLSDrwMirrorType;
     function  GetOffsetX: double;
     function  GetOffsetY: double;
     function  GetScaleX: double;
     function  GetScaleY: double;
     procedure SetAlignment(const Value: TST_RectAlignment);
     procedure SetMirrorType(const Value: TXLSDrwMirrorType);
     procedure SetOffsetX(const Value: double);
     procedure SetOffsetY(const Value: double);
     procedure SetScale(const Value: double);
     procedure SetScaleX(const Value: double);
     procedure SetScaleY(const Value: double);
protected
     FTile: TCT_TileInfoProperties;
public
     constructor Create(ATile: TCT_TileInfoProperties);

     property OffsetX: double read GetOffsetX write SetOffsetX;
     property OffsetY: double read GetOffsetY write SetOffsetY;
     property Scale: double write SetScale;
     property ScaleX: double read GetScaleX write SetScaleX;
     property ScaleY: double read GetScaleY write SetScaleY;
     property Alignment: TST_RectAlignment read GetAlignment write SetAlignment;
     property MirrorType: TXLSDrwMirrorType read GetMirrorType write SetMirrorType;
     end;

type TXLSDrwFillPicture = class(TXLSDrwFill)
private
     function  GetRorateWithShape: boolean;
     function  GetTransparency: double;
     procedure SetRorateWithShape(const Value: boolean);
     procedure SetTransparency(const Value: double);
     function  GetBrightness: double;
     function  GetContrast: double;
     procedure SetBrightness(const Value: double);
     procedure SetContrast(const Value: double);
     function  GetTilePicture: boolean;
     procedure SetTilePicture(const Value: boolean);
protected
     FFilename: AxUCString;
     FBlip    : TCT_BlipFillProperties;
     FTile    : TXLSDrwFillPictTileOpts;
     FStretch : TXLSDrwFillPictStretchOpts;

     function GetFillType: TXLSDrwFillType; override;
public
     constructor Create(ABlip: TCT_BlipFillProperties);
     destructor Destroy; override;

     // Loads the Filename picture;
     function  LoadPicture: boolean; overload;
     function  LoadPicture(const APictureName: AxUCString): boolean; overload;

     property Filename       : AxUCString read FFilename write FFilename;
     property TilePicture    : boolean read GetTilePicture write SetTilePicture;
     property RorateWithShape: boolean read GetRorateWithShape write SetRorateWithShape;
     property Transparency   : double read GetTransparency write SetTransparency;
     property Brightness     : double read GetBrightness write SetBrightness;
     property Contrast       : double read GetContrast write SetContrast;
     end;

type TXLSDrwFillAutomatic = class(TXLSDrwFill)
protected
     function GetFillType: TXLSDrwFillType; override;
public
     end;

type TXLSDrwFillChoice = class(TXLSDrwShapeProp)
private
     function  GetAsGradient: TXLSDrwFillGradient;
     function  GetAsNone: TXLSDrwFillNone;
     function  GetAsPicture: TXLSDrwFillPicture;
     function  GetAsSolid: TXLSDrwFillSolid;
     function  GetFillType: TXLSDrwFillType;
     procedure SetFillType(const Value: TXLSDrwFillType);
protected
     FFillChoice: TEG_FillProperties;
     FFill: TXLSDrwFill;
public
     constructor Create(AFillChoice: TEG_FillProperties);
     destructor Destroy; override;

     function RGB_(ADefault: longword): longword;

     property FillType: TXLSDrwFillType read GetFillType write SetFillType;
     property AsNone: TXLSDrwFillNone read GetAsNone;
     property AsSolid: TXLSDrwFillSolid read GetAsSolid;
     property AsGradient: TXLSDrwFillGradient read GetAsGradient;
     property AsPicture: TXLSDrwFillPicture read GetAsPicture;
     // Automatiic = No fill assigned.
     end;

type TXLSDrwLineFillChoice = class(TXLSDrwShapeProp)
private
     function  GetAsGradient: TXLSDrwFillGradient;
     function  GetAsNone: TXLSDrwFillNone;
     function  GetAsSolid: TXLSDrwFillSolid;
     function  GetFillType: TXLSDrwFillType;
     procedure SetFillType(const Value: TXLSDrwFillType);
protected
     FFillChoice: TEG_FillProperties;
     FFill: TXLSDrwFill;
public
     constructor Create(AFillChoice: TEG_FillProperties);

     property FillType: TXLSDrwFillType read GetFillType write SetFillType;
     property AsNone: TXLSDrwFillNone read GetAsNone;
     property AsSolid: TXLSDrwFillSolid read GetAsSolid;
     property AsGradient: TXLSDrwFillGradient read GetAsGradient;
     end;

type TXLSDrwLineEndStyle = class(TXLSDrwShapeProp)
private
     function  GetEndType: TST_LineEndType;
     function  GetLen: TST_LineEndLength;
     function  GetSize: TST_LineEndWidth;
     procedure SetEndType(const Value: TST_LineEndType);
     procedure SetLen(const Value: TST_LineEndLength);
     procedure SetSize(const Value: TST_LineEndWidth);
protected
     FStyle: TCT_LineEndProperties;
public
     constructor Create(AStyle: TCT_LineEndProperties);

     property EndType: TST_LineEndType read GetEndType write SetEndType;
     property Size: TST_LineEndWidth read GetSize write SetSize;
     property Len: TST_LineEndLength read GetLen write SetLen;
     end;

type TXLSDrwLineStyle = class(TXLSDrwShapeProp)
private
     function  GetCapType: TST_LineCap;
     function  GetCompoundType: TST_CompoundLine;
     function  GetDashType: TST_PresetLineDashVal;
     function  GetJoinType: TXLSDrwLineJoinType;
     function  GetWidth: double;
     procedure SetCapType(const Value: TST_LineCap);
     procedure SetCompoundType(const Value: TST_CompoundLine);
     procedure SetDashType(const Value: TST_PresetLineDashVal);
     procedure SetJoinType(const Value: TXLSDrwLineJoinType);
     procedure SetWidth(const Value: double);
     procedure SetHasBeginArrow(const Value: boolean);
     procedure SetHasEndArrow(const Value: boolean);
     function  GetHasBeginArrow: boolean;
     function  GetHasEndArrow: boolean;
protected
     FFill      : TXLSDrwFillChoice;
     FLineStyle : TCT_LineProperties;
     FIsOpenLine: boolean;
     FBeginArrow: TXLSDrwLineEndStyle;
     FEndArrow: TXLSDrwLineEndStyle;
public
     constructor Create(ALineStyle: TCT_LineProperties; AIsOpenLine: boolean);
     destructor Destroy; override;

     function RGB_(ADefault: longword): longword;

     function  CanHaveArrows: boolean;

     property Fill         : TXLSDrwFillChoice read FFill;
     property Width        : double read GetWidth write SetWidth;
     property CompoundType : TST_CompoundLine read GetCompoundType write SetCompoundType;
     property DashType     : TST_PresetLineDashVal read GetDashType write SetDashType;
     property CapType      : TST_LineCap read GetCapType write SetCapType;
     property JoinType     : TXLSDrwLineJoinType read GetJoinType write SetJoinType;
     property HasBeginArrow: boolean read GetHasBeginArrow write SetHasBeginArrow;
     property HasEndArrow  : boolean read GetHasEndArrow write SetHasEndArrow;
     property BeginArrow   : TXLSDrwLineEndStyle read FBeginArrow;
     property EndArrow     : TXLSDrwLineEndStyle read FEndArrow;
     end;

type TXLSDrwShapeProperies = class(TXLSDrwShapeProp)
private
     function  GetHasLine: boolean;
     procedure SetHasLine(const Value: boolean);
protected
     FSpPr    : TCT_ShapeProperties;
     FOwnsSpPr: boolean;
     FLine    : TXLSDrwLineStyle;
     FFill    : TXLSDrwFillChoice;
public
     constructor Create(ASpPr: TCT_ShapeProperties);
     destructor Destroy; override;

     function FillRGB(ADefault: longword): longword;

     function Assigned: boolean;

     property HasLine: boolean read GetHasLine write SetHasLine;
     property Line   : TXLSDrwLineStyle read FLine;
     property Fill   : TXLSDrwFillChoice read FFill;

     property SpPr   : TCT_ShapeProperties read FSpPr;
     end;

type TXLSDrwTextFont = class(TXLSDrwShapeProp)
private
     function  GetBold: boolean;
     function  GetItalic: boolean;
     function  GetName: AxUCString;
     function  GetSize: double;
     function  GetUnderline: TST_TextUnderlineType;
     procedure SetBold(const Value: boolean);
     procedure SetItalic(const Value: boolean);
     procedure SetName(const Value: AxUCString);
     procedure SetSize(const Value: double);
     procedure SetUnderline(const Value: TST_TextUnderlineType);
     procedure SetColor(const Value: TColor);
protected
     FProps: TCT_TextCharacterProperties;
     FColor: TColor;
     FOwnsProps: boolean;
public
     constructor Create(AProps: TCT_TextCharacterProperties);
     destructor Destroy; override;

     function  Assigned: boolean;

     property Name: AxUCString read GetName write SetName;
     property Size: double read GetSize write SetSize;
     property Bold: boolean read GetBold write SetBold;
     property Italic: boolean read GetItalic write SetItalic;
     property Underline: TST_TextUnderlineType read GetUnderline write SetUnderline;
     property Color: TColor read FColor write SetColor;
     end;

     TXLSDrwTextRuns = class;

     TXLSDrwTextRun = class(TXLSDrwShapeProp)
private
     function  GetText: AxUCString;
     procedure SetText(const Value: AxUCString);
protected
     FRun: TEG_TextRun;
     FFont: TXLSDrwTextFont;
public
     constructor Create(ARun: TEG_TextRun);
     destructor Destroy; override;

     procedure AddFont;

     // If this run is a line break, #13 is returned.
     property Text: AxUCString read GetText write SetText;
     property Font: TXLSDrwTextFont read FFont;
     end;

     TXLSDrwTextPara = class;

     TXLSDrwTextRuns = class(TXLSDrwShapeProp)
private
     function GetItems(Index: integer): TXLSDrwTextRun;
protected
     FParent: TXLSDrwTextPara;
     FItems: TObjectList;
     FRuns: TEG_TextRunXpgList;

     function  Add(ARun: TEG_TextRun): TXLSDrwTextRun; overload;
     function  Add: TXLSDrwTextRun; overload;
public
     constructor Create(AParent: TXLSDrwTextPara; ARuns: TEG_TextRunXpgList);
     destructor Destroy; override;

     procedure Clear;
     function  Count: integer;

     function  AddText(AText: AxUCString): TXLSDrwTextRun;
     procedure AddLineBreak(ARun: TXLSDrwTextRun);
     procedure Delete(const AIndex: integer);

     property Items[Index: integer]: TXLSDrwTextRun read GetItems; default;
     end;

     TXLSDrwTextParas = class;

     TXLSDrwTextPara = class(TXLSDrwShapeProp)
private
     function  GetPlainText: AxUCString;
     procedure SetPlainText(const Value: AxUCString);
     function  GetAlign: TXLSDrwTextAlign;
     procedure SetAlign(const Value: TXLSDrwTextAlign);
protected
     FParent: TXLSDrwTextParas;
     FPara: TCT_TextParagraph;
     FRuns: TXLSDrwTextRuns;
public
     constructor Create(AParent: TXLSDrwTextParas; APara: TCT_TextParagraph);
     destructor Destroy; override;

     procedure Clear;

     function  AddText(const AText: AxUCString): TXLSDrwTextRun;
     procedure AddLineBreak;

     property PlainText: AxUCString read GetPlainText write SetPlainText;
     property Runs: TXLSDrwTextRuns read FRuns;
     property Align: TXLSDrwTextAlign read GetAlign write SetAlign;
     end;

     TXLSDrwTextBody = class;

     TXLSDrwTextParas = class(TXLSDrwShapeProp)
private
     function  GetItems(Index: integer): TXLSDrwTextPara;
protected
     FParent: TXLSDrwTextBody;
     FItems: TObjectList;
     FParas: TCT_TextParagraphXpgList;

     function Add(APara: TCT_TextParagraph): TXLSDrwTextPara; overload;
public
     constructor Create(AParent: TXLSDrwTextBody; AParas: TCT_TextParagraphXpgList);
     destructor Destroy; override;

     function  Add: TXLSDrwTextPara; overload;
     function  LastFont: TXLSDrwTextFont;

     function  Count: integer;
     procedure Clear;

     property Items[Index: integer]: TXLSDrwTextPara read GetItems; default;
     end;

     TXLSDrwTextBody = class(TXLSDrwShapeProp)
protected
     FBody: TCT_TextBody;
     FParas: TXLSDrwTextParas;
     FDefaultFont: TXLSDrwTextFont;

     function  GetPlainText: AxUCString;
     procedure SetPlainText(const Value: AxUCString);
public
     constructor Create(ABody: TCT_TextBody);
     destructor Destroy; override;

     property Paras: TXLSDrwTextParas read FParas;
     // The default font is only for the first text run in the first paragraph.
     property DefaultFont: TXLSDrwTextFont read FDefaultFont;
     property PlainText: AxUCString read GetPlainText write SetPlainText;
     end;

     TXLSDrwText = class(TXLSDrwShapeProp)
private
     function  GetPlainText: AxUCString;
     function  GetReF: TXLSRelCells;
     procedure SetPlainText(const Value: AxUCString);
     procedure SetRef(const Value: TXLSRelCells);
protected
     FTx  : TCT_Tx;
     FBody: TXLSDrwTextBody;
     FRef : TCT_StrRef;

     procedure SetType(ARef: boolean);
public
     constructor Create(ATx: TCT_Tx);

     property Ref      : TXLSRelCells read GetReF write SetRef;
     property PlainText: AxUCString read GetPlainText write SetPlainText;
     property Body     : TXLSDrwTextBody read FBody;
     end;

type TXLSDrawing = class;

     TXLSDrawingObject = class(TObject)
protected
     FOwner: TXLSDrawing;
public
     constructor Create(AOwner: TXLSDrawing);
     end;

     TXLSDrwTwoPosShape = class(TXLSDrawingObject)
private
     function  GetCol1: integer;
     function  GetCol1Offs: double;
     function  GetCol2: integer;
     function  GetCol2Offs: double;
     function  GetRow1: integer;
     function  GetRow1Offs: double;
     function  GetRow2: integer;
     function  GetRow2Offs: double;
     procedure SetCol1(const Value: integer);
     procedure SetCol1Offs(const Value: double);
     procedure SetCol2(const Value: integer);
     procedure SetCol2Offs(const Value: double);
     procedure SetRow1(const Value: integer);
     procedure SetRow1Offs(const Value: double);
     procedure SetRow2(const Value: integer);
     procedure SetRow2Offs(const Value: double);
protected
     FAnchor: TCT_TwoCellAnchor;

     procedure SetAnchor(AAnchor: TCT_TwoCellAnchor); virtual;
public
     procedure SetArea(AArea: AxUCString);

     function  Hit(const ACol,ARow: integer): boolean;

     property Anchor  : TCT_TwoCellAnchor read FAnchor;

     property Col1    : integer read GetCol1     write SetCol1;
     property Col1Offs: double  read GetCol1Offs write SetCol1Offs;
     property Col2    : integer read GetCol2     write SetCol2;
     property Col2Offs: double  read GetCol2Offs write SetCol2Offs;
     property Row1    : integer read GetRow1     write SetRow1;
     property Row1Offs: double  read GetRow1Offs write SetRow1Offs;
     property Row2    : integer read GetRow2     write SetRow2;
     property Row2Offs: double  read GetRow2Offs write SetRow2Offs;
     end;

     TXLSDrawingImage = class(TXLSDrwTwoPosShape)
private
     function  GetPositioning: TDrwOptImagePositioning;
     procedure SetPositioning(const Value: TDrwOptImagePositioning);
     function  GetImageType: TXc12ImageType;
     function  GetImageTypeStr: AxUCString;
     function  GetUniqueName: AxUCString;
protected
     FBitmap: TBitmap;

     procedure SetAnchor(AAnchor: TCT_TwoCellAnchor); override;
public
     destructor Destroy; override;

     function  CreateBitmap: TBitmap;
     procedure CacheBitmap;
     procedure SaveToStream(AStream: TStream);
     // If filename don't have dot and extension, it is added,
     // AFilename = 'MyImage' ==> diskfile = MyImage.jpg
     procedure SaveToFile(AFilename: AxUCString);
     // Used by (html) export.
     procedure TagForWriting;
     function  IsTaggedForWriting: boolean;

     property Positioning: TDrwOptImagePositioning read GetPositioning write SetPositioning;
     property ImageType: TXc12ImageType read GetImageType;
     property ImageTypeStr: AxUCString read GetImageTypeStr;
     property UniqueName: AxUCString read GetUniqueName;

     property Bitmap: TBitmap read FBitmap;
     end;

     TXLSDrawingShape = class(TXLSDrwTwoPosShape)
private

protected
     FShapeType: TST_ShapeType;

     procedure SetAnchor(AAnchor: TCT_TwoCellAnchor); override;
public
     property ShapeType: TST_ShapeType read FShapeType write FShapeType;
     end;

     TXLSDrawingWordArt = class(TXLSDrwTwoPosShape)
protected
     FAnchor: TCT_OneCellAnchor;

     procedure SetAnchor(AAnchor: TCT_OneCellAnchor); reintroduce;
public
     constructor Create;
     destructor Destroy; override;
     end;

     TXLSDrawingTextBox = class(TXLSDrwTwoPosShape)
private
     function  GetPlainText: AxUCString;
     procedure SetPlainText(const Value: AxUCString);
protected
     FCurrIndex: integer;

     procedure SetAnchor(AAnchor: TCT_TwoCellAnchor); override;
public
     constructor Create(AOwner: TXLSDrawing);

     property PlainText: AxUCString read GetPlainText write SetPlainText;
     end;

     TXLSDrawingObjects = class(TXLSDrawingObject)
protected
     FItems: TObjectList;
public
     constructor Create(AOwner: TXLSDrawing);
     destructor Destroy; override;

     procedure Clear;
     function  Count: integer;
     end;

     TXLSDrawingImages = class(TXLSDrawingObjects)
protected
     function  GetItems(Index: integer): TXLSDrawingImage;
     function  Add(AOwner: TXLSDrawing): TXLSDrawingImage;
public
     function  Find(const ACol,ARow: integer): TXLSDrawingImage;
     function  FindTopLeft(const ACol,ARow: integer): TXLSDrawingImage;
     procedure MaxDimension(out ACol,ARow: integer);

     property Items[Index: integer]: TXLSDrawingImage read GetItems; default;
     end;

     TXLSDrawingShapes = class(TXLSDrawingObjects)
protected
     function  GetItems(Index: integer): TXLSDrawingShape;
     function  Add: TXLSDrawingShape;
public

     property Items[Index: integer]: TXLSDrawingShape read GetItems; default;
     end;

     TXLSDrawingWordArts = class(TXLSDrawingObjects)
protected
     function  GetItems(Index: integer): TXLSDrawingWordArt;
     function  Add: TXLSDrawingWordArt;
public

     property Items[Index: integer]: TXLSDrawingWordArt read GetItems; default;
     end;

     TXLSDrawingTextBoxes = class(TXLSDrawingObjects)
protected
     FBIFFTextBox: TXLSDrawingTextBox;

     function  GetItems(Index: integer): TXLSDrawingTextBox;
     function  Add(AAnchor: TCT_TwoCellAnchor): TXLSDrawingTextBox; overload;
public
     destructor Destroy; override;

     function  Add: TXLSDrawingTextBox; overload;

     property Items[Index: integer]: TXLSDrawingTextBox read GetItems; default;
     end;

     TXLSDrawingChart = class(TXLSDrwTwoPosShape)
private
     function GetChartSpace: TCT_ChartSpace;
protected
     FChartSpace: TCT_ChartSpace;
public
     constructor Create(AOwner: TXLSDrawing; AChart: TXPGDocXLSXChart);
     destructor Destroy; override;

     property ChartSpace: TCT_ChartSpace read GetChartSpace;
     end;

     TXLSDrawingCharts = class(TXLSDrawingObject)
private
     function GetItems(Index: integer): TXLSDrawingChart;
protected
     FErrors: TXLSErrorManager;
     FOwner : TXLSDrawing;
     FItems : TObjectList;

     function  MakeChartPrologue(ACol, ARow: integer): TXPGDocXLSXChart;
     procedure MakeChartEpilogue(AChart : TXPGDocXLSXChart; AAxId: TCT_UnsignedIntXpgList);
     procedure Make3D(AChart: TCT_Chart);
public
     constructor Create(AErrors: TXLSErrorManager; AOwner: TXLSDrawing);
     destructor Destroy; override;

     procedure Clear;
     function Count: integer;

     function  Find(const ACol,ARow: integer): TXLSDrawingChart; overload;

     function _FileAdd(AChart: TXPGDocXLSXChart): TXLSDrawingChart;

     procedure GetChartFont(ATxt: TCT_TextBody; AFont: TXc12Font); overload;
     procedure GetChartFont(ARPr: TCT_TextCharacterProperties; AFont: TXc12Font); overload;

     procedure SetChartFont(ATxt: TCT_TextBody; AFont: TXc12Font); overload;
     procedure SetChartFont(ARPr: TCT_TextCharacterProperties; AFont: TXc12Font); overload;

     function MakeBarChart(ASrcArea: TXLSRelCells; ACol,ARow: integer; AHasSeriesNames: boolean; A3D: boolean = False): TCT_ChartSpace;
     function MakeAreaChart(ASrcArea: TXLSRelCells; ACol,ARow: integer;  A3D: boolean = False): TCT_ChartSpace;
     function MakeLineChart(ASrcArea: TXLSRelCells; ACol,ARow: integer; A3D: boolean = False): TCT_ChartSpace;
     function MakeRadarChart(ASrcArea: TXLSRelCells; ACol,ARow: integer): TCT_ChartSpace;
     function MakeScatterChart(ASrcArea: TXLSRelCells; ACol,ARow: integer): TCT_ChartSpace;
     function MakeDoughnutChart(ASrcArea: TXLSRelCells; ACol,ARow: integer): TCT_ChartSpace;
     function MakePieChart(ASrcArea: TXLSRelCells; ACol,ARow: integer; A3D: boolean = False): TCT_ChartSpace;
     function MakeOfPieChart(ASrcArea: TXLSRelCells; ACol,ARow: integer): TCT_ChartSpace;
     function MakeSurfaceChart(ASrcArea: TXLSRelCells; ACol,ARow: integer): TCT_ChartSpace;
     function MakeBubbleChart(ASrcArea: TXLSRelCells; ACol,ARow: integer; A3D: boolean = False): TCT_ChartSpace;

     property Items[Index: integer]: TXLSDrawingChart read GetItems; default;
     end;

     TXLSDrawingEditor = class(TObject)
protected
public
     end;

     TXLSDrawingEditorShape = class(TXLSDrawingEditor)
protected
     FAnchor: TCT_TwoCellAnchor;
     FShapeProperies: TXLSDrwShapeProperies;
public
     constructor Create(AAnchor: TCT_TwoCellAnchor);
     destructor Destroy; override;

     property ShapeProperies: TXLSDrwShapeProperies read FShapeProperies;
     end;

     TXLSDrawingEditorTextBox = class(TXLSDrawingEditor)
protected
     FEx12Body: TCT_TextBody;
     FBody: TXLSDrwTextBody;
public
     constructor Create(AEx12Body: TCT_TextBody);
     destructor Destroy; override;

     property Body: TXLSDrwTextBody read FBody;
     end;

     TXLSDrawingEditorImage = class(TXLSDrawingEditor)
private
     function  GetCmHeight: double;
     function  GetCmWidth: double;
     function  GetInchHeight: double;
     function  GetInchWidth: double;
     function  GetOriginalHeight: integer;
     function  GetOriginalWidth: integer;
     procedure SetCmHeight(const Value: double);
     procedure SetCmWidth(const Value: double);
     procedure SetInchHeight(const Value: double);
     procedure SetInchWidth(const Value: double);
protected
     FOwner      : TXLSDrawing;

     FImage      : TXLSDrawingImage;
     FBlip       : TCT_Blip;

     FPixelWidth : integer;
     FPixelHeight: integer;
     FCmWidth    : double;
     FCmHeight   : double;

     FKeepAspect : boolean;

     procedure CalcSize;
     procedure ScaleWidth(const APercent: double);
     procedure ScaleHeight(const APercent: double);
public
     constructor Create(AOwner: TXLSDrawing; AImage: TXLSDrawingImage);

     procedure Scale(const APercent: double);

     property OriginalWidth: integer read GetOriginalWidth;
     property OriginalHeight: integer read GetOriginalHeight;
     property CmWidth: double read GetCmWidth write SetCmWidth;
     property CmHeight: double read GetCmHeight write SetCmHeight;
     property InchWidth: double read GetInchWidth write SetInchWidth;
     property InchHeight: double read GetInchHeight write SetInchHeight;

     property KeepAspect: boolean read FKeepAspect write FKeepAspect;
     end;

     TXLSDrawingEditorChart = class(TXLSDrawingEditor)
private
     function  GetValues: AxUCString;
     procedure SetValues(const Value: AxUCString);
     function  GetCategories: AxUCString;
     procedure SetCategories(const Value: AxUCString);
protected
     FChart: TCT_Chart;
public
     constructor Create(AChart: TXLSDrawingChart);

     property Values: AxUCString read GetValues write SetValues;
     property Categories: AxUCString read GetCategories write SetCategories;
     end;

     TXLSDrawing = class(TObject)
protected
     // Seems to at least have to be 2 for the first shape on an empty drawing.
     // Set to 1000 now to be on the safe side.
     FMaxShapeId : integer;

     FManager    : TXc12Manager;
     FDrw        : TCT_Drawing;
     FGrManager  : TXc12GraphicManager;

     FErrors     : TXLSErrorManager;
     FColumns    : TXLSColumns;
     FRows       : TXLSRows;

     FImages     : TXLSDrawingImages;
     FShapes     : TXLSDrawingShapes;
     FWordArts   : TXLSDrawingWordArts;
     FTextBoxes  : TXLSDrawingTextBoxes;
     FCharts     : TXLSDrawingCharts;

     FAddImage97 : TAddImageEvent97;
{$ifdef XLS_BIFF}
     FBIFFDrawing: TDrawingObjects;
{$endif}
     // Used when creating TCT_XXX objects;
     FDocBase    : TXPGDocBase;

     procedure CheckShapeId(AObjChoice: TEG_ObjectChoices);
     function  GetShapeId: integer;

     function  DoAddImage(AAnchor: TCT_TwoCellAnchor): TXLSDrawingImage;
     function  DoAddShape(AAnchor: TCT_TwoCellAnchor): TXLSDrawingShape;
     function  DoAddTextBox(AAnchor: TCT_TwoCellAnchor): TXLSDrawingTextBox;
     function  DoAddWordArt(AAnchor: TCT_OneCellAnchor): TXLSDrawingWordArt;
     function  DoAddChart(AAnchor: TCT_TwoCellAnchor): TXLSDrawingChart;

     function  DoInsertImage(AImage: TXc12GraphicImage; AWidth, AHeight, ACol, ARow: integer; const AColOffset, ARowOffset: double): TXLSDrawingImage;
public
     constructor Create(AManager: TXc12Manager; ADrw: TCT_Drawing; AErrors: TXLSErrorManager);
     destructor Destroy; override;

     procedure Clear;

     procedure SetProps(ACols: TXLSColumns; ARows: TXLSRows);
     procedure AfterRead;

     // Return object must be destroyed.
     function  CreateShapeProps: TXLSDrwShapeProperies;

     // AImageName must contain the image type, eg: Image.jpg, Image.jpeg, Image.png
     function  InsertImage(const AFilename: AxUCString; const ACol, ARow: integer; const AColOffset, ARowOffset: double; const AScale: double = 1): TXLSDrawingImage; overload;
     function  InsertImage(const AImageName: AxUCString; AImageStream: TStream; const ACol, ARow: integer; const AColOffset, ARowOffset: double; const AScale: double = 1): TXLSDrawingImage; overload;

{$ifdef XLS_BIFF}
     function  InsertImage97(const AImageName: AxUCString; AImageStream: TStream; APicture97: TDrwPicture): TXLSDrawingImage;
{$endif}

     function  InsertTextBox(const AText: AxUCString; const ACol1, ARow1, ACol2, ARow2: integer): TXLSDrawingTextBox; overload;
     function  InsertTextBox(const AText: AxUCString; const AArea: AxUCString): TXLSDrawingTextBox; overload;

     function  InsertShape(const AShapeType: TST_ShapeType; const ACol1, ARow1: integer; const ACol1Offset, ARow1Offset: double; const ACol2, ARow2: integer; const ACol2Offset, ARow2Offset: double): TXLSDrawingShape;

     function  EditTextBox(ATextBox: TXLSDrawingTextBox): TXLSDrawingEditorTextBox;
     function  EditShape(AShape: TXLSDrwTwoPosShape): TXLSDrawingEditorShape;
     function  EditImage(AImage: TXLSDrawingImage): TXLSDrawingEditorImage;
     function  EditChart(AChart: TXLSDrawingChart): TXLSDrawingEditorChart;

     procedure Copy(Col1,Row1,Col2,Row2: integer; DestCol,DestRow: integer);
     procedure Move(const Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     procedure Delete(ACol1,ARow1,ACol2,ARow2: integer);

     property Xc12Drw  : TCT_Drawing read FDrw;
     property Images   : TXLSDrawingImages read FImages;
     property Shapes   : TXLSDrawingShapes read FShapes;
     property WordArts : TXLSDrawingWordArts read FWordArts;
     property TextBoxes: TXLSDrawingTextBoxes read FTextBoxes;
     property Charts   : TXLSDrawingCharts read FCharts;

     property OnAddImage97: TAddImageEvent97 read FAddImage97 write FAddImage97;
{$ifdef XLS_BIFF}
     property BIFFDrawing: TDrawingObjects read FBIFFDrawing write FBIFFDrawing;
{$endif}
     end;

procedure SetChartFont(AFont: TXc12Font; AValAx: TCT_ValAx); overload;
procedure SetChartFont(AFont: TXc12Font; ACatAx: TCT_CatAx); overload;
procedure SetChartFont(AFont: TXc12Font; ALegend: TCT_Legend); overload;
procedure SetChartFont(AFont: TXc12Font; ATitle: TCT_Title); overload;
procedure SetChartFont(AFont: TXc12Font; AChart: TCT_ChartSpace); overload;

implementation

procedure DoSetChartFont(AFont: TXc12Font; ATxPr: TCT_TextBody);
begin
  ATxPr.Create_Paras;
  ATxPr.Paras.Add;
  ATxPr.Paras[0].Create_PPr;
  ATxPr.Paras[0].PPr.Create_defRPr;
  ATxPr.Paras[0].PPr.defRPr.Create_Latin;

  ATxPr.Paras[0].PPr.defRPr.Latin.Typeface := AFont.Name;
  ATxPr.Paras[0].PPr.defRPr.Sz := Round(AFont.Size * 100);
  ATxPr.Paras[0].PPr.defRPr.B := xfsBold in AFont.Style;
  ATxPr.Paras[0].PPr.defRPr.I := xfsItalic in AFont.Style;

  if AFont.Color.ARGB <> 0 then begin
    ATxPr.Paras[0].PPr.defRPr.FillProperties.Create_SolidFill;
    ATxPr.Paras[0].PPr.defRPr.FillProperties.SolidFill.ColorChoice.Create_SrgbClr;
    ATxPr.Paras[0].PPr.defRPr.FillProperties.SolidFill.ColorChoice.SrgbClr.Val := RevRGB(AFont.Color.ARGB);
  end;
end;

procedure SetChartFont(AFont: TXc12Font; AValAx: TCT_ValAx);
begin
  DoSetChartFont(AFont,AValAx.Shared.Create_TxPr);
end;

procedure SetChartFont(AFont: TXc12Font; ACatAx: TCT_CatAx);
begin
  DoSetChartFont(AFont,ACatAx.Shared.Create_TxPr);
end;

procedure SetChartFont(AFont: TXc12Font; ALegend: TCT_Legend); overload;
begin
  DoSetChartFont(AFont,ALegend.Create_TxPr);
end;

procedure SetChartFont(AFont: TXc12Font; ATitle: TCT_Title); overload;
begin
  DoSetChartFont(AFont,ATitle.Create_TxPr);
end;

procedure SetChartFont(AFont: TXc12Font; AChart: TCT_ChartSpace); overload;
begin
  DoSetChartFont(AFont,AChart.Create_TxPr);
end;

function PixelsToEMU(APPI,APixels: int64): int64;
begin
  Result := (914400 * APixels) div APPI;
end;

function EMUToPixels(APPI,AEMU: int64): int64;
begin
  Result := (AEMU * APPI) div 914400;
end;

{ TXLSDrawing }

procedure TXLSDrawing.AfterRead;
var
  i: integer;
  Two: TCT_TwoCellAnchor;
  One: TCT_OneCellAnchor;
begin
  for i := 0 to FDrw.Anchors.TwoCellAnchors.Count - 1 do begin
    Two := FDrw.Anchors.TwoCellAnchors[i];
    CheckShapeId(Two.Objects);

    if (Two.Objects.Pic <> Nil) and Two.Objects.Pic.Available and (Two.Objects.Pic.BlipFill.Blip <> Nil) then
      DoAddImage(Two)
    else if (Two.Objects.Sp <> Nil) and Two.Objects.Sp.NvSpPr.CNvSpPr.TxBox then
      DoAddTextBox(Two)
    else if (Two.Objects.GraphicFrame <> Nil) and (Two.Objects.GraphicFrame.IsChart) then
      DoAddChart(Two)
    // If it not fit anywhere else, its a shape.
    else if (Two.Objects.Sp <> Nil) and Two.Objects.Sp.Available and (Two.Objects.Sp.SpPr.Geometry.PrstGeom <> NIl) and Two.Objects.Sp.SpPr.Geometry.PrstGeom.Available then
      DoAddShape(Two);
  end;

  for i := 0 to FDrw.Anchors.OneCellAnchors.Count - 1 do begin
    One := FDrw.Anchors.OneCellAnchors[i];
    CheckShapeId(One.Objects);

    if (One.Objects.Sp <> Nil) and One.Objects.Sp.Available and
       (One.Objects.Sp.TxBody <> Nil) and One.Objects.Sp.TxBody.BodyPr.Available and
       (One.Objects.Sp.TxBody.BodyPr.Scene3d <> Nil) and One.Objects.Sp.TxBody.BodyPr.Scene3d.Available then
      DoAddWordArt(One);
  end;

  for i := 0 to FDrw.Anchors.AbsoluteAnchors.Count - 1 do begin
    CheckShapeId(FDrw.Anchors.AbsoluteAnchors[i].Objects);
  end;
end;

procedure TXLSDrawing.CheckShapeId(AObjChoice: TEG_ObjectChoices);
begin
  if AObjChoice.Sp <> Nil then begin
    if (AObjChoice.Sp.NvSpPr <> Nil) and (AObjChoice.Sp.NvSpPr.CNvPr <> Nil) then
      FMaxShapeId := Max(FMaxShapeId,AObjChoice.Sp.NvSpPr.CNvPr.Id);
  end
  else if AObjChoice.GrpSp <> Nil then begin
    if (AObjChoice.GrpSp.NvGrpSpPr <> Nil) and (AObjChoice.GrpSp.NvGrpSpPr.CNvPr <> Nil) then
      FMaxShapeId := Max(FMaxShapeId,AObjChoice.GrpSp.NvGrpSpPr.CNvPr.Id);
  end
  else if AObjChoice.Pic <> Nil then begin
    if (AObjChoice.Pic.NvPicPr <> Nil) and (AObjChoice.Pic.NvPicPr.CNvPr <> Nil) then
      FMaxShapeId := Max(FMaxShapeId,AObjChoice.Pic.NvPicPr.CNvPr.Id);
  end
end;

procedure TXLSDrawing.Clear;
begin
  FImages.Clear;
  FShapes.Clear;
  FWordArts.Clear;

  FMaxShapeId := 1000;
end;

procedure TXLSDrawing.Copy(Col1, Row1, Col2, Row2, DestCol, DestRow: integer);
begin

end;

constructor TXLSDrawing.Create(AManager: TXc12Manager; ADrw: TCT_Drawing; AErrors: TXLSErrorManager);
begin
  FManager := AManager;
  FGrManager := FManager.GrManager;
  FDrw := ADrw;
  FErrors := AErrors;

  FMaxShapeId := 1000;

  FImages := TXLSDrawingImages.Create(Self);
  FShapes := TXLSDrawingShapes.Create(Self);
  FWordArts := TXLSDrawingWordArts.Create(Self);
  FTextBoxes := TXLSDrawingTextBoxes.Create(Self);
  FCharts := TXLSDrawingCharts.Create(FErrors,Self);

  FDocBase := TXPGDocBase.Create(FGrManager);
end;

function TXLSDrawing.CreateShapeProps: TXLSDrwShapeProperies;
var
  SpPr: TCT_ShapeProperties;
begin
  SpPr := TCT_ShapeProperties.Create(FDocBase);
  SpPr.Create_Ln;

  Result := TXLSDrwShapeProperies.Create(SpPr);
  Result.FOwnsSpPr := True;
end;

procedure TXLSDrawing.Delete(ACol1, ARow1, ACol2, ARow2: integer);
begin
end;

destructor TXLSDrawing.Destroy;
begin
  FImages.Free;
  FShapes.Free;
  FWordArts.Free;
  FTextBoxes.Free;

  FCharts.Free;

  FDocBase.Free;

  inherited;
end;

function TXLSDrawing.DoAddChart(AAnchor: TCT_TwoCellAnchor): TXLSDrawingChart;
begin
  if (AAnchor.Objects.GraphicFrame <> Nil) and
     (AAnchor.Objects.GraphicFrame.Graphic <> Nil) and
     (AAnchor.Objects.GraphicFrame.Graphic.GraphicData <> Nil) and
     (AAnchor.Objects.GraphicFrame.Graphic.GraphicData.Chart <> Nil) then begin

    Result := FCharts._FileAdd(AAnchor.Objects.GraphicFrame.Graphic.GraphicData.Chart);
    Result.SetAnchor(AAnchor);
  end
  else
    Result := Nil;
end;

function TXLSDrawing.DoAddImage(AAnchor: TCT_TwoCellAnchor): TXLSDrawingImage;
begin
  Result := FImages.Add(Self);
  Result.SetAnchor(AAnchor);
end;

function TXLSDrawing.DoAddShape(AAnchor: TCT_TwoCellAnchor): TXLSDrawingShape;
begin
  Result := FShapes.Add;
  if AAnchor.Objects.Sp = Nil then begin
    AAnchor.Objects.Create_Sp;
    AAnchor.Objects.Sp.Clear;
  end;
  Result.SetAnchor(AAnchor);
end;

function TXLSDrawing.DoAddTextBox(AAnchor: TCT_TwoCellAnchor): TXLSDrawingTextBox;
begin
  Result := FTextBoxes.Add(AAnchor);
end;

function TXLSDrawing.DoAddWordArt(AAnchor: TCT_OneCellAnchor): TXLSDrawingWordArt;
begin
  Result := FWordArts.Add;
  Result.FOwner := Self;
  Result.SetAnchor(AAnchor);
end;

function TXLSDrawing.DoInsertImage(AImage: TXc12GraphicImage; AWidth, AHeight, ACol, ARow: integer; const AColOffset, ARowOffset: double): TXLSDrawingImage;
var
  Anchor: TCT_TwoCellAnchor;
  W,H: integer;
  PPIX,PPIY: integer;
  CO,RO: double;
begin
  CO := Fork(AColOffset,0,1);
  RO := Fork(ARowOffset,0,1);

  FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  W := Round(FColumns[ACol].PixelWidth * CO);
  H := Round(FRows[ARow].PixelHeight * RO);
  if W > AWidth then
    W := AWidth;
  if H > AHeight then
    H := AHeight;
  Anchor := FDrw.Anchors.TwoCellAnchors.Add;
  Anchor.From.Col := ACol;
  Anchor.From.ColOff := PixelsToEMU(PPIX,W);
  Anchor.From.Row := ARow;
  Anchor.From.RowOff := PixelsToEMU(PPIY,H);

  Inc(AWidth,W);
  Inc(AHeight,H);
  while (AWidth - FColumns[ACol].PixelWidth) >= 0 do begin
    Dec(AWidth,FColumns[ACol].PixelWidth);
    Inc(ACol);
  end;
  while (AHeight - FRows[ARow].PixelHeight) >= 0 do begin
    Dec(AHeight,FRows[ARow].PixelHeight);
    Inc(ARow);
  end;
  Anchor.To_.Col := ACol;
  Anchor.To_.ColOff := PixelsToEMU(PPIX,AWidth);
  Anchor.To_.Row := ARow;
  Anchor.To_.RowOff := PixelsToEMU(PPIY,AHeight);

  Anchor.Objects.Create_Pic;
  Anchor.Objects.Pic.Clear;

  Anchor.Objects.Pic.Create_NvPicPr;
  Anchor.Objects.Pic.NvPicPr.Create_CNvPr;
  Anchor.Objects.Pic.NvPicPr.CNvPr.Id := AImage.Index + 1;
  Anchor.Objects.Pic.NvPicPr.CNvPr.Name := Format('Picture %d',[AImage.Index + 1]);
  Anchor.Objects.Pic.NvPicPr.CNvPr.Descr := AImage.Descr;

  Anchor.Objects.Pic.NvPicPr.Create_CNvPicPr;
  Anchor.Objects.Pic.NvPicPr.CNvPicPr.Create_A_PicLocks;
  Anchor.Objects.Pic.NvPicPr.CNvPicPr.A_PicLocks.NoChangeAspect := True;

  Anchor.Objects.Pic.Create_BlipFill;
  Anchor.Objects.Pic.BlipFill.Create_Blip;
  Anchor.Objects.Pic.BlipFill.Blip.Image := AImage;
  Anchor.Objects.Pic.BlipFill.FillModeProperties.Create_Stretch;

  Anchor.Objects.Pic.Create_SpPr;
  Anchor.Objects.Pic.SpPr.Geometry.Create_PrstGeom;
  Anchor.Objects.Pic.SpPr.Geometry.PrstGeom.Prst := ststRect;

  if Assigned(FAddImage97) then
    FAddImage97(Anchor);

  Result := DoAddImage(Anchor);
end;

function TXLSDrawing.EditChart(AChart: TXLSDrawingChart): TXLSDrawingEditorChart;
begin
  Result := TXLSDrawingEditorChart.Create(AChart);
end;

function TXLSDrawing.EditImage(AImage: TXLSDrawingImage): TXLSDrawingEditorImage;
begin
  Result := TXLSDrawingEditorImage.Create(Self,AImage);
end;

function TXLSDrawing.EditShape(AShape: TXLSDrwTwoPosShape): TXLSDrawingEditorShape;
begin
  if AShape.FAnchor.Objects.Sp <> Nil then
    Result := TXLSDrawingEditorShape.Create(AShape.FAnchor)
  else begin
    FErrors.Error('',XLSWARN_SHAPE_CAN_NOT_EDIT);
    Result := Nil;
  end;
end;

function TXLSDrawing.EditTextBox(ATextBox: TXLSDrawingTextBox): TXLSDrawingEditorTextBox;
begin
  if (ATextBox.FAnchor.Objects.Sp <> Nil) and (ATextBox.FAnchor.Objects.Sp.TxBody <> Nil) then
    Result := TXLSDrawingEditorTextBox.Create(ATextBox.FAnchor.Objects.Sp.TxBody)
  else begin
    FErrors.Error('',XLSWARN_SHAPE_EDIT_TEXTBOX);
    Result := Nil;
  end;
end;

function TXLSDrawing.GetShapeId: integer;
begin
  Result := FMaxShapeId;
  Inc(FMaxShapeId);
end;

function TXLSDrawing.InsertImage(const AFilename: AxUCString; const ACol, ARow: integer; const AColOffset, ARowOffset: double; const AScale: double): TXLSDrawingImage;
var
  Image: TXc12GraphicImage;
begin
  Image := FGrManager.Images.LoadFromFile(AFilename);

  if Image = Nil then
    raise XLSRWException.Create('Image file is not valid')
  else
    Result := DoInsertImage(Image,Round(Image.Width * AScale),Round(Image.Heigh * AScale),ACol,ARow,AColOffset,ARowOffset);
end;

function TXLSDrawing.InsertImage(const AImageName: AxUCString; AImageStream: TStream; const ACol, ARow: integer; const AColOffset, ARowOffset: double; const AScale: double): TXLSDrawingImage;
var
  Image: TXc12GraphicImage;
begin
  Image := FGrManager.Images.LoadFromStream(AImageStream,AImageName);

  Result := DoInsertImage(Image,Round(Image.Width * AScale),Round(Image.Heigh * AScale),ACol,ARow,AColOffset,ARowOffset);
end;

{$ifdef XLS_BIFF}
function TXLSDrawing.InsertImage97(const AImageName: AxUCString; AImageStream: TSTream; APicture97: TDrwPicture): TXLSDrawingImage;
var
  Anchor: TCT_TwoCellAnchor;
  Image : TXc12GraphicImage;
  W,H   : integer;
  PPIX,
  PPIY  : integer;
begin
  Image := FGrManager.Images.LoadFromStream(AImageStream,AImageName);

  if Image <> Nil then begin
    FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

    W := Round(FColumns[APicture97.Col1].PixelWidth * APicture97.Col1Offset);
    H := Round(FRows[APicture97.Row1].PixelHeight * APicture97.Row1Offset);

    Anchor := FDrw.Anchors.TwoCellAnchors.Add;
    Anchor.From.Col := APicture97.Col1;
    Anchor.From.ColOff := PixelsToEMU(PPIX,W);;
    Anchor.From.Row := APicture97.Row1;
    Anchor.From.RowOff := PixelsToEMU(PPIY,H);;

    W := Round(FColumns[APicture97.Col2].PixelWidth * APicture97.Col2Offset);
    H := Round(FRows[APicture97.Row2].PixelHeight * APicture97.Row2Offset);

    Anchor.To_.Col := APicture97.Col2;
    Anchor.To_.ColOff := PixelsToEMU(PPIX,W);;
    Anchor.To_.Row := APicture97.Row2;
    Anchor.To_.RowOff := PixelsToEMU(PPIY,H);;

    Anchor.Objects.Create_Pic;
    Anchor.Objects.Pic.Clear;

    Anchor.Objects.Pic.Create_NvPicPr;
    Anchor.Objects.Pic.NvPicPr.Create_CNvPr;
    Anchor.Objects.Pic.NvPicPr.CNvPr.Id := Image.Index + 1;
    Anchor.Objects.Pic.NvPicPr.CNvPr.Name := Format('Picture %d',[Image.Index + 1]);
    Anchor.Objects.Pic.NvPicPr.CNvPr.Descr := Image.Descr;

    Anchor.Objects.Pic.NvPicPr.Create_CNvPicPr;
    Anchor.Objects.Pic.NvPicPr.CNvPicPr.Create_A_PicLocks;
    Anchor.Objects.Pic.NvPicPr.CNvPicPr.A_PicLocks.NoChangeAspect := True;

    Anchor.Objects.Pic.Create_BlipFill;
    Anchor.Objects.Pic.BlipFill.Create_Blip;
    Anchor.Objects.Pic.BlipFill.Blip.Image := Image;
    Anchor.Objects.Pic.BlipFill.FillModeProperties.Create_Stretch;

    Anchor.Objects.Pic.Create_SpPr;
    Anchor.Objects.Pic.SpPr.Geometry.Create_PrstGeom;
    Anchor.Objects.Pic.SpPr.Geometry.PrstGeom.Prst := ststRect;

    Result := DoAddImage(Anchor);
  end
  else
    Result := Nil;
end;
{$endif}

function TXLSDrawing.InsertShape(const AShapeType: TST_ShapeType; const ACol1, ARow1: integer; const ACol1Offset, ARow1Offset: double; const ACol2, ARow2: integer; const ACol2Offset, ARow2Offset: double): TXLSDrawingShape;
begin
  Result := DoAddShape(FDrw.Anchors.TwoCellAnchors.Add);

  Result.ShapeType := AShapeType;

  Result.Col1 := ACol1;
  Result.Col1Offs := ACol1Offset;
  Result.Row1 := ARow1;
  Result.Row1Offs := ARow1Offset;

  Result.Col2 := ACol2;
  Result.Col2Offs := ACol2Offset;
  Result.Row2 := ARow2;
  Result.Row2Offs := ARow2Offset;

  Result.FAnchor.Objects.Sp.Create_NvSpPr;
  Result.FAnchor.Objects.Sp.Create_NvSpPr.Create_CNvPr;

  Result.FAnchor.Objects.Sp.NvSpPr.CNvPr.Id := 2;
  Result.FAnchor.Objects.Sp.NvSpPr.CNvPr.Name := 'Shape';

  Result.FAnchor.Objects.Sp.Create_SpPr;
  Result.FAnchor.Objects.Sp.SpPr.Geometry.Create_PrstGeom;

  Result.FAnchor.Objects.Sp.SpPr.Geometry.PrstGeom.Prst := AShapeType;

  Result.FAnchor.Objects.Sp.Create_Style;
  Result.FAnchor.Objects.Sp.Style.Create_A_LnRef;
  Result.FAnchor.Objects.Sp.Style.Create_A_FillRef;
  Result.FAnchor.Objects.Sp.Style.Create_A_EffectRef;
  Result.FAnchor.Objects.Sp.Style.Create_A_FontRef;

  Result.FAnchor.Objects.Sp.Style.A_LnRef.Idx := 2;
  Result.FAnchor.Objects.Sp.Style.A_LnRef.A_EG_ColorChoice.Create_SchemeClr;
  Result.FAnchor.Objects.Sp.Style.A_LnRef.A_EG_ColorChoice.SchemeClr.Val := stscvAccent1;
  Result.FAnchor.Objects.Sp.Style.A_LnRef.A_EG_ColorChoice.SchemeClr.ColorTransform.AsInteger['shade'] := 50000;

  Result.FAnchor.Objects.Sp.Style.A_FillRef.Idx := 1;
  Result.FAnchor.Objects.Sp.Style.A_FillRef.A_EG_ColorChoice.Create_SchemeClr;
  Result.FAnchor.Objects.Sp.Style.A_FillRef.A_EG_ColorChoice.SchemeClr.Val := stscvAccent1;

  Result.FAnchor.Objects.Sp.Style.A_EffectRef.Idx := 0;
  Result.FAnchor.Objects.Sp.Style.A_EffectRef.A_EG_ColorChoice.Create_SchemeClr;
  Result.FAnchor.Objects.Sp.Style.A_EffectRef.A_EG_ColorChoice.SchemeClr.Val := stscvAccent1;

  Result.FAnchor.Objects.Sp.Style.A_FontRef.Idx := stfciMinor;
  Result.FAnchor.Objects.Sp.Style.A_FontRef.A_EG_ColorChoice.Create_SchemeClr;
  Result.FAnchor.Objects.Sp.Style.A_FontRef.A_EG_ColorChoice.SchemeClr.Val := stscvLt1;
end;

function TXLSDrawing.InsertTextBox(const AText,AArea: AxUCString): TXLSDrawingTextBox;
var
  C1,R1,C2,R2: integer;
begin
  AreaStrToColRow(AArea,C1,R1,C2,R2);
  Result := InsertTextBox(AText,C1,R1,C2,R2);
end;

function TXLSDrawing.InsertTextBox(const AText: AxUCString; const ACol1, ARow1, ACol2, ARow2: integer): TXLSDrawingTextBox;
var
  Ed: TXLSDrawingEditorTextBox;
begin
  Result := FTextBoxes.Add;

  Result.Col1 := ACol1;
  Result.Col1Offs := 0;
  Result.Row1 := ARow1;
  Result.Row1Offs := 0;
  Result.Col2 := ACol2;
  Result.Col2Offs := 0;
  Result.Row2 := ARow2;
  Result.Row2Offs := 0;

  Ed := EditTextBox(Result);
  try
    Ed.Body.PlainText := AText;
  finally
    Ed.Free;
  end;
end;

procedure TXLSDrawing.Move(const Col1, Row1, Col2, Row2, DestCol, DestRow: integer);
begin
end;

procedure TXLSDrawing.SetProps(ACols: TXLSColumns; ARows: TXLSRows);
begin
  FColumns := ACols;
  FRows := ARows;
end;

{ TXLSDrawingObjectImages }

procedure TXLSDrawingObjects.Clear;
begin
  FItems.Clear;
end;

function TXLSDrawingObjects.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSDrawingObjects.Create(AOwner: TXLSDrawing);
begin
  inherited Create(AOwner);
  FItems := TObjectList.Create;
end;

destructor TXLSDrawingObjects.Destroy;
begin
  FItems.Free;
  inherited;
end;

{ TXLSDrawingImage }

procedure TXLSDrawingImage.CacheBitmap;
begin
  if FBitmap = Nil then
    FBitmap := CreateBitmap;
end;

function TXLSDrawingImage.CreateBitmap: TBitmap;
begin
  if FAnchor.Objects.Pic.BlipFill.Blip.Image <> Nil then
    Result := FAnchor.Objects.Pic.BlipFill.Blip.Image.CreateBitmap
  else
    Result := Nil;
end;

destructor TXLSDrawingImage.Destroy;
begin
  if FBitmap <> Nil then
    FBitmap.Free;

  inherited;
end;

function TXLSDrawingImage.GetImageType: TXc12ImageType;
begin
  Result := FAnchor.Objects.Pic.BlipFill.Blip.Image.ImageType;
end;

function TXLSDrawingImage.GetImageTypeStr: AxUCString;
begin
  Result := FAnchor.Objects.Pic.BlipFill.Blip.Image.TypeExt;
end;

function TXLSDrawingImage.GetPositioning: TDrwOptImagePositioning;
begin
  case FAnchor.EditAs of
    steaTwoCell : Result := doipMoveAndSize;
    steaOneCell : Result := doipMoveButDoNotSize;
    steaAbsolute: Result := doipDoNotMoveOrSize;
    else          Result := doipMoveAndSize;
  end;
end;

function TXLSDrawingImage.GetUniqueName: AxUCString;
begin
  Result := FAnchor.Objects.Pic.BlipFill.Blip.Image.UniqueName;
end;

function TXLSDrawingImage.IsTaggedForWriting: boolean;
begin
  Result := not FAnchor.Objects.Pic.BlipFill.Blip.Image.Written;
end;

procedure TXLSDrawingImage.SaveToFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  if CPos('.',AFilename) < 1 then
    Stream := TFileStream.Create(AFilename + '.' + GetImageTypeStr,fmCreate)
  else
    Stream := TFileStream.Create(AFilename,fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXLSDrawingImage.SaveToStream(AStream: TStream);
var
  Image: TXc12GraphicImage;
begin
  Image := FAnchor.Objects.Pic.BlipFill.Blip.Image;
  AStream.Write(Image.Image^,Image.ImageSize);
end;

procedure TXLSDrawingImage.SetAnchor(AAnchor: TCT_TwoCellAnchor);
begin
  inherited SetAnchor(AAnchor);
end;

procedure TXLSDrawingImage.SetPositioning(const Value: TDrwOptImagePositioning);
begin
  case Value of
    doipMoveButDoNotSize: FAnchor.EditAs := steaOneCell;
    doipMoveAndSize     : FAnchor.EditAs := steaTwoCell;
    doipDoNotMoveOrSize : FAnchor.EditAs := steaAbsolute;
  end;
end;

procedure TXLSDrawingImage.TagForWriting;
begin
  FAnchor.Objects.Pic.BlipFill.Blip.Image.Written := False;
end;

{ TXLSDrawingImages }

function TXLSDrawingImages.Add(AOwner: TXLSDrawing): TXLSDrawingImage;
begin
  Result := TXLSDrawingImage.Create(AOwner);
  FItems.Add(Result);
end;

function TXLSDrawingImages.Find(const ACol, ARow: integer): TXLSDrawingImage;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Hit(ACol,ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSDrawingImages.FindTopLeft(const ACol, ARow: integer): TXLSDrawingImage;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].Col1 = ACol) and (Items[i].Row1 = ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSDrawingImages.GetItems(Index: integer): TXLSDrawingImage;
begin
  Result := TXLSDrawingImage(FItems.Items[Index]);
end;

procedure TXLSDrawingImages.MaxDimension(out ACol, ARow: integer);
var
  i: integer;
begin
  ACol := 0;
  ARow := 0;
  for i := 0 to Count - 1 do begin
    ACol := Max(ACol,Items[i].Col2);
    ARow := Max(ACol,Items[i].Row2);
  end;
end;

{ TXLSDrawingWordArt }

constructor TXLSDrawingWordArt.Create;
begin

end;

destructor TXLSDrawingWordArt.Destroy;
begin
  inherited;
end;

procedure TXLSDrawingWordArt.SetAnchor(AAnchor: TCT_OneCellAnchor);
begin
  FAnchor := AAnchor;
end;

{ TXLSDrawingWordArts }

function TXLSDrawingWordArts.Add: TXLSDrawingWordArt;
begin
  Result := TXLSDrawingWordArt.Create;
  FItems.Add(Result);
end;

function TXLSDrawingWordArts.GetItems(Index: integer): TXLSDrawingWordArt;
begin
  Result := TXLSDrawingWordArt(FItems.Items[Index]);
end;

{ TXLSDrawingText }


{ TXLSDrawingShape }
procedure TXLSDrawingShape.SetAnchor(AAnchor: TCT_TwoCellAnchor);
begin
  inherited SetAnchor(AAnchor);
end;

{ TXLSDrawingShapes }

function TXLSDrawingShapes.Add: TXLSDrawingShape;
begin
  Result := TXLSDrawingShape.Create(FOwner);
  FItems.Add(Result);
end;

function TXLSDrawingShapes.GetItems(Index: integer): TXLSDrawingShape;
begin
  Result := TXLSDrawingShape(FItems[Index]);
end;

{ TXLSDrwTwoPosShape }

function TXLSDrwTwoPosShape.GetCol1: integer;
begin
  Result := FAnchor.From.Col;
end;

function TXLSDrwTwoPosShape.GetCol1Offs: double;
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  if FOwner.FColumns[FAnchor.From.Col].Width > 0  then
    Result := EMUToPixels(PPIX,FAnchor.From.ColOff) / FOwner.FColumns[FAnchor.From.Col].PixelWidth
  else
    Result := 0;
end;

function TXLSDrwTwoPosShape.GetCol2: integer;
begin
  Result := FAnchor.To_.Col;
end;

function TXLSDrwTwoPosShape.GetCol2Offs: double;
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  if FOwner.FColumns[FAnchor.From.Col].Width > 0  then
    Result := EMUToPixels(PPIX,FAnchor.To_.ColOff) / FOwner.FColumns[FAnchor.To_.Col].PixelWidth
  else
    Result := 0;
end;

function TXLSDrwTwoPosShape.GetRow1: integer;
begin
  Result := FAnchor.From.Row;
end;

function TXLSDrwTwoPosShape.GetRow1Offs: double;
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  if FOwner.FRows[FAnchor.From.Row].Height > 0  then
    Result := EMUToPixels(PPIY,FAnchor.From.RowOff) / FOwner.FRows[FAnchor.From.Row].PixelHeight
  else
    Result := 0;
end;

function TXLSDrwTwoPosShape.GetRow2: integer;
begin
  Result := FAnchor.To_.Row;
end;

function TXLSDrwTwoPosShape.GetRow2Offs: double;
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  if FOwner.FRows[FAnchor.To_.Row].Height > 0  then
    Result := EMUToPixels(PPIY,FAnchor.To_.RowOff) / FOwner.FRows[FAnchor.To_.Row].PixelHeight
  else
    Result := 0;
end;

function TXLSDrwTwoPosShape.Hit(const ACol, ARow: integer): boolean;
begin
  Result := (ACol >= Col1) and (ACol <= Col2) and (ARow >= Row1) and (ARow <= Row2);
end;

procedure TXLSDrwTwoPosShape.SetAnchor(AAnchor: TCT_TwoCellAnchor);
begin
  FAnchor := AAnchor;
end;

procedure TXLSDrwTwoPosShape.SetArea(AArea: AxUCString);
var
  C1,R1,C2,R2: integer;
begin
  AreaStrToColRow(AArea,C1,R1,C2,R2);
  Col1 := C1;
  Row1 := R1;
  Col2 := C2;
  Row2 := R2;
end;

procedure TXLSDrwTwoPosShape.SetCol1(const Value: integer);
begin
  FAnchor.From.Col := Value;
end;

procedure TXLSDrwTwoPosShape.SetCol1Offs(const Value: double);
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  FAnchor.From.ColOff := Round(PixelsToEMU(PPIX,FOwner.FColumns[FAnchor.From.Col].PixelWidth) * Fork(Value,0,1));
end;

procedure TXLSDrwTwoPosShape.SetCol2(const Value: integer);
begin
  FAnchor.To_.Col := Value;
end;

procedure TXLSDrwTwoPosShape.SetCol2Offs(const Value: double);
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  FAnchor.To_.ColOff := Round(PixelsToEMU(PPIX,FOwner.FColumns[FAnchor.To_.Col].PixelWidth) * Fork(Value,0,1));
end;

procedure TXLSDrwTwoPosShape.SetRow1(const Value: integer);
begin
  FAnchor.From.Row := Value;
end;

procedure TXLSDrwTwoPosShape.SetRow1Offs(const Value: double);
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  FAnchor.From.RowOff := Round(PixelsToEMU(PPIY,FOwner.FRows[FAnchor.From.Row].PixelHeight) * Fork(Value,0,1));
end;

procedure TXLSDrwTwoPosShape.SetRow2(const Value: integer);
begin
  FAnchor.To_.Row := Value;
end;

procedure TXLSDrwTwoPosShape.SetRow2Offs(const Value: double);
var
  PPIX,PPIY: integer;
begin
  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  FAnchor.To_.RowOff := Round(PixelsToEMU(PPIY,FOwner.FRows[FAnchor.To_.Row].PixelHeight) * Fork(Value,0,1));
end;

{ TXLSDrawingObject }

constructor TXLSDrawingObject.Create(AOwner: TXLSDrawing);
begin
  FOwner := AOwner;
end;

{ TXLSDrawingTextBox }

constructor TXLSDrawingTextBox.Create(AOwner: TXLSDrawing);
begin
  inherited Create(AOwner);
end;

function TXLSDrawingTextBox.GetPlainText: AxUCString;
var
  Body: TXLSDrwTextBody;
begin
{$ifdef XLS_BIFF}
  if FOwner.BIFFDrawing <> Nil then
    Result := FOwner.BIFFDrawing.Texts[FCurrIndex].Text
  else begin
{$endif}
    if (FAnchor.Objects.Sp <> Nil) and (FAnchor.Objects.Sp.TxBody <> Nil) then begin
      Body := TXLSDrwTextBody.Create(FAnchor.Objects.Sp.TxBody);
      try
        Result := Body.PlainText;
      finally
        Body.Free;
      end;
    end
    else
      Result := '';
{$ifdef XLS_BIFF}
  end;
{$endif}
end;

procedure TXLSDrawingTextBox.SetAnchor(AAnchor: TCT_TwoCellAnchor);
begin
  inherited SetAnchor(AAnchor);

  FAnchor := AAnchor;
end;

procedure TXLSDrawingTextBox.SetPlainText(const Value: AxUCString);
var
  Body: TXLSDrwTextBody;
begin
{$ifdef XLS_BIFF}
  if FOwner.BIFFDrawing <> Nil then
    FOwner.BIFFDrawing.Texts[FCurrIndex].Text := Value
  else begin
{$endif}
    if (FAnchor.Objects.Sp <> Nil) and (FAnchor.Objects.Sp.TxBody <> Nil) then begin
      Body := TXLSDrwTextBody.Create(FAnchor.Objects.Sp.TxBody);
      try
        Body.PlainText := Value;
      finally
        Body.Free;
      end;
    end;
{$ifdef XLS_BIFF}
  end;
{$endif}
end;

{ TXLSDrawingTextBoxes }

function TXLSDrawingTextBoxes.Add: TXLSDrawingTextBox;
var
  Anchor: TCT_TwoCellAnchor;
begin
  Anchor := FOwner.FDrw.Anchors.TwoCellAnchors.Add;

  Anchor.Objects.Create_Sp;

  Anchor.Objects.Sp.Create_NvSpPr;
  Anchor.Objects.Sp.NvSpPr.Create_CNvSpPr;
  Anchor.Objects.Sp.NvSpPr.Create_CNvPr;
  Anchor.Objects.Sp.NvSpPr.CNvPr.Id := FOwner.FMaxShapeId + Count + 1;
  Anchor.Objects.Sp.NvSpPr.CNvPr.Name := Format('TextBox %d',[Count + 1]);
  Anchor.Objects.Sp.NvSpPr.CNvSpPr.TxBox := True;

  Anchor.Objects.Sp.Create_SpPr;
  Anchor.Objects.Sp.SpPr.Geometry.Create_PrstGeom;
  Anchor.Objects.Sp.SpPr.Geometry.PrstGeom.Prst := ststRect;
  Anchor.Objects.Sp.SpPr.FillProperties.Create_SolidFill;
  Anchor.Objects.Sp.SpPr.FillProperties.SolidFill.ColorChoice.Create_SchemeClr;
  Anchor.Objects.Sp.SpPr.FillProperties.SolidFill.ColorChoice.SchemeClr.Val := stscvLt1;

  Anchor.Objects.Sp.SpPr.Create_Ln;
  Anchor.Objects.Sp.SpPr.Ln.W := 9525;
  Anchor.Objects.Sp.SpPr.Ln.Cmpd := stclSng;
  Anchor.Objects.Sp.SpPr.Ln.LineFillProperties.Create_SolidFill;
  Anchor.Objects.Sp.SpPr.Ln.LineFillProperties.SolidFill.ColorChoice.Create_SchemeClr;
  Anchor.Objects.Sp.SpPr.Ln.LineFillProperties.SolidFill.ColorChoice.SchemeClr.Val := stscvLt1;
  Anchor.Objects.Sp.SpPr.Ln.LineFillProperties.SolidFill.ColorChoice.SchemeClr.ColorTransform.AsInteger['shade'] := 50000;

  Anchor.Objects.Sp.Create_TxBody;

  Anchor.Objects.Sp.TxBody.Create_BodyPr;
  Anchor.Objects.Sp.TxBody.Create_Paras;

  Result := TXLSDrawingTextBox.Create(FOwner);
  Result.SetAnchor(Anchor);
  FItems.Add(Result);
end;

destructor TXLSDrawingTextBoxes.Destroy;
begin
  if FBIFFTextBox <> Nil then
    FBIFFTextBox.Free;
  inherited;
end;

function TXLSDrawingTextBoxes.Add(AAnchor: TCT_TwoCellAnchor): TXLSDrawingTextBox;
begin
  Result := TXLSDrawingTextBox.Create(FOwner);
  Result.SetAnchor(AAnchor);
  FItems.Add(Result);
end;

function TXLSDrawingTextBoxes.GetItems(Index: integer): TXLSDrawingTextBox;
begin
{$ifdef XLS_BIFF}
  if FOwner.FBIFFDrawing <> Nil then begin
    if FBIFFTextBox = Nil then
      FBIFFTextBox := TXLSDrawingTextBox.Create(FOwner);
    Result := FBIFFTextBox;
    Result.FCurrIndex := Index;
  end
  else
{$endif}
    Result := TXLSDrawingTextBox(FItems.Items[Index]);
end;

{ TXLSDrwFillNone }

constructor TXLSDrwFillNone.Create(AFill: TCT_NoFillProperties);
begin
  FFill := AFill;
end;

function TXLSDrwFillNone.GetFillType: TXLSDrwFillType;
begin
  Result := xdftNone;
end;

{ TXLSDrwFillSolid }

constructor TXLSDrwFillSolid.Create(AFill: TCT_SolidColorFillProperties);
begin
  FFill := AFill;
  FColor := TXLSDrwColor.Create(FFill.ColorChoice);
end;

destructor TXLSDrwFillSolid.Destroy;
begin
  FColor.Free;
  inherited;
end;

function TXLSDrwFillSolid.GetFillType: TXLSDrwFillType;
begin
  Result := xdftSolid;
end;

{ TXLSDrwColor }

constructor TXLSDrwColor.Create(AColorChoice: TEG_ColorChoice);
begin
  FColorChoice := AColorChoice;

  AssignColorTransform;

  if FColorChoice.CheckAssigned <= 0 then
    XlsColorType;
end;

function TXLSDrwColor.GetAsARGB: longword;
begin
  Result := GetAsRGB + (Round(GetTransparency) * $FF) shl 24;
end;

function TXLSDrwColor.GetAsHSL: longword;
begin
  if XlsColorType = xdxctHSL then
    Result := FColorChoice.HslClr.HSL
  else
    raise XLSRWException.Create('Color is not HSL');
end;

function TXLSDrwColor.GetAsPreset: TST_PresetColorVal;
begin
  if XlsColorType = xdxctPreset then
    Result := FColorChoice.PrstClr.Val
  else
    raise XLSRWException.Create('Color is not Preset');
end;

function TXLSDrwColor.GetAsRGB: longword;
begin
  Result := $00000000;
  case XlsColorType of
    xdxctRGB    : Result := FColorChoice.SrgbClr.Val;
    xdxctCRGB   : Result := FColorChoice.SCrgbClr.RGB;
    xdxctHSL    : Result := HSLToRGB(FColorChoice.HslClr.HSL);
    xdxctSystem : begin
{$ifdef BABOON}
      Result := $FFFFFF;
{$else}
      Result := vcl.Graphics.ColorToRGB(clWindow);
{$endif}
    end;
    xdxctScheme : begin
      case FColorChoice.SchemeClr.Val of
        stscvBg1: Result := $00FFFFFF; // ???
        stscvTx1: Result := $00000000; // ???
        stscvBg2: Result := $00FFFFFF; // ???
        stscvTx2: Result := $00000000; // ???
        stscvAccent1: Result := Xc12DefColorScheme[cscAccent1];
        stscvAccent2: Result := Xc12DefColorScheme[cscAccent2];
        stscvAccent3: Result := Xc12DefColorScheme[cscAccent3];
        stscvAccent4: Result := Xc12DefColorScheme[cscAccent4];
        stscvAccent5: Result := Xc12DefColorScheme[cscAccent5];
        stscvAccent6: Result := Xc12DefColorScheme[cscAccent6];
        stscvHlink: Result := Xc12DefColorScheme[cscHLink];
        stscvFolHlink: Result := Xc12DefColorScheme[cscFolHLink];
        stscvPhClr: Result := $00FFFFFF; // ???
        stscvDk1: Result := Xc12DefColorScheme[cscDk1];
        stscvLt1: Result := Xc12DefColorScheme[cscLt1];
        stscvDk2: Result := Xc12DefColorScheme[cscDk2];
        stscvLt2: Result := Xc12DefColorScheme[cscLt2];
      end;
    end;
// {$message warn '__TODO__ Do not use HTMLColors'}
    xdxctPreset : Result := $FFFFFF;
//    xdxctPreset : Result := TColorToRGB(Longword(HTMLColors.Objects[Integer(FColorChoice.PrstClr.Val)]));
  end;

  if FColorTransform <> Nil then
    Result := FColorTransform.Apply(Result);
end;

function TXLSDrwColor.GetAsScheme: TST_SchemeColorVal;
begin
  if XlsColorType = xdxctScheme then
    Result := FColorChoice.SchemeClr.Val
  else
    raise XLSRWException.Create('Color is not Scheme');
end;

function TXLSDrwColor.GetAsSystem: TST_SystemColorVal;
begin
  if XlsColorType = xdxctSystem then
    Result := FColorChoice.SysClr.Val
  else
    raise XLSRWException.Create('Color is not System');
end;

function TXLSDrwColor.GetAsTColor: TColor;
begin
  Result := RGBToTColor(GetAsRGB);
end;

function TXLSDrwColor.GetColorType: TXLSDrwColorType;
begin
  if (FColorChoice.ScrgbClr <> Nil) or (FColorChoice.SrgbClr <> Nil) then
    Result := xdctRGB
  else if (FColorChoice.HslClr <> Nil) then
    Result := xdctHSL
  else if (FColorChoice.PrstClr <> Nil) then
    Result := xdctPreset
  else if (FColorChoice.SchemeClr <> Nil) then
    Result := xdctScheme
  else if (FColorChoice.SysClr <> Nil) then
    Result := xdctSystem
  else begin
    SetAsScheme(stscvAccent1);
    Result := xdctScheme
  end;
end;

function TXLSDrwColor.GetTransparency: double;
begin
  Result := FColorTransform.AsDouble['alpha'];
{$ifdef DELPHI_5}
  try
    Result := Result / 1000;
  except
    Result := 0
  end;
{$else}
  if IsNaN(Result) then
    Result := 0
  else
    Result := Result / 1000;
{$endif}
end;

procedure TXLSDrwColor.SetAsARGB(const Value: longword);
begin
  SetTransparency(((Value and $FF000000) shr 24) / 255);
  case XlsColorType of
    xdxctRGB    : FColorChoice.SrgbClr.Val := Value and $00FF0000;
    xdxctCRGB   : FColorChoice.ScrgbClr.RGB := Value and $00FF0000;
    xdxctHSL,
    xdxctSystem,
    xdxctScheme,
    xdxctPreset : begin

    end;
  end;
end;

procedure TXLSDrwColor.SetAsHSL(const Value: longword);
begin
  FColorChoice.Create_HslClr;
  FColorChoice.HslClr.HSL := Value;
end;

procedure TXLSDrwColor.SetAsPreset(const Value: TST_PresetColorVal);
begin
  FColorChoice.Create_PrstClr;
  FColorChoice.PrstClr.Val := Value;
end;

procedure TXLSDrwColor.SetAsRGB(const Value: longword);
begin
  FColorChoice.Create_SrgbClr;
  FColorChoice.SrgbClr.Val := Value;
end;

procedure TXLSDrwColor.SetAsScheme(const Value: TST_SchemeColorVal);
begin
  FColorChoice.Create_SchemeClr;
  FColorChoice.SchemeClr.Val := Value;
end;

procedure TXLSDrwColor.SetAsSystem(const Value: TST_SystemColorVal);
begin
  FColorChoice.Create_SysClr;
  FColorChoice.SysClr.Val := Value;
end;

procedure TXLSDrwColor.SetAsTColor(const Value: TColor);
begin
  SetAsRGB(TColorToRGB(Value));
end;

procedure TXLSDrwColor.AssignColorTransform;
begin
  case XlsColorType of
    xdxctRGB    : FColorTransform := FColorChoice.SrgbClr.ColorTransform;
    xdxctCRGB   : FColorTransform := FColorChoice.ScrgbClr.ColorTransform;
    xdxctHSL    : FColorTransform := FColorChoice.HslClr.ColorTransform;
    xdxctSystem : FColorTransform := FColorChoice.SysClr.ColorTransform;
    xdxctScheme : FColorTransform := FColorChoice.SchemeClr.ColorTransform;
    xdxctPreset : FColorTransform := FColorChoice.PrstClr.ColorTransform;
    else begin
      SetAsScheme(stscvAccent1);
      FColorTransform := FColorChoice.SchemeClr.ColorTransform;
    end;
  end;
end;

procedure TXLSDrwColor.SetTransparency(const Value: double);
begin
  FColorTransform.AsDouble['alpha'] := Fork(Value,0,1) * 1000;
end;

function TXLSDrwColor.XlsColorType: TXLSDrwXMLColorType;
begin
  if (FColorChoice.ScrgbClr <> Nil) then
    Result := xdxctCRGB
  else if (FColorChoice.SrgbClr <> Nil) then
    Result := xdxctRGB
  else if (FColorChoice.HslClr <> Nil) then
    Result := xdxctHSL
  else if (FColorChoice.PrstClr <> Nil) then
    Result := xdxctPreset
  else if (FColorChoice.SchemeClr <> Nil) then
    Result := xdxctScheme
  else if (FColorChoice.SysClr <> Nil) then
    Result := xdxctSystem
  else begin
    SetAsScheme(stscvAccent1);
    Result := xdxctScheme
  end;
end;

{ TXLSDrwFillPicture }

constructor TXLSDrwFillPicture.Create(ABlip: TCT_BlipFillProperties);
begin
  FBlip := ABlip;

  FBlip.Create_Blip;

  if FBlip.FillModeProperties.Stretch <> Nil then
    FStretch := TXLSDrwFillPictStretchOpts.Create(FBlip.FillModeProperties.Stretch)
  else begin
    FBlip.FillModeProperties.Create_Tile;
    FTile := TXLSDrwFillPictTileOpts.Create(FBlip.FillModeProperties.Tile);
  end;
end;

destructor TXLSDrwFillPicture.Destroy;
begin

  inherited;
end;

function TXLSDrwFillPicture.GetBrightness: double;
begin
  if FBlip.Blip.Lum <> Nil then
    Result := FBlip.Blip.Lum.Bright / 100000
  else
    Result := 0;
end;

function TXLSDrwFillPicture.GetContrast: double;
begin
  if FBlip.Blip.Lum <> Nil then
    Result := FBlip.Blip.Lum.Contrast / 100000
  else
    Result := 0;
end;

function TXLSDrwFillPicture.GetFillType: TXLSDrwFillType;
begin
  Result := xdftPicture;
end;

function TXLSDrwFillPicture.GetRorateWithShape: boolean;
begin
  Result := FBlip.RotWithShape;
end;

function TXLSDrwFillPicture.GetTilePicture: boolean;
begin
  Result := FTile <> Nil;
end;

function TXLSDrwFillPicture.GetTransparency: double;
begin
  if FBlip.Blip.AlphaModFix <> Nil then
    Result := FBlip.Blip.AlphaModFix.Amt / 100000
  else
    Result := 0;
end;

function TXLSDrwFillPicture.LoadPicture: boolean;
begin
  LoadPicture(FFilename);

  Result := True;
end;

function TXLSDrwFillPicture.LoadPicture(const APictureName: AxUCString): boolean;
var
  Image: TXc12GraphicImage;
begin
  Image := FBlip.Blip.Owner.GrManager.Images.FindByFileId(APictureName);
  if Image = Nil then
    Image := FBlip.Blip.Owner.GrManager.Images.LoadFromFile(APictureName);

  Result := Image <> Nil;
  if Result then
    FBlip.Blip.Image := Image;
end;

procedure TXLSDrwFillPicture.SetBrightness(const Value: double);
begin
  FBlip.Blip.Create_Lum;
  FBlip.Blip.Lum.Bright := Round(Value * 100000);

  if (FBlip.Blip.Lum.Contrast = 0) and (FBlip.Blip.Lum.Bright = 0) then
    FBlip.Blip.Free_Lum;
end;

procedure TXLSDrwFillPicture.SetContrast(const Value: double);
begin
  FBlip.Blip.Create_Lum;
  FBlip.Blip.Lum.Contrast := Round(Value * 100000);

  if (FBlip.Blip.Lum.Contrast = 0) and (FBlip.Blip.Lum.Bright = 0) then
    FBlip.Blip.Free_Lum;
end;

procedure TXLSDrwFillPicture.SetRorateWithShape(const Value: boolean);
begin
  FBlip.RotWithShape := Value;
end;

procedure TXLSDrwFillPicture.SetTilePicture(const Value: boolean);
begin
  if Value then begin
    if FTile = Nil then begin
      FreeAndNil(FStretch);
      FBlip.FillModeProperties.Create_Tile;
      FTile := TXLSDrwFillPictTileOpts.Create(FBlip.FillModeProperties.Tile);
    end;
  end
  else begin
    if FStretch = Nil then begin
      FreeAndNil(FTile);
      FBlip.FillModeProperties.Create_Stretch;
      FStretch := TXLSDrwFillPictStretchOpts.Create(FBlip.FillModeProperties.Stretch);
    end;
  end;
end;

procedure TXLSDrwFillPicture.SetTransparency(const Value: double);
begin
  if Value > 0 then begin
    FBlip.Blip.Create_AlphaModFix;
    FBlip.Blip.AlphaModFix.Amt := Round(Fork(Value,0,1) * 100000);
  end
  else
    FBlip.Blip.Free_AlphaModFix;
end;

{ TXLSDrwFillGradientStop }

constructor TXLSDrwFillGradientStop.Create(AStop: TCT_GradientStop);
begin
  FStop := AStop;
  FColor := TXLSDrwColor.Create(FStop.ColorChoice);
end;

destructor TXLSDrwFillGradientStop.Destroy;
begin
  FColor.Free;
  inherited;
end;

function TXLSDrwFillGradientStop.GetStopPosition: double;
begin
  Result := FStop.Pos / 100000;
end;

function TXLSDrwFillGradientStop.GetTransparency: double;
begin
  Result := FColor.Transparency;
end;

procedure TXLSDrwFillGradientStop.SetStopPosition(const Value: double);
begin
  FStop.Pos := Round(Value * 100000);
end;

procedure TXLSDrwFillGradientStop.SetTransparency(const Value: double);
begin
  FColor.Transparency := Value;
end;

{ TXLSDrwFillGradientStops }

function TXLSDrwFillGradientStops.Add: TXLSDrwFillGradientStop;
begin
  Result := TXLSDrwFillGradientStop.Create(FGSList.Gs.Add);
  FItems.Add(Result);
end;

procedure TXLSDrwFillGradientStops.Clear;
begin
  FGSList.Clear;
  FItems.Clear;

  SetDefault;
end;

function TXLSDrwFillGradientStops.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSDrwFillGradientStops.Create(AGSList: TCT_GradientStopList);
var
  i: integer;
begin
  FGSList := AGSList;
  FItems := TObjectList.Create;

  for i := 0 to AGSList.Gs.Count - 1 do
    FItems.Add(TXLSDrwFillGradientStop.Create(FGSList.Gs[i]));

  if FItems.Count < 2 then
    Clear;
end;

destructor TXLSDrwFillGradientStops.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXLSDrwFillGradientStops.GetItems(const Index: integer): TXLSDrwFillGradientStop;
begin
  Result := TXLSDrwFillGradientStop(FItems[Index]);
end;

procedure TXLSDrwFillGradientStops.Remove(const AIndex: integer);
begin
  FItems.Delete(AIndex);
  FGSList.Gs.Delete(AIndex);
end;

procedure TXLSDrwFillGradientStops.SetDefault;
var
  GS: TXLSDrwFillGradientStop;
begin
  GS := Add;
  GS.StopPosition := 0;
  GS.Color.AsRGB := $0000FF;

  GS := Add;
  GS.StopPosition := 1;
  GS.Color.AsRGB := $A0A0FF;
end;

{ TXLSDrwFillGradient }

constructor TXLSDrwFillGradient.Create(AFill: TCT_GradientFillProperties);
begin
  FFill := AFill;
  FStops := TXLSDrwFillGradientStops.Create(FFill.GsLst);
end;

destructor TXLSDrwFillGradient.Destroy;
begin
  FStops.Free;
  inherited;
end;

function TXLSDrwFillGradient.GetAngle: double;
begin
  if FFill.ShadeProperties.Lin <> Nil then
    Result := FFill.ShadeProperties.Lin.Ang / 60000
  else
    Result := 0;
end;

function TXLSDrwFillGradient.GetDirection: TXLSDrwGradientFillDirection;
begin
  if (FFill.ShadeProperties.Path <> Nil) and (FFill.ShadeProperties.Path.Path in [stpstCircle,stpstRect]) and (FFill.ShadeProperties.Path.FillToRect <> Nil) then begin
    if FFill.ShadeProperties.Path.FillToRect.Equal(100000,0,0,100000) and FFill.TileRect.Equal(0,-100000,-100000,0) then
      Result := xdgfdTopLeft
    else if FFill.ShadeProperties.Path.FillToRect.Equal(0,100000,0,100000) and FFill.TileRect.Equal(-100000,0,-100000,0) then
      Result := xdgfdTopRight
    else if FFill.ShadeProperties.Path.FillToRect.Equal(0,100000,100000,0) and FFill.TileRect.Equal(-100000,0,0,-100000) then
      Result := xdgfdBottomLeft
    else if FFill.ShadeProperties.Path.FillToRect.Equal(100000,0,100000,0) and FFill.TileRect.Equal(0,-100000,0,-100000) then
      Result := xdgfdBottomRight
    else if FFill.ShadeProperties.Path.FillToRect.Equal(50000,50000,50000,50000) and FFill.TileRect.Equal(0,0,0,0) then
      Result := xdgfdCentre
    else
      Result := xdgfdNone;
  end
  else
    Result := xdgfdNone;
end;

function TXLSDrwFillGradient.GetFillType: TXLSDrwFillType;
begin
  Result := xdftGradient;
end;

function TXLSDrwFillGradient.GetGradientType: TXLSDrwGradientFillType;
begin
  if FFill.ShadeProperties.Lin <> Nil then
    Result := xdgftLinear
  else if FFill.ShadeProperties.Path <> Nil then begin
    case FFill.ShadeProperties.Path.Path of
      stpstShape : Result := xdgftPath;
      stpstCircle: Result := xdgftRadial;
      stpstRect  : Result := xdgftRectangular;
      else         Result := xdgftPath;
    end;
  end
  else begin
    FFill.ShadeProperties.Create_Lin;
    Result := xdgftLinear
  end;
end;

function TXLSDrwFillGradient.GetRotateWithShape: boolean;
begin
  Result := FFill.RotWithShape;
end;

procedure TXLSDrwFillGradient.SetAngle(const Value: double);
begin
  if FFill.ShadeProperties.Lin <> Nil then
    FFill.ShadeProperties.Lin.Ang := Round(Value * 60000);
end;

procedure TXLSDrwFillGradient.SetDirection(const Value: TXLSDrwGradientFillDirection);
begin
  case GetGradientType of
    xdgftLinear: begin
      FFill.ShadeProperties.Create_Lin;
      case Value of
        xdgfdNone       : ;
        xdgfdCentre     : SetAngle(0);
        xdgfdTopLeft    : SetAngle(45);
        xdgfdTopRight   : SetAngle(135);
        xdgfdBottomLeft : SetAngle(315);
        xdgfdBottomRight: SetAngle(225);
      end;
    end;
    xdgftRadial,
    xdgftRectangular: begin
      FFill.ShadeProperties.Clear;
      FFill.ShadeProperties.Create_Path;
      FFill.ShadeProperties.Path.Create_FillToRect;
      if Value in [xdgfdNone,xdgfdCentre] then
        FFill.Free_TileRect
      else
        FFill.Create_TileRect;
      case Value of
        xdgfdNone       : begin
          FFill.ShadeProperties.Clear;
        end;
        xdgfdCentre     : FFill.ShadeProperties.Path.FillToRect.SetValues(50000,50000,50000,50000);
        xdgfdTopLeft    : begin
          FFill.TileRect.SetValues(0,-100000,-100000,0);
          FFill.ShadeProperties.Path.FillToRect.SetValues(100000,0,0,100000);
        end;
        xdgfdTopRight   : begin
          FFill.TileRect.SetValues(-100000,0,-100000,0);
          FFill.ShadeProperties.Path.FillToRect.SetValues(0,100000,0,100000);
        end;
        xdgfdBottomLeft : begin
          FFill.TileRect.SetValues(-100000,0,0,-100000);
          FFill.ShadeProperties.Path.FillToRect.SetValues(0,100000,100000,0);
        end;
        xdgfdBottomRight: begin
          FFill.TileRect.SetValues(0,-100000,0,-100000);
          FFill.ShadeProperties.Path.FillToRect.SetValues(100000,0,100000,0);
        end;
      end;
    end;
    xdgftPath: ;
  end;
end;

procedure TXLSDrwFillGradient.SetGradientType(const Value: TXLSDrwGradientFillType);
begin
  if Value <> GetGradientType then begin
    FFill.ShadeProperties.Clear;
    case Value of
      xdgftLinear     : begin
        FFill.Free_TileRect;
        FFill.ShadeProperties.Create_Lin;
      end;
      xdgftRadial     : begin
        FFill.ShadeProperties.Create_Path;
        FFill.ShadeProperties.Path.Path := stpstCircle;
      end;
      xdgftRectangular: begin
        FFill.ShadeProperties.Create_Path;
        FFill.ShadeProperties.Path.Path := stpstRect;
      end;
      xdgftPath       : begin
        FFill.ShadeProperties.Create_Path;
        FFill.ShadeProperties.Path.Path := stpstShape;
      end;
    end;
  end;
end;

procedure TXLSDrwFillGradient.SetRotateWithShape(const Value: boolean);
begin
  FFill.RotWithShape := Value;
end;

{ TXLSDrwFillPictTileOpts }

constructor TXLSDrwFillPictTileOpts.Create(ATile: TCT_TileInfoProperties);
begin
  FTile := ATile;
end;

function TXLSDrwFillPictTileOpts.GetAlignment: TST_RectAlignment;
begin
  Result := FTile.Algn;
end;

function TXLSDrwFillPictTileOpts.GetMirrorType: TXLSDrwMirrorType;
begin
  Result := TXLSDrwMirrorType(FTile.Flip);
end;

function TXLSDrwFillPictTileOpts.GetOffsetX: double;
begin
  Result := FTile.Tx / EMU_PER_PT;
end;

function TXLSDrwFillPictTileOpts.GetOffsetY: double;
begin
  Result := FTile.Ty / EMU_PER_PT;
end;

function TXLSDrwFillPictTileOpts.GetScaleX: double;
begin
  Result := FTile.Sx / 100000;
end;

function TXLSDrwFillPictTileOpts.GetScaleY: double;
begin
  Result := FTile.Sy / 100000;
end;

procedure TXLSDrwFillPictTileOpts.SetAlignment(const Value: TST_RectAlignment);
begin
  FTile.Algn := Value;
end;

procedure TXLSDrwFillPictTileOpts.SetMirrorType(const Value: TXLSDrwMirrorType);
begin
  FTile.Flip := TST_TileFlipMode(Value);
end;

procedure TXLSDrwFillPictTileOpts.SetOffsetX(const Value: double);
begin
  FTile.Tx := Round(Value * EMU_PER_PT);
end;

procedure TXLSDrwFillPictTileOpts.SetOffsetY(const Value: double);
begin
  FTile.Ty := Round(Value * EMU_PER_PT);
end;

procedure TXLSDrwFillPictTileOpts.SetScale(const Value: double);
begin
  SetScaleX(Value);
  SetScaleY(Value);
end;

procedure TXLSDrwFillPictTileOpts.SetScaleX(const Value: double);
begin
  FTile.Sx := Round(Value * 100000);
end;

procedure TXLSDrwFillPictTileOpts.SetScaleY(const Value: double);
begin
  FTile.Sy := Round(Value * 100000);
end;

{ TXLSDrwFillPictStretchOpts }

constructor TXLSDrwFillPictStretchOpts.Create(AStretch: TCT_StretchInfoProperties);
begin
  FStretch := AStretch;
end;

function TXLSDrwFillPictStretchOpts.GetOffsetBottom: double;
begin
  Result := FStretch.FillRect.B / 100000;
end;

function TXLSDrwFillPictStretchOpts.GetOffsetLeft: double;
begin
  Result := FStretch.FillRect.L / 100000;
end;

function TXLSDrwFillPictStretchOpts.GetOffsetRight: double;
begin
  Result := FStretch.FillRect.R / 100000;
end;

function TXLSDrwFillPictStretchOpts.GetOffsetTop: double;
begin
  Result := FStretch.FillRect.T / 100000;
end;

procedure TXLSDrwFillPictStretchOpts.SetOffsetBottom(const Value: double);
begin
  FStretch.FillRect.B := Round(Value * 100000);
end;

procedure TXLSDrwFillPictStretchOpts.SetOffsetLeft(const Value: double);
begin
  FStretch.FillRect.L := Round(Value * 100000);
end;

procedure TXLSDrwFillPictStretchOpts.SetOffsetRight(const Value: double);
begin
  FStretch.FillRect.R := Round(Value * 100000);
end;

procedure TXLSDrwFillPictStretchOpts.SetOffsetTop(const Value: double);
begin
  FStretch.FillRect.T := Round(Value * 100000);
end;

{ TXLSDrwFillChoice }

constructor TXLSDrwFillChoice.Create(AFillChoice: TEG_FillProperties);
begin
  FFillChoice := AFillChoice;
  if FFillChoice.SolidFill <> Nil then
    SetFillType(xdftSolid)
  else if FFillChoice.GradFill <> Nil then
    SetFillType(xdftGradient)
  else if FFillChoice.BlipFill <> Nil then
    SetFillType(xdftPicture)
  else if FFillChoice.NoFill <> Nil then
    SetFillType(xdftNone)
  else
    SetFillType(xdftAutomatic);
end;

destructor TXLSDrwFillChoice.Destroy;
begin
  if FFill <> Nil then
    FFill.Free;
  inherited;
end;

function TXLSDrwFillChoice.GetAsGradient: TXLSDrwFillGradient;
begin
  if FFill is TXLSDrwFillGradient then
    Result := TXLSDrwFillGradient(FFill)
  else
    raise XLSRWException.Create('Fill is not Gradient Fill');
end;

function TXLSDrwFillChoice.GetAsNone: TXLSDrwFillNone;
begin
  if FFill is TXLSDrwFillNone then
    Result := TXLSDrwFillNone(FFill)
  else
    raise XLSRWException.Create('Fill is not None Fill');
end;

function TXLSDrwFillChoice.GetAsPicture: TXLSDrwFillPicture;
begin
  if FFill is TXLSDrwFillPicture then
    Result := TXLSDrwFillPicture(FFill)
  else
    raise XLSRWException.Create('Fill is not Pictuer Fill');
end;

function TXLSDrwFillChoice.GetAsSolid: TXLSDrwFillSolid;
begin
  if FFill is TXLSDrwFillSolid then
    Result := TXLSDrwFillSolid(FFill)
  else
    raise XLSRWException.Create('Fill is not Solid Fill');
end;

function TXLSDrwFillChoice.GetFillType: TXLSDrwFillType;
begin
  Result := FFill.FillType;
end;

function TXLSDrwFillChoice.RGB_(ADefault: longword): longword;
begin
  Result := ADefault;

  case FFill.FillType of
    xdftNone    : Result := XLS_COLOR_NONE;
    xdftSolid   : Result := AsSolid.Color.AsRGB;
    xdftGradient: ;
    xdftPicture : ;
  end;
end;

procedure TXLSDrwFillChoice.SetFillType(const Value: TXLSDrwFillType);
begin
  if (FFill = Nil) or (GetFillType <> Value) then begin
    if FFill <> Nil then
      FFill.Free;
    case Value of
      xdftNone    : begin
        FFillChoice.Create_NoFill;
        FFill := TXLSDrwFillNone.Create(FFillChoice.NoFill);
      end;
      xdftSolid   : begin
        FFillChoice.Create_SolidFill;
        FFill := TXLSDrwFillSolid.Create(FFillChoice.SolidFill);
//        AsSolid.Color.AsRGB := $00FFFFFF;
      end;
      xdftGradient: begin
        FFillChoice.Create_GradFill;
        FFill := TXLSDrwFillGradient.Create(FFillChoice.GradFill);
      end;
      xdftPicture : begin
        FFillChoice.Create_BlipFill;
        FFill := TXLSDrwFillPicture.Create(FFillChoice.BlipFill);
      end;
      xdftAutomatic : FFill := TXLSDrwFillAutomatic.Create;
    end;
  end;
end;

{ TXLSDrwLineFillChoice }

constructor TXLSDrwLineFillChoice.Create(AFillChoice: TEG_FillProperties);
begin
  FFillChoice := AFillChoice;
  if FFillChoice.SolidFill <> Nil then
    SetFillType(xdftSolid)
  else if FFillChoice.GradFill <> Nil then
    SetFillType(xdftGradient)
  else
    SetFillType(xdftNone);
end;

function TXLSDrwLineFillChoice.GetAsGradient: TXLSDrwFillGradient;
begin
  if FFill is TXLSDrwFillGradient then
    Result := TXLSDrwFillGradient(FFill)
  else
    raise XLSRWException.Create('Fill is not Gradient Fill');
end;

function TXLSDrwLineFillChoice.GetAsNone: TXLSDrwFillNone;
begin
  if FFill is TXLSDrwFillNone then
    Result := TXLSDrwFillNone(FFill)
  else
    raise XLSRWException.Create('Fill is not No Fill');
end;

function TXLSDrwLineFillChoice.GetAsSolid: TXLSDrwFillSolid;
begin
  if FFill is TXLSDrwFillSolid then
    Result := TXLSDrwFillSolid(FFill)
  else
    raise XLSRWException.Create('Fill is not Solid Fill');
end;

function TXLSDrwLineFillChoice.GetFillType: TXLSDrwFillType;
begin
  Result := FFill.FillType;
end;

procedure TXLSDrwLineFillChoice.SetFillType(const Value: TXLSDrwFillType);
begin
  if GetFillType <> Value then begin
    FFill.Free;
    case Value of
      xdftNone    : begin
        FFillChoice.Create_NoFill;
        FFill := TXLSDrwFillNone.Create(FFillChoice.NoFill);
      end;
      xdftSolid   : begin
        FFillChoice.Create_SolidFill;
        FFill := TXLSDrwFillSolid.Create(FFillChoice.SolidFill);
      end;
      xdftGradient: begin
        FFillChoice.Create_GradFill;
        FFill := TXLSDrwFillGradient.Create(FFillChoice.GradFill);
      end;
      xdftPicture : raise XLSRWException.Create('Lines can not have picture fill');
    end;
  end;
end;

{ TXLSDrwLineStyle }

function TXLSDrwLineStyle.CanHaveArrows: boolean;
begin
  Result := FIsOpenLine;
end;

constructor TXLSDrwLineStyle.Create(ALineStyle: TCT_LineProperties; AIsOpenLine: boolean);
begin
  FLineStyle := ALineStyle;
  FIsOpenLine := AIsOpenLine;

  if ALineStyle.LineFillProperties.NoFill = Nil then
    FFill := TXLSDrwFillChoice.Create(ALineStyle.LineFillProperties);
end;

destructor TXLSDrwLineStyle.Destroy;
begin
  if FFill <> Nil then
    FFill.Free;

  inherited;
end;

function TXLSDrwLineStyle.GetCapType: TST_LineCap;
begin
  Result := FLineStyle.Cap;
end;

function TXLSDrwLineStyle.GetCompoundType: TST_CompoundLine;
begin
  Result := FLineStyle.Cmpd;
end;

function TXLSDrwLineStyle.GetDashType: TST_PresetLineDashVal;
begin
  if FLineStyle.LineDashProperties.PrstDash <> Nil then
    Result := FLineStyle.LineDashProperties.PrstDash.Val
  else
    Result := stpldvSolid;
end;

function TXLSDrwLineStyle.GetHasBeginArrow: boolean;
begin
  Result := FBeginArrow <> Nil;
end;

function TXLSDrwLineStyle.GetHasEndArrow: boolean;
begin
  Result := FEndArrow <> Nil;
end;

function TXLSDrwLineStyle.GetJoinType: TXLSDrwLineJoinType;
begin
  if FLineStyle.LineJoinProperties.Round <> Nil then
    Result :=  xdljtRound
  else if FLineStyle.LineJoinProperties.Bevel <> Nil then
    Result :=  xdljtBevel
  else if FLineStyle.LineJoinProperties.Miter <> Nil then
    Result :=  xdljtMiter
  else
    Result :=  xdljtNone;
end;

function TXLSDrwLineStyle.GetWidth: double;
begin
  Result := FLineStyle.W / EMU_PER_PT;
end;

function TXLSDrwLineStyle.RGB_(ADefault: longword): longword;
begin
  if FFill <> Nil then
    Result := FFill.RGB_(ADefault)
  else
    Result := ADefault;
end;

procedure TXLSDrwLineStyle.SetCapType(const Value: TST_LineCap);
begin
  FLineStyle.Cap := Value;
end;

procedure TXLSDrwLineStyle.SetCompoundType(const Value: TST_CompoundLine);
begin
  FLineStyle.Cmpd := Value;
end;

procedure TXLSDrwLineStyle.SetDashType(const Value: TST_PresetLineDashVal);
begin
  if FLineStyle.LineDashProperties.PrstDash = Nil then
    FLineStyle.LineDashProperties.Create_PrstDash;
    FLineStyle.LineDashProperties.PrstDash.Val := Value;
end;

procedure TXLSDrwLineStyle.SetHasBeginArrow(const Value: boolean);
begin
  if CanHaveArrows then begin
    if Value then begin
      FLineStyle.Create_HeadEnd;
      FBeginArrow := TXLSDrwLineEndStyle.Create(FLineStyle.HeadEnd);
    end
    else begin
      FreeAndNil(FBeginArrow);
      FLineStyle.Free_HeadEnd;
    end;
  end
  else
    raise XLSRWException.Create('This line can not have arrows');
end;

procedure TXLSDrwLineStyle.SetHasEndArrow(const Value: boolean);
begin
  if CanHaveArrows then begin
    if Value then begin
      FLineStyle.Create_TailEnd;
      FEndArrow := TXLSDrwLineEndStyle.Create(FLineStyle.TailEnd);
    end
    else begin
      FreeAndNil(FEndArrow);
      FLineStyle.Free_TailEnd;
    end;
  end
  else
    raise XLSRWException.Create('This line can not have arrows');
end;

procedure TXLSDrwLineStyle.SetJoinType(const Value: TXLSDrwLineJoinType);
begin
  case Value of
    xdljtNone : FLineStyle.LineJoinProperties.Clear;
    xdljtRound: FLineStyle.LineJoinProperties.Create_Round;
    xdljtBevel: FLineStyle.LineJoinProperties.Create_Bevel;
    xdljtMiter: FLineStyle.LineJoinProperties.Create_Miter;
  end;
end;

procedure TXLSDrwLineStyle.SetWidth(const Value: double);
begin
  FLineStyle.W := Round(Value * EMU_PER_PT);
end;

{ TXLSDrwLineEndStyle }

constructor TXLSDrwLineEndStyle.Create(AStyle: TCT_LineEndProperties);
begin
  FStyle := AStyle;
end;

function TXLSDrwLineEndStyle.GetEndType: TST_LineEndType;
begin
  Result := FStyle.Type_;
end;

function TXLSDrwLineEndStyle.GetLen: TST_LineEndLength;
begin
  Result := FStyle.Len;
end;

function TXLSDrwLineEndStyle.GetSize: TST_LineEndWidth;
begin
  Result := FStyle.W;
end;

procedure TXLSDrwLineEndStyle.SetEndType(const Value: TST_LineEndType);
begin
  FStyle.Type_ := Value;
end;

procedure TXLSDrwLineEndStyle.SetLen(const Value: TST_LineEndLength);
begin
  FStyle.Len := Value;
end;

procedure TXLSDrwLineEndStyle.SetSize(const Value: TST_LineEndWidth);
begin
  FStyle.W := Value;
end;

{ TXLSDrwFont }

function TXLSDrwTextFont.Assigned: boolean;
begin
  Result := FProps.CheckAssigned > 0;
end;

constructor TXLSDrwTextFont.Create(AProps: TCT_TextCharacterProperties);
begin
  FOwnsProps := AProps = Nil;
  if FOwnsProps then
    FProps := TCT_TextCharacterProperties.Create(Nil)
  else
    FProps := AProps;
end;

destructor TXLSDrwTextFont.Destroy;
begin
  if FProps.FillProperties.CheckAssigned <= 0 then
    FProps.FillProperties.Clear;
  if FOwnsProps then
    FProps.Free;
  inherited;
end;

function TXLSDrwTextFont.GetBold: boolean;
begin
  Result := FProps.B;
end;

function TXLSDrwTextFont.GetItalic: boolean;
begin
  Result := FProps.I;
end;

function TXLSDrwTextFont.GetName: AxUCString;
begin
  if FProps.Latin <> Nil then
    Result := FProps.Latin.Typeface
  else
    Result := '';
end;

function TXLSDrwTextFont.GetSize: double;
begin
  Result := FProps.Sz / 100;
end;

function TXLSDrwTextFont.GetUnderline: TST_TextUnderlineType;
begin
  Result := FProps.U;
end;

procedure TXLSDrwTextFont.SetBold(const Value: boolean);
begin
  FProps.B := Value;
end;

procedure TXLSDrwTextFont.SetColor(const Value: TColor);
begin
  FColor := Value;

  FProps.FillProperties.Create_SolidFill;
  FProps.FillProperties.SolidFill.ColorChoice.Create_SrgbClr;
  FProps.FillProperties.SolidFill.ColorChoice.SrgbClr.Val := TColorToRGB(FColor);
end;

procedure TXLSDrwTextFont.SetItalic(const Value: boolean);
begin
  FProps.I := Value;
end;

procedure TXLSDrwTextFont.SetName(const Value: AxUCString);
begin
  FProps.Create_Latin;
  FProps.Latin.Typeface := Value;
end;

procedure TXLSDrwTextFont.SetSize(const Value: double);
begin
  FProps.Sz := Round(Value * 100);
end;

procedure TXLSDrwTextFont.SetUnderline(const Value: TST_TextUnderlineType);
begin
  FProps.U := Value;
end;

{ TXLSDrwTextRun }

procedure TXLSDrwTextRun.AddFont;
begin
  if FFont = Nil then begin
    FRun.Run.Create_RPr;
    FFont := TXLSDrwTextFont.Create(FRun.Run.RPr);
  end;
end;

constructor TXLSDrwTextRun.Create(ARun: TEG_TextRun);
begin
  FRun := ARun;
  if FRun.Run <> Nil then begin
    FRun.Run.Create_RPr;
    FFont := TXLSDrwTextFont.Create(FRun.Run.RPr);
  end
  else if FRun.Br <> Nil then begin
    FRun.Br.Create_RPr;
    FFont := TXLSDrwTextFont.Create(FRun.Br.RPr);
  end;
end;

destructor TXLSDrwTextRun.Destroy;
begin
  FFont.Free;
  inherited;
end;

function TXLSDrwTextRun.GetText: AxUCString;
begin
  if FRun.Run <> Nil then
    Result := FRun.Run.T
  else if FRun.Br <> Nil then
    Result := #13
  else
    Result := '';
end;

procedure TXLSDrwTextRun.SetText(const Value: AxUCString);
begin
  if (Value <> '') and (Value[1] = #13) then begin
    if FRun.Run <> Nil then
      FRun.ConvertToBreak;
    FRun.Create_Br;
  end
  else begin
    if FRun.Br <> Nil then
      FRun.ConvertToRun;
    FRun.Create_R;
    FRun.Run.T := Value;
  end;
end;

{ TXLSDrwTextRuns }

function TXLSDrwTextRuns.Add: TXLSDrwTextRun;
begin
  Result := Add(FRuns.Add);
end;

function TXLSDrwTextRuns.Add(ARun: TEG_TextRun): TXLSDrwTextRun;
begin
  Result := TXLSDrwTextRun.Create(ARun);
  FItems.Add(Result);
end;

procedure TXLSDrwTextRuns.AddLineBreak(ARun: TXLSDrwTextRun);
var
  Run: TXLSDrwTextRun;
  RPr: TCT_TextCharacterProperties;
begin
  Rpr := Nil;
  if ARun <> Nil then begin
    if ARun.FRun.Run <> Nil then
      Rpr := ARun.FRun.Run.RPr
    else if ARun.FRun.Fld <> Nil then
      Rpr := ARun.FRun.Br.RPr;
  end;

  Run := Add;
  Run.FRun.Create_Br;
  if RPr <> Nil then
    Run.FRun.Br.Create_RPr.AssignFont(RPr);
end;

function TXLSDrwTextRuns.AddText(AText: AxUCString): TXLSDrwTextRun;
var
  LastFont: TXLSDrwTextFont;
begin
  LastFont := FParent.FParent.LastFont;

  Result := Add;
  Result.FRun.Create_R;
  Result.FRun.Run.T := AText;
  if LastFont <> Nil then
    Result.FRun.Run.Create_RPr.AssignFont(LastFont.FProps);
end;

procedure TXLSDrwTextRuns.Clear;
begin
  FItems.Clear;
  FRuns.Clear;
end;

function TXLSDrwTextRuns.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSDrwTextRuns.Create(AParent: TXLSDrwTextPara; ARuns: TEG_TextRunXpgList);
var
  i: integer;
begin
  FParent := AParent;
  FRuns := ARuns;

  FItems := TObjectList.Create;

  for i := 0 to FRuns.Count - 1 do
    Add(FRuns[i]);
end;

procedure TXLSDrwTextRuns.Delete(const AIndex: integer);
begin
  FRuns.Delete(AIndex);
  FItems.Delete(AIndex);
end;

destructor TXLSDrwTextRuns.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXLSDrwTextRuns.GetItems(Index: integer): TXLSDrwTextRun;
begin
  Result := TXLSDrwTextRun(FItems.Items[Index]);
end;

{ TXLSDrwTextPara }

procedure TXLSDrwTextPara.AddLineBreak;
begin

end;

function TXLSDrwTextPara.AddText(const AText: AxUCString): TXLSDrwTextRun;
begin
  Result := FRuns.AddText(AText);
end;

procedure TXLSDrwTextPara.Clear;
begin
  FPara.Clear;
  FRuns.Clear;
end;

constructor TXLSDrwTextPara.Create(AParent: TXLSDrwTextParas; APara: TCT_TextParagraph);
begin
  FParent := AParent;
  FPara := APara;
  FRuns := TXLSDrwTextRuns.Create(Self,FPara.TextRuns);
end;

destructor TXLSDrwTextPara.Destroy;
begin
  FRuns.Free;
  inherited;
end;

function TXLSDrwTextPara.GetAlign: TXLSDrwTextAlign;
begin
  if (FPara.PPr <> Nil) then
    Result := TXLSDrwTextAlign(FPara.PPr.Algn)
  else
    Result := xdtaLeft;
end;

function TXLSDrwTextPara.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FRuns.Count - 1 do
    Result := Result + FRuns[i].Text;
end;


procedure TXLSDrwTextPara.SetAlign(const Value: TXLSDrwTextAlign);
begin
  FPara.Create_PPr;
  FPara.PPr.Algn := TST_TextAlignType(Value);
end;

procedure TXLSDrwTextPara.SetPlainText(const Value: AxUCString);
begin
  Clear;
  FRuns.AddText(Value);
end;

{ TXLSDrwTextParas }

function TXLSDrwTextParas.Add(APara: TCT_TextParagraph): TXLSDrwTextPara;
begin
  Result := TXLSDrwTextPara.Create(Self,APara);
  FItems.Add(Result);
end;

function TXLSDrwTextParas.Add: TXLSDrwTextPara;
begin
  Result := Add(FParas.Add);
end;

procedure TXLSDrwTextParas.Clear;
begin
  FItems.Clear;
  FParas.Clear;
end;

function TXLSDrwTextParas.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSDrwTextParas.Create(AParent: TXLSDrwTextBody; AParas: TCT_TextParagraphXpgList);
var
  i: integer;
begin
  FParent := AParent;
  FParas := AParas;

  FItems := TObjectList.Create;

  for i := 0 to FParas.Count - 1 do
    Add(FParas[i]);
end;

destructor TXLSDrwTextParas.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXLSDrwTextParas.GetItems(Index: integer): TXLSDrwTextPara;
begin
  Result := TXLSDrwTextPara(FItems.Items[Index]);
end;

function TXLSDrwTextParas.LastFont: TXLSDrwTextFont;
var
  i,j: integer;
begin
  Result := Nil;
  for i := 0 to FItems.Count - 1 do begin
    for j := 0 to Items[i].Runs.Count - 1 do begin
      if Items[i].Runs[j].Font <> Nil then
        Result := Items[i].Runs[j].Font;
    end;
  end;
  if (Result = Nil) and FParent.DefaultFont.Assigned then
    Result := FParent.DefaultFont;
end;

{ TXLSDrwTextBody }

constructor TXLSDrwTextBody.Create(ABody: TCT_TextBody);
begin
  FBody := ABody;
  FDefaultFont := TXLSDrwTextFont.Create(Nil);
  FBody.Create_Paras;
  FParas := TXLSDrwTextParas.Create(Self,FBody.Paras);
end;

destructor TXLSDrwTextBody.Destroy;
begin
  FParas.Free;
  FDefaultFont.Free;
  inherited;
end;

function TXLSDrwTextBody.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FParas.Count - 1 do
    Result := Result + FParas[i].PlainText + #13;
  if Result <> '' then
    Result := Copy(Result,1,Length(Result) - 1);
end;

procedure TXLSDrwTextBody.SetPlainText(const Value: AxUCString);
var
  Para: TXLSDrwTextPara;
begin
  FParas.Clear;

  Para := FParas.Add(FBody.Paras.Add);
  Para.PlainText := Value;
end;

{ TDrwEditorTextBox }

constructor TXLSDrawingEditorTextBox.Create(AEx12Body: TCT_TextBody);
begin
  FEx12Body := AEx12Body;
  FBody := TXLSDrwTextBody.Create(FEx12Body);
end;

destructor TXLSDrawingEditorTextBox.Destroy;
begin
  FBody.Free;
  inherited;
end;

{ TXLSDrwShapeProperies }

function TXLSDrwShapeProperies.Assigned: boolean;
begin
  Result := FSpPr.CheckAssigned > 0;
end;

constructor TXLSDrwShapeProperies.Create(ASpPr: TCT_ShapeProperties);
begin
  FSpPr := ASpPr;
  FOwnsSpPr := False;

  if FSpPr.Ln <> Nil then
    // TODO Check if it is an open line
    FLine := TXLSDrwLineStyle.Create(FSpPr.Ln,False);
  FFill := TXLSDrwFillChoice.Create(FSpPr.FillProperties);
end;

destructor TXLSDrwShapeProperies.Destroy;
begin
  if FLine <> Nil then
    FLine.Free;
  FFill.Free;

  if FOwnsSpPr then
    FSpPr.Free;

  inherited;
end;

function TXLSDrwShapeProperies.FillRGB(ADefault: longword): longword;
begin
  if FFill <> Nil then
    Result := FFill.RGB_(ADefault)
  else
    Result := ADefault;
end;

function TXLSDrwShapeProperies.GetHasLine: boolean;
begin
  Result := FLine <> Nil;
end;

procedure TXLSDrwShapeProperies.SetHasLine(const Value: boolean);
begin
  if Value and (FLine = Nil) then begin
    FSpPr.Create_Ln;
    // TODO Check if it is an open line
    FLine := TXLSDrwLineStyle.Create(FSpPr.Ln,False);
  end
  else if not Value and (FLine <> Nil) then begin
    FSpPr.Free_Ln;
    FreeAndNil(FLine);
  end;
end;

{ TDrwEditorShape }

constructor TXLSDrawingEditorShape.Create(AAnchor: TCT_TwoCellAnchor);
begin
  FAnchor := AAnchor;
  FAnchor.Objects.Sp.Create_SpPr;
  FShapeProperies := TXLSDrwShapeProperies.Create(FAnchor.Objects.Sp.SpPr);
end;

destructor TXLSDrawingEditorShape.Destroy;
begin
  FShapeProperies.Free;
  inherited;
end;

{ TXLSDrawingEditorImage }

procedure TXLSDrawingEditorImage.CalcSize;
var
  i: integer;
  PPIX,PPIY: integer;
begin
  FPixelWidth := 0;
  FPixelHeight := 0;

  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  for i := FImage.Col1 to FImage.Col2 do
    Inc(FPixelWidth,FOwner.FColumns[i].PixelWidth);

  if FImage.Col1Offs > 0 then
    Dec(FPixelWidth,Round(FImage.Col1Offs * FOwner.FColumns[FImage.Col1].PixelWidth));
  if FImage.Col2Offs > 0 then
    Dec(FPixelWidth,FOwner.FColumns[FImage.Col2].PixelWidth - Round(FImage.Col2Offs * FOwner.FColumns[FImage.Col2].PixelWidth));

  for i := FImage.Row1 to FImage.Row2 do
    Inc(FPixelHeight,FOwner.FRows[i].PixelHeight);

  if FImage.Row1Offs > 0 then
    Dec(FPixelHeight,Round(FImage.Row1Offs * FOwner.FRows[FImage.Row1].PixelHeight));
  if FImage.Row2Offs > 0 then
    Dec(FPixelHeight,FOwner.FRows[FImage.Row2].PixelHeight - Round(FImage.Row2Offs * FOwner.FRows[FImage.Row2].PixelHeight));

  FCmWidth := FPixelWidth / (PPIX / 2.54);
  FCmHeight := FPixelHeight / (PPIY / 2.54);
end;

constructor TXLSDrawingEditorImage.Create(AOwner: TXLSDrawing; AImage: TXLSDrawingImage);
begin
  FOwner := AOwner;
  FImage := AImage;
  FBlip := FImage.FAnchor.Objects.Pic.BlipFill.Blip;

  FKeepAspect := True;

  CalcSize;
end;

function TXLSDrawingEditorImage.GetCmHeight: double;
begin
  Result := FCmHeight;
end;

function TXLSDrawingEditorImage.GetCmWidth: double;
begin
  Result := FCmWidth;
end;

function TXLSDrawingEditorImage.GetInchHeight: double;
begin
  Result := FCmHeight / 2.54;
end;

function TXLSDrawingEditorImage.GetInchWidth: double;
begin
  Result := FCmWidth / 2.54;
end;

function TXLSDrawingEditorImage.GetOriginalHeight: integer;
begin
  Result := FBlip.Image.Heigh;
end;

function TXLSDrawingEditorImage.GetOriginalWidth: integer;
begin
  Result := FBlip.Image.Width;
end;

procedure TXLSDrawingEditorImage.Scale(const APercent: double);
begin
  ScaleWidth(APercent);
  ScaleHeight(APercent);

  CalcSize;
end;

procedure TXLSDrawingEditorImage.ScaleHeight(const APercent: double);
var
  H: integer;
  PPIX,PPIY: integer;
  RO: double;
  Row: integer;
begin
  FPixelHeight := Round(FPixelHeight * APercent);

  RO := FImage.Row1Offs;

  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  Row := FImage.Row1;

  H := Round(FOwner.FRows[Row].PixelHeight * RO);
  if H > FPixelHeight then
    H := FPixelHeight;
  FImage.FAnchor.From.Row := Row;
  FImage.FAnchor.From.RowOff := PixelsToEMU(PPIY,H);

  Dec(FPixelHeight,H);
  while (FPixelHeight - FOwner.FRows[Row].PixelHeight) >= 0 do begin
    Dec(FPixelHeight,FOwner.FRows[Row].PixelHeight);
    Inc(Row);
  end;
  FImage.FAnchor.To_.Row := Row;
  FImage.FAnchor.To_.RowOff := PixelsToEMU(PPIY,FPixelHeight);
end;

procedure TXLSDrawingEditorImage.ScaleWidth(const APercent: double);
var
  W: integer;
  PPIX,PPIY: integer;
  CO: double;
  Col: integer;
begin
  FPixelWidth  := Round(FPixelWidth * APercent);

  CO := FImage.Col1Offs;

  FOwner.FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  Col := FImage.Col1;

  W := Round(FOwner.FColumns[Col].PixelWidth * CO);
  if W > FPixelWidth then
    W := FPixelWidth;
  FImage.FAnchor.From.Col := Col;
  FImage.FAnchor.From.ColOff := PixelsToEMU(PPIX,W);

  Dec(FPixelWidth,W);
  while (FPixelWidth - FOwner.FColumns[Col].PixelWidth) >= 0 do begin
    Dec(FPixelWidth,FOwner.FColumns[Col].PixelWidth);
    Inc(Col);
  end;
  FImage.FAnchor.To_.Col := Col;
  FImage.FAnchor.To_.ColOff := PixelsToEMU(PPIX,FPixelWidth);
end;

procedure TXLSDrawingEditorImage.SetCmHeight(const Value: double);
begin
  if KeepAspect then
    ScaleWidth(Value / FCmHeight);
  ScaleHeight(Value / FCmHeight);

  CalcSize;
end;

procedure TXLSDrawingEditorImage.SetCmWidth(const Value: double);
begin
  ScaleWidth(Value / FCmWidth);
  if KeepAspect then
    ScaleHeight(Value / FCmWidth);

  CalcSize;
end;

procedure TXLSDrawingEditorImage.SetInchHeight(const Value: double);
begin
  SetCmHeight(Value * 2.54);
end;

procedure TXLSDrawingEditorImage.SetInchWidth(const Value: double);
begin
  SetCmWidth(Value * 2.54);
end;

{ TXLSDrawingCharts }

function TXLSDrawingCharts._FileAdd(AChart: TXPGDocXLSXChart): TXLSDrawingChart;
begin
  Result := TXLSDrawingChart.Create(FOwner,AChart);
  FItems.Add(Result);
end;

procedure TXLSDrawingCharts.Clear;
begin
  FItems.Clear;
end;

function TXLSDrawingCharts.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXLSDrawingCharts.Create(AErrors: TXLSErrorManager; AOwner: TXLSDrawing);
begin
  FErrors := AErrors;
  FOwner := AOwner;
  FItems := TObjectList.Create;
end;

destructor TXLSDrawingCharts.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXLSDrawingCharts.Find(const ACol, ARow: integer): TXLSDrawingChart;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    if Items[i].Hit(ACol,ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXLSDrawingCharts.GetChartFont(ATxt: TCT_TextBody; AFont: TXc12Font);
begin
  if(ATxt <> Nil) and (ATxt.Paras <> Nil) and (ATxt.Paras.Count > 0) and
    (ATxt.Paras[0].PPr <> Nil) and (ATxt.Paras[0].PPr.DefRPr <> Nil) then
    GetChartFont(ATxt.Paras[0].PPr.DefRPr,AFont);
end;

procedure TXLSDrawingCharts.GetChartFont(ARPr: TCT_TextCharacterProperties; AFont: TXc12Font);
var
  Fill: TXLSDrwColor;
begin
  if ARPr <> Nil then begin
    if ARPr.Sz < 100000 then
      AFont.Size := ARPr.Sz / 100;
    if ARPr.B then
      AFont.Style := AFont.Style + [xfsBold];
    if ARPr.I then
      AFont.Style := AFont.Style + [xfsItalic];

    if (ARPr.Latin <> Nil) and (ARPr.Latin.Typeface <> '') then
      AFont.Name := ARPr.Latin.Typeface;

     if ARPr.FillProperties.SolidFill <> Nil then begin
       Fill := TXLSDrwColor.Create(ARPr.FillProperties.SolidFill.ColorChoice);
       try
         AFont.Color := RGBColorToXc12(RevRGB(Fill.AsRGB));
       finally
         Fill.Free;
       end;
     end;
  end;
end;

function TXLSDrawingCharts.GetItems(Index: integer): TXLSDrawingChart;
begin
  Result := TXLSDrawingChart(FItems[Index]);
end;

procedure TXLSDrawingCharts.Make3D(AChart: TCT_Chart);
begin
  AChart.Create_View3D;
  AChart.View3D.Create_RAngAx;
  AChart.View3D.RAngAx.Val := True;
end;

function TXLSDrawingCharts.MakeAreaChart(ASrcArea: TXLSRelCells; ACol, ARow: integer; A3D: boolean = False): TCT_ChartSpace;
var
  i     : integer;
  Chart : TXPGDocXLSXChart;
  Shared: TEG_AreaChartShared;
  Area  : TCT_AreaChart;
  Area3D: TCT_Area3DChart;
  AxId  : TCT_UnsignedIntXpgList;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  if A3D then begin
    Make3D(Chart.Root.ChartSpace.Chart);

    Area3D := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultArea3DChart;
    Shared := Area3D.Shared;
    AxId := Area3D.AxId;
  end
  else begin
    Area := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultAreaChart;
    Shared := Area.Shared;
    AxId := Area.AxId;
  end;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Shared.AddSerie(ASrcArea.CloneCol(i));

  MakeChartEpilogue(Chart,AxId);

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeBarChart(ASrcArea: TXLSRelCells; ACol,ARow: integer; AHasSeriesNames: boolean; A3D: boolean = False): TCT_ChartSpace;
var
  i     : integer;
  Chart : TXPGDocXLSXChart;
  Shared: TEG_BarChartShared;
  AxId  : TCT_UnsignedIntXpgList;
  Bar   : TCT_BarChart;
  Bar3D : TCT_Bar3DChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  if A3D then begin
    Make3D(Chart.Root.ChartSpace.Chart);

    Bar3D := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultBar3DChart;
    Shared := Bar3D.Shared;
    AxId := Bar3D.AxId;
  end
  else begin
    Bar := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultBarChart;
    Shared := Bar.Shared;

    AxId := Bar.AxId;
  end;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Shared.AddSerie(ASrcArea.CloneCol(i),AHasSeriesNames);

  MakeChartEpilogue(Chart,AxId);

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeBubbleChart(ASrcArea: TXLSRelCells; ACol, ARow: integer; A3D: boolean = False): TCT_ChartSpace;
var
  r     : integer;
  Chart : TXPGDocXLSXChart;
  Blub  : TCT_BubbleChart;
  AxId  : TCT_UnsignedIntXpgList;
  XChart: TXLSDrawingChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  Blub := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultBubbleChart;
  AxId := Blub.AxId;

  if ASrcArea.Rows >= 3 then begin
    r := ASrcArea.Row1 + 1;
    while r <= ASrcArea.Row2 do begin
      Blub.AddSerie(ASrcArea.CloneRow(ASrcArea.Row1),ASrcArea.CloneRow(r),ASrcArea.CloneRow(r + 1),A3D);
      Inc(r,2);
    end;
  end;

  Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultCrossValAx;

  AxId.Add.Val := Chart.Root.ChartSpace.Chart.PlotArea.ValAxis[0].Shared.AxId.Val;
  AxId.Add.Val := Chart.Root.ChartSpace.Chart.PlotArea.ValAxis[1].Shared.AxId.Val;

  XChart := _FileAdd(Chart);

  XChart.SetAnchor(TCT_TwoCellAnchor(FOwner.FDrw.Anchors.TwoCellAnchors.Last));

  Result := Chart.ChartSpace;
end;

procedure TXLSDrawingCharts.MakeChartEpilogue(AChart: TXPGDocXLSXChart; AAxId: TCT_UnsignedIntXpgList);
var
  CatAx : TCT_CatAx;
  ValAx : TCT_ValAx;
  XChart: TXLSDrawingChart;
begin
  CatAx := AChart.Root.ChartSpace.Chart.PlotArea.CreateDefaultCatAx;
  ValAx := AChart.Root.ChartSpace.Chart.PlotArea.CreateDefaultValAx;

  ValAx.Shared.CrossAx.Val := CatAx.Shared.AxId.Val;
  CatAx.Shared.CrossAx.Val := ValAx.Shared.AxId.Val;

  AAxId.Add.Val := CatAx.Shared.AxId.Val;
  AAxId.Add.Val := ValAx.Shared.AxId.Val;

  XChart := _FileAdd(AChart);

  XChart.SetAnchor(TCT_TwoCellAnchor(FOwner.FDrw.Anchors.TwoCellAnchors.Last));
end;

function TXLSDrawingCharts.MakeChartPrologue(ACol, ARow: integer): TXPGDocXLSXChart;
var
  S     : AxUCString;
  Anchor: TCT_TwoCellAnchor;
begin
  S := 'Chart ' + IntToStr(FItems.Count + 1);

  Anchor := FOwner.FDrw.Anchors.TwoCellAnchors.Add(ACol,ARow,ACol + 7,ARow + 14);
  Result := Anchor.MakeChart;
  Anchor.Objects.GraphicFrame.CreateDefault(FItems.Count + 1,S);

  Result.Root.CreateDefault;

  Result.Root.ChartSpace.CreateDefault;

  Result.Root.ChartSpace.Chart.CreateDefault;
end;

function TXLSDrawingCharts.MakeDoughnutChart(ASrcArea: TXLSRelCells; ACol, ARow: integer): TCT_ChartSpace;
var
  i     : integer;
  Chart : TXPGDocXLSXChart;
  Nut   : TCT_DoughnutChart;
  XChart: TXLSDrawingChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  Nut := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultDoughnutChart;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Nut.AddSerie(ASrcArea.CloneCol(i));
//  for i := ASrcArea.Col1 to ASrcArea.Col2 do
//    Nut.AddSerie('Sheet1!' + AreaToRefStrAbs(i,ASrcRow1,i,ASrcRow2));

  XChart := _FileAdd(Chart);

  XChart.SetAnchor(TCT_TwoCellAnchor(FOwner.FDrw.Anchors.TwoCellAnchors.Last));

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeLineChart(ASrcArea: TXLSRelCells; ACol, ARow: integer; A3D: boolean = False): TCT_ChartSpace;
var
  i     : integer;
  Chart : TXPGDocXLSXChart;
  Line  : TCT_LineChart;
  Line3D: TCT_Line3DChart;
  Shared: TEG_LineChartShared;
  AxId  : TCT_UnsignedIntXpgList;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  if A3D then begin
    Make3D(Chart.Root.ChartSpace.Chart);

    Line3D := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultLine3DChart;
    Shared := Line3D.Shared;
    AxId := Line3D.AxId;
  end
  else begin
    Line := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultLineChart;
    Shared := Line.Shared;
    AxId := Line.AxId;
  end;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Shared.AddSerie(ASrcArea.CloneCol(i));

  MakeChartEpilogue(Chart,AxId);

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeOfPieChart(ASrcArea: TXLSRelCells; ACol, ARow: integer): TCT_ChartSpace;
var
  i     : integer;
  Chart : TXPGDocXLSXChart;
  Shared: TEG_PieChartShared;
  OfPie : TCT_OfPieChart;
  XChart: TXLSDrawingChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  OfPie := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultOfPieChart;

  Shared := OfPie.Shared;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Shared.AddSerie(ASrcArea.CloneCol(i));

  XChart := _FileAdd(Chart);

  XChart.SetAnchor(TCT_TwoCellAnchor(FOwner.FDrw.Anchors.TwoCellAnchors.Last));

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakePieChart(ASrcArea: TXLSRelCells; ACol, ARow: integer; A3D: boolean = False): TCT_ChartSpace;
var
  i    : integer;
  Chart: TXPGDocXLSXChart;
  Shared: TEG_PieChartShared;
  Pie   : TCT_PieChart;
  Pie3D : TCT_Pie3DChart;
  XChart: TXLSDrawingChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  if A3D then begin
    Make3D(Chart.Root.ChartSpace.Chart);

    Pie3D := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultPie3DChart;
    Shared := Pie3D.Shared;
  end
  else begin
    Pie := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultPieChart;

    Shared := Pie.Shared;
  end;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Shared.AddSerie(ASrcArea.CloneCol(i));

  XChart := _FileAdd(Chart);

  XChart.SetAnchor(TCT_TwoCellAnchor(FOwner.FDrw.Anchors.TwoCellAnchors.Last));

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeRadarChart(ASrcArea: TXLSRelCells; ACol, ARow: integer): TCT_ChartSpace;
var
  i    : integer;
  Chart: TXPGDocXLSXChart;
  Radar: TCT_RadarChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  Radar := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultRadarChart;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Radar.AddSerie(ASrcArea.CloneCol(i));

  MakeChartEpilogue(Chart,Radar.AxId);

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeScatterChart(ASrcArea: TXLSRelCells; ACol, ARow: integer): TCT_ChartSpace;
var
  i      : integer;
  Chart  : TXPGDocXLSXChart;
  Scatter: TCT_ScatterChart;
  ValAx1 : TCT_ValAx;
  ValAx2 : TCT_ValAx;
  XSrc   : TXLSRelCells;
  XChart : TXLSDrawingChart;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  Scatter := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultScatterChart;

  XSrc := ASrcArea.CloneCol(0);

  if ASrcArea.Col1 = ASrcArea.Col2 then
     Chart.Root.ChartSpace.Chart.PlotArea.ScatterChart.AddSerie(XSrc,XSrc)
  else begin
    for i := ASrcArea.Col1 + 1 to ASrcArea.Col2 do
      Chart.Root.ChartSpace.Chart.PlotArea.ScatterChart.AddSerie(XSrc,ASrcArea.CloneCol(i));
  end;

  ValAx1 := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultValAx;
  ValAx1.Shared.AxPos.Val := stapB;
  ValAx1.Shared.Remove_MajorGridlines;
  ValAx2 := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultValAx;

  ValAx1.Shared.CrossAx.Val := ValAx2.Shared.AxId.Val;
  ValAx2.Shared.CrossAx.Val := ValAx1.Shared.AxId.Val;

  Scatter.AxId.Add.Val := ValAx1.Shared.AxId.Val;
  Scatter.AxId.Add.Val := ValAx2.Shared.AxId.Val;

  XChart := _FileAdd(Chart);

  XChart.SetAnchor(TCT_TwoCellAnchor(FOwner.FDrw.Anchors.TwoCellAnchors.Last));

  Result := Chart.ChartSpace;
end;

function TXLSDrawingCharts.MakeSurfaceChart(ASrcArea: TXLSRelCells; ACol, ARow: integer): TCT_ChartSpace;
var
  i      : integer;
  Chart  : TXPGDocXLSXChart;
  Surface: TCT_Surface3DChart;
  SerAx  : TCT_SerAx;
begin
  Chart := MakeChartPrologue(ACol,ARow);

  Chart.Root.ChartSpace.Chart.Create_View3D;
  Chart.Root.ChartSpace.Chart.View3D.Create_Perspective.Val := 30;

  Surface := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultSurface3DChart;

  for i := ASrcArea.Col1 to ASrcArea.Col2 do
    Surface.Shared.AddSerie(ASrcArea.CloneCol(i));

  MakeChartEpilogue(Chart,Surface.AxId);

  SerAx := Chart.Root.ChartSpace.Chart.PlotArea.CreateDefaultSerAx;
  SerAx.Shared.CrossAx.Val := Chart.Root.ChartSpace.Chart.PlotArea.ValAxis[0].Shared.AxId.Val;

  Surface.AxId.Add.Val := SerAx.Shared.AxId.Val;

  Result := Chart.ChartSpace;
end;

procedure TXLSDrawingCharts.SetChartFont(ARPr: TCT_TextCharacterProperties; AFont: TXc12Font);
begin
  ARPr.Sz := Round(AFont.Size * 100);
  ARPr.B := xfsBold in AFont.Style;
  ARPr.I := xfsItalic in AFont.Style;

  ARPr.Create_Latin;
  ARPr.Latin.Typeface := AFont.Name;

  if Xc12ColorToRGB(AFont.Color) <> $000000 then begin
    ARPr.FillProperties.Create_SolidFill;
    ARPr.FillProperties.SolidFill.ColorChoice.Create_SrgbClr;
    ARPr.FillProperties.SolidFill.ColorChoice.SrgbClr.Val := Xc12ColorToRGB(AFont.Color);
  end;
end;

procedure TXLSDrawingCharts.SetChartFont(ATxt: TCT_TextBody; AFont: TXc12Font);
var
  P: TCT_TextParagraph;
begin
  ATxt.Create_Paras;
  P := ATxt.Paras.Add;
  P.Create_PPr;
  P.PPr.Create_DefRPr;
  SetChartFont(ATxt.Paras[0].PPr.DefRPr,AFont);
end;

{ TXLSDrawingChart }

constructor TXLSDrawingChart.Create(AOwner: TXLSDrawing; AChart: TXPGDocXLSXChart);
begin
  inherited Create(AOwner);

  FChartSpace := AChart.Root.ChartSpace;
end;

destructor TXLSDrawingChart.Destroy;
begin
  inherited;
end;

function TXLSDrawingChart.GetChartSpace: TCT_ChartSpace;
begin
  Result := FChartSpace;
end;

{ TXLSDrawingEditorChart }

constructor TXLSDrawingEditorChart.Create(AChart: TXLSDrawingChart);
begin
  FChart := AChart.FChartSpace.Chart;
end;

function TXLSDrawingEditorChart.GetCategories: AxUCString;
begin
  Result := TCT_BarSer(FChart.PlotArea.BarChart.Shared.Series[0]).Cat.NumRef.F;
end;

function TXLSDrawingEditorChart.GetValues: AxUCString;
begin
  Result := TCT_BarSer(FChart.PlotArea.BarChart.Shared.Series[0]).Val.NumRef.F;
end;

procedure TXLSDrawingEditorChart.SetCategories(const Value: AxUCString);
begin
  TCT_BarSer(FChart.PlotArea.BarChart.Shared.Series[0]).Cat.NumRef.F := Value;
end;

procedure TXLSDrawingEditorChart.SetValues(const Value: AxUCString);
var
  Ser: TCT_BarSer;
begin
  Ser := TCT_BarSer(FChart.PlotArea.BarChart.Shared.Series[0]);
  Ser.Val.NumRef.F := Value;
end;

{ TXLSDrwFillAutomatic }

function TXLSDrwFillAutomatic.GetFillType: TXLSDrwFillType;
begin
  Result := xdftAutomatic;
end;

{ TXLSDrwText }

constructor TXLSDrwText.Create(ATx: TCT_Tx);
begin
  FTx := ATx;
end;

function TXLSDrwText.GetPlainText: AxUCString;
begin
  if FBody <> Nil then
    Result := FBody.PlainText
  else
    Result := '';
end;

function TXLSDrwText.GetReF: TXLSRelCells;
begin
  if FRef <> Nil then
    Result := FRef.RCells
  else
    Result := Nil;
end;

procedure TXLSDrwText.SetPlainText(const Value: AxUCString);
begin
  SetType(False);

  FBody.PlainText := Value;
end;

procedure TXLSDrwText.SetRef(const Value: TXLSRelCells);
begin
  SetType(True);
  if FRef.RCells <> Nil then
    FRef.RCells.Free;
  FRef.RCells := Value;
end;

procedure TXLSDrwText.SetType(ARef: boolean);
begin
  if ARef then begin
    if FBody <> Nil then begin
      FBody.Free;
      FBody := Nil;
    end;
    FRef := FTx.Create_StrRef;
  end
  else begin
    if FRef <> Nil then begin
      FTx.Delete_StrRef;
      FRef := Nil;
    end;
    FTx.Create_Rich;
    FBody := TXLSDrwTextBody.Create(FTx.Rich);
  end;
end;

end.
