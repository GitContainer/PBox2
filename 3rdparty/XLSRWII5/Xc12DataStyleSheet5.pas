unit Xc12DataStyleSheet5;

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

{$ifdef BABOON}
  {$undef XLS_XSS}
{$endif}

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
{$ifdef MSWINDOWS}
     Windows,
{$endif}
{$ifdef DELPHI_XE3_OR_LATER}
     UITypes,
{$endif}
     xpgPUtils, xpgParseDrawingCommon,
     XLSUtils5, Xc12Utils5, Xc12Common5;

// TODO Don't works correct.
{.$define STYLE_USE_HASHTABLES}

// Temp
type TDiagLines = (dlNone,dlUp,dlDown,dlBoth);

const XC12_STYLES_COMPACT_CNT = 25;

type TXc12CellBorderStyle = (cbsNone,cbsThin,cbsMedium,cbsDashed,cbsDotted,
                             cbsThick,cbsDouble,cbsHair,cbsMediumDashed,
                             cbsDashDot,cbsMediumDashDot,cbsDashDotDot,
                             cbsMediumDashDotDot,cbsSlantedDashDot);

type TXc12CellBorderOption = (ecboDiagonalDown,ecboDiagonalUp,ecboOutline);
     TXc12CellBorderOptions = set of TXc12CellBorderOption;


type TXc12FillPattern = (efpNone,efpSolid,efpMediumGray,efpDarkGray,efpLightGray,
                         efpDarkHorizontal,efpDarkVertical,efpDarkDown,efpDarkUp,
                         efpDarkGrid,efpDarkTrellis,efpLightHorizontal,
                         efpLightVertical,efpLightDown,efpLightUp,efpLightGrid,
                         efpLightTrellis,efpGray125,efpGray0625);

type TXc12GradientFillType = (gftLinear,gftPath);
type TXc12GradientFillCorners = (gfcLeft,gfcRight,gfcTop,gfcBottom);

type TXc12HorizAlignment = (chaGeneral,     //* No alignment.
                            chaLeft,        //* Left alignment
                            chaCenter,      //* Center alignment.
                            chaRight,       //* Right alignment.
                            chaFill,        //* Fill's the entire cell with the text or character. Like: 'XXXXXXXXXX'.
                            chaJustify,     //* Justify's the word space to fit the text in the cell.
                            chaCenterContinuous,//* Don't know what this is. CenterContinuous in Excel 2007?

                            chaDistributed
                            );

//* Vertical alignment of text in cells.
type TXc12VertAlignment = (cvaTop,        //* Top alignment.
                           cvaCenter,     //* Center alignment.
                           cvaBottom,     //* Bottom alignment.
                           cvaJustify,    //* Justify's the line space to fit the text in the cell.
                           cvaDistributed //* Text vertical distributed in cell.
                           );

type TXc12AlignmentOption = (foWrapText,    //* Wrap text in cells.
                             foShrinkToFit, //* Shrink text to fit horizontal cell space. The result is
                                     //* that the font size is changed so the text fit's the cell's horizontal size.
                             foJustifyLastLine
                             );
     TXc12AlignmentOptions = set of TXc12AlignmentOption;

type TXc12ReadOrder = (xroContextDependent,xroLeftToRight,xroRightToLeft);

type TXc12FontStyle = (xfsBold,xfsItalic,xfsStrikeOut,xfsOutline,xfsShadow,xfsCondense,xfsExtend);
type TXc12FontStyles = set of TXc12FontStyle;

type TXc12SubSuperscript = (xssNone,xssSuperscript,xssSubscript);

type TXc12Underline = (xulNone,xulSingle,xulDouble,xulSingleAccount,xulDoubleAccount);

type TXc12FontScheme = (efsNone,efsMajor,efsMinor);

type TXc12CellProtection = (cpLocked, //* Cell is locked. This does not mean that the cell value not can be
                                      //* changed. To prevent the cell from being changed, the worksheet has to be locked.
                            cpHidden  //* Cell value is hidden.
                            );
     TXc12CellProtections = set of TXc12CellProtection;

type TXc12ApplyFormat = (eafNumberFormat,eafFont,eafFill,eafBorder,eafAlignment,eafProtection);
     TXc12ApplyFormats = set of TXc12ApplyFormat;

type TXc12TableStyleType =  (sttstWholeTable,sttstHeaderRow,sttstTotalRow,sttstFirstColumn,sttstLastColumn,sttstFirstRowStripe,sttstSecondRowStripe,sttstFirstColumnStripe,sttstSecondColumnStripe,sttstFirstHeaderCell,sttstLastHeaderCell,sttstFirstTotalCell,sttstLastTotalCell,sttstFirstSubtotalColumn,sttstSecondSubtotalColumn,sttstThirdSubtotalColumn,sttstFirstSubtotalRow,sttstSecondSubtotalRow,sttstThirdSubtotalRow,sttstBlankRow,sttstFirstColumnSubheading,sttstSecondColumnSubheading,sttstThirdColumnSubheading,sttstFirstRowSubheading,sttstSecondRowSubheading,sttstThirdRowSubheading,sttstPageFieldLabels,sttstPageFieldValues);

const XLS_NUMFMT_STD_DATE      = 14;
const XLS_NUMFMT_STD_TIME      = 20;
const XLS_NUMFMT_STD_DATETIME  = 22;
const ExcelStandardNumFormats: array[0..49] of AxUCString = (
{ 00} '',
{ 01} '0',
{ 02} '0.00',
{ 03} '#,##0',
{ 04} '#,##0.00',
{ 05} '',
{ 06} '',
{ 07} '',
{ 08} '',
{ 09} '0%',
{ 10} '0.00%',
{ 11} '0.00E+00',
{ 12} '# ?/?',
{ 13} '# ??/??',
{ 14} 'm/d/yy',          // Localized date format in Excel.
{ 15} 'd-mmm-y',
{ 16} 'd-mmm',
{ 17} 'mmmm-yy',
{ 18} 'h:mm AM/PM',
{ 19} 'h:mm:ss AM/PM',
{ 20} 'h:mm',            // Localized time format in Excel.
{ 21} 'h:mm:SS',
{ 22} 'm/d/yy h:mm',     // Localized date-time format in Excel.
{ 23} '', // Format 23 - 26 are unknown.
{ 24} '',
{ 25} '',
{ 26} '',
{ 27} '', // Format 27 - 36 are localized.
{ 28} '',
{ 29} '',
{ 30} '',
{ 31} '',
{ 32} '',
{ 33} '',
{ 34} '',
{ 35} '',
{ 36} '',
{ 37} '_(#,##0_);(#,##0)',
{ 38} '_(#,##0_);[Red](#,##0)',
{ 39} '_(#,##0.00_);(#,##0.00)',
{ 40} '_(#,##0.00_);[Red](#,##0.00)',
{ 41} '',
{ 42} '',
{ 43} '',
{ 44} '',
{ 45} 'mm:ss',
{ 46} '[h]:mm:ss',
{ 47} 'mm:ss.0',
{ 48} '# #0.0E+0',
{ 49} '@');

const Excel40StandardNumFormats: array[0..33] of AxUCString = (
{01} 'General',
{02} '0',
{03} '0.00',
{04} '#,##0',
{05} '#,##0.00',
{06} '#,##0\ _$;\-#,##0\ _$',
{07} '#,##0\ _$;[Red]\-#,##0\ _$',
{08} '#,##0.00\ _$;\-#,##0.00\ _$',
{09} '#,##0.00\ _$;[Red]\-#,##0.00\ _$',
{10} '#,##0\ "kr";\-#,##0\ "kr"',
{11} '#,##0\ "kr";[Red]\-#,##0\ "kr"',
{12} '#,##0.00\ "kr";\-#,##0.00\ "kr"',
{13} '#,##0.00\ "kr";[Red]\-#,##0.00\ "kr"',
{14} '0%',
{15} '0.00%',
{16} '0.00E+00',
{17} '#" "?/?',
{18} '#" "??/??',
{19} 'yyyy/mm/dd',
{20} 'dd/mmm/yy',
{21} 'dd/mmm',
{22} 'mmm/yy',
{23} 'h:mm\ AM/PM',
{24} 'h:mm:ss\ AM/PM',
{25} 'hh:mm',
{26} 'hh:mm:ss',
{27} 'yyyy/mm/dd\ hh:mm',
{28} '##0.0E+0',
{29} 'mm:ss',
{30} '@',
{31} '_-* #,##0\ "$"_-;\-* #,##0\ "$"_-;_-* "-"\ "$"_-;_-@_-',
{32} '_-* #,##0\ _$_-;\-* #,##0\ _$_-;_-* "-"\ _$_-;_-@_-',
{33} '_-* #,##0.00\ "kr"_-;\-* #,##0.00\ "kr"_-;_-* "-"??\ "kr"_-;_-@_-',
{34} '_-* #,##0.00\ _$_-;\-* #,##0.00\ _$_-;_-* "-"??\ _$_-;_-@_-');

type TXLSPPIEvent = procedure(out APPIX,APPYY: integer) of object;

type TXc12DataStyleSheet = class;

     TXc12BorderPr = class(TObject)
private
     function  GetColorRGB: TXc12RGBColor;
     procedure SetColorRGB(const Value: TXc12RGBColor);
protected
     FColor: TXc12Color;
     FStyle: TXc12CellBorderStyle;
public
     procedure Clear;

     procedure Assign(ABorderPr: TXc12BorderPr);
     function  Equal(ABorderPr: TXc12BorderPr): boolean;
     function  CalcHash: longword;

     property Color: TXc12Color read FColor write FColor;
     property ColorRGB: TXc12RGBColor read GetColorRGB write SetColorRGB;
     property Style: TXc12CellBorderStyle read FStyle write FStyle;
     end;


     TXc12ColorObj = class(TXLSStyleObject)
private
     function GetColor: TXc12Color;
protected
     FColor: TXc12Color;

     procedure CalcHash; override;
public
     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12ColorObj);

     property Color: TXc12Color read GetColor;
     end;

     TXc12ColorObjs = class(TXLSStyleObjectList)
protected
     function GetItems(Index: integer): TXc12ColorObj;
public
     constructor Create;

     procedure SetIsDefault; override;
     procedure Clear; reintroduce;

     procedure FixupColors;

     function  Add: TXc12ColorObj; overload;
     function  Add(AColor: TXc12Color): TXc12ColorObj; overload;

     property Items[Index: integer]: TXc12ColorObj read GetItems; default;
     end;

     TXc12IndexColorObj = class(TXLSStyleObject)
private
     function GetRGB: longword;
     procedure SetRGB(const Value: longword);
protected
     FRGB: longword;

     procedure CalcHash; override;
public
     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12IndexColorObj);

     property RGB: longword read GetRGB write SetRGB;
     end;

     TXc12IndexColorObjs = class(TXLSStyleObjectList)
protected
     function GetItems(Index: integer): TXc12IndexColorObj;
public
     constructor Create;

     procedure SetIsDefault; override;
     procedure Clear; reintroduce;

     function  Add: TXc12IndexColorObj; overload;
     function  Add(ARGB: integer): TXc12IndexColorObj; overload;

     property Items[Index: integer]: TXc12IndexColorObj read GetItems; default;
     end;

     TXc12GradientStop = class(TObject)
private
     function  GetColor: TXc12Color;
     function  GetPosition: double;
     procedure SetColor(const Value: TXc12Color);
     procedure SetPosition(const Value: double);
protected
     FColor: TXc12Color;
     FPosition: double;

     function CalcHash: longword;
public
     property Color: TXc12Color read GetColor write SetColor;
     property Position: double read GetPosition write SetPosition;
     end;

     TXc12GradientStopList = class(TObjectList)
private
     function GetItems(Index: integer): TXc12GradientStop;
protected
     function CalcHash: longword;
public
     function Equal(AList: TXc12GradientStopList): boolean;

     function Add: TXc12GradientStop; overload;
     function Add(AColor: TXc12Color; APosition: double): TXc12GradientStop; overload;

     property Items[Index: integer]: TXc12GradientStop read GetItems; default;
     end;

     TXc12GradientFill = class(TXLSStyleObject)
private
     FStops: TXc12GradientStopList;
     FGradientType: TXc12GradientFillType;
     FDegree: double;
     FLeft: double;
     FRight: double;
     FTop: double;
     FBottom: double;
     FCorners: array[gfcLeft..gfcBottom] of boolean;

     procedure SetBottom(const Value: double);
     procedure SetLeft(const Value: double);
     procedure SetRight(const Value: double);
     procedure SetTop(const Value: double);
     function  GetCorners(Index: TXc12GradientFillCorners): boolean;
     procedure SetCorners(Index: TXc12GradientFillCorners; const Value: boolean);
     function  GetBottom: double;
     function  GetDegree: double;
     function  GetGradientType: TXc12GradientFillType;
     function  GetLeft: double;
     function  GetRight: double;
     function  GetStops: TXc12GradientStopList;
     function  GetTop: double;
     procedure SetDegree(const Value: double);
     procedure SetGradientType(const Value: TXc12GradientFillType);
protected
     procedure CalcHash; override;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;
     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12GradientFill);

     property Stops: TXc12GradientStopList read GetStops;
     property GradientType: TXc12GradientFillType read GetGradientType write SetGradientType;
     property Degree: double read GetDegree write SetDegree;
     property Left: double read GetLeft write SetLeft;
     property Right: double read GetRight write SetRight;
     property Top: double read GetTop write SetTop;
     property Bottom: double read GetBottom write SetBottom;
     property Corners[Index: TXc12GradientFillCorners]: boolean read GetCorners write SetCorners;
     end;

     TXc12Fill = class(TXLSStyleObject)
private
     function  GetIsGradientFill: boolean;
     procedure SetIsGradientFill(const Value: boolean);
     function  GetBgColor: TXc12Color;
     function  GetFgColor: TXc12Color;
     function  GetGradientFill: TXc12GradientFill;
     function  GetPatternType: TXc12FillPattern;
     procedure SetBgColor(const Value: TXc12Color);
     procedure SetFgColor(const Value: TXc12Color);
     procedure SetPatternType(const Value: TXc12FillPattern);
     function  GetPBgColor: PXc12Color;
     function  GetPFgColor: PXc12Color;
protected
     FGradientFill: TXc12GradientFill;
     FPatternType: TXc12FillPattern;
     FFgColor: TXc12Color;
     FBgColor: TXc12Color;

     procedure CalcHash; override;
public
     destructor Destroy; override;

     function AsString(const AIndent: integer = 0): AxUCString; override;

     procedure Clear;

     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12Fill);

     property IsGradientFill: boolean read GetIsGradientFill write SetIsGradientFill;
     property GradientFill: TXc12GradientFill read GetGradientFill;
     property PatternType: TXc12FillPattern read GetPatternType write SetPatternType;
     property FgColor: TXc12Color read GetFgColor write SetFgColor;
     property BgColor: TXc12Color read GetBgColor write SetBgColor;
     property PFgColor: PXc12Color read GetPFgColor;
     property PBgColor: PXc12Color read GetPBgColor;
     end;

     TXc12Fills = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12Fill;
protected
public
     procedure SetIsDefault; override;
     procedure Clear; reintroduce;

     function DefaultFill: TXc12Fill;

     function  Find(AFill: TXc12Fill): TXc12Fill;
     function  Add: TXc12Fill; overload;
     procedure Add(const AFill: TXc12Fill); overload;

     procedure FixupColors;

     property Items[Index: integer]: TXc12Fill read GetItems; default;
     end;

     TXc12Border = class(TXLSStyleObject)
private
     function  GetOptions: TXc12CellBorderOptions;
     procedure SetOptions(const Value: TXc12CellBorderOptions);
protected
     FOptions      : TXc12CellBorderOptions;
     FLeft         : TXc12BorderPr;
     FTop          : TXc12BorderPr;
     FRight        : TXc12BorderPr;
     FBottom       : TXc12BorderPr;
     FDiagonal     : TXc12BorderPr;
     FHorizontal   : TXc12BorderPr;
     FVertical     : TXc12BorderPr;

     procedure CalcHash; override;
     function  AsStringBorder(const AIndent: integer; const ABorder: TXc12BorderPr): AxUCString;
public
     constructor Create(AOwner: TXLSStyleObjectList);
     destructor Destroy; override;

     function AsString(const AIndent: integer = 0): AxUCString; override;

     procedure Clear;

     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12Border);

     property Options   : TXc12CellBorderOptions read GetOptions write SetOptions;
     property Left      : TXc12BorderPr read FLeft;
     property Top       : TXc12BorderPr read FTop;
     property Right     : TXc12BorderPr read FRight;
     property Bottom    : TXc12BorderPr read FBottom;
     property Diagonal  : TXc12BorderPr read FDiagonal;
     property Horizontal: TXc12BorderPr read FHorizontal;
     property Vertical  : TXc12BorderPr read FVertical;
     end;

     TXc12Borders = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12Border;
public
     constructor Create;

     procedure SetIsDefault; override;
     procedure Clear; reintroduce;

     function  DefaultBorder: TXc12Border;

     function  Find(ABorder: TXc12Border): TXc12Border;
     function  Add: TXc12Border; overload;

     property Items[Index: integer]: TXc12Border read GetItems; default;
     end;

     TXc12CellAlignment = class(TObject)
private
     function  GetHorizAlignment: TXc12HorizAlignment;
     function  GetIndent: integer;
     function  GetOptions: TXc12AlignmentOptions;
     function  GetReadingOrder: TXc12ReadOrder;
     function  GetRelativeIndent: integer;
     function  GetRotation: integer;
     function  GetVertAlignment: TXc12VertAlignment;
     procedure SetHorizAlignment(const Value: TXc12HorizAlignment);
     procedure SetIndent(const Value: integer);
     procedure SetOptions(const Value: TXc12AlignmentOptions);
     procedure SetReadingOrder(const Value: TXc12ReadOrder);
     procedure SetRelativeIndent(const Value: integer);
     procedure SetRotation(const Value: integer);
     procedure SetVertAlignment(const Value: TXc12VertAlignment);
protected
     FHorizAlignment: TXc12HorizAlignment;
     FVertAlignment : TXc12VertAlignment;
     FRotation      : integer;
     FOptions       : TXc12AlignmentOptions;
     FIndent        : integer;
     FRelativeIndent: integer;
     FReadingOrder  : TXc12ReadOrder;

     FHash          : longword;
     FHashValid     : boolean;

     function  HashKey: longword;
     procedure CalcHash;
public
     constructor Create;

     function  AsString: AxUCString;

     function  Equal(AObject: TXc12CellAlignment): boolean;
     procedure Assign(AObject: TXc12CellAlignment);

     function  IsWrapText: boolean;

     property HorizAlignment: TXc12HorizAlignment read GetHorizAlignment write SetHorizAlignment;
     property VertAlignment: TXc12VertAlignment read GetVertAlignment write SetVertAlignment;
     property Rotation: integer read GetRotation write SetRotation;
     property Options: TXc12AlignmentOptions read GetOptions write SetOptions;
     property Indent: integer read GetIndent write SetIndent;
     property RelativeIndent: integer read GetRelativeIndent write SetRelativeIndent;
     property ReadingOrder: TXc12ReadOrder read GetReadingOrder write SetReadingOrder;
     end;

     TXc12NumberFormat = class(TXLSStyleObject)
private
     procedure SetValue(const Value: AxUCString);
     function  GetValue: AxUCString;
protected
     FValue: AxUCString;
     FIsDateTime: boolean;
     // True if a standard number format is redefined.
     FStdRedefined: boolean;

     procedure CalcHash; override;
public
     function AsString(const AIndent: integer = 0): AxUCString; override;

     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12NumberFormat);

     function Assigned: boolean;
     function IsDateTime: boolean;
     function IsDate: boolean;
     function IsTime: boolean;

     property Value: AxUCString read GetValue write SetValue;
     property StdRedefined: boolean read FStdRedefined write FStdRedefined;
     end;

     TXc12NumberFormats = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12NumberFormat;
protected
     procedure AddDefault(AValue: AxUCString; ALocked: boolean);
public
     constructor Create;

     function  Find(ANumFmt: TXc12NumberFormat): TXc12NumberFormat;
     procedure Clear; reintroduce;
     procedure SetIsDefault; override;

     function  DefaultNumFmt: TXc12NumberFormat;

     function Add: TXc12NumberFormat; overload;
     function Add(AValue: AxUCString; AIndex: integer): TXc12NumberFormat; overload;

     function BuiltInCount: integer;

     property Items[Index: integer]: TXc12NumberFormat read GetItems; default;
     end;

     TXc12Font = class(TXLSStyleObject)
private
     function  GetCharset: integer;
     function  GetColor: TXc12Color;
     function  GetFamily: integer;
     function  GetName: AxUCString;
     function  GetScheme: TXc12FontScheme;
     function  GetSize: double;
     function  GetStyle: TXc12FontStyles;
     function  GetSubSuperscript: TXc12SubSuperscript;
     function  GetUnderline: TXc12Underline;
     procedure SetCharset(const Value: integer);
     procedure SetColor(const Value: TXc12Color);
     procedure SetFamily(const Value: integer);
     procedure SetName(const Value: AxUCString);
     procedure SetScheme(const Value: TXc12FontScheme);
     procedure SetSize(const Value: double);
     procedure SetStyle(const Value: TXc12FontStyles);
     procedure SetSubSuperscript(const Value: TXc12SubSuperscript);
     procedure SetUnderline(const Value: TXc12Underline);
     function  GetPColor: PXc12Color;
protected
     FName          : AxUCString;
     FCharset       : integer;
     FFamily        : integer;
     FStyle         : TXc12FontStyles;
     FUnderline     : TXc12Underline;
     FSubSuperscript: TXc12SubSuperscript;
     FColor         : TXc12Color;
     FSize          : double;
     FScheme        : TXc12FontScheme;
{$ifdef XLS_XSS}
     FHandle        : longword;
     FLF            : LOGFONT;
{$endif}
     procedure CalcHash; override;
public
     destructor Destroy; override;

     function AsString(const AIndent: integer = 0): AxUCString; override;

     procedure Clear;

     procedure Use;

     function  Equal(AItem: TXLSStyleObject): boolean; overload; override;
     function  EqualTFont(AItem: TFont): boolean;
     procedure Assign(AItem: TXc12Font); overload;
     procedure Assign(AItem: TFont); overload;
     procedure Assign(ARPr: TCT_TextCharacterProperties); overload;
     procedure CopyToTFont(ADest: TFont);

     function  AsCSS(const AWriteSelector: boolean = True): AxUCString;
     function  CSSSelector: AxUCString;
{$ifdef XLS_XSS}
     function  Width: integer;
{$endif}
     // TODO Not correct calculation.
{$ifdef XLS_XSS}
     procedure CopyToLOGFONT(var Dest: LOGFONT; var FontColor: longword);
     function  GetHandle(APixelsPerInch: integer): integer;
{$endif}
     procedure DeleteHandle;

     property Name: AxUCString read GetName write SetName;
     property Charset: integer read GetCharset write SetCharset;
     property Family: integer read GetFamily write SetFamily;
     property Style: TXc12FontStyles read GetStyle write SetStyle;
     property Underline: TXc12Underline read GetUnderline write SetUnderline;
     property SubSuperscript: TXc12SubSuperscript read GetSubSuperscript write SetSubSuperscript;
     property Color: TXc12Color read GetColor write SetColor;
     property PColor: PXc12Color read GetPColor;
     property Size: double read GetSize write SetSize;
     property Scheme: TXc12FontScheme read GetScheme write SetScheme;
     end;

     TXc12Fonts = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12Font;
     function GetDefaultFont: TXc12Font;
protected
     // Font zero max witdth of 0..9, used by column widths.
     FStdFontWidth: integer;
     FStdFontHeight: integer;
{$ifdef XLS_XSS}
     procedure CalcStdFontWidthHeight;
{$endif}
public
     constructor Create;
     destructor Destroy; override;

     procedure FixupColors;

     procedure SetIsDefault; override;
     procedure Clear; reintroduce;
{$ifdef XLS_XSS}
     procedure DeleteHandles;
{$endif}
     function  MakeCommentFont(const ABold: boolean): TXc12Font;

     function  Find(AFont: TFont): TXc12Font; overload;
     function  Find(AFont: TXc12Font): TXc12Font; overload;

     function  Add: TXc12Font; overload;
     procedure Add(AFont: TXc12Font); overload;

     property Items[Index: integer]: TXc12Font read GetItems; default;
     property DefaultFont: TXc12Font read GetDefaultFont;
     end;

     PXc12FontRun = ^TXc12FontRun;
     TXc12FontRun = record
     // Index is zero-relative.
     Index: integer;
     Font: TXc12Font;
     end;

     TXc12DynFontRunArray = array of TXc12FontRun;
     TXc12FontRunArray = array[0..$FFFF] of TXc12FontRun;
     PXc12FontRunArray = ^TXc12FontRunArray;

     TXc12PhoneticRun = record
     Sb: integer;
     Eb: integer;
     Text: AxUCString;
     end;

     TXc12DynPhoneticRunArray = array of TXc12PhoneticRun;
     TXc12PhoneticRunArray = array[0..$FFFF] of TXc12PhoneticRun;
     PXc12PhoneticRunArray = ^TXc12PhoneticRunArray;

     TXc12XF = class(TXLSStyleObject)
private
     procedure SetBorder(const Value: TXc12Border);
     procedure SetFill(const Value: TXc12Fill);
     procedure SetFont(const Value: TXc12Font);
     procedure SetNumFmt(const Value: TXc12NumberFormat);
     procedure SetXF(const Value: TXc12XF);
     function  GetAlignment: TXc12CellAlignment;
     function  GetApply: TXc12ApplyFormats;
     function  GetBorder: TXc12Border;
     function  GetFill: TXc12Fill;
     function  GetFont: TXc12Font;
     function  GetPivotButton: boolean;
     function  GetProtection: TXc12CellProtections;
     function  GetQuotePrefix: boolean;
     function  GetXF: TXc12XF;
     procedure SetApply(const Value: TXc12ApplyFormats);
     procedure SetPivotButton(const Value: boolean);
     procedure SetProtection(const Value: TXc12CellProtections);
     procedure SetQuotePrefix(const Value: boolean);
     function  GetNumFmt: TXc12NumberFormat;
protected
     FAlignment  : TXc12CellAlignment;
     FProtection : TXc12CellProtections;
     FNumFmt     : TXc12NumberFormat;
     FFont       : TXc12Font;
     FBorder     : TXc12Border;
     FFill       : TXc12Fill;
     FXF         : TXc12XF;
     FApply      : TXc12ApplyFormats;
     FQuotePrefix: boolean;
     FPivotButton: boolean;

     FE97XfData1 : word;

     function  GetCSSBorder(const ABorder: TXc12BorderPr): AxUCString;
     procedure CalcHash; override;
public
     constructor Create(AOwner: TXLSStyleObjectList);
     destructor Destroy; override;

     function  AsString(const AIndent: integer = 0): AxUCString; override;

     function  IsDefault: boolean;

     function  AsCSS: AxUCString;
     function  BordersAsCSS: AxUCString;
     function  CSSSelector: AxUCString;

     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(const AStyle: TXc12XF);
     procedure AssignProps(const AStyle: TXc12XF);

     property Alignment: TXc12CellAlignment read GetAlignment;
     property Protection: TXc12CellProtections read GetProtection write SetProtection;
     property NumFmt: TXc12NumberFormat read GetNumFmt write SetNumFmt;
     property Font: TXc12Font read GetFont write SetFont;
     property Fill: TXc12Fill read GetFill write SetFill;
     property Border: TXc12Border read GetBorder write SetBorder;
     property XF: TXc12XF read GetXF write SetXF;
     property Apply: TXc12ApplyFormats read GetApply write SetApply;

     property E97XfData1: word read FE97XfData1 write FE97XfData1;

     property QuotePrefix: boolean read GetQuotePrefix write SetQuotePrefix;
     property PivotButton: boolean read GetPivotButton write SetPivotButton;
     end;

     TXc12XFs = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12XF;
protected
     FOwner: TXc12DataStyleSheet;
public
     constructor Create(AOwner: TXc12DataStyleSheet);
     destructor Destroy; override;

     function  DefaultXF: TXc12XF;
     procedure SetIsDefault; override;
     procedure Clear(AOwner: TXc12DataStyleSheet); reintroduce; overload;

     function  AddOrGetFree: TXc12XF;
     procedure FreeAndNil(AXF: TXc12XF);

     function  Add: TXc12XF; overload;
     function  Add(AStyle: TXc12XF): integer; overload;

     function  CopyAndAdd(AStyle: TXc12XF): TXc12XF;

     function  Find(AStyle: TXc12XF): TXc12XF;

     function  AddExtern(const AStyle: TXc12XF): TXc12XF;

     property Items[Index: integer]: TXc12XF read GetItems; default;
     end;

     TXc12DXF = class(TXLSStyleObject)
private
     function  GetAlignment: TXc12CellAlignment;
     function  GetBorder: TXc12Border;
     function  GetFill: TXc12Fill;
     function  GetFont: TXc12Font;
     function  GetNumFmt: TXc12NumberFormat;
     function  GetProtection: TXc12CellProtections;
protected
     FFont: TXc12Font;
     FNumFmt: TXc12NumberFormat;
     FFill: TXc12Fill;
     FBorder: TXc12Border;
     FAlignment: TXc12CellAlignment;
     FProtection: TXc12CellProtections;

     procedure CalcHash; override;
public
     destructor Destroy; override;

     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12DXF);

     function AddFont: TXc12Font;
     function AddNumFmt: TXc12NumberFormat;
     function AddFill: TXc12Fill;
     function AddAlignment: TXc12CellAlignment;
     function AddBorder: TXc12Border;

     property Font: TXc12Font read GetFont;
     property NumFmt: TXc12NumberFormat read GetNumFmt;
     property Fill: TXc12Fill read GetFill;
     property Alignment: TXc12CellAlignment read GetAlignment;
     property Border: TXc12Border read GetBorder;
     property Protection: TXc12CellProtections read GetProtection write FProtection;
     end;

     TXc12DXFs = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12DXF;
protected
     FOwner: TXc12DataStyleSheet;
public
     constructor Create(AOwner: TXc12DataStyleSheet);

     procedure SetIsDefault; override;
     procedure Clear; reintroduce;

     function  Find(ADXF: TXc12DXF): TXc12DXF;
     function  Add: TXc12DXF; overload;
     function  Add(const AStyle: TXc12DXF): integer; overload;

     property Owner: TXc12DataStyleSheet read FOwner;
     property Items[Index: integer]: TXc12DXF read GetItems; default;
     end;

     TXc12CellStyle = class(TXLSStyleObject)
private
     function  GetBuiltInId: integer;
     function  GetCustomBuiltIn: boolean;
     function  GetHidden: boolean;
     function  GetLevel: integer;
     function  GetName: AxUCString;
     function  GetXF: TXc12XF;
     procedure SetBuiltInId(const Value: integer);
     procedure SetCustomBuiltIn(const Value: boolean);
     procedure SetHidden(const Value: boolean);
     procedure SetLevel(const Value: integer);
     procedure SetName(const Value: AxUCString);
     procedure SetXF(const Value: TXc12XF);
protected
     FBuiltInId: integer;
     FCustomBuiltIn: boolean;
     FHidden: boolean;
     FLevel: integer;
     FName: AxUCString;
     FXF: TXc12XF;

     procedure CalcHash; override;
public
     function  Equal(AItem: TXLSStyleObject): boolean; override;
     procedure Assign(AItem: TXc12CellStyle);

     property BuiltInId: integer read GetBuiltInId write SetBuiltInId;
     property CustomBuiltIn: boolean read GetCustomBuiltIn write SetCustomBuiltIn;
     property Hidden: boolean read GetHidden write SetHidden;
     property Level: integer read GetLevel write SetLevel;
     property Name: AxUCString read GetName write SetName;
     property XF: TXc12XF read GetXF write SetXF;
     end;

     TXc12CellStyles = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12CellStyle;
protected
public
     constructor Create;

     procedure SetIsDefault; override;
     procedure Clear; reintroduce;

     function  Find(AName: AxUCString): TXc12CellStyle;
     function  Add: TXc12CellStyle;
     function  AddOrGet(AName: AxUCString): TXc12CellStyle;

     property Items[Index: integer]: TXc12CellStyle read GetItems; default;
     end;

     TXc12TableStyleElement = class(TObject)
private
     function  GetDXF: TXc12DXF;
     function  GetSize: integer;
     function  GetType: TXc12TableStyleType;
     procedure SetDXf(const Value: TXc12DXF);
     procedure SetSize(const Value: integer);
     procedure SetType(const Value: TXc12TableStyleType);
protected
     FType: TXc12TableStyleType;
     FSize: integer;
     FDXF: TXc12DXF;
public
     constructor Create;

     procedure Clear;
     function  Equal(AItem: TXc12TableStyleElement): boolean;
     procedure Assign(AItem: TXc12TableStyleElement);

     property Type_: TXc12TableStyleType read GetType write SetType;
     property Size: integer read GetSize write SetSize;
     property DXF: TXc12DXF read GetDXF write SetDXf;
     end;

     TXc12TableStyleElements = class(TObjectList)
protected
     function GetItems(Index: integer): TXc12TableStyleElement;
public
     constructor Create;

     function Add: TXc12TableStyleElement;

     property Items[Index: integer]: TXc12TableStyleElement read GetItems; default;
     end;

     TXc12TableStyle = class(TXLSStyleObject)
private
     function  GetName: AxUCString;
     function  GetPivot: boolean;
     function  GetTable: boolean;
     procedure SetName(const Value: AxUCString);
     procedure SetPivot(const Value: boolean);
     procedure SetTable(const Value: boolean);
     function  GetTableStyleElements: TXc12TableStyleElements;
protected
     FName: AxUCString;
     FPivot: boolean;
     FTable: boolean;
     FTableStyleElements: TXc12TableStyleElements;

     procedure CalcHash; override;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure Assign(AItem: TXc12TableStyle);
     function  Equal(AItem: TXLSStyleObject): boolean; override;

     property Name: AxUCString read GetName write SetName;
     property Pivot: boolean read GetPivot write SetPivot;
     property Table: boolean read GetTable write SetTable;
     property TableStyleElements: TXc12TableStyleElements read GetTableStyleElements;
     end;

     TXc12TableStyles = class(TXLSStyleObjectList)
private
     function GetItems(Index: integer): TXc12TableStyle;
protected
     FDefaultTableStyle: AxUCString;
     FDefaultPivotStyle: AxUCString;
public
     constructor Create;

     procedure Clear; reintroduce;

     procedure SetIsDefault; override;

     function Add: TXc12TableStyle;

     property DefaultTableStyle: AxUCString read FDefaultTableStyle write FDefaultTableStyle;
     property DefaultPivotStyle: AxUCString read FDefaultPivotStyle write FDefaultPivotStyle;
     property Items[Index: integer]: TXc12TableStyle read GetItems; default;
     end;

     TXc12XFEditor = class(TObject)
private
     procedure SetFontSize(const Value: double);
     procedure SetFontCharset(const Value: TFontCharset);
     procedure SetFontColorRGB(const Value: TXc12RGBColor);
     procedure SetFontFamily(const Value: integer);
     procedure SetFontName(const Value: AxUCString);
     procedure SetFontStyle(const Value: TXc12FontStyles);
     procedure SetFontSubSuperScript(const Value: TXc12SubSuperscript);
     procedure SetFontUnderline(const Value: TXc12Underline);

     procedure SetFillColorRGB(const Value: TXc12RGBColor);
     procedure SetFillBgColorRGB(const Value: TXc12RGBColor);
     procedure SetFillColor(const Value: TXc12Color);
     procedure SetFillBgColor(const Value: TXc12Color);
     procedure SetFillPattern(const Value: TXc12FillPattern);

     procedure SetNumberFormat(const Value: AxUCString);

     procedure SetBorderBottomColor(const Value: TXc12Color);
     procedure SetBorderBottomColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderDiagColor(const Value: TXc12Color);
     procedure SetBorderDiagColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderLeftColor(const Value: TXc12Color);
     procedure SetBorderLeftColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderRightColor(const Value: TXc12Color);
     procedure SetBorderRightColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderRightStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderTopColor(const Value: TXc12Color);
     procedure SetBorderTopColorRGB(const Value: TXc12RGBColor);
     procedure SetBorderTopStyle(const Value: TXc12CellBorderStyle);

     procedure SetProtectionHidden(const Value: boolean);
     procedure SetProtectionLocked(const Value: boolean);

     procedure SetAlignHoriz(const Value: TXc12HorizAlignment);
     procedure SetAlignIndent(const Value: integer);
     procedure SetAlignOptions(const Value: TXc12AlignmentOptions);
     procedure SetAlignReadingOrder(const Value: TXc12ReadOrder);
     procedure SetAlignRelativeIndent(const Value: integer);
     procedure SetAlignRotation(const Value: integer);
     procedure SetAlignVert(const Value: TXc12VertAlignment);

protected
     FStyles        : TXc12DataStyleSheet;

     FXF            : TXc12XF;

     FOrigXF        : TXc12XF;

     FCmpXF         : TXc12XF;
     FCmpNumFmt     : TXc12NumberFormat;
     FCmpFont       : TXc12Font;
     FCmpFill       : TXc12Fill;
     FCmpBorder     : TXc12Border;

     FProtectChanged: boolean;
     FAlignChanged  : boolean;
     FNumFmtChanged : boolean;
     FFontChanged   : boolean;
     FFillChanged   : boolean;
     FBorderChanged : boolean;

     FCompactCount  : integer;

     procedure DoFontChanged;
     procedure DoFillChanged;
     procedure DoBorderChanged;
     procedure DoAlignChanged;

     procedure DoCompactStyles(AStyles: TXLSStyleObjectList);
     procedure DoLockStyles(AStyles: TXLSStyleObjectList);
public
     constructor Create(AStyleSheet: TXc12DataStyleSheet);
     destructor Destroy; override;

     function  UseDefault: TXc12XF;

     procedure UseStyle(const AIndex: integer); overload;  //{$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure UseStyle(AXF: TXc12XF); overload;           //{$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure FreeStyle(const AIndex: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure FreeStyle(AXF: TXc12XF); overload;          {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DecStyle(const AIndex: integer);            {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure UseFont(AFont: TXc12Font);                  {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure UseFonts(AFonts: TXc12DynFontRunArray);
     function  UseRequiredFont(AFont: TXc12Font): TXc12Font;

     procedure FreeFont(AFont: TXc12Font);                 {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure FreeFonts(AFonts: TXc12DynFontRunArray); overload;
     procedure FreeFonts(AFonts: PXc12FontRunArray; const ACount: integer); overload;

     procedure UseStyleStyle(const AIndex: integer);       {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure BeginEdit(AXF: TXc12XF);
     function  EndEdit: TXc12XF;

     procedure CompactStyles;
     //* Locks all styles, formats, fonts, fills and borders.
     //* Used after loading default Excel 97 file.
     procedure LockAll;

     property FontName            : AxUCString write SetFontName;
     property FontCharset         : TFontCharset write SetFontCharset;
     property FontFamily          : integer write SetFontFamily;
     property FontColorRGB        : TXc12RGBColor write SetFontColorRGB;
     property FontSize            : double write SetFontSize;
     property FontStyle           : TXc12FontStyles write SetFontStyle;
     property FontSubSuperScript  : TXc12SubSuperscript write SetFontSubSuperScript;
     property FontUnderline       : TXc12Underline write SetFontUnderline;

     property FillColor           : TXc12Color write SetFillColor;
     property FillBgColor         : TXc12Color write SetFillBgColor;
     property FillColorRGB        : TXc12RGBColor write SetFillColorRGB;
     property FillBgColorRGB      : TXc12RGBColor write SetFillBgColorRGB;
     property FillPattern         : TXc12FillPattern write SetFillPattern;

     property NumberFormat        : AxUCString write SetNumberFormat;

     property BorderLeftColor     : TXc12Color write SetBorderLeftColor;
     property BorderLeftColorRGB  : TXc12RGBColor write SetBorderLeftColorRGB;
     property BorderLeftStyle     : TXc12CellBorderStyle write SetBorderLeftStyle;

     property BorderTopColor      : TXc12Color write SetBorderTopColor;
     property BorderTopColorRGB   : TXc12RGBColor write SetBorderTopColorRGB;
     property BorderTopStyle      : TXc12CellBorderStyle write SetBorderTopStyle;

     property BorderRightColor    : TXc12Color write SetBorderRightColor;
     property BorderRightColorRGB : TXc12RGBColor write SetBorderRightColorRGB;
     property BorderRightStyle    : TXc12CellBorderStyle write SetBorderRightStyle;

     property BorderBottomColor   : TXc12Color write SetBorderBottomColor;
     property BorderBottomColorRGB: TXc12RGBColor write SetBorderBottomColorRGB;
     property BorderBottomStyle   : TXc12CellBorderStyle write SetBorderBottomStyle;

     property BorderDiagColor     : TXc12Color write SetBorderDiagColor;
     property BorderDiagColorRGB  : TXc12RGBColor write SetBorderDiagColorRGB;
     property BorderDiagStyle     : TXc12CellBorderStyle write SetBorderDiagStyle;

     property ProtectionLocked    : boolean write SetProtectionLocked;
     property ProtectionHidden    : boolean write SetProtectionHidden;

     property AlignHoriz          : TXc12HorizAlignment write SetAlignHoriz;
     property AlignVert           : TXc12VertAlignment write SetAlignVert;
     property AlignRotation       : integer write SetAlignRotation;
     property AlignOptions        : TXc12AlignmentOptions write SetAlignOptions;
     property AlignIndent         : integer write SetAlignIndent;
     property AlignRelIndent      : integer write SetAlignRelativeIndent;
     property AlignReadingOrder   : TXc12ReadOrder write SetAlignReadingOrder;
     end;

     TXc12DataStyleSheet = class(TXc12Data)
private
     function  GetStdFontWidth: integer;
     function  GetStdFontHeight: integer;
protected
     FXFs         : TXc12XFs;
     FDXFs        : TXc12DXFs;
     FStyleXFs    : TXc12XFs;
     FNumFmts     : TXc12NumberFormats;
     FFonts       : TXc12Fonts;
     FFills       : TXc12Fills;
     FBorders     : TXc12Borders;
     FStyles      : TXc12CellStyles;
     FTableStyles : TXc12TableStyles;
     FMruColors   : TXc12ColorObjs;
     FInxColors   : TXc12IndexColorObjs;

     FXFEditor    : TXc12XFEditor;

     FPPIEvent    : TXLSPPIEvent;

     function  FindEqualStyle(const AStyle: integer): TXc12XF; overload;
     function  FindEqualStyle(const AStyle: TXc12XF): TXc12XF; overload;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure PixelsPerInchXY(out AX,AY: integer);

     procedure SetXFDefaultValues(AXF: TXc12XF);

     procedure AfterRead;
     procedure BeforeWrite;

     property XFs          : TXc12XFs read FXFs;
     property DXFs         : TXc12DXFs read FDXFs;
     property StyleXFs     : TXc12XFs read FStyleXFs;
     property NumFmts      : TXc12NumberFormats read FNumFmts;
     property Fonts        : TXc12Fonts read FFonts;
     property Fills        : TXc12Fills read FFills;
     property Borders      : TXc12Borders read FBorders;
     property Styles       : TXc12CellStyles read FStyles;
     property TableStyles  : TXc12TableStyles read FTableStyles;
     property MruColors    : TXc12ColorObjs read FMruColors;
     property InxColors    : TXc12IndexColorObjs read FInxColors;

     property XFEditor     : TXc12XFEditor read FXFEditor;

     property StdFontWidth : integer read GetStdFontWidth;
     property StdFontHeight: integer read GetStdFontHeight;

     property OnPixelsPerInch: TXLSPPIEvent read FPPIEvent write FPPIEvent;
     end;

implementation

function Xc12ColorAsString(const AIndent: integer; const AColor: TXc12Color): AxUCString;
begin
  Result := MakeIndent(AIndent) + 'ARGB: ' + Format('%.8X',[AColor.ARGB]) + #13 +
            MakeIndent(AIndent) + 'Tint: ' + Format('%.2f',[AColor.Tint]) + #13 +
            MakeIndent(AIndent) + 'Color type:' + #13;
  case AColor.ColorType of
    exctAuto      : Result := Result + MakeIndent(AIndent) + 'Auto' + #13;
    exctIndexed   : Result := Result + MakeIndent(AIndent) + 'Indexed: ' + IntToStr(Integer(AColor.Indexed)) + #13;
    exctRgb       : Result := Result + MakeIndent(AIndent) + 'RGB: ' + Format('%.8X',[AColor.OrigRGB]) + #13;
    exctTheme     : Result := Result + MakeIndent(AIndent) + 'Theme: ' + IntToStr(Integer(AColor.Theme)) + #13;
    exctUnassigned: Result := Result + MakeIndent(AIndent) + '!UNASSIGNED!' + #13;
  end;
end;

{ TXc12GradientStopList }

function TXc12GradientStopList.Add: TXc12GradientStop;
var
  Item: TXc12GradientStop;
begin
  Item := TXc12GradientStop.Create;
  Result := Item;
  inherited Add(Item);
end;

function TXc12GradientStopList.Add(AColor: TXc12Color; APosition: double): TXc12GradientStop;
begin
  Result := Add;
  Result.Color := AColor;
  Result.Position := APosition;
end;

function TXc12GradientStopList.CalcHash: longword;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Result := Result xor Items[i].CalcHash;
end;

function TXc12GradientStopList.Equal(AList: TXc12GradientStopList): boolean;
var
  i: integer;
begin
  Result := Count = AList.Count;
  if Result then begin
    for i := 0 to Count - 1 do begin
      Result := Xc12ColorEqual(Items[i].Color,AList[i].Color) and SameValue(Items[i].Position,AList[i].Position,0.000001);
      if not Result then
        Exit;
    end;
  end;
end;

function TXc12GradientStopList.GetItems(Index: integer): TXc12GradientStop;
begin
  Result := TXc12GradientStop(inherited Items[Index]);
end;

{ TXc12GradientFill }

procedure TXc12GradientFill.Assign(AItem: TXc12GradientFill);
var
  i: integer;
begin
  FStops.Clear;
  for i := 0 to AItem.Stops.Count - 1 do
    FStops.Add(AItem.Stops[i].Color,AItem.Stops[i].Position);
  FDegree := AItem.Degree;
  Left := AItem.Left;
  Right := AItem.Right;
  Top := AItem.Top;
  Bottom := AItem.Bottom;
end;

procedure TXc12GradientFill.CalcHash;
begin

end;

procedure TXc12GradientFill.Clear;
begin
  FStops.Clear;
  FGradientType := gftLinear;
  FDegree := 0;
  FLeft := 0;
  FRight := 0;
  FTop := 0;
  FBottom := 0;
  FCorners[gfcLeft] := False;
  FCorners[gfcRight] := False;
  FCorners[gfcTop] := False;
  FCorners[gfcBottom] := False;
end;

constructor TXc12GradientFill.Create;
begin
  FStops := TXc12GradientStopList.Create;
end;

destructor TXc12GradientFill.Destroy;
begin
  FStops.Free;
  inherited;
end;

function TXc12GradientFill.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := FStops.Equal(TXc12GradientFill(AItem).Stops);
  if Result then begin
     Result := (FGradientType = TXc12GradientFill(AItem).GradientType) and
     SameValue(FDegree,TXc12GradientFill(AItem).Degree,0.000001) and
     SameValue(FLeft,TXc12GradientFill(AItem).Left,0.000001) and
     SameValue(FRight,TXc12GradientFill(AItem).Right,0.000001) and
     SameValue(FTop,TXc12GradientFill(AItem).Top,0.000001) and
     SameValue(FBottom,TXc12GradientFill(AItem).Bottom,0.000001);
  end;
end;

function TXc12GradientFill.GetBottom: double;
begin
  Result := FBottom;
end;

function TXc12GradientFill.GetCorners(Index: TXc12GradientFillCorners): boolean;
begin
  Result := FCorners[Index];
end;

function TXc12GradientFill.GetDegree: double;
begin
  Result := FDegree;
end;

function TXc12GradientFill.GetGradientType: TXc12GradientFillType;
begin
  Result := FGradientType;
end;

function TXc12GradientFill.GetLeft: double;
begin
  Result := FLeft;
end;

function TXc12GradientFill.GetRight: double;
begin
  Result := FRight;
end;

function TXc12GradientFill.GetStops: TXc12GradientStopList;
begin
  Result := FStops;
end;

function TXc12GradientFill.GetTop: double;
begin
  Result := FTop;
end;

procedure TXc12GradientFill.SetBottom(const Value: double);
begin
  FBottom := Value;
  FCorners[gfcBottom] := True;
end;

procedure TXc12GradientFill.SetCorners(Index: TXc12GradientFillCorners; const Value: boolean);
begin
  FCorners[Index] := Value;
end;

procedure TXc12GradientFill.SetDegree(const Value: double);
begin
  FDegree := Value;
end;

procedure TXc12GradientFill.SetGradientType(const Value: TXc12GradientFillType);
begin
  FGradientType := Value;
end;

procedure TXc12GradientFill.SetLeft(const Value: double);
begin
  FLeft := Value;
  FCorners[gfcLeft] := True;
end;

procedure TXc12GradientFill.SetRight(const Value: double);
begin
  FRight := Value;
  FCorners[gfcRight] := True;
end;

procedure TXc12GradientFill.SetTop(const Value: double);
begin
  FTop := Value;
  FCorners[gfcTop] := True;
end;

{ TXc12Fill }

procedure TXc12Fill.Assign(AItem: TXc12Fill);
begin
  if AItem.IsGradientFill then begin
    IsGradientFill := True;
    FGradientFill.Assign(TXc12GradientFill(AItem));
  end
  else begin
    FPatternType := AItem.PatternType;
    FFgColor := AItem.FgColor;
    FBgColor := AItem.BgColor;
  end;
end;

function TXc12Fill.AsString(const AIndent: integer): AxUCString;
var
  S: AxUCString;
begin
  if IsGradientFill then
    Result := MakeIndent(AIndent) + '[Gradient - not shown]' + #13
  else begin
    Result := MakeIndent(AIndent) + 'Pattern:' + #31;
    case FPatternType of
      efpNone           : S := 'None';
      efpSolid          : S := 'Solid';
      efpMediumGray     : S := 'MediumGray';
      efpDarkGray       : S := 'DarkGray';
      efpLightGray      : S := 'LightGray';
      efpDarkHorizontal : S := 'DarkHorizontal';
      efpDarkVertical   : S := 'DarkVertical';
      efpDarkDown       : S := 'DarkDown';
      efpDarkUp         : S := 'DarkUp';
      efpDarkGrid       : S := 'DarkGrid';
      efpDarkTrellis    : S := 'DarkTrellis';
      efpLightHorizontal: S := 'LightHorizontal';
      efpLightVertical  : S := 'LightVertical';
      efpLightDown      : S := 'LightDown';
      efpLightUp        : S := 'LightUp';
      efpLightGrid      : S := 'LightGrid';
      efpLightTrellis   : S := 'LightTrellis';
      efpGray125        : S := 'Gray125';
      efpGray0625       : S := 'Gray0625';
    end;
    Result := Result + MakeIndent(AIndent + 1) + S + #13;

    Result := Result + 'Foreground color:' + #13 + Xc12ColorAsString(AIndent + 1,FFgColor);
    Result := Result + 'Background color:' + #13 + Xc12ColorAsString(AIndent + 1,FBgColor);
  end;
end;

procedure TXc12Fill.CalcHash;
begin
  if IsGradientFill then begin
    FGradientFill.CalcHash;
    FHash := FGradientFill.FHash;
  end
  else
    FHash := Xc12ColorHash(FFgColor) + Xc12ColorHash(FBgColor) + Longword(FPatternType)
end;

procedure TXc12Fill.Clear;
begin
  SetIsGradientFill(False);
  FPatternType := efpNone;
  FFgColor.ColorType := exctAuto;
  FBgColor.ColorType := exctAuto;
end;

destructor TXc12Fill.Destroy;
begin
  SetIsGradientFill(False);
  inherited;
end;

function TXc12Fill.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := not (IsGradientFill xor TXc12Fill(AItem).IsGradientFill);
  if Result then begin
    if IsGradientFill then
      Result := FGradientFill.Equal(TXc12Fill(AItem).GradientFill)
    else begin
       Result := (FPatternType = TXc12Fill(AItem).PatternType) and
                 Xc12ColorEqual(FFgColor,TXc12Fill(AItem).FgColor) and
                 Xc12ColorEqual(FBgColor,TXc12Fill(AItem).BgColor);

    end;
  end;
end;

function TXc12Fill.GetBgColor: TXc12Color;
begin
  Result := FBgColor;
end;

function TXc12Fill.GetFgColor: TXc12Color;
begin
  Result := FFgColor;
end;

function TXc12Fill.GetGradientFill: TXc12GradientFill;
begin
  Result := FGradientFill;
end;

function TXc12Fill.GetIsGradientFill: boolean;
begin
  Result := FGradientFill <> Nil;
end;

function TXc12Fill.GetPatternType: TXc12FillPattern;
begin
  Result := FPatternType;
end;

function TXc12Fill.GetPBgColor: PXc12Color;
begin
  Result := @FBgColor;
end;

function TXc12Fill.GetPFgColor: PXc12Color;
begin
  Result := @FFgColor;
end;

procedure TXc12Fill.SetBgColor(const Value: TXc12Color);
begin
  FBgColor := Value;
end;

procedure TXc12Fill.SetFgColor(const Value: TXc12Color);
begin
  FFgColor := Value;
end;

procedure TXc12Fill.SetIsGradientFill(const Value: boolean);
begin
  if Value and (FGradientFill = Nil) then
    FGradientFill := TXc12GradientFill.Create
  else if not Value and (FGradientFill <> Nil) then begin
    FGradientFill.Free;
    FGradientFill := Nil;
  end;
end;

procedure TXc12Fill.SetPatternType(const Value: TXc12FillPattern);
begin
  FPatternType := Value;
end;

{ TXc12Fills }

function TXc12Fills.Add: TXc12Fill;
begin
  Result := TXc12Fill.Create(Self);
  inherited Add(Result);
end;

procedure TXc12Fills.Add(const AFill: TXc12Fill);
begin
  AFill.FOwner := Self;
  inherited Add(AFill);
end;

procedure TXc12Fills.Clear;
begin
  inherited;

  SetIsDefault;
end;

function TXc12Fills.DefaultFill: TXc12Fill;
begin
  Result := Items[0];
end;

function TXc12Fills.Find(AFill: TXc12Fill): TXc12Fill;
{$ifndef STYLE_USE_HASHTABLES}
var
  i: integer;
{$endif}
begin
{$ifdef STYLE_USE_HASHTABLES}
  Result := TXc12Fill(inherited Find(AFill));
{$else}
  for i := 0 to Count - 1 do begin
    if Items[i].Equal(AFill) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
{$endif}
end;

procedure TXc12Fills.FixupColors;
var
  i,j: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FgColor.ColorType = exctIndexed then
      Items[i].FFgColor.ARGB := Xc12IndexColorPalette[Integer(Items[i].FFgColor.Indexed)];
    if Items[i].FBgColor.ColorType = exctIndexed then
      Items[i].FBgColor.ARGB := Xc12IndexColorPalette[Integer(Items[i].FBgColor.Indexed)];

    if Items[i].IsGradientFill then begin
      for j := 0 to Items[i].GradientFill.GetStops.Count - 1 do begin
        if Items[i].GradientFill.Stops[j].Color.ColorType = exctIndexed then
          Items[i].GradientFill.Stops[j].FColor.ARGB := Xc12IndexColorPalette[Integer(Items[i].GradientFill.Stops[j].FColor.Indexed)];
      end;
    end;
  end;
end;

function TXc12Fills.GetItems(Index: integer): TXc12Fill;
begin
  Result := TXc12Fill(inherited Items[Index]);
end;

procedure TXc12Fills.SetIsDefault;
var
  F: TXc12Fill;
begin
  F := Add;
  F.Locked := True;
  F.PatternType := efpNone;

  F := Add;
  F.Locked := True;
  F.PatternType := efpGray125;
end;

{ TXc12Borders }

function TXc12Borders.Add: TXc12Border;
begin
  Result := TXc12Border.Create(Self);
  inherited Add(Result);
end;

procedure TXc12Borders.Clear;
begin
  inherited;

  SetIsDefault;
end;

constructor TXc12Borders.Create;
begin
  inherited Create;
end;

function TXc12Borders.DefaultBorder: TXc12Border;
begin
  Result := Items[0];
end;

function TXc12Borders.Find(ABorder: TXc12Border): TXc12Border;
{$ifndef STYLE_USE_HASHTABLES}
var
  i: integer;
{$endif}
begin
{$ifdef STYLE_USE_HASHTABLES}
  Result := TXc12Border(inherited Find(ABorder));
{$else}
  for i := 0 to Count - 1 do begin
    if Items[i].Equal(ABorder) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
{$endif}
end;

function TXc12Borders.GetItems(Index: integer): TXc12Border;
begin
  Result := TXc12Border(inherited Items[Index]);
end;

procedure TXc12Borders.SetIsDefault;
var
  B: TXc12Border;
begin
  B := Add;
  B.Locked := True;
end;

{ TXc12XF }

function TXc12XF.AsCSS: AxUCString;
var
  TempDS: AxUCChar;
begin
  TempDS := FormatSettings.DecimalSeparator;
  try
    FormatSettings.DecimalSeparator := '.';

    Result := CSSSelector + ' {';

    if FFont.Index <> XLS_STYLE_DEFAULT_FONT then
      Result := Result + FFont.AsCSS(False);

    Result := Result + BordersAsCSS;

    if not ((FFill.FgColor.ColorType = exctAuto) or ((FFill.FgColor.ColorType = exctIndexed) and (FFill.FgColor.Indexed = xcAutomatic))) then
      Result := Result + 'background:' + Format(' #%.6X;',[Xc12ColorToRGB(FFill.FgColor)]);

    if FAlignment.HorizAlignment <> chaGeneral then begin
      case FAlignment.HorizAlignment of
        chaLeft            : Result := Result + ' text-align:left;';
        chaCenter          : Result := Result + ' text-align:center;';
        chaRight           : Result := Result + ' text-align:right;';
        chaFill            : ;
        chaJustify         : Result := Result + ' text-align:justify;';
        chaCenterContinuous: Result := Result + ' text-align:center;';
        chaDistributed     : Result := Result + ' text-align:center;';
      end;
    end;

    if FFont.SubSuperscript <> xssNone then begin
      case FFont.SubSuperscript of
        xssSuperscript: Result := Result + ' vertical-align:sup;';
        xssSubscript  : Result := Result + ' vertical-align:sub;';
      end;
    end
    else if FAlignment.VertAlignment <> cvaBottom then begin
      case FAlignment.VertAlignment of
        cvaTop        : Result := Result + ' vertical-align:top;';
        cvaCenter     : Result := Result + ' vertical-align:middle;';
        cvaJustify    : Result := Result + ' vertical-align:top;';
        cvaDistributed: Result := Result + ' vertical-align:middle;';
      end;
    end;

    Result := Result + '}';
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

procedure TXc12XF.Assign(const AStyle: TXc12XF);
begin
  FAlignment.Assign(AStyle.FAlignment);

  FProtection := AStyle.FProtection;

  FNumFmt := AStyle.FNumFmt;

  FFont := AStyle.FFont;

  FFill := AStyle.FFill;

  FBorder := AStyle.FBorder;

  FXF := AStyle.XF;

  FQuotePrefix := AStyle.FQuotePrefix;
  FPivotButton := AStyle.FPivotButton;
end;

procedure TXc12XF.AssignProps(const AStyle: TXc12XF);
begin
  FAlignment.Assign(AStyle.FAlignment);
  FProtection := AStyle.FProtection;
  FQuotePrefix := AStyle.FQuotePrefix;
  FPivotButton := AStyle.FPivotButton;
end;

function TXc12XF.AsString(const AIndent: integer = 0): AxUCString;
begin
  Result := MakeIndent(AIndent) + 'This XF: ' + IntToStr(FIndex) + #13 +
            MakeIndent(AIndent) + 'NumberFormat: ' + IntToStr(FNumFmt.Index) + #13 + FNumFmt.AsString(AIndent + 1) + #13 +
            MakeIndent(AIndent) + 'Font: '         + IntToStr(FFont.Index)   + #13 + FFont.AsString(AIndent + 1) + #13 +
            MakeIndent(AIndent) + 'Fill: '         + IntToStr(FFill.Index)   + #13 + FFill.AsString(AIndent + 1) + #13 +
            MakeIndent(AIndent) + 'Border: '       + IntToStr(FBorder.Index) + #13 + FBorder.AsString(AIndent + 1) + #13 +
            MakeIndent(AIndent) + 'Alignment: ' + #13 + FAlignment.AsString;
  if FXF <> Nil then
    Result := Result + MakeIndent(AIndent) + 'XF: ' + IntToStr(FXF.Index) + #13;
  if FProtection <> [] then begin
    Result := Result + MakeIndent(AIndent) + 'Protection:' + #13;
    if cpLocked in FProtection then
      Result := Result + MakeIndent(AIndent + 1) + 'Locked' + #13;
    if cpHidden in FProtection then
      Result := Result + MakeIndent(AIndent + 1) + 'Hidden' + #13;
  end;
  if FQuotePrefix then
    Result := Result + MakeIndent(AIndent) + 'QuotePrfix: True' + #13;
  if FPivotButton then
    Result := Result + MakeIndent(AIndent) + 'PivotButton: True' + #13;
end;

function TXc12XF.BordersAsCSS: AxUCString;
begin
  Result := '';
  if FBorder.FTop.Style <> cbsNone then
    Result := Result + 'border-top:' + GetCSSBorder(FBorder.FTop);
  if FBorder.FRight.Style <> cbsNone then
    Result := Result + 'border-right:' + GetCSSBorder(FBorder.FRight);
  if FBorder.FBottom.Style <> cbsNone then
    Result := Result + 'border-bottom:' + GetCSSBorder(FBorder.FBottom);
  if FBorder.FLeft.Style <> cbsNone then
    Result := Result + 'border-left:' + GetCSSBorder(FBorder.FLeft);
end;

procedure TXc12XF.CalcHash;
begin
  FHash := FAlignment.HashKey;
  FHash := FHash xor Longword(Byte(FProtection));
  FHash := FHash xor FNumFmt.HashKey;
  FHash := FHash xor FFont.HashKey;
  FHash := FHash xor FFill.HashKey;
  FHash := FHash xor FBorder.HashKey;
  if FXF <> Nil then
    FHash := FHash xor FXF.HashKey;
  Inc(FHash,Longword(FQuotePrefix));
  Inc(FHash,Longword(FPivotButton));

  FHashValid := True;
end;

constructor TXc12XF.Create(AOwner: TXLSStyleObjectList);
begin
  inherited Create(AOwner);

  FAlignment := TXc12CellAlignment.Create;
  FProtection := [cpLocked];
end;

function TXc12XF.CSSSelector: AxUCString;
begin
  Result := 'xf' + IntToStr(Index + 1);
end;

destructor TXc12XF.Destroy;
begin
  FAlignment.Free;
  FFont := Nil;
  FNumFmt := Nil;
  FFill := Nil;
  FBorder := Nil;
  FXF := Nil;
  inherited;
end;

function TXc12XF.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := FAlignment.Equal(TXc12XF(AItem).FAlignment)    and
           (FProtection = TXc12XF(AItem).FProtection)      and
            FNumFmt.Equal(TXc12XF(AItem).FNumFmt)          and
            FFont.Equal(TXc12XF(AItem).FFont)              and
            FFill.Equal(TXc12XF(AItem).FFill)              and
            FBorder.Equal(TXc12XF(AItem).FBorder)          and
           (FQuotePrefix = TXc12XF(AItem).FQuotePrefix)    and
           (FPivotButton = TXc12XF(AItem).FPivotButton);
  if Result and (FXF <> Nil) and (TXc12XF(AItem).FXF <> Nil) then
    Result := FXF.Index = TXc12XF(AItem).FXF.Index;
end;

function TXc12XF.GetAlignment: TXc12CellAlignment;
begin
  Result := FAlignment;
end;

function TXc12XF.GetApply: TXc12ApplyFormats;
begin
  Result := FApply;
end;

function TXc12XF.GetBorder: TXc12Border;
begin
  Result := FBorder;
end;

function TXc12XF.GetCSSBorder(const ABorder: TXc12BorderPr): AxUCString;
begin
  if ABorder.Style = cbsNone then
    Result := 'none'
  else begin
    if ABorder.Style in [cbsMedium,cbsMediumDashed,cbsMediumDashDot,cbsMediumDashDotDot] then
      Result := '1pt'
    else if ABorder.Style in [cbsThick,cbsDouble] then
      Result := '1.5pt'
    else
      Result := '.5pt';

    case ABorder.Style of
      cbsThin            : Result := Result + ' solid';
      cbsMedium          : Result := Result + ' solid';
      cbsDashed          : Result := Result + ' dashed';
      cbsDotted          : Result := Result + ' dotted';
      cbsThick           : Result := Result + ' solid';
      cbsDouble          : Result := Result + ' double';
      cbsHair            : Result := Result + ' solid';
      cbsMediumDashed    : Result := Result + ' dashed';
      cbsDashDot         : Result := Result + ' dashed';
      cbsMediumDashDot   : Result := Result + ' dashed';
      cbsDashDotDot      : Result := Result + ' dotted';
      cbsMediumDashDotDot: Result := Result + ' dotted';
      cbsSlantedDashDot  : Result := Result + ' dashed';
    end;

    if ABorder.Color.ColorType = exctAuto then
      Result := Result + ' windowtext;'
    else
      Result := Result + Format(' #%.6X;',[Xc12ColorToRGB(ABorder.Color)]);
  end;
end;

function TXc12XF.GetFill: TXc12Fill;
begin
  Result := FFill;
end;

function TXc12XF.GetFont: TXc12Font;
begin
  Result := FFont;
end;

function TXc12XF.GetNumFmt: TXc12NumberFormat;
begin
  Result := FNumFmt;
end;

function TXc12XF.GetPivotButton: boolean;
begin
  Result := FPivotButton;
end;

function TXc12XF.GetProtection: TXc12CellProtections;
begin
  Result := FProtection;
end;

function TXc12XF.GetQuotePrefix: boolean;
begin
  Result := FQuotePrefix;
end;

function TXc12XF.GetXF: TXc12XF;
begin
  Result := FXF;
end;

function TXc12XF.IsDefault: boolean;
begin
  Result := FIndex = XLS_STYLE_DEFAULT_XF;
end;

procedure TXc12XF.SetApply(const Value: TXc12ApplyFormats);
begin
  FApply := Value;
end;

procedure TXc12XF.SetBorder(const Value: TXc12Border);
begin
  FBorder := Value;
end;

procedure TXc12XF.SetFill(const Value: TXc12Fill);
begin
  FFill := Value;
end;

procedure TXc12XF.SetFont(const Value: TXc12Font);
begin
  FFont := Value;
end;

procedure TXc12XF.SetNumFmt(const Value: TXc12NumberFormat);
begin
  FNumFmt := Value;
end;

procedure TXc12XF.SetPivotButton(const Value: boolean);
begin
  if Value <> FPivotButton then begin
    FPivotButton := Value;
    FHashValid := False;
  end;
end;

procedure TXc12XF.SetProtection(const Value: TXc12CellProtections);
begin
  if Value <> FProtection then begin
    FProtection := Value;
    FHashValid := False;
  end;
end;

procedure TXc12XF.SetQuotePrefix(const Value: boolean);
begin
  if Value <> FQuotePrefix then begin
    FQuotePrefix := Value;
    FHashValid := False;
  end;
end;

procedure TXc12XF.SetXF(const Value: TXc12XF);
begin
  if TXc12XF(Value) <> Self then
    FXF := Value
  else
    FXF := Nil;
end;

{ TXc12DataStyleSheet }

procedure TXc12DataStyleSheet.AfterRead;
begin
  FFills.FixupColors;
  FMRUColors.FixupColors;
  FFonts.FixupColors;
end;

procedure TXc12DataStyleSheet.BeforeWrite;
begin
  // Force compact.
  FXFEditor.FCompactCount := XC12_STYLES_COMPACT_CNT;
  FXFEditor.CompactStyles;

  FNumFmts.Enumerate;
  FFonts.Enumerate;
  FFills.Enumerate;
  FBorders.Enumerate;
  FStyles.Enumerate;
  FTableStyles.Enumerate;
  FMruColors.Enumerate;
  FInxColors.Enumerate;
  FXFs.Enumerate;
  FStyleXFs.Enumerate;
  FDXFs.Enumerate;
end;

procedure TXc12DataStyleSheet.Clear;
begin
  FNumFmts.Clear;
  FFonts.Clear;
  FFills.Clear;
  FBorders.Clear;
  FStyles.Clear;
  FTableStyles.Clear;
  FMruColors.Clear;
  FInxColors.Clear;
  FXFs.Clear(Self);
  FStyleXFs.Clear(Self);
  FDXFs.Clear;
end;

constructor TXc12DataStyleSheet.Create;
begin
  inherited Create;

  FXFs := TXc12XFs.Create(Self);
  FDXFs := TXc12DXFs.Create(Self);
  FStyleXFs := TXc12XFs.Create(Self);
  FNumFmts := TXc12NumberFormats.Create;
  FFonts := TXc12Fonts.Create;
  FFills := TXc12Fills.Create;
  FBorders := TXc12Borders.Create;
  FStyles := TXc12CellStyles.Create;
  FTableStyles := TXc12TableStyles.Create;
  FMruColors := TXc12ColorObjs.Create;
  FInxColors := TXc12IndexColorObjs.Create;

  FXFEditor := TXc12XFEditor.Create(Self);

  Clear;
end;

destructor TXc12DataStyleSheet.Destroy;
begin
  FXFs.Free;
  FDXFs.Free;
  FStyleXFs.Free;
  FNumFmts.Free;
  FFonts.Free;
  FFills.Free;
  FBorders.Free;
  FStyles.Free;
  FTableStyles.Free;
  FMruColors.Free;
  FInxColors.Free;
  FXFEditor.Free;
  inherited;
end;

function TXc12DataStyleSheet.FindEqualStyle(const AStyle: TXc12XF): TXc12XF;
var
  i: integer;
begin
  if AStyle.Index = XLS_STYLE_DEFAULT_XF then
    Result := AStyle
  else begin
    for i := 0 to FXFs.Count - 1 do begin
      if (i <> AStyle.Index) and FXFs[i].Equal(AStyle) then begin
        Result := FXFs[i];
        Exit;
      end;
    end;
    Result := Nil;
  end;
end;

function TXc12DataStyleSheet.FindEqualStyle(const AStyle: integer): TXc12XF;
var
  i: integer;
  S: TXc12XF;
begin
  if AStyle = XLS_STYLE_DEFAULT_XF then
    Result := FXFs[XLS_STYLE_DEFAULT_XF]
  else begin
    S := FXFs[AStyle];
    for i := 0 to FXFs.Count - 1 do begin
      if (i <> AStyle) and FXFs[i].Equal(S) then begin
        Result := FXFs[i];
        Exit;
      end;
    end;
    Result := Nil;
  end;
end;

function TXc12DataStyleSheet.GetStdFontHeight: integer;
begin
  Result := FFonts.FStdFontHeight;
end;

function TXc12DataStyleSheet.GetStdFontWidth: integer;
begin
  Result := FFonts.FStdFontWidth;
end;

procedure TXc12DataStyleSheet.PixelsPerInchXY(out AX, AY: integer);
{$ifdef MSWINDOWS}
var
  DC: HDC;
{$endif}
begin
  if Assigned(FPPIEvent) then
    FPPIEvent(AX,AY)
  else begin
{$ifdef MACOS}
{$message warn '__TODO__ Requires obtaining info from the target OS'}
    AX := 96;
    AY := 96;
{$else}
    DC := GetDC(0);
    try
      AX := GetDeviceCaps(DC, LOGPIXELSX);
      AY := GetDeviceCaps(DC, LOGPIXELSY);
    finally
      ReleaseDC(0,DC);
    end;
{$endif}
  end;
end;

procedure TXc12DataStyleSheet.SetXFDefaultValues(AXF: TXc12XF);
begin
  AXF.Border := FBorders[0];
  AXF.Fill := FFills[0];
  AXF.Font := FFonts[0];
  AXF.NumFmt := FNumFmts[0];
  AXF.FXF := Nil;
end;

{ TXc12CellStyles }

function TXc12CellStyles.Add: TXc12CellStyle;
begin
  Result := TXc12CellStyle.Create(Self);
  inherited Add(Result);
end;

function TXc12CellStyles.AddOrGet(AName: AxUCString): TXc12CellStyle;
begin
  Result := Find(AName);
  if Result = Nil then begin
    Result := Add;
    Result.Name := AName;
  end;
end;

procedure TXc12CellStyles.Clear;
begin
  inherited;
  SetIsDefault;
end;

constructor TXc12CellStyles.Create;
begin
  inherited Create;
end;

function TXc12CellStyles.Find(AName: AxUCString): TXc12CellStyle;
var
  i: integer;
begin
  AName := Lowercase(AName);

  for i := 0 to Count - 1 do begin
    if Lowercase(Items[i].Name) = AName then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12CellStyles.GetItems(Index: integer): TXc12CellStyle;
begin
  Result := TXc12CellStyle(inherited Items[Index]);
end;

procedure TXc12CellStyles.SetIsDefault;
var
  Style: TXc12CellStyle;
begin
  Style := Add;
  Style.Name := 'Normal';
  Style.Locked := True;
end;

{ TXc12ColorObjs }

function TXc12ColorObjs.Add: TXc12ColorObj;
begin
  Result := TXc12ColorObj.Create(Self);
  inherited Add(Result);
end;

function TXc12ColorObjs.Add(AColor: TXc12Color): TXc12ColorObj;
begin
  Result := Add;
  Result.FColor := AColor;
end;

procedure TXc12ColorObjs.Clear;
begin
  inherited;

  SetIsDefault;
end;

constructor TXc12ColorObjs.Create;
begin
  inherited Create;
end;

procedure TXc12ColorObjs.FixupColors;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Color.ColorType = exctIndexed then
      Items[i].FColor.ARGB := Xc12IndexColorPalette[Integer(Items[i].Color.Indexed)];
  end;
end;

function TXc12ColorObjs.GetItems(Index: integer): TXc12ColorObj;
begin
  Result := TXc12ColorObj(inherited Items[Index]);
end;

procedure TXc12ColorObjs.SetIsDefault;
begin

end;

{ TXc12IndexColorObjs }

function TXc12IndexColorObjs.Add: TXc12IndexColorObj;
begin
  Result := TXc12IndexColorObj.Create(Self);
  inherited Add(Result);
end;

function TXc12IndexColorObjs.Add(ARGB: integer): TXc12IndexColorObj;
begin
  Result := Add;
  Result.RGB := ARGB;
end;

procedure TXc12IndexColorObjs.Clear;
begin
  inherited;

  SetIsDefault;
end;

constructor TXc12IndexColorObjs.Create;
begin
  inherited Create;
end;

function TXc12IndexColorObjs.GetItems(Index: integer): TXc12IndexColorObj;
begin
  Result := TXc12IndexColorObj(inherited Items[Index]);
end;

procedure TXc12IndexColorObjs.SetIsDefault;
begin

end;

{ TEx12DXFs }

function TXc12DXFs.Add: TXc12DXF;
begin
  Result := TXc12DXF.Create(Self);
  inherited Add(Result);
end;

function TXc12DXFs.Add(const AStyle: TXc12DXF): integer;
begin
  inherited Add(AStyle);
  Result := AStyle.Index;
end;

procedure TXc12DXFs.Clear;
begin
  inherited;

  SetIsDefault;
end;

constructor TXc12DXFs.Create(AOwner: TXc12DataStyleSheet);
begin
  inherited Create;

  FOwner := AOwner;
end;

function TXc12DXFs.Find(ADXF: TXc12DXF): TXc12DXF;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Equal(ADXF) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12DXFs.GetItems(Index: integer): TXc12DXF;
begin
  Result := TXc12DXF(inherited Items[Index]);
end;

procedure TXc12DXFs.SetIsDefault;
begin

end;

{ TXc12Fonts }

function TXc12Fonts.Add: TXc12Font;
begin
  Result := TXc12Font.Create(Self);
  inherited Add(Result);
end;

procedure TXc12Fonts.Add(AFont: TXc12Font);
begin
  AFont.FOwner := Self;
  inherited Add(AFont);
end;

{$ifdef XLS_XSS}
procedure TXc12Fonts.CalcStdFontWidthHeight;
var
  i: integer;
  S: AxUCString;
  Sz: TSize;
  DC: HDC;
begin
  FStdFontWidth := 0;
  DC := CreateIC('DISPLAY',Nil,Nil,Nil);
  try
    SelectObject(DC,Items[0].GetHandle(GetDeviceCaps(DC,LOGPIXELSY)));
    for i := 1 to 10 do begin
      S := StrDigits[i];
      Windows.GetTextExtentPoint32W(DC,PWideChar(S),1,Sz);
      if Sz.cx > FStdFontWidth then
        FStdFontWidth := Sz.cx;
    end;
    FStdFontHeight := Items[0].FLF.lfHeight;
  finally
    DeleteDC(DC);
  end;
end;
{$endif}

procedure TXc12Fonts.Clear;
begin
  inherited Clear;

  SetIsDefault;
end;

constructor TXc12Fonts.Create;
begin
  inherited Create;
end;

{$ifdef XLS_XSS}
procedure TXc12Fonts.DeleteHandles;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].DeleteHandle;
end;
{$endif}

destructor TXc12Fonts.Destroy;
begin
  inherited;
end;

function TXc12Fonts.Find(AFont: TXc12Font): TXc12Font;
{$ifndef STYLE_USE_HASHTABLES}
var
  i: integer;
{$endif}
begin
{$ifdef STYLE_USE_HASHTABLES}
  Result := TXc12Font(inherited Find(AFont));
{$else}
  for i := 0 to Count - 1 do begin
    if Items[i].Equal(AFont) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
{$endif}
end;

function TXc12Fonts.Find(AFont: TFont): TXc12Font;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].EqualTFont(AFont) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

procedure TXc12Fonts.FixupColors;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Color.ColorType = exctIndexed then
      Items[i].FColor.ARGB := Xc12IndexColorPalette[Integer(Items[i].FColor.Indexed)];
  end;
end;

function TXc12Fonts.GetDefaultFont: TXc12Font;
begin
  Result := Items[0];
end;

function TXc12Fonts.GetItems(Index: integer): TXc12Font;
begin
  Result := TXc12Font(inherited Items[Index]);
end;

function TXc12Fonts.MakeCommentFont(const ABold: boolean): TXc12Font;
var
  Font: TXc12Font;
begin
  Font := TXc12Font.Create(Nil);
  Font.Name := 'Tahoma';
  Font.Size := 8;
  Font.Color := RGBColorToXc12($00000000);
  Font.Family := 2;
  if ABold then
    Font.Style := [xfsBold];

  Result := Find(Font);
  if Result <> Nil then
    Font.Free
  else begin
    Add(Font);
    Result := Font;
  end;
  Result.Locked := True;
end;

procedure TXc12Fonts.SetIsDefault;
var
  Font: TXc12Font;
begin
  Font := Add;
  Font.Name := 'Calibri';
  Font.Size := 11;
  Font.FColor.ColorType := exctTheme;
  Font.FColor.Theme := cscDk1;
  Font.FColor.Tint := 0;
  Font.Family := 2;
//  Font.Scheme := efsMinor;
  Font.Scheme := efsNone;
  Font.Locked := True;
end;

{ TXc12NumberFormats }

function TXc12NumberFormats.Add(AValue: AxUCString; AIndex: integer): TXc12NumberFormat;
var
  i: integer;
begin
  if AIndex <= High(ExcelStandardNumFormats) then begin
    Result := Items[AIndex];
    Result.Value := AValue;
    Result.StdRedefined := True;
  end
  else begin
    for i := Count - 1 to AIndex do
      Add;
    Result := Items[AIndex];
    Result.Value := AValue;
  end;
end;

procedure TXc12NumberFormats.AddDefault(AValue: AxUCString; ALocked: boolean);
var
  Fmt: TXc12NumberFormat;
begin
  Fmt := Add;
  Fmt.Value := AValue;
  Fmt.Locked := ALocked;
end;

function TXc12NumberFormats.BuiltInCount: integer;
begin
  Result := Length(ExcelStandardNumFormats);
end;

function TXc12NumberFormats.Add: TXc12NumberFormat;
begin
  Result := TXc12NumberFormat.Create(Self);
  inherited Add(Result);
end;

procedure TXc12NumberFormats.Clear;
begin
  inherited Clear;
  SetIsDefault;
end;

constructor TXc12NumberFormats.Create;
begin
  inherited Create;
end;

function TXc12NumberFormats.DefaultNumFmt: TXc12NumberFormat;
begin
  Result := Items[0];
end;

function TXc12NumberFormats.Find(ANumFmt: TXc12NumberFormat): TXc12NumberFormat;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Equal(ANumFmt) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12NumberFormats.GetItems(Index: integer): TXc12NumberFormat;
begin
  Result := TXc12NumberFormat(inherited Items[Index]);
end;

procedure TXc12NumberFormats.SetIsDefault;
var
  i: integer;
begin
  for i := 0 to High(ExcelStandardNumFormats) do
    AddDefault(ExcelStandardNumFormats[i],True);
end;

{ TXc12XFs }

function TXc12XFs.Add: TXc12XF;
begin
  Result := TXc12XF.Create(Self);
  inherited Add(Result);
end;

function TXc12XFs.Add(AStyle: TXc12XF): integer;
begin
  inherited Add(AStyle);
  Result := AStyle.Index;
end;

function TXc12XFs.AddExtern(const AStyle: TXc12XF): TXc12XF;
begin
  Result := Add;

  Result.Fill := FOwner.Fills.Find(AStyle.Fill);
  if Result.Fill = Nil then begin
    Result.Fill := FOwner.Fills.Add;
    Result.Fill.Assign(AStyle.Fill);
  end;

  Result.Alignment.Assign(AStyle.Alignment);
  Result.Protection := AStyle.Protection;

  Result.NumFmt  := FOwner.NumFmts.Find(AStyle.NumFmt);
  if Result.NumFmt = Nil then begin
    Result.NumFmt := FOwner.NumFmts.Add;
    Result.NumFmt.Assign(AStyle.NumFmt);
  end;

  Result.Font  := FOwner.Fonts.Find(AStyle.Font);
  if Result.Font = Nil then begin
    Result.Font := FOwner.Fonts.Add;
    Result.Font.Assign(AStyle.Font);
  end;

  Result.Border  := FOwner.Borders.Find(AStyle.Border);
  if Result.Border = Nil then begin
    Result.Border := FOwner.Borders.Add;
    Result.Border.Assign(AStyle.Border);
  end;

  Result.QuotePrefix := AStyle.QuotePrefix;
  Result.PivotButton := AStyle.PivotButton;
end;

function TXc12XFs.AddOrGetFree: TXc12XF;
var
  i: integer;
begin
  for i := XLS_STYLE_DEFAULT_XF_COUNT to Count - 1 do begin
    if (Items[i] = Nil) or (not Items[i].Locked and (Items[i].RefCount <= 0)) then begin
      if Items[i] <> Nil then
        Result := Items[i]
      else begin
         Result := TXc12XF.Create(Self);
        inherited Items[i] := Result;
      end;
      Enumerate;
      FOwner.SetXFDefaultValues(Result);

      Exit;
    end;
  end;

  Result := TXc12XF.Create(Self);
  inherited Add(Result);
end;

procedure TXc12XFs.Clear(AOwner: TXc12DataStyleSheet);
var
  i: integer;
begin
  inherited Clear;

  SetIsDefault;
  for i := 0 to Count - 1 do
    AOwner.SetXFDefaultValues(Items[i]);
end;

function TXc12XFs.CopyAndAdd(AStyle: TXc12XF): TXc12XF;
begin
  Result := AddOrGetFree;

  if Result.FAlignment = Nil then begin
    Result.FAlignment := TXc12CellAlignment.Create;
    Result.FAlignment.Assign(AStyle.FAlignment);
  end;

  Result.FProtection := AStyle.FProtection;

  if FOwner.NumFmts.DefaultNumFmt.Equal(AStyle.FNumFmt) then
    Result.FNumFmt := FOwner.NumFmts.DefaultNumFmt
  else begin
    Result.FNumFmt := FOwner.NumFmts.Add;
    Result.FNumFmt.Assign(AStyle.FNumFmt);
  end;

  if FOwner.Fonts.DefaultFont.Equal(AStyle.FFont) then
    Result.FFont := FOwner.Fonts.DefaultFont
  else begin
    Result.FFont := FOwner.Fonts.Add;
    Result.FFont.Assign(AStyle.FFont);
  end;

  if FOwner.Borders.DefaultBorder.Equal(AStyle.FBorder) then
    Result.FBorder := FOwner.Borders.DefaultBorder
  else begin
    Result.FBorder := FOwner.Borders.Add;
    Result.FBorder.Assign(AStyle.FBorder);
  end;

  if FOwner.Fills.DefaultFill.Equal(AStyle.FFill) then
    Result.FFill := FOwner.Fills.DefaultFill
  else begin
    Result.FFill := FOwner.Fills.Add;
    Result.FFill.Assign(AStyle.FFill);
  end;

  Result.FApply := AStyle.FApply;

  Result.FQuotePrefix := AStyle.FQuotePrefix;
  Result.FPivotButton := AStyle.FPivotButton
end;

constructor TXc12XFs.Create(AOwner: TXc12DataStyleSheet);
begin
  inherited Create;
  FOwner := AOwner;
end;

function TXc12XFs.DefaultXF: TXc12XF;
begin
  Result := Items[XLS_STYLE_DEFAULT_XF];
end;

destructor TXc12XFs.Destroy;
begin
  FIsDestroying := True;
  Clear;
  inherited;
end;

function TXc12XFs.Find(AStyle: TXc12XF): TXc12XF;
{$ifndef STYLE_USE_HASHTABLES}
var
  i: integer;
{$endif}
begin
{$ifdef STYLE_USE_HASHTABLES}
  Result := TXc12XF(inherited Find(AStyle));
{$else}
  for i := 0 to Count - 1 do begin
    if (Items[i] <> Nil) and Items[i].Equal(AStyle) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
{$endif}
end;

procedure TXc12XFs.FreeAndNil(AXF: TXc12XF);
begin
  // List owns objects. Free is called by list.
  inherited Items[AXF.Index] := Nil;
end;

function TXc12XFs.GetItems(Index: integer): TXc12XF;
begin
  Result := TXc12XF(inherited Items[Index]);
end;

procedure TXc12XFs.SetIsDefault;
var
  XF: TXc12XF;
begin
  XF := Add;
  XF.Locked := True;
end;

{ TXc12NumberFormat }

procedure TXc12NumberFormat.Assign(AItem: TXc12NumberFormat);
begin
  if not FLocked then
    FValue := AItem.Value;
end;

function TXc12NumberFormat.Assigned: boolean;
begin
  Result := FValue <> '';
  if Result and (FIndex <= High(ExcelStandardNumFormats)) then
    Result := FValue <> ExcelStandardNumFormats[FIndex];
end;

function TXc12NumberFormat.AsString(const AIndent: integer): AxUCString;
begin
  if FValue <> '' then
    Result := MakeIndent(AIndent) + FValue
  else
    Result := MakeIndent(AIndent) + '@';
end;

procedure TXc12NumberFormat.CalcHash;
begin
  FHash := XLSCalcCRC32(PByteArray(FValue),Length(FValue) * 2);
end;

function TXc12NumberFormat.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := FValue = TXc12NumberFormat(AItem).Value;
end;

function TXc12NumberFormat.GetValue: AxUCString;
begin
  Result := FValue;
end;

function TXc12NumberFormat.IsDate: boolean;
begin
  Result := (FIndex >= 14) and (FIndex <= 17);
end;

function TXc12NumberFormat.IsDateTime: boolean;
begin
  Result := (FIndex >= 14) and (FIndex <= 22);
end;

function TXc12NumberFormat.IsTime: boolean;
begin
  Result := (FIndex >= 18) and (FIndex <= 21);
end;

procedure TXc12NumberFormat.SetValue(const Value: AxUCString);
begin
  FValue := Value;
end;

{ TXc12Font }

procedure TXc12Font.Assign(AItem: TXc12Font);
begin
  FName := AItem.Name;
  FCharset := AItem.Charset;
  FFamily := AItem.Family;
  FStyle := AItem.Style;
  FUnderline := AItem.Underline;
  FSubSuperscript := AItem.SubSuperscript;
  FColor := AItem.Color;
  FSize := AItem.Size;
  FScheme := AItem.Scheme;
end;

function TXc12Font.AsCSS(const AWriteSelector: boolean = True): AxUCString;
var
  TempDS: AxUCChar;
begin
  TempDS := FormatSettings.DecimalSeparator;
  try
    FormatSettings.DecimalSeparator := '.';

    if AWriteSelector then
      Result := CSSSelector + ' {'
    else
      Result := '';

    Result := Result + Format('font-family:%s;',[FName]);
    Result := Result + Format('font-size:%.1fpt;',[FSize]);
    Result := Result + Format('color:#%.6x;',[Xc12ColorToRGB(FColor)]);

    if xfsBold in FStyle then
      Result := Result + 'font-weight:bold;';
    if xfsItalic in FStyle then
      Result := Result + 'font-style:italic;';

    if AWriteSelector then
      Result := Result + '}';
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

procedure TXc12Font.Assign(AItem: TFont);
begin
{$ifdef XLS_XSS}
  FName := AItem.Name;
  FCharset := AItem.Charset;

  FStyle := [];
  if fsBold in AItem.Style then
    FStyle := FStyle + [xfsBold];
  if fsItalic in AItem.Style then
    FStyle := FStyle + [xfsItalic];
  if fsStrikeOut in AItem.Style then
    FStyle := FStyle + [xfsStrikeOut];

  if fsUnderline in AItem.Style then
    FUnderline := xulSingle
  else
    FUnderline := xulNone;

  FSubSuperscript := xssNone;

  FColor := RGBColorToXc12(AItem.Color);

  FScheme := efsNone;

  FSize := AItem.Size;

  FCharset := 0;
  FFamily := 0;
{$endif}
end;

procedure TXc12Font.Assign(ARPr: TCT_TextCharacterProperties);
//var
//  Fill: TXLSDrwColor;
begin
  if ARPr <> Nil then begin
    if ARPr.Sz < 100000 then
      FSize := ARPr.Sz / 100;
    if ARPr.B then
      FStyle := FStyle + [xfsBold];
    if ARPr.I then
      FStyle := FStyle + [xfsItalic];

    case ARPr.U of
      sttutNone       : ;
      sttutDbl        : FUnderline := xulDouble;
      sttutSng        : FUnderline := xulSingle;
    end;

    if (ARPr.Latin <> Nil) and (ARPr.Latin.Typeface <> '') then
      FName := ARPr.Latin.Typeface;

//     if ARPr.FillProperties.SolidFill <> Nil then begin
//       Fill := TXLSDrwColor.Create(ARPr.FillProperties.SolidFill.ColorChoice);
//       try
//         AFont.Color := RGBColorToXc12(RevRGB(Fill.AsRGB));
//       finally
//         Fill.Free;
//       end;
//     end;
  end;

end;

function TXc12Font.AsString(const AIndent: integer): AxUCString;
begin
  Result := MakeIndent(AIndent) + 'Name   : ' + FName + #13 +
            MakeIndent(AIndent) + 'Size   : ' + FloatToStr(FSize) + #13;
  if FCharSet <> 0 then
    Result := Result + MakeIndent(AIndent) + 'CharSet: ' + IntToStr(FCharSet) + #13;
  if FFamily <> 0 then
    Result := Result + MakeIndent(AIndent) + 'Family : ' + IntToStr(FFamily) + #13;

  if FStyle <> [] then begin
    Result := Result + MakeIndent(AIndent) + 'Style  :' + #13;
    if xfsBold in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'Bold' + #13;
    if xfsItalic in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'Italic' + #13;
    if xfsStrikeOut in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'StrikeOut' + #13;
    if xfsOutline in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'Outline' + #13;
    if xfsShadow in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'Shadow' + #13;
    if xfsCondense in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'Condense' + #13;
    if xfsExtend in FStyle then
      Result := Result + MakeIndent(AIndent + 1) + 'Extend' + #13;
  end;
  case FUnderline of
    xulNone         : ;
    xulSingle       : Result := Result + MakeIndent(AIndent) + 'Underline: Single' + #13;
    xulDouble       : Result := Result + MakeIndent(AIndent) + 'Underline: Double' + #13;
    xulSingleAccount: Result := Result + MakeIndent(AIndent) + 'Underline: SingleAccount' + #13;
    xulDoubleAccount: Result := Result + MakeIndent(AIndent) + 'Underline: DoubleAccount' + #13;
  end;
  case FSubSuperscript of
    xssNone       : ;
    xssSuperscript: Result := Result + MakeIndent(AIndent) + 'SubSuperscript: Superscript' + #13;
    xssSubscript  : Result := Result + MakeIndent(AIndent) + 'SubSuperscript: Subscript' + #13;
  end;
  case FScheme of
    efsNone : ;
    efsMajor: Result := Result + MakeIndent(AIndent) + 'Scheme : Major' + #13;
    efsMinor: Result := Result + MakeIndent(AIndent) + 'Scheme : Minor' + #13;
  end;
  Result := Result + MakeIndent(AIndent) + 'Color  :' + #13 + Xc12ColorAsString(AIndent + 1,FColor);
end;

procedure TXc12Font.Clear;
begin
  FName := '';
  FCharset := 0;
  FFamily := 0;
  FStyle := [];
  FUnderline := xulNone;
  FSubSuperscript := xssNone;
  FColor.ColorType := exctAuto;
  FSize := 0;
  FScheme := efsNone;
end;

{$ifdef XLS_XSS}
procedure TXc12Font.CopyToLOGFONT(var Dest: LOGFONT; var FontColor: longword);
var
  i: integer;
begin
  FillChar(Dest,SizeOf(LOGFONT),#0);
  Dest.lfCharSet := Charset;
  for i := 1 to Length(FName) do
    Dest.lfFaceName[i - 1] := Char(FName[i]);
  for i := Length(FName) to 31 do
    Dest.lfFaceName[i] := #0;
  // Height is not correct. Shall be in pixels.
  Dest.lfHeight := Round(Size * 20);
  Dest.lfItalic := Byte(xfsItalic in Style);
  Dest.lfUnderline := Byte(Underline = xulSingle);
  if xfsBold in Style then
    Dest.lfWeight := 700
  else
    Dest.lfWeight := 400;
  Dest.lfStrikeOut := Byte(xfsStrikeOut in Style);
  FontColor := FColor.ARGB;
end;
{$endif}

procedure TXc12Font.CopyToTFont(ADest: TFont);
begin
{$ifdef XLS_XSS}
  ADest.Charset := FCharset;
  ADest.Name := FName;
  ADest.Size := Round(Size);
  ADest.Style := [];
  if Underline <> xulNone then
    ADest.Style := [fsUnderline];
  if xfsBold in Style then
    ADest.Style := ADest.Style + [fsBold];
  if xfsItalic in Style then
    ADest.Style := ADest.Style + [fsItalic];
  if xfsStrikeOut in Style then
    ADest.Style := ADest.Style + [fsStrikeOut];
  ADest.Color := RevRGB(FColor.ARGB);
{$endif}
end;

function TXc12Font.CSSSelector: AxUCString;
begin
  Result := 'Font' + IntToStr(Index + 1);
end;

procedure TXc12Font.DeleteHandle;
begin
{$ifdef XLS_XSS}
  if FHandle <> 0 then
    DeleteObject(FHandle);
  FHandle := 0;
{$endif}
end;

destructor TXc12Font.Destroy;
begin
{$ifdef XLS_XSS}
  DeleteHandle;
{$endif}
  inherited;
end;

function TXc12Font.EqualTFont(AItem: TFont): boolean;
var
  FS: TFontStyles;
begin
  FS := [];

  if (xfsShadow in FStyle) or (xfsCondense in FStyle) or (xfsExtend in FSTyle) then begin
    Result := False;
    Exit;
  end;

  if not ((FUnderline = xulNone) or (FUnderline = xulSingle)) then begin
    Result := False;
    Exit;
  end;

  if FUnderline = xulSingle then
    FS := FS + [{$ifdef BABOON}TFontStyle.{$endif}fsUnderline];
  if xfsBold in FStyle then
    FS := FS + [{$ifdef BABOON}TFontStyle.{$endif}fsBold];
  if xfsItalic in FStyle then
    FS := FS + [{$ifdef BABOON}TFontStyle.{$endif}fsItalic];
  if xfsStrikeOut in FStyle then
    FS := FS + [{$ifdef BABOON}TFontStyle.{$endif}fsStrikeOut];

{$ifdef BABOON}
  Result := (AnsiUppercase(FName) = AnsiUppercase(AItem.Family)) and
{$else}
  Result := (AnsiUppercase(FName) = AnsiUppercase(AItem.Name)) and
            (FCharset = AItem.Charset) and
{$endif}
            (FS = AItem.Style) and
            (FSubSuperscript = xssNone) and
{$ifndef BABOON}
            (FColor.ColorType = exctRgb) and (FColor.ARGB = TColorToRGB(AItem.Color)) and
{$endif}
            (FSize = AItem.Size);
end;

procedure TXc12Font.Use;
begin
  IncRef;
end;

function TXc12Font.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := (AnsiUppercase(FName) = AnsiUppercase(TXc12Font(AItem).Name)) and
            (FCharset = TXc12Font(AItem).Charset) and
            (FFamily = TXc12Font(AItem).Family) and
            (FStyle = TXc12Font(AItem).Style) and
            (FUnderline = TXc12Font(AItem).Underline) and
            (FSubSuperscript = TXc12Font(AItem).SubSuperscript) and
            Xc12ColorEqual(FColor,TXc12Font(AItem).Color) and
            (FSize = TXc12Font(AItem).Size) and
            (FScheme = TXc12Font(AItem).Scheme);

//  Result := (AnsiUppercase(FName) = AnsiUppercase(TXc12Font(AItem).Name)) and
////            (FStyle = TXc12Font(AItem).Style) and
////            (FUnderline = TXc12Font(AItem).Underline) and
////            (FSubSuperscript = TXc12Font(AItem).SubSuperscript) and
////            Xc12ColorEqual(FColor,TXc12Font(AItem).Color) and
//            (FSize = TXc12Font(AItem).Size);
end;

function TXc12Font.GetCharset: integer;
begin
  Result := FCharset;
end;

function TXc12Font.GetColor: TXc12Color;
begin
  Result := FColor;
end;

function TXc12Font.GetFamily: integer;
begin
  Result := FFamily;
end;

{$ifdef XLS_XSS}
function TXc12Font.GetHandle(APixelsPerInch: integer): integer;
var
  Cl: longword;
begin
  if FHandle = 0 then begin
    CopyToLOGFONT(FLF,Cl);
//    FLF.lfHeight := -Round((FLF.lfHeight / 20) * PixelsPerInch / 72);
//    FLF.lfQuality  := NONANTIALIASED_QUALITY;
    if APixelsPerInch > 0 then
      FLF.lfHeight := -MulDiv(FLF.lfHeight div 20,APixelsPerInch,72);
    FLF.lfOutPrecision := OUT_TT_ONLY_PRECIS;
    FHandle := CreateFontIndirect(FLF);
    if FHandle = 0 then
      FHandle := GetStockObject(SYSTEM_FONT);
  end;
  Result := FHandle;
end;
{$endif}

function TXc12Font.GetName: AxUCString;
begin
  Result := FName;
end;

function TXc12Font.GetPColor: PXc12Color;
begin
  Result := @FColor;
end;

function TXc12Font.GetScheme: TXc12FontScheme;
begin
  Result := FScheme;
end;

function TXc12Font.GetSize: double;
begin
  Result := FSize;
end;

function TXc12Font.GetStyle: TXc12FontStyles;
begin
  Result := FStyle;
end;

function TXc12Font.GetSubSuperscript: TXc12SubSuperscript;
begin
  Result := FSubSuperscript;
end;

function TXc12Font.GetUnderline: TXc12Underline;
begin
  Result := FUnderline;
end;

procedure TXc12Font.CalcHash;
var
  i: integer;
begin
  FHash := 0;

  for i := 1 to Length(FName) do
    Inc(FHash,Word(FName[i]));
  Inc(FHash,FCharset);
  Inc(FHash,FFamily);
  Inc(FHash,Byte(FStyle));
  Inc(FHash,Integer(FUnderline));
  Inc(FHash,Integer(FSubSuperscript));
  Inc(FHash,Xc12ColorHash(FColor));
  Inc(FHash,Round(FSize * 100));
  Inc(FHash,Integer(FScheme));
end;

procedure TXc12Font.SetCharset(const Value: integer);
begin
  FCharset := Value;
  FHashvalid := False;
end;

procedure TXc12Font.SetColor(const Value: TXc12Color);
begin
  FColor := Value;
  FHashvalid := False;
end;

procedure TXc12Font.SetFamily(const Value: integer);
begin
  FFamily := Value;
  FHashvalid := False;
  DeleteHandle;
end;

procedure TXc12Font.SetName(const Value: AxUCString);
begin
  FName := Value;
  FHashvalid := False;
  DeleteHandle;
end;

procedure TXc12Font.SetScheme(const Value: TXc12FontScheme);
begin
  FScheme := Value;
  FHashvalid := False;
end;

procedure TXc12Font.SetSize(const Value: double);
begin
  DeleteHandle;
  FSize := Value;
{$ifdef XLS_XSS}
  if (FOwner <> Nil) and (FIndex = 0) and (Value > 0) then
    TXc12Fonts(FOwner).CalcStdFontWidthHeight;
{$endif}
  FHashvalid := False;
end;

procedure TXc12Font.SetStyle(const Value: TXc12FontStyles);
begin
  FStyle := Value;
  FHashvalid := False;
  DeleteHandle;
end;

procedure TXc12Font.SetSubSuperscript(const Value: TXc12SubSuperscript);
begin
  FSubSuperscript := Value;
  FHashvalid := False;
  DeleteHandle;
end;

procedure TXc12Font.SetUnderline(const Value: TXc12Underline);
begin
  FUnderline := Value;
  FHashvalid := False;
end;

{$ifdef XLS_XSS}
function TXc12Font.Width: integer;
var
  i: integer;
  Sz: TSize;
  DC: HDC;
begin
  Result := 0;
  DC := CreateIC('DISPLAY',Nil,Nil,Nil);
  try
    SelectObject(DC,GetHandle(GetDeviceCaps(DC,LOGPIXELSY)));
    for i := 1 to 10 do begin
      Windows.GetTextExtentPoint32W(DC,LPCWSTR(@StrDigits[i]),1,Sz);
      if Sz.cy > Result then
        Result := Sz.cx;
    end;
  finally
    DeleteDC(DC);
  end;
end;
{$endif}

{ TXc12Border }

procedure TXc12Border.Assign(AItem: TXc12Border);
begin
  FOptions := AItem.Options;
  FLeft.Assign(AItem.FLeft);
  FTop.Assign(AItem.FTop);
  FRight.Assign(AItem.FRight);
  FBottom.Assign(AItem.FBottom);
  FDiagonal.Assign(AItem.FDiagonal);
  FHorizontal.Assign(AItem.FHorizontal);
  FVertical.Assign(AItem.FVertical);
end;

function TXc12Border.AsString(const AIndent: integer): AxUCString;
begin
  Result := '';
  if FOptions <> [] then begin
    Result := Result + MakeIndent(AIndent) + 'Options: ' + #13;
    if ecboDiagonalDown in FOptions then
      Result := Result + MakeIndent(AIndent + 1) + 'DiagonalDown' + #13;
    if ecboDiagonalUp in FOptions then
      Result := Result + MakeIndent(AIndent + 1) + 'DiagonalUp' + #13;
    if ecboOutline in FOptions then
      Result := Result + MakeIndent(AIndent + 1) + 'Outline' + #13;
  end;
  if FLeft.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Left: ' + #13 + AsStringBorder(AIndent + 1,FLeft);
  if FTop.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Top: ' + #13 + AsStringBorder(AIndent + 1,FTop);
  if FRight.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Right: ' + #13 + AsStringBorder(AIndent + 1,FRight);
  if FBottom.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Bottom: ' + #13 + AsStringBorder(AIndent + 1,FBottom);
  if FDiagonal.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Diagonal: ' + #13 + AsStringBorder(AIndent + 1,FDiagonal);
  if FHorizontal.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Horizontal: ' + #13 + AsStringBorder(AIndent + 1,FHorizontal);
  if FVertical.Style <> cbsNone then
    Result := Result + MakeIndent(AIndent) + 'Vertical: ' + #13 + AsStringBorder(AIndent + 1,FVertical);
end;

function TXc12Border.AsStringBorder(const AIndent: integer; const ABorder: TXc12BorderPr): AxUCString;
var
  S: AxUCString;
begin
  case ABorder.Style of
    cbsNone            : S := 'None';
    cbsThin            : S := 'Thin';
    cbsMedium          : S := 'Medium';
    cbsDashed          : S := 'Dashed';
    cbsDotted          : S := 'Dotted';
    cbsThick           : S := 'Thick';
    cbsDouble          : S := 'Double';
    cbsHair            : S := 'Hair';
    cbsMediumDashed    : S := 'MediumDashed';
    cbsDashDot         : S := 'DashDot';
    cbsMediumDashDot   : S := 'MediumDashDot';
    cbsDashDotDot      : S := 'DashDotDot';
    cbsMediumDashDotDot: S := 'MediumDashDotDot';
    cbsSlantedDashDot  : S := 'SlantedDashDot';
  end;
  Result := Result + MakeIndent(AIndent) + 'Style: ' + S + #13;
  Result := Result + MakeIndent(AIndent) + 'Color: ' + #13 + Xc12ColorAsString(AIndent + 1,ABorder.Color);
end;

procedure TXc12Border.CalcHash;
begin
  FHash := Byte(FOptions) xor FLeft.CalcHash xor FTop.CalcHash xor FRight.CalcHash xor
           FBottom.CalcHash xor FDiagonal.CalcHash xor FHorizontal.CalcHash xor FVertical.CalcHash;
end;

procedure TXc12Border.Clear;
begin
  FOptions := [];
  FLeft.Clear;
  FTop.Clear;
  FRight.Clear;
  FBottom.Clear;
  FDiagonal.Clear;
  FHorizontal.Clear;
  FVertical.Clear;
end;

constructor TXc12Border.Create(AOwner: TXLSStyleObjectList);
begin
  inherited Create(AOwner);

  FLeft := TXc12BorderPr.Create;
  FTop := TXc12BorderPr.Create;
  FRight := TXc12BorderPr.Create;
  FBottom := TXc12BorderPr.Create;
  FDiagonal := TXc12BorderPr.Create;
  FHorizontal := TXc12BorderPr.Create;
  FVertical := TXc12BorderPr.Create;
end;

destructor TXc12Border.Destroy;
begin
  FLeft.Free;
  FTop.Free;
  FRight.Free;
  FBottom.Free;
  FDiagonal.Free;
  FHorizontal.Free;
  FVertical.Free;

  inherited;
end;

function TXc12Border.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := (FOptions = TXc12Border(AItem).Options) and
             FLeft.Equal(TXc12Border(AItem).FLeft) and
             FTop.Equal(TXc12Border(AItem).FTop) and
             FRight.Equal(TXc12Border(AItem).FRight) and
             FBottom.Equal(TXc12Border(AItem).FBottom) and
             FDiagonal.Equal(TXc12Border(AItem).FDiagonal) and
             FHorizontal.Equal(TXc12Border(AItem).FHorizontal) and
             FVertical.Equal(TXc12Border(AItem).FVertical);
end;

function TXc12Border.GetOptions: TXc12CellBorderOptions;
begin
  Result := FOptions;
end;

procedure TXc12Border.SetOptions(const Value: TXc12CellBorderOptions);
begin
  FOptions := Value;
end;

{ TXc12ColorObj }

procedure TXc12ColorObj.Assign(AItem: TXc12ColorObj);
begin
  FColor := AItem.Color;
end;

procedure TXc12ColorObj.CalcHash;
begin
  FHash := Xc12ColorHash(FColor);
end;

function TXc12ColorObj.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := Xc12ColorEqual(FColor,TXc12ColorObj(AItem).Color);
end;

function TXc12ColorObj.GetColor: TXc12Color;
begin
  Result := FColor;
end;

{ TXc12IndexColorObj }

procedure TXc12IndexColorObj.Assign(AItem: TXc12IndexColorObj);
begin
  FRGB := AItem.RGB;
end;

procedure TXc12IndexColorObj.CalcHash;
begin
  FHash := FRGB;
end;

function TXc12IndexColorObj.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := FRGB = TXc12IndexColorObj(AItem).RGB;
end;

function TXc12IndexColorObj.GetRGB: longword;
begin
  Result := FRGB;
end;

procedure TXc12IndexColorObj.SetRGB(const Value: longword);
begin
  FRGB := Value;
end;

{ TXc12CellAlignent }

procedure TXc12CellAlignment.Assign(AObject: TXc12CellAlignment);
begin
  FHorizAlignment := AObject.FHorizAlignment;
  FVertAlignment := AObject.FVertAlignment;
  FRotation := AObject.FRotation;
  FOptions := AObject.FOptions;
  FIndent := AObject.FIndent;
  FRelativeIndent := AObject.FRelativeIndent;
  FReadingOrder := AObject.FReadingOrder;
end;

function TXc12CellAlignment.AsString: AxUCString;
begin
  Result := '';
end;

procedure TXc12CellAlignment.CalcHash;
begin
  FHash := Longword(FHorizAlignment);
  Inc(FHash,Longword(FVertAlignment));
  Inc(FHash,FRotation);
  Inc(FHash,Longword(Byte(FOptions)));
  Inc(FHash,FIndent);
  Inc(FHash,FRelativeIndent);
  Inc(FHash,Longword(FReadingOrder));

  FHashValid := True;
end;

constructor TXc12CellAlignment.Create;
begin
  FVertAlignment := cvaBottom;
end;

function TXc12CellAlignment.Equal(AObject: TXc12CellAlignment): boolean;
begin
  Result := (FHorizAlignment = AObject.FHorizAlignment) and
            (FVertAlignment = AObject.FVertAlignment) and
            (FRotation = AObject.FRotation) and
            (FOptions = AObject.FOptions) and
            (FIndent = AObject.FIndent) and
            (FRelativeIndent = AObject.FRelativeIndent) and
            (FReadingOrder = AObject.FReadingOrder);
end;

function TXc12CellAlignment.GetHorizAlignment: TXc12HorizAlignment;
begin
  Result := FHorizAlignment;
end;

function TXc12CellAlignment.GetIndent: integer;
begin
  Result := FIndent;
end;

function TXc12CellAlignment.GetOptions: TXc12AlignmentOptions;
begin
  Result := FOptions;
end;

function TXc12CellAlignment.GetReadingOrder: TXc12ReadOrder;
begin
  Result := FReadingOrder;
end;

function TXc12CellAlignment.GetRelativeIndent: integer;
begin
  Result := FRelativeIndent;
end;

function TXc12CellAlignment.GetRotation: integer;
begin
  Result := FRotation;
end;

function TXc12CellAlignment.GetVertAlignment: TXc12VertAlignment;
begin
  Result := FVertAlignment;
end;

function TXc12CellAlignment.HashKey: longword;
begin
  if not FHashValid then
    CalcHash;
  Result := FHash;
end;

function TXc12CellAlignment.IsWrapText: boolean;
begin
  Result := ((foWrapText in FOptions) or (FHorizAlignment in [chaJustify,chaDistributed])) and (FRotation = 0);
end;

procedure TXc12CellAlignment.SetHorizAlignment(const Value: TXc12HorizAlignment);
begin
  FHorizAlignment := Value;
end;

procedure TXc12CellAlignment.SetIndent(const Value: integer);
begin
  FIndent := Value;
end;

procedure TXc12CellAlignment.SetOptions(const Value: TXc12AlignmentOptions);
begin
  FOptions := Value;
end;

procedure TXc12CellAlignment.SetReadingOrder(const Value: TXc12ReadOrder);
begin
  FReadingOrder := Value;
end;

procedure TXc12CellAlignment.SetRelativeIndent(const Value: integer);
begin
  FRelativeIndent := Value;
end;

procedure TXc12CellAlignment.SetRotation(const Value: integer);
begin
  FRotation := Value;
end;

procedure TXc12CellAlignment.SetVertAlignment(const Value: TXc12VertAlignment);
begin
  FVertAlignment := Value;
end;

{ TXc12DXF }


function TXc12DXF.AddAlignment: TXc12CellAlignment;
begin
  if FAlignment = Nil then
    FAlignment := TXc12CellAlignment.Create;
  Result := FAlignment;
end;

function TXc12DXF.AddBorder: TXc12Border;
begin
  if FBorder = Nil then
    FBorder := TXc12Border.Create(Nil);
  Result := FBorder;
end;

function TXc12DXF.AddFill: TXc12Fill;
begin
  if FFill = Nil then
    FFill := TXc12Fill.Create(Nil);
  Result := FFill;
end;

function TXc12DXF.AddFont: TXc12Font;
begin
  if FFont = Nil then
    FFont := TXc12Font.Create(Nil);
  Result := FFont;
end;

function TXc12DXF.AddNumFmt: TXc12NumberFormat;
begin
  if FNumFmt = Nil then
    FNumFmt := TXc12NumberFormat.Create(Nil);
  Result := FNumFmt;
end;

procedure TXc12DXF.Assign(AItem: TXc12DXF);
begin
  if AItem.Font <> Nil then
    AddFont.Assign(AItem.Font);

  if AItem.FNumFmt <> Nil then
    AddNumFmt.Assign(AItem.FNumFmt);

  if AItem.FFill <> Nil then
    AddFill.Assign(AItem.FFill);

  if AItem.FBorder <> Nil then
    AddBorder.Assign(AItem.FBorder);

  if AItem.FAlignment <> Nil then
    AddAlignment.Assign(AItem.FAlignment);

  FProtection := AItem.FProtection;
end;

procedure TXc12DXF.CalcHash;
begin
  FHash := 0;
  if FFont <> Nil then begin
    FFont.CalcHash;
    FHash := FHash xor FFont.FHash;
  end;
  if FNumFmt <> Nil then begin
    FNumFmt.CalcHash;
    FHash := FHash xor FNumFmt.FHash;
  end;
  if FFill <> Nil then begin
    FFill.CalcHash;
    FHash := FHash xor FFill.FHash;
  end;
  if FBorder <> Nil then begin
    FBorder.CalcHash;
    FHash := FHash xor FBorder.FHash;
  end;
  if FAlignment <> Nil then begin
    FAlignment.CalcHash;
    FHash := FHash xor FAlignment.FHash;
  end;
  FHash := FHash + Byte(FProtection);
end;

destructor TXc12DXF.Destroy;
begin
  if FFont <> Nil then
    FFont.Free;
  if FNumFmt <> Nil then
    FNumFmt.Free;
  if FFill <> Nil then
    FFill.Free;
  if FBorder <> Nil then
    FBorder.Free;
  if FAlignment <> Nil then
    FAlignment.Free;
end;

function TXc12DXF.Equal(AItem: TXLSStyleObject): boolean;
begin
  raise XLSRWException.Create('TODO');
end;

function TXc12DXF.GetAlignment: TXc12CellAlignment;
begin
  Result := FAlignment;
end;

function TXc12DXF.GetBorder: TXc12Border;
begin
  Result := FBorder;
end;

function TXc12DXF.GetFill: TXc12Fill;
begin
  Result := FFill;
end;

function TXc12DXF.GetFont: TXc12Font;
begin
  Result := FFont;
end;

function TXc12DXF.GetNumFmt: TXc12NumberFormat;
begin
  Result := FNumFmt;
end;

function TXc12DXF.GetProtection: TXc12CellProtections;
begin
  Result := FProtection;
end;

{ TXc12CellStyle }

procedure TXc12CellStyle.Assign(AItem: TXc12CellStyle);
begin
  FBuiltInId := AItem.BuiltInId;
  FCustomBuiltIn := AItem.CustomBuiltIn;
  FHidden := AItem.Hidden;
  FLevel := AItem.Level;
  FName := AItem.Name;
  FXF := AItem.XF;
end;

procedure TXc12CellStyle.CalcHash;
begin
  FHash := Longword(FBuiltInId) +
           Longword(FCustomBuiltIn) +
           Longword(FHidden) +
           Longword(FLevel) +
           XLSCalcCRC32(PByteArray(FName),Length(FName) * 2) xor
           FXF.FHash;
end;

function TXc12CellStyle.Equal(AItem: TXLSStyleObject): boolean;
begin
  Result := (FBuiltInId = TXc12CellStyle(AItem).BuiltInId) and
            (FCustomBuiltIn = TXc12CellStyle(AItem).CustomBuiltIn) and
            (FHidden = TXc12CellStyle(AItem).Hidden) and
            (FLevel = TXc12CellStyle(AItem).Level) and
            (FName = TXc12CellStyle(AItem).Name) and
            FXF.Equal(TXc12CellStyle(AItem).XF);
end;

function TXc12CellStyle.GetBuiltInId: integer;
begin
  Result := FBuiltInId;
end;

function TXc12CellStyle.GetCustomBuiltIn: boolean;
begin
  Result := FCustomBuiltIn;
end;

function TXc12CellStyle.GetHidden: boolean;
begin
  Result := FHidden;
end;

function TXc12CellStyle.GetLevel: integer;
begin
  Result := FLevel;
end;

function TXc12CellStyle.GetName: AxUCString;
begin
  Result := FName;
end;

function TXc12CellStyle.GetXF: TXc12XF;
begin
  Result := FXF
end;

procedure TXc12CellStyle.SetBuiltInId(const Value: integer);
begin
  FBuiltInId := Value;
end;

procedure TXc12CellStyle.SetCustomBuiltIn(const Value: boolean);
begin
  FCustomBuiltIn := Value;
end;

procedure TXc12CellStyle.SetHidden(const Value: boolean);
begin
  FHidden := Value;
end;

procedure TXc12CellStyle.SetLevel(const Value: integer);
begin
  FLevel := Value;
end;

procedure TXc12CellStyle.SetName(const Value: AxUCString);
begin
  FName := Value;
end;

procedure TXc12CellStyle.SetXF(const Value: TXc12XF);
begin
  FXF := Value;
end;

{ TXc12TableStyleElements }

function TXc12TableStyleElements.Add: TXc12TableStyleElement;
var
  Item: TXc12TableStyleElement;
begin
  Item := TXc12TableStyleElement.Create;
  Result := Item;
  TObjectList(Self).Add(Item);
end;

constructor TXc12TableStyleElements.Create;
begin
  inherited Create;
end;

function TXc12TableStyleElements.GetItems(Index: integer): TXc12TableStyleElement;
begin
  Result := TXc12TableStyleElement(inherited Items[Index]);
end;

{ TXc12_TableStyle }

procedure TXc12TableStyle.Assign(AItem: TXc12TableStyle);
begin
  raise XLSRWException.Create('TODO');
end;

procedure TXc12TableStyle.CalcHash;
begin
  raise XLSRWException.Create('TODO');
end;

procedure TXc12TableStyle.Clear;
begin
  FTableStyleElements.Clear;
  FName := '';
  FPivot := True;
  FTable := True;
end;

constructor TXc12TableStyle.Create;
begin
  FTableStyleElements := TXc12TableStyleElements.Create;
  Clear;
end;

destructor TXc12TableStyle.Destroy;
begin
  FTableStyleElements.Free;
  inherited;
end;

function TXc12TableStyle.Equal(AItem: TXLSStyleObject): boolean;
begin
  raise XLSRWException.Create('TODO');
end;

function TXc12TableStyle.GetName: AxUCString;
begin
  Result := FName;
end;

function TXc12TableStyle.GetPivot: boolean;
begin
  Result := FPivot;
end;

function TXc12TableStyle.GetTable: boolean;
begin
  Result := FTable;
end;

function TXc12TableStyle.GetTableStyleElements: TXc12TableStyleElements;
begin
  Result := FTableStyleElements;
end;

procedure TXc12TableStyle.SetName(const Value: AxUCString);
begin
  FName := Value;
end;

procedure TXc12TableStyle.SetPivot(const Value: boolean);
begin
  FPivot := Value;
end;

procedure TXc12TableStyle.SetTable(const Value: boolean);
begin
  FTable := Value;
end;

{ TXc12TableStyles }

function TXc12TableStyles.Add: TXc12TableStyle;
begin
  Result := TXc12TableStyle.Create;
  inherited Add(Result);
end;

procedure TXc12TableStyles.Clear;
begin
  inherited Clear;
  FDefaultTableStyle := '';
  FDefaultPivotStyle := '';
end;

constructor TXc12TableStyles.Create;
begin
  inherited Create;
end;

function TXc12TableStyles.GetItems(Index: integer): TXc12TableStyle;
begin
  Result := TXc12TableStyle(inherited Items[Index]);
end;

procedure TXc12TableStyles.SetIsDefault;
begin
  inherited;

end;

{ TXc12TableStyleElement }

procedure TXc12TableStyleElement.Assign(AItem: TXc12TableStyleElement);
begin
  raise XLSRWException.Create('TODO');
end;

procedure TXc12TableStyleElement.Clear;
begin
  FType := TXc12TableStyleType(XPG_UNKNOWN_ENUM);
  FSize := 1;
  FDXF := Nil;
end;

constructor TXc12TableStyleElement.Create;
begin
  Clear;
end;

function TXc12TableStyleElement.Equal(AItem: TXc12TableStyleElement): boolean;
begin
  raise XLSRWException.Create('TODO');
end;

function TXc12TableStyleElement.GetDXF: TXc12DXF;
begin
  Result := FDXF;
end;

function TXc12TableStyleElement.GetSize: integer;
begin
  Result := FSize;
end;

function TXc12TableStyleElement.GetType: TXc12TableStyleType;
begin
  Result := FType;
end;

procedure TXc12TableStyleElement.SetDXf(const Value: TXc12DXF);
begin
  FDXF := Value;
end;

procedure TXc12TableStyleElement.SetSize(const Value: integer);
begin
  FSize := Value;
end;

procedure TXc12TableStyleElement.SetType(const Value: TXc12TableStyleType);
begin
  FType := Value;
end;

{ TXc12GradientStop }

function TXc12GradientStop.CalcHash: longword;
begin
  Result := Xc12ColorHash(FColor) xor Longword(PIntegerArray(@FPosition)[0]) xor Longword(PIntegerArray(@FPosition)[1]);
end;

function TXc12GradientStop.GetColor: TXc12Color;
begin
  Result := FColor;
end;

function TXc12GradientStop.GetPosition: double;
begin
  Result := FPosition;
end;

procedure TXc12GradientStop.SetColor(const Value: TXc12Color);
begin
  FColor := Value;
end;

procedure TXc12GradientStop.SetPosition(const Value: double);
begin
  FPosition := Value;
end;

{ TXc12XFEditor }

procedure TXc12XFEditor.BeginEdit(AXF: TXc12XF);
begin
  if AXF <> Nil then begin
    if not AXF.Locked then begin
      // Lock the style in order to prevent it from being deleted in CompactStyles.
      AXF.Locked := True;
      CompactStyles;
      AXF.Locked := False;
    end
    else
      CompactStyles;

    FOrigXF := AXF;
  end
  else
    FOrigXF := FStyles.XFs[XLS_STYLE_DEFAULT_XF];

//  FCmpXF := FOrigXF;
//  FCmpNumFmt := FOrigXF.NumFmt;
//  FCmpFont := FOrigXF.Font;
//  FCmpFill := FOrigXF.Fill;
//  FCmpBorder := FOrigXF.Border;

  FCmpXF := FXF;
  FCmpNumFmt := FXF.NumFmt;
  FCmpFont := FXF.Font;
  FCmpFill := FXF.Fill;
  FCmpBorder := FXF.Border;

  FCmpXF.AssignProps(FOrigXF);
  FCmpNumFmt.Assign(FOrigXF.NumFmt);
  FCmpFont.Assign(FOrigXF.Font);
  FCmpFill.Assign(FOrigXF.Fill);
  FCmpBorder.Assign(FOrigXF.Border);

  FProtectChanged := False;
  FAlignChanged := False;
  FNumFmtChanged := False;
  FFontChanged := False;
  FFillChanged := False;
  FBorderChanged := False;
end;

procedure TXc12XFEditor.CompactStyles;
var
  i: integer;
begin
  Inc(FCompactCount);
  if FCompactCount < XC12_STYLES_COMPACT_CNT then
    Exit;

  FCompactCount := 0;

  // Don't compact number formats.
  // The Id is used in pivot tables and this must then be changed as well.
//  DoCompactStyles(FStyles.NumFmts);
  DoCompactStyles(FStyles.Fonts);
  DoCompactStyles(FStyles.Fills);
  DoCompactStyles(FStyles.Borders);

  i := FStyles.XFs.Count - 1;

  while (FStyles.XFs.Count > 0) and (FStyles.XFs[i] <> Nil) and not FStyles.XFs[i].Locked and (FStyles.XFs[i].RefCount <= 0) do begin
    FStyles.XFs.Delete(i);
    Dec(i);
  end;

  for i := 0 to FStyles.XFs.Count - 1 do begin
    if (FStyles.XFs[i] <> Nil) and (FStyles.XFs[i].RefCount <= 0) and not FStyles.XFs[i].Locked then
      FStyles.XFs.FreeAndNil(FStyles.XFs[i]);
  end;
end;

constructor TXc12XFEditor.Create(AStyleSheet: TXc12DataStyleSheet);
begin
  FStyles := AStyleSheet;

  FXF := TXc12XF.Create(Nil);
  FXF.NumFmt := TXc12NumberFormat.Create(Nil);
  FXF.Font := TXc12Font.Create(Nil);
  FXF.Fill := TXc12Fill.Create(Nil);
  FXF.Border := TXc12Border.Create(Nil);
end;

procedure TXc12XFEditor.DecStyle(const AIndex: integer);
var
  XF: TXc12XF;
begin
  XF := FStyles.XFs[AIndex];

  XF.NumFmt.DecRef;
  XF.Font.DecRef;
  XF.Fill.DecRef;
  XF.Border.DecRef;
  XF.DecRef;
{$ifdef _AXOLOT_DEBUG}
  if (XF.NumFmt.RefCount < 0) or (XF.Font.RefCount < 0) or (XF.Fill.RefCount < 0) or (XF.Border.RefCount < 0) or (XF.RefCount < 0) then
    raise XLSRWException.Create('Ref count < 0');
{$endif}
end;

destructor TXc12XFEditor.Destroy;
begin
  FXF.NumFmt.Free;
  FXF.Font.Free;
  FXF.Fill.Free;
  FXF.Border.Free;
  FXF.Free;
  inherited;
end;

procedure TXc12XFEditor.DoAlignChanged;
begin
  FXF.Alignment.Assign(FCmpXF.Alignment);

  FAlignChanged := True;
end;

procedure TXc12XFEditor.DoBorderChanged;
begin
  FXF.Border.Assign(FCmpBorder);
  FCmpBorder := FXF.Border;

  FBorderChanged := True;
end;

procedure TXc12XFEditor.DoCompactStyles(AStyles: TXLSStyleObjectList);
var
  i: integer;
begin
  i := 0;
  while i < AStyles.Count - 1 do begin
    if (AStyles.StyleItems[i] <> Nil) and not AStyles.StyleItems[i].Locked and (AStyles.StyleItems[i].RefCount <= 0) then
      AStyles.Delete(i)
    else
      Inc(i);
  end;
end;

procedure TXc12XFEditor.DoFillChanged;
begin
  FXF.Fill.Assign(FCmpFill);
  FCmpFill := FXF.Fill;
  FCmpFill.FPatternType := efpSolid;

  FFillChanged := True;
end;

procedure TXc12XFEditor.DoFontChanged;
begin
  FXF.Font.Assign(FCmpFont);
  FCmpFont := FXF.Font;

  FFontChanged := True;
end;

procedure TXc12XFEditor.DoLockStyles(AStyles: TXLSStyleObjectList);
var
  i: integer;
begin
  for i := 0 to AStyles.Count - 1 do
    AStyles.StyleItems[i].Locked := True;
end;

function TXc12XFEditor.EndEdit: TXc12XF;
var
  SearchXF: TXc12XF;
begin
  SearchXF := TXc12XF.Create(Nil);
  try
    SearchXF.AssignProps(FCmpXF);

    SearchXF.NumFmt := FCmpNumFmt;
    SearchXF.Font := FCmpFont;
    SearchXF.Fill := FCmpFill;
    SearchXF.Border := FCmpBorder;

    Result := FStyles.XFs.Find(SearchXF);

    if Result = Nil then begin
      Result := FStyles.XFs.AddOrGetFree;
      Result.AssignProps(FCmpXF);

      if FProtectChanged then
        Result.Apply := Result.Apply + [eafProtection];

      if FAlignChanged then begin
        Result.Alignment.Assign(FCmpXF.Alignment);
        Result.Apply := Result.Apply + [eafAlignment];
      end;

      if FNumFmtChanged then begin
        Result.NumFmt := FStyles.NumFmts.Find(FCmpNumFmt);
        if Result.NumFmt = Nil then begin
          Result.NumFmt := FStyles.NumFmts.Add;
          Result.NumFmt.Assign(FCmpNumFmt);
        end;
        Result.Apply := Result.Apply + [eafNumberFormat];
      end
      else
        Result.NumFmt := FOrigXF.NumFmt;

      if FFontChanged then begin
        Result.Font := FStyles.Fonts.Find(FCmpFont);
        if Result.Font = Nil then begin
          Result.Font := FStyles.Fonts.Add;
          Result.Font.Assign(FCmpFont);
        end;
        Result.Apply := Result.Apply + [eafFont];
      end
      else
        Result.Font := FOrigXF.Font;

      if FFillChanged then begin
        Result.Fill := FStyles.Fills.Find(FCmpFill);
        if Result.Fill = Nil then begin
          Result.Fill := FStyles.Fills.Add;
          Result.Fill.Assign(FCmpFill);
        end;
        Result.Apply := Result.Apply + [eafFill];
      end
      else
        Result.Fill := FOrigXF.Fill;

      if FBorderChanged then begin
        Result.Border := FStyles.Borders.Find(FCmpBorder);
        if Result.Border = Nil then begin
          Result.Border := FStyles.Borders.Add;
          Result.Border.Assign(FCmpBorder);
        end;
        Result.Apply := Result.Apply + [eafBorder];
      end
      else
        Result.Border := FOrigXF.Border;

    end;
  finally
    SearchXF.Free;
  end;
end;

procedure TXc12XFEditor.FreeStyle(const AIndex: integer);
begin
  FreeStyle(FStyles.XFs[AIndex]);
end;

procedure TXc12XFEditor.FreeFont(AFont: TXc12Font);
begin
  AFont.DecRef;
  // Don't test for less than zero. Some Excel 97 files has got RefCount = -1.
  // Probably caused by not incrementing the ref count when the font is used
  // somewhere, possibly charts.
  if not AFont.Locked and (AFont.RefCount = 0) then
    FStyles.Fonts.Delete(AFont.Index);
end;

procedure TXc12XFEditor.FreeFonts(AFonts: TXc12DynFontRunArray);
var
  i: integer;
begin
  for i := 0 to High(AFonts) do begin
    if AFonts[i].Font <> Nil then
      FreeFont(AFonts[i].Font);
  end;
end;

procedure TXc12XFEditor.FreeFonts(AFonts: PXc12FontRunArray; const ACount: integer);
var
  i: integer;
begin
  for i := 0 to ACount - 1 do
    FreeFont(AFonts[i].Font);
end;

procedure TXc12XFEditor.FreeStyle(AXF: TXc12XF);
begin
  // ***************************************************************************
  // ** Do not delete any objects here, that causes all kind of problems when **
  // ** indexes got messed up while a style is edited. Do all deletion in     **
  // ** CompactFormats when it's safe, such as before formatting starts,      **
  // ** XFEditor.BeginEdit().                                                 **
  // ***************************************************************************

  AXF.NumFmt.DecRef;
//  if not AXF.NumFmt.Locked and (AXF.NumFmt.RefCount <= 0) then
//    FStyles.NumFmts.Delete(AXF.NumFmt.Index);

  AXF.Font.DecRef;
//  if not AXF.Font.Locked and (AXF.Font.RefCount <= 0) then
//    FStyles.Fonts.Delete(AXF.Font.Index);

  AXF.Fill.DecRef;
//  if not AXF.Fill.Locked and (AXF.Fill.RefCount <= 0) then
//    FStyles.Fills.Delete(AXF.Fill.Index);

  AXF.Border.DecRef;
//  if not AXF.Border.Locked and (AXF.Border.RefCount <= 0) then
//    FStyles.Borders.Delete(AXF.Border.Index);

  AXF.DecRef;
end;

procedure TXc12XFEditor.LockAll;
begin
  DoLockStyles(FStyles.NumFmts);
  DoLockStyles(FStyles.Fonts);
  DoLockStyles(FStyles.Fills);
  DoLockStyles(FStyles.Borders);
end;

procedure TXc12XFEditor.SetFillColor(const Value: TXc12Color);
begin
  if not FFillChanged and not Xc12ColorEqual(Value,FCmpFill.FgColor) then
    DoFillChanged;
  FCmpFill.FgColor := Value;
end;

procedure TXc12XFEditor.SetFillColorRGB(const Value: TXc12RGBColor);
begin
  if not FFillChanged and (Value <> Xc12ColorToRGB(FCmpFill.FgColor)) then
    DoFillChanged;
  FCmpFill.FgColor := RGBColorToXc12(Value);
end;

procedure TXc12XFEditor.SetAlignHoriz(const Value: TXc12HorizAlignment);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.HorizAlignment) then
    DoAlignChanged;
  FCmpXF.Alignment.HorizAlignment := Value;
end;

procedure TXc12XFEditor.SetAlignIndent(const Value: integer);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.Indent) then
    DoAlignChanged;
  FCmpXF.Alignment.Indent := Value;
end;

procedure TXc12XFEditor.SetAlignOptions(const Value: TXc12AlignmentOptions);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.Options) then
    DoAlignChanged;
  FCmpXF.Alignment.Options := Value;
end;

procedure TXc12XFEditor.SetAlignReadingOrder(const Value: TXc12ReadOrder);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.ReadingOrder) then
    DoAlignChanged;
  FCmpXF.Alignment.ReadingOrder := Value;
end;

procedure TXc12XFEditor.SetAlignRelativeIndent(const Value: integer);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.RelativeIndent) then
    DoAlignChanged;
  FCmpXF.Alignment.RelativeIndent := Value;
end;

procedure TXc12XFEditor.SetAlignRotation(const Value: integer);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.Rotation) then
    DoAlignChanged;
  FCmpXF.Alignment.Rotation := Value;
end;

procedure TXc12XFEditor.SetAlignVert(const Value: TXc12VertAlignment);
begin
  if not FAlignChanged and (Value <> FCmpXF.Alignment.VertAlignment) then
    DoAlignChanged;
  FCmpXF.Alignment.VertAlignment := Value;
end;

procedure TXc12XFEditor.SetBorderBottomColor(const Value: TXc12Color);
begin
  if not FBorderChanged and not Xc12ColorEqual(Value,FCmpBorder.FBottom.Color) then
    DoBorderChanged;
  FCmpBorder.FBottom.Color := Value;
end;

procedure TXc12XFEditor.SetBorderBottomColorRGB(const Value: TXc12RGBColor);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FBottom.ColorRGB) then
    DoBorderChanged;
  FCmpBorder.FBottom.ColorRGB := Value;
end;

procedure TXc12XFEditor.SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FBottom.Style) then
    DoBorderChanged;
  FCmpBorder.FBottom.Style := Value;
end;

procedure TXc12XFEditor.SetBorderDiagColor(const Value: TXc12Color);
begin
  if not FBorderChanged and not Xc12ColorEqual(Value,FCmpBorder.FDiagonal.Color) then
    DoBorderChanged;
  FCmpBorder.FDiagonal.Color := Value;
end;

procedure TXc12XFEditor.SetBorderDiagColorRGB(const Value: TXc12RGBColor);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FDiagonal.ColorRGB) then
    DoBorderChanged;
  FCmpBorder.FDiagonal.ColorRGB := Value;
end;

procedure TXc12XFEditor.SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FDiagonal.Style) then
    DoBorderChanged;
  FCmpBorder.FDiagonal.Style := Value;
end;

procedure TXc12XFEditor.SetBorderLeftColor(const Value: TXc12Color);
begin
  if not FBorderChanged and not Xc12ColorEqual(Value,FCmpBorder.FLeft.Color) then
    DoBorderChanged;
  FCmpBorder.FLeft.Color := Value;
end;

procedure TXc12XFEditor.SetBorderLeftColorRGB(const Value: TXc12RGBColor);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FLeft.ColorRGB) then
    DoBorderChanged;
  FCmpBorder.FLeft.ColorRGB := Value;
end;

procedure TXc12XFEditor.SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FLeft.Style) then
    DoBorderChanged;
  FCmpBorder.FLeft.Style := Value;
end;

procedure TXc12XFEditor.SetBorderRightColor(const Value: TXc12Color);
begin
  if not FBorderChanged and not Xc12ColorEqual(Value,FCmpBorder.FRight.Color) then
    DoBorderChanged;
  FCmpBorder.FRight.Color := Value;
end;

procedure TXc12XFEditor.SetBorderRightColorRGB(const Value: TXc12RGBColor);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FRight.ColorRGB) then
    DoBorderChanged;
  FCmpBorder.FRight.ColorRGB := Value;
end;

procedure TXc12XFEditor.SetBorderRightStyle(const Value: TXc12CellBorderStyle);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FRight.Style) then
    DoBorderChanged;
  FCmpBorder.FRight.Style := Value;
end;

procedure TXc12XFEditor.SetBorderTopColor(const Value: TXc12Color);
begin
  if not FBorderChanged and not Xc12ColorEqual(Value,FCmpBorder.FTop.Color) then
    DoBorderChanged;
  FCmpBorder.FTop.Color := Value;
end;

procedure TXc12XFEditor.SetBorderTopColorRGB(const Value: TXc12RGBColor);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FTop.ColorRGB) then
    DoBorderChanged;
  FCmpBorder.FTop.ColorRGB := Value;
end;

procedure TXc12XFEditor.SetBorderTopStyle(const Value: TXc12CellBorderStyle);
begin
  if not FBorderChanged and (Value <> FCmpBorder.FTop.Style) then
    DoBorderChanged;
  FCmpBorder.FTop.Style := Value;
end;

procedure TXc12XFEditor.SetFillBgColor(const Value: TXc12Color);
begin
  if not FFillChanged and not Xc12ColorEqual(Value,FCmpFill.BgColor) then
    DoFillChanged;
  FCmpFill.BgColor := Value;
end;

procedure TXc12XFEditor.SetFillBgColorRGB(const Value: TXc12RGBColor);
begin
  if not FFillChanged and (Value <> Xc12ColorToRGB(FCmpFill.BgColor)) then
    DoFillChanged;
  FCmpFill.BgColor := RGBColorToXc12(Value);
end;

procedure TXc12XFEditor.SetFillPattern(const Value: TXc12FillPattern);
begin
  if not FFillChanged and (Value <> FCmpFill.PatternType) then
    DoFillChanged;
  FCmpFill.PatternType := Value;
end;

procedure TXc12XFEditor.SetFontCharset(const Value: TFontCharset);
begin
  if not FFontChanged and (Value <> FCmpFont.CharSet) then
    DoFontChanged;
  FCmpFont.CharSet := Value;
end;

procedure TXc12XFEditor.SetFontColorRGB(const Value: TXc12RGBColor);
begin
  if not FFontChanged and not Xc12ColorEqual(FCmpFont.Color,Value) then
    DoFontChanged;
  FCmpFont.Color := RGBColorToXc12(Value);
end;

procedure TXc12XFEditor.SetFontFamily(const Value: integer);
begin
  if not FFontChanged and (Value <> FCmpFont.Family) then
    DoFontChanged;
  FCmpFont.Family := Value;
end;

procedure TXc12XFEditor.SetFontName(const Value: AxUCString);
begin
  if not FFontChanged and not SameText(Value,FCmpFont.Name) then
    DoFontChanged;
  FCmpFont.Name := Value;
end;

procedure TXc12XFEditor.SetFontSize(const Value: double);
begin
  if not FFontChanged and (Value <> FCmpFont.Size) then
    DoFontChanged;
  FCmpFont.Size := Value;
end;

procedure TXc12XFEditor.SetFontStyle(const Value: TXc12FontStyles);
begin
  if not FFontChanged and (Value <> FCmpFont.Style) then
    DoFontChanged;
  FCmpFont.Style := Value;
end;

procedure TXc12XFEditor.SetFontSubSuperScript(const Value: TXc12SubSuperscript);
begin
  if not FFontChanged and (Value <> FCmpFont.SubSuperscript) then
    DoFontChanged;
  FCmpFont.SubSuperscript := Value;
end;

procedure TXc12XFEditor.SetFontUnderline(const Value: TXc12Underline);
begin
  if not FFontChanged and (Value <> FCmpFont.Underline) then
    DoFontChanged;
  FCmpFont.Underline := Value;
end;

procedure TXc12XFEditor.SetNumberFormat(const Value: AxUCString);
begin
  if Trim(Value) <> '' then begin
    if not FNumFmtChanged and (Value <> FCmpNumFmt.Value) then begin
      FXF.NumFmt.Assign(FCmpNumFmt);
      FCmpNumFmt := FXF.NumFmt;

      FNumFmtChanged := True;
    end;
    FCmpNumFmt.Value := Value;
  end;
end;

procedure TXc12XFEditor.SetProtectionHidden(const Value: boolean);
var
  V: TXc12CellProtections;
begin
  V := FCmpXF.Protection;
  if Value then
    V := V + [cpHidden]
  else
    V := V - [cpHidden];

  if not FProtectChanged and (FCmpXF.Protection <> V) then
    FProtectChanged := True;
  FCmpXF.Protection := V;
end;

procedure TXc12XFEditor.SetProtectionLocked(const Value: boolean);
var
  V: TXc12CellProtections;
begin
  V := FCmpXF.Protection;
  if Value then
    V := V + [cpLocked]
  else
    V := V - [cpLocked];

  if not FProtectChanged and (FCmpXF.Protection <> V) then
    FProtectChanged := True;
  FCmpXF.Protection := V;
end;

function TXc12XFEditor.UseDefault: TXc12XF;
begin
  Result := FStyles.XFs[XLS_STYLE_DEFAULT_XF];
  UseStyle(Result);
end;

procedure TXc12XFEditor.UseFont(AFont: TXc12Font);
begin
  AFont.IncRef;
end;

procedure TXc12XFEditor.UseFonts(AFonts: TXc12DynFontRunArray);
var
  i: integer;
begin
  for i := 0 to High(AFonts) do
    AFonts[i].Font.IncRef;
end;

function TXc12XFEditor.UseRequiredFont(AFont: TXc12Font): TXc12Font;
begin
  Result := FStyles.Fonts.Find(AFont);
  if Result <> Nil then
    Result.IncRef
  else
    raise XLSRWException.Create('Can not find font');
end;

procedure TXc12XFEditor.UseStyle(AXF: TXc12XF);
begin
  AXF.IncRef;
  AXF.NumFmt.IncRef;
  AXF.Font.IncRef;
  AXF.Fill.IncRef;
  AXF.Border.IncRef;
end;

procedure TXc12XFEditor.UseStyleStyle(const AIndex: integer);
var
  XF: TXc12XF;
begin
  XF := FStyles.StyleXFs[AIndex];

  XF.IncRef;
  XF.NumFmt.IncRef;
  XF.Font.IncRef;
  XF.Fill.IncRef;
  XF.Border.IncRef;
end;

procedure TXc12XFEditor.UseStyle(const AIndex: integer);
var
  XF: TXc12XF;
begin
  XF := FStyles.XFs[AIndex];

  XF.IncRef;
  XF.NumFmt.IncRef;
  XF.Font.IncRef;
  XF.Fill.IncRef;
  XF.Border.IncRef;
end;

{ TXc12BorderPr }

procedure TXc12BorderPr.Assign(ABorderPr: TXc12BorderPr);
begin
  FColor := ABorderPr.Color;
  FStyle := ABorderPr.Style;
end;

function TXc12BorderPr.CalcHash: longword;
begin
  Result := Xc12ColorHash(FColor) + Longword(FStyle);
end;

procedure TXc12BorderPr.Clear;
begin
  FColor.ColorType := exctAuto;
  FStyle := cbsNone;
end;

function TXc12BorderPr.Equal(ABorderPr: TXc12BorderPr): boolean;
begin
  Result := Xc12ColorEqual(FColor,ABorderPr.Color) and (FStyle = ABorderPr.Style);
end;

function TXc12BorderPr.GetColorRGB: TXc12RGBColor;
begin
  Result := Xc12ColorToRGB(FColor);
end;

procedure TXc12BorderPr.SetColorRGB(const Value: TXc12RGBColor);
begin
  FColor := RGBColorToXc12(Value);
end;

end.
