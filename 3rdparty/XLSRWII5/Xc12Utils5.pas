unit Xc12Utils5;

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
     XLSUtils5;

const XLS_STYLE_DEFAULT_XF          = 0;
const XLS_STYLE_DEFAULT_FONT        = 0;
const XLS_STYLE_DEFAULT_XF_COUNT    = 1;
const XLS_STYLE_DEFAULT_FONT_COUNT  = 1;
// From number of XF in XLS_DEFAULT_FILE_97 (Xc12DefaultData5).
const XLS_STYLE_DEFAULT_XF_97       = 15;
const XLS_STYLE_DEFAULT_XF_COUNT_97 = 21;

const XLS_MAXROWS   = 1 shl 20;
const XLS_MAXCOLS   = 1 shl 14;
const XLS_MAXROW    = XLS_MAXROWS - 1;
const XLS_MAXCOL    = XLS_MAXCOLS - 1;

const XLS_MAXCOL_97 = 255;
const XLS_MAXROW_97 = 65535;

const XLS_MAXSHEETS = 255;

// True for Excel 97. Not tested with Excel 2007.
const MAXSELROW = $3FFF;

const MAX_EXCEL_ARGCOUNT = 254;

const MAX_EXCEL_STRSZ = $7FFF - 1;

// Empiric.
const XLS_DEFAULT_COLWIDTH  = 2340;
const XLS_DEFAULT_ROWHEIGHT = 15 * 20;

const XLS_DEFAULT_ROWHEIGHT_FLAG = $FFFF;

const COL_ABSFLAG     = $8000;
const ROW_ABSFLAG     = $80000000;
const NOT_COL_ABSFLAG = $7FFF;
const NOT_ROW_ABSFLAG = $7FFFFFFF;

const XLS_COLOR_DEFAULT_COMMENT = $FFFFE1;
const XLS_COLOR_AUTO            = $FF000000;
const XLS_COLOR_NONE            = $FE000000;

type TExcelVersion = (xvNone,xvExcelUnknown,xvExcel21,xvExcel30,xvExcel40,xvExcel50,xvExcel97,xvExcel2007,xvOffice2007Encrypted);

type TXLSCellType = (xctNone,               //* Undefined cell.
                     xctBlank,              //* Blank cell.
                     xctBoolean,            //* Boolean cell.
                     xctError,              //* Cell with an error value.
                     xctString,             //* String cell (index into SST).
                     xctFloat,              //* Floating point cell.
                     xctFloatFormula,       //* Formula cell where the result of the formula is a numeric value.
                     xctStringFormula,      //* Formula cell where the result of the formula is a string value.
                     xctBooleanFormula,     //* Formula cell where the result of the formula is a boolean value.
                     xctErrorFormula        //* Formula cell where the result of the formula is an error value.
                     );

const XLSCellTypeFormulas = [xctFloatFormula..xctErrorFormula];
const XLSCellTypeNum = [xctFloat,xctFloatFormula];

type TXc12RId = string;

type TXc12RefEvent = procedure (Col,Row: integer) of object;
type TXc12CreateObjectEvent = procedure (AClass: TClass; out AObject: TObject) of object;

type TXLSHyperlinkType = (xhltUnknown, xhltURL, xhltFile, xhltUNC, xhltWorkbook);

type TXc12Visibility = (x12vVisible,x12vHidden,x12vVeryHidden);

//* Built in names that defines special purpose areas.
type TXc12BuiltInName = (bnConsolidateArea,
                         bnExtract,
                         bnDatabase,
                         bnCriteria,
                         bnPrintArea,
                         bnPrintTitles,
                         bnSheetTitle,
                         bnFilterDatabase,
                         bnNone
                         );

                                                   // xsntError = sheet is deleted, #REF! error
type TXc12SimpleNameType = (xsntNone,xsntRef,xsntArea,xsntError);

const Xc12FormulaOpt_ACA  = $01;
const Xc12FormulaOpt_BX   = $02;
const Xc12FormulaOpt_CA   = $04;

const Xc12FormulaTableOpt_DEL1 = $01;
const Xc12FormulaTableOpt_DEL2 = $02;
const Xc12FormulaTableOpt_DT2D = $04;
const Xc12FormulaTableOpt_DTR  = $08;

type TXc12CellError = (errUnknown,
                       errNull,  //* Empty intersection.
                       errDiv0,  //* Division by null.
                       errValue, //* Illegal value.
                       errRef,   //* Illegal reference. Possibly to a deleted worksheet.
                       errName,  //* Unknown name.
                       errNum,   //* Illegal number.
                       errNA,    //* Not available error.
                       errGettingData);

type PXc12CellError = ^TXc12CellError;

const Xc12CellErrorNames: array[TXc12CellError] of AxUCString =
('_ERROR_ERROR_','#NULL!','#DIV/0!','#VALUE!','#REF!','#NAME?','#NUM!','#N/A','#GETTING_DATA');

type TXc12PaperSize = (
  psNone, psLetter, psLetterSmall, psTabloid, psLedger, psLegal, psStatement,
  psExecutive, psA3, psA4, psA4Small, psA5, psB4, psB5, psFolio, psQuarto,
  ps10X14, ps11X17, psNote, psEnv9, psEnv10, psEnv11, psEnv12, psEnv14,
  psCSheet, psDSheet, psESheet, psEnvDL, psEnvC5, psEnvC3, psEnvC4, psEnvC6,
  psEnvC65, psEnvB4, psEnvB5, psEnvB6, psEnvItaly, psEnvMonarch, psEnvPersonal,
  psFanfoldUS, psFanfoldStdGerman, psFanfoldLglGerman, psISO_B4, psJapanesePostcard,
  ps9X11, ps10X11, ps15X11, psEnvInvite, psReserved48, psReserved49, psLetterExtra,
  psLegalExtra, psTabloidExtra, psA4Extra, psLetterTransverse, psA4Transverse,
  psLetterExtraTransverse, psAPlus, psBPlus, psLetterPlus, psA4Plus, psA5Transverse,
  psB5transverse, psA3Extra, psA5Extra, psB5Extra, psA2, psA3Transverse, psA3ExtraTransverse);

const Xc12PaperSizeNames: array[TXc12PaperSize] of string = (
  'Undefined', 'Letter', 'LetterSmall', 'Tabloid', 'Ledger', 'Legal', 'Statement',
  'Executive', 'A3', 'A4', 'A4Small', 'A5', 'B4', 'B5', 'Folio', 'Quarto',
  '10X14', '11X17', 'Note', 'Env9', 'Env10', 'Env11', 'Env12', 'Env14',
  'CSheet', 'DSheet', 'ESheet', 'EnvDL', 'EnvC5', 'EnvC3', 'EnvC4', 'EnvC6',
  'EnvC65', 'EnvB4', 'EnvB5', 'EnvB6', 'EnvItaly', 'EnvMonarch', 'EnvPersonal',
  'FanfoldUS', 'FanfoldStdGerman', 'FanfoldLglGerman', 'ISO_B4', 'JapanesePostcard',
  '9X11', '10X11', '15X11', 'EnvInvite', 'Reserved48', 'Reserved49', 'LetterExtra',
  'LegalExtra', 'TabloidExtra', 'A4Extra', 'LetterTransverse', 'A4Transverse',
  'LetterExtraTransverse', 'APlus', 'BPlus', 'LetterPlus', 'A4Plus', 'A5Transverse',
  'B5transverse', 'A3Extra', 'A5Extra', 'B5Extra', 'A2', 'A3Transverse', 'A3ExtraTransverse');

// TODO Find some better value.
const XLSCOLOR_AUTO = $F0000000;

type TXc12RGBColor = longword;

type TXc12IndexColor = (
     xc0,xc1,xc2,xc3,xc4,xc5,xc6,xc7,
     xcBlack,xcWhite,xcRed,xcBrightGreen,xcBlue,xcYellow,xcPink,xcTurquoise,
     xcDarkRed,xcGreen,xcDarkBlue,xcBrownGreen,xcViolet,xcBlueGreen,xcGray25,xcGray50,
     xc24,xc25,xc26,xc27,xc28,xc29,xc30,xc31,
     xc32,xc33,xc34,xc35,xc36,xc37,xc38,xc39,
     xcSky,xcPaleTurquois,xcPaleGreen,xcLightYellow,xcPaleSky,xcRose,xcLilac,xcLightBrown,
     xcDarkSky,xcDarkTurquois,xcGrass,xcGold,xcLightOrange,xcOrange,xcDarkBlueGray,xcGray40,
     xcDarkGreenGray,xcEmerald,xcDarkGreen,xcOlive,xcBrown,xcCherry,xcIndigo,xcGray80,
     xcAutomatic);

// For TXc12Colors that not are index colors.
const XLSCOLOR_INDEX_NOT_INDEX = xc0;

const TXc12DefaultIndexColorPalette: array[0..65] of integer = (
$000000, $FFFFFF, $0000FF, $00FF00, $FF0000, $00FFFF, $FF00FF, $FFFF00,
$000000, $FFFFFF, $0000FF, $00FF00, $FF0000, $00FFFF, $FF00FF, $FFFF00,
$000080, $008000, $800000, $008080, $800080, $808000, $C0C0C0, $808080,
$FF9999, $663399, $CCFFFF, $FFFFCC, $660066, $8080FF, $CC6600, $FFCCCC,
$800000, $FF00FF, $00FFFF, $FFFF00, $800080, $000080, $808000, $FF0000,
$FFCC00, $FFFFCC, $CCFFCC, $99FFFF, $FFCC99, $CC99FF, $FF99CC, $99CCFF,
$FF6633, $CCCC33, $00CC99, $00CCFF, $0099FF, $0066FF, $996666, $969696,
$663300, $669933, $003300, $003333, $003399, $663399, $993333, $333333,
$0F000000,$0F000000);

const TXc12DefaultIndexColorPaletteRGB: array[0..65] of integer = (
$000000, $FFFFFF, $FF0000, $00FF00, $0000FF, $FFFF00, $FF00FF, $00FFFF,
$000000, $FFFFFF, $FF0000, $00FF00, $0000FF, $FFFF00, $FF00FF, $00FFFF,
$800000, $008000, $000080, $808000, $800080, $008080, $C0C0C0, $808080,
$9999FF, $993366, $FFFFCC, $CCFFFF, $660066, $FF8080, $0066CC, $CCCCFF,
$000080, $FF00FF, $FFFF00, $00FFFF, $800080, $800000, $008080, $0000FF,
$00CCFF, $CCFFFF, $CCFFCC, $FFFF99, $99CCFF, $FF99CC, $CC99FF, $FFCC99,
$3366FF, $33CCCC, $99CC00, $FFCC00, $FF9900, $FF6600, $666699, $969696,
$003366, $339966, $003300, $333300, $993300, $993366, $333399, $333333,
$0F000000,$0F000000);

var Xc12IndexColorPalette: array[0..High(TXc12DefaultIndexColorPalette)] of integer;

var Xc12IndexColorPaletteRGB: array[0..High(TXc12DefaultIndexColorPalette)] of integer;

type TXc12ClrSchemeColor = (cscLt1,cscDk1,cscLt2,cscDk2,cscAccent1,cscAccent2,
                            cscAccent3,cscAccent4,cscAccent5,cscAccent6,
                            cscHLink,cscFolHLink,cscExtList);

// Shall match name in xsd scheme
const TXc12ClrSchemeColorName: array[TXc12ClrSchemeColor] of AxUCString = (
      'lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5',
      'accent6','hlink','folHlink','extList');

const Xc12DefColorSchemeVals: array[TXc12ClrSchemeColor] of longword = (
      $FFFFFF,$000000,$E1ECEE,$7D491F,$BD814F,$4D50C0,$59BB9B,
      $A26480,$C6AC4B,$4696F7,$FF0000,$800080,$7F7F7F);

const Xc12DefColorSchemeValsRGB: array[TXc12ClrSchemeColor] of longword = (
      $FFFFFF,$000000,$EEECE1,$1F497D,$4F81BD,$C0504D,$9BBB59,
      $8064A2,$4BACC6,$F79646,$0000FF,$800080,$7F7F7F);

var Xc12DefColorScheme   : array[TXc12ClrSchemeColor] of longword;
var Xc12DefColorSchemeRGB: array[TXc12ClrSchemeColor] of longword;

type TXc12ColorType = (exctAuto,exctIndexed,exctRgb,exctTheme,exctUnassigned);

type PXc12Color = ^TXc12Color;
     TXc12Color = record
     ARGB: TXc12RGBColor;
     Tint: double;
     case ColorType: TXc12ColorType of
       exctAuto:    (Auto: boolean);
       exctIndexed: (Indexed: TXc12IndexColor);
       exctRgb:     (OrigRGB: TXc12RGBColor);
       exctTheme:   (Theme: TXc12ClrSchemeColor);
     end;

type TRGBRec = packed record
    case Integer of
      0: (Value: longword);
      1: (Red, Green, Blue: Byte);
      2: (R, G, B, Flag: Byte);
      3: (Index: Word);
  end;

type THLSRec = record
     H,L,S: byte;
     end;

type PXLS3dCompactRef = ^TXLS3dCompactRef;
     TXLS3dCompactRef = record
     SheetId: word;
     Col: word;
     Row: longword;
     end;

type PXLS3dCellArea = ^TXLS3dCellArea;
     TXLS3dCellArea = record
     // A negative index means the sheet is deleted and the area is an error ref.
     SheetIndex: integer;
     Row1,Row2: integer;
     Col1,Col2: integer;
     end;

// TODO Rename to TRecCellArea
type PXLSCellArea = ^TXLSCellArea;
     TXLSCellArea = record
     Row1,Row2: integer;
     Col1,Col2: integer;
     end;

type PXLSCellRef = ^TXLSCellRef;
     TXLSCellRef = record
     Row: integer;
     Col: integer;
     end;

function  Xc12ColorToRGB(const Color: TXc12Color; AAutoColor: TXc12RGBColor = $00000000): TXc12RGBColor; overload; {$ifdef D2006PLUS} inline; {$endif}
function  Xc12ColorToRGB(const Color: PXc12Color; AAutoColor: TXc12RGBColor = $00000000): TXc12RGBColor; overload; {$ifdef D2006PLUS} inline; {$endif}

function  IndexColorToXc12(const Color: integer): TXc12Color; overload; {$ifdef D2006PLUS} inline; {$endif}
function  IndexColorToXc12(const Color: TXc12IndexColor): TXc12Color; overload; {$ifdef D2006PLUS} inline; {$endif}

function  RGBColorToXc12(const ARGB: TXc12RGBColor): TXc12Color; overload;
function  RGBColorToXc12(const ARGB: TXc12RGBColor; Color: PXc12Color): TXc12Color; overload;

function  ThemeColorToXc12(const AScheme: TXc12ClrSchemeColor; const ATint: double): TXc12Color; overload
procedure ThemeColorToXc12(const AColor: PXc12Color; const AScheme: TXc12ClrSchemeColor; const ATint: double); overload

function  ThemeColorToRGB(const AScheme: TXc12ClrSchemeColor; const ATint: double): longword;

function  HLSToRGB(const PHLS: THLSRec): TRGBRec; overload;
function  HLSToRGB(const AHue,ALum,ASat: double): longword; overload;
function  RGBToHLS(const PRGB: TRGBRec): THLSRec; overload;
procedure RGBToHLS(const ARGB: longword; out AHue,ALum,ASat: double); overload;

function  Xc12ColorEqual(const C1,C2: TXc12Color): boolean; overload;
function  Xc12ColorEqual(const ColorXc12: TXc12Color; const ColorRGB: TXc12RGBColor): boolean; overload;
function  Xc12ColorEqual(const ColorXc12: PXc12Color; const ColorRGB: TXc12RGBColor): boolean; overload;
function  Xc12ColorEqual(const ColorXc12: TXc12Color; const ColorIndex: TXc12IndexColor): boolean; overload;

function  Xc12ColorIsIndexColor(const AColor: TXc12Color): boolean; {$ifdef D2006PLUS} inline; {$endif}
function  Xc12ColorHash(const AColor: TXc12Color): longword;
function  MakeXc12ColorAuto(const ADefault: longword = $000000): TXc12Color;

function  AreaIsAssigned(const Area: TXLSCellArea): boolean;
procedure AreaStrToCellArea(const S: AxUCString; out Area: TXLSCellArea); overload;
function  AreaStrToCellArea(const S: AxUCString): TXLSCellArea; overload;
procedure ClearCellArea(out Area: TXLSCellArea);
procedure SetCellArea(out Area: TXLSCellArea; C1,R1,C2,R2: integer); overload;
procedure SetCellArea(out Area: TXLSCellArea; C,R: integer); overload;
function  SetCellArea(C,R: integer): TXLSCellArea; overload;
function  SetCellArea(C1,R1,C2,R2: integer): TXLSCellArea; overload;
function  CellAreaAssigned(Area: TXLSCellArea): boolean;
function  CellAreaToAreaStr(Area: TXLSCellArea): AxUCString;
function  IntersectCellArea(A1, A2: TXLSCellArea; out Dest: TXLSCellArea): boolean;
procedure ExtendCellArea(A1, A2: TXLSCellArea; out Dest: TXLSCellArea); overload;
procedure ExtendCellArea(var AArea: TXLSCellArea; C1,R1,C2,R2: integer); overload;
function  AreaIsRef(AArea: TXLSCellArea): boolean;

function  AreaToRefStr(Col1,Row1,Col2,Row2: integer): AxUCString; overload;
function  AreaToRefStr(Col1,Row1,Col2,Row2: integer; AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean): AxUCString; overload;
function  AreaToRefStrAbs(Col1,Row1,Col2,Row2: integer): AxUCString;
function  ColRowToRefStr(ACol,ARow: integer): AxUCString; overload;
function  ColRowToRefStr(ACol,ARow: integer; AbsCol,AbsRow: boolean): AxUCString; overload;
function  ColToRefStr(ACol: integer; const AAbsolute: boolean): AxUCString;
function  RowToRefStr(ARow: integer; const AAbsolute: boolean): AxUCString;
function  ColsToRefStr(ACol1,ACol2: integer; const AAbsolute: boolean): AxUCString;
function  RowsToRefStr(ARow1,ARow2: integer; const AAbsolute: boolean): AxUCString;
function  ShortAreaToRefStr(Col1,Row1,Col2,Row2: integer; AAbsolute: boolean = False): AxUCString;

function  ColRowToRefStrEnc(ACol,ARow: integer): AxUCString; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function  AreaToRefStrEnc(Col1,Row1,Col2,Row2: integer): AxUCString; // {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

procedure ClipAreaToExtent(var C1,R1,C2,R2: integer);
function  InsideExtent(const C,R: integer): boolean; overload;
function  InsideExtent(const C1,R1,C2,R2: integer): boolean; overload;

function  RefStrToColRow(const S: AxUCString; out Col,Row: integer): boolean; overload;
function  RefStrToColRow(const S: AxUCString; out Col,Row: integer; var AbsCol,AbsRow: boolean): boolean; overload;
function  AreaStrToColRow(const S: AxUCString; out Col1,Row1,Col2,Row2: integer; var AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean): boolean; overload;
function  AreaStrToColRow(S: AxUCString; out Col1,Row1,Col2,Row2: integer): boolean; overload;
function  IsAreaStr(S: AxUCString): boolean;
function  ColStrToCol(const S: AxUCString; out Col: integer): boolean;
function  RowStrToRow(const S: AxUCString; out Row: integer): boolean;

function  ColStrToCol97(const S: AxUCString; out Col: integer): boolean;
function  RowStrToRow97(const S: AxUCString; out Row: integer): boolean;
function  RefStrToColRow97(const S: AxUCString; out Col,Row: integer): boolean;
function  AreaStrToColRow97(const S: AxUCString; out Col1,Row1,Col2,Row2: integer): boolean;

function  R1C1Index(const S: AxUCString; out AIndex: integer; out ARel: boolean): boolean;
function  R1C1ToColRow(const S: AxUCString; out Col,Row: integer; out RelC,RelR: boolean): boolean;
function  R1C1ToArea(const S: AxUCString; out Col1,Row1,Col2,Row2: integer; out RelC1,RelR1,RelC2,RelR2: boolean): boolean;

function  ErrorTextToCellError(const S: AxUCString): TXc12CellError;
function  CellErrorToErrorText(AError: TXc12CellError): AxUCString;

procedure NormalizeArea(var C1,R1,C2,R2: integer);
function  AreaInsideSheet(const C1,R1,C2,R2: integer): boolean;
function  ClipAreaToSheet(var C1,R1,C2,R2: integer): boolean;

function TColorToClosestIndexColor(Color: TColor): TXc12IndexColor;
function LightenColor(AColor: longword; AValue: double): longword;

function  GetBasePtgs(APtg: byte): byte; {$ifdef D2006PLUS} inline; {$endif}

var
  G_StrTRUE,G_StrFALSE: AxUCString;

implementation

const EX12_AUTO_COLOR = $0F000000;

function Lighten(Lum: integer; Value: double): integer;
begin
// Lum‘ = Lum * (1.0-tint) + (HLSMAX – HLSMAX * (1.0-tint))
  Result := Round(Lum * (1.0 - Value) + (255 - 255 * (1.0 - Value)));
end;

function Darken(Lum: integer; Value: double): integer;
begin
// Lum’ = Lum * (1.0 + tint)
  Result := Round(Lum * (1.0 + Value));
end;


function Xc12ColorToRGB(const Color: PXc12Color; AAutoColor: TXc12RGBColor = $00000000): longword;
var
  RGB: TRGBRec;
  HLS: THLSRec;
//  RRec: TRGBRec;
//  T: Byte;
begin
  case Color.ColorType of
    exctAuto   : Result := AAutoColor;
    exctIndexed: begin
      if Integer(Color.Indexed) > High(Xc12IndexColorPalette) then
        Result := AAutoColor
      else
        Result := Xc12IndexColorPalette[Integer(Color.Indexed)];
    end;
    exctRgb    : begin
      Result := Color.OrigRGB and $00FFFFFF;
// Don't reverse colors. This function shall return RGB.
//      RRec.Value := Result;
//      T := RRec.R;
//      RRec.R := RRec.B;
//      RRec.B := T;
//      Result := RRec.Value;
    end;
    exctTheme  : begin
      if Color.Tint = 0 then
        Result := Xc12DefColorScheme[Color.Theme]
      else begin
        RGB.Value := Xc12DefColorScheme[Color.Theme];
        HLS := RGBToHLS(RGB);
        if Color.Tint > 0 then
          HLS.L := Lighten(HLS.L,Color.Tint)
        else if Color.Tint < 0 then
          HLS.L := Darken(HLS.L,Color.Tint);
        RGB := HLSToRGB(HLS);
        Result := RGB.Value;
      end;
    end
    else       Result := $000000;
  end;
end;

function LightenColor(AColor: longword; AValue: double): longword;
var
  RGB: TRGBRec;
  HLS: THLSRec;
begin
  RGB.Value := AColor;
  HLS := RGBToHLS(RGB);

  if AValue > 0 then
    HLS.L := Lighten(HLS.L,AValue)
  else if AValue < 0 then
    HLS.L := Darken(HLS.L,AValue);

  RGB := HLSToRGB(HLS);

  Result := RGB.Value;
end;

function Xc12ColorToRGB(const Color: TXc12Color; AAutoColor: TXc12RGBColor = $00000000): longword;
begin
  Result := Xc12ColorToRGB(@Color,AAutoColor);
end;

function IndexColorToXc12(const Color: integer): TXc12Color;
begin
  Result.ColorType := exctIndexed;
  Result.Tint := 0;
  if (Color < 0) or (Color > Integer(High(TXc12IndexColor))) then begin
    Result.Indexed := xcAutomatic;
    Result.ARGB := XLSCOLOR_AUTO;
  end
  else begin
    Result.Indexed := TXc12IndexColor(Color);
    Result.ARGB := Xc12IndexColorPalette[Color];
  end;
end;

function  IndexColorToXc12(const Color: TXc12IndexColor): TXc12Color;
begin
  Result := IndexColorToXc12(Longword(Color));
end;

function Max3(P1,P2,P3: byte): byte;
begin
  if (P1 > P2) then begin
    if (P1 > P3) then begin
      Result := P1;
    end else begin
      Result := P3;
    end;
  end else if P2 > P3 then begin
    result := P2;
  end else result := P3;
end;

function Min3(P1,P2,P3: byte): byte;
begin
  if (P1 < P2) then begin
    if (P1 < P3) then begin
      Result := P1;
    end else begin
      Result := P3;
    end;
  end else if P2 < P3 then begin
    result := P2;
  end else result := P3;
end;

procedure RGBToHLS(const ARGB: longword; out AHue,ALum,ASat: double); overload;
var
  RGB: TRGBRec;
  HLS: THLSRec;
begin
  RGB.Value := ARGB;

  HLS := RGBToHLS(RGB);

  AHue := HLS.H / 255;
  ALum := HLS.L / 255;
  ASat := HLS.S / 255;
end;

function RGBToHLS(const PRGB: TRGBRec): THLSRec;
Var
  LR,LG,LB,LMin,LMax,LDif : byte;
  LH,LL,LS,LSum : integer;
begin
  LR := PRGB.R;
  LG := PRGB.G;
  LB := PRGB.B;
  LMin := min3(LR,LG,LB);
  LMax := max3(LR,LG,LB);
  LDif := LMax - LMin;
  LSum := LMax + LMin;
  LL := LSum shr 1;
  if LMin = LMax then begin
    LH := 0;
    LS := 0;
    Result.H := LH;
    Result.L := LL;
    Result.S := LS;
    exit;
  end;
  If LL < 128 then LS := LDif shl 8 div LSum
              else LS := LDif shl 8 div (512 - LSum);
  if LS > 255 then LS := 255;
  If LR = LMax then LH := (LG - LB) shl 8 div LDif
  else If LG = LMax then LH := 512 + (LB - LR) shl 8 div LDif
  else LH := 1024 + (LR - LG) shl 8 div LDif;
  Result.H := LH div 6;
  Result.L := LL;
  Result.S := LS;
end;

function  HLSToRGB(const AHue,ALum,ASat: double): longword;
var
  RGB: TRGBRec;
  HLS: THLSRec;
begin
  HLS.H := Round(AHue * 255);
  HLS.L := Round(ALum * 255);
  HLS.S := Round(ASat * 255);

  RGB := HLSToRGB(HLS);

  Result := RGB.Value;
end;

function HLSToRGB(const PHLS: THLSRec): TRGBRec;
Var
  LH,LL,LS : byte;
  LR,LG,LB,L1,L2,LDif,L6Dif : integer;
begin
  LH := PHLS.H;
  LL := PHLS.L;
  LS := PHLS.S;
  if LS = 0 then begin
    Result.R := LL;
    Result.G := LL;
    Result.B := LL;
    Exit;
  end;
  If LL < 128 then L2 := LL * (256 + LS) shr 8
              else L2 := LL + LS - LL * LS shr 8;
  L1 := LL shl 1 - L2;
  LDif := L2 - L1;
  L6Dif := LDif * 6;
  LR := LH + 85;
  if LR < 0 then Inc(LR, 256);
  if LR > 256 then Dec(LR, 256);
  If LR < 43 then LR := L1 + L6Dif * LR shr 8
  Else if LR < 128 then LR := L2
  Else if LR < 171 then LR := L1 + L6Dif * (170 - LR) shr 8
  Else LR := L1;
  if LR > 255 then LR := 255;
  LG := LH;
  if LG < 0 then Inc(LG, 256);
  if LG > 256 then Dec(LG, 256);
  If LG < 43 then LG := L1 + L6Dif * LG shr 8
  Else if LG < 128 then LG := L2
  Else if LG < 171 then LG := L1 + L6Dif * (170 - LG) shr 8
  Else LG := L1;
  if LG > 255 then LG := 255;
  LB := LH - 85;
  if LB < 0 then Inc(LB, 256);
  if LB > 256 then Dec(LB, 256);
  If LB < 43 then LB := L1 + L6Dif * LB shr 8
  Else if LB < 128 then LB := L2
  Else if LB < 171 then LB := L1 + L6Dif * (170 - LB) shr 8
  Else LB := L1;
  if LB > 255 then LB := 255;
  Result.R := LR;
  Result.G := LG;
  Result.B := LB;
end;

function RGBToColor(PR,PG,PB: Integer): TColor;
begin
  Result := TColor((PB shl 16) + (PG shl 8) + PR);
end;

function ColorToRGB(PColor: TColor): TRGBRec;
begin
  Result.R :=  PColor and $000000FF;
  Result.G := (PColor and $0000FF00) shr 8;
  Result.B := (PColor and $00FF0000) shr 16;
end;

function RGBToCol(PRGB: TRGBRec): TColor;
begin
  Result := RGBToColor(PRGB.R,PRGB.G,PRGB.B);
end;

function Xc12ColorEqual(const C1,C2: TXc12Color): boolean;
begin
  Result := C1.ColorType = C2.ColorType;
  if Result then begin
    case C1.ColorType of
      exctAuto:       Result := True;
      exctIndexed:    Result := C1.Indexed = C2.Indexed;
      exctRgb:        Result := C1.OrigRGB = C2.OrigRGB;
      exctTheme:      Result := (C1.Theme = C2.Theme) and SameValue(C1.Tint,C2.Tint,0.0001);
      exctUnassigned: Result := False;
    end;
  end;
end;

function  Xc12ColorEqual(const ColorXc12: PXc12Color; const ColorRGB: TXc12RGBColor): boolean; overload;
begin
  Result := (ColorXc12.ColorType = exctRgb) and (ColorXc12.OrigRGB = ColorRGB);
end;

function Xc12ColorEqual(const ColorXc12: TXc12Color; const ColorRGB: longword): boolean;
begin
  Result := (ColorXc12.ColorType = exctRgb) and (ColorXc12.OrigRGB = ColorRGB);
end;

function Xc12ColorEqual(const ColorXc12: TXc12Color; const ColorIndex: TXc12IndexColor): boolean;
begin
  Result := (ColorXc12.ColorType = exctIndexed) and (ColorXc12.Indexed = ColorIndex);
end;

function  Xc12ColorIsIndexColor(const AColor: TXc12Color): boolean;
begin
  Result := AColor.ColorType = exctIndexed;
end;

function  Xc12ColorHash(const AColor: TXc12Color): longword;
begin
  Result := Longword(AColor.ColorType);
  case AColor.ColorType of
    exctAuto      : Inc(Result);
    exctIndexed   : Inc(Result,Integer(AColor.Indexed));
    exctRgb       : Inc(Result,Integer(AColor.OrigRGB));
    exctTheme     : begin
      Inc(Result,Integer(AColor.Theme));
      Inc(Result,Round(AColor.Tint * 10000));
    end;
    exctUnassigned: ;
  end;
end;

function MakeXc12ColorAuto(const ADefault: longword = $000000): TXc12Color;
begin
  Result.ColorType := exctAuto;
  Result.ARGB := ADefault;
end;

function  AreaIsAssigned(const Area: TXLSCellArea): boolean;
begin
  Result := (Area.Col1 >= 0) and (Area.Row1 >= 0);
end;

procedure AreaStrToCellArea(const S: AxUCString; out Area: TXLSCellArea);
begin
  if S = '' then begin
    Area.Col1 := 0;
    Area.Row1 := 0;
    Area.Col2 := 0;
    Area.Row2 := 0;
  end
  else if CPos(':',S) < 1 then begin
    RefStrToColRow(S,Area.Col1,Area.Row1);
    Area.Col2 := Area.Col1;
    Area.Row2 := Area.Row1
  end
  else
    AreaStrToColRow(S,Area.Col1,Area.Row1,Area.Col2,Area.Row2);
end;

function  AreaStrToCellArea(const S: AxUCString): TXLSCellArea;
begin
  if S = '' then begin
    Result.Col1 := -1;
    Result.Row1 := -1;
    Result.Col2 := -1;
    Result.Row2 := -1;
  end
  else if CPos(':',S) < 1 then begin
    RefStrToColRow(S,Result.Col1,Result.Row1);
    Result.Col2 := Result.Col1;
    Result.Row2 := Result.Row1
  end
  else
    AreaStrToColRow(S,Result.Col1,Result.Row1,Result.Col2,Result.Row2);
end;

procedure ClearCellArea(out Area: TXLSCellArea);
begin
  Area.Col1 := -1;
  Area.Row1 := -1;
  Area.Col2 := -1;
  Area.Row2 := -1;
end;

procedure SetCellArea(out Area: TXLSCellArea; C,R: integer);
begin
  Area.Col1 := C;
  Area.Row1 := R;
  Area.Col2 := C;
  Area.Row2 := R;
end;

procedure SetCellArea(out Area: TXLSCellArea; C1,R1,C2,R2: integer);
begin
  Area.Col1 := C1;
  Area.Row1 := R1;
  Area.Col2 := C2;
  Area.Row2 := R2;
end;

function  SetCellArea(C,R: integer): TXLSCellArea; overload;
begin
  Result.Col1 := C;
  Result.Row1 := R;
  Result.Col2 := C;
  Result.Row2 := R;
end;

function  SetCellArea(C1,R1,C2,R2: integer): TXLSCellArea; overload;
begin
  Result.Col1 := C1;
  Result.Row1 := R1;
  Result.Col2 := C2;
  Result.Row2 := R2;
end;

function CellAreaAssigned(Area: TXLSCellArea): boolean;
begin
  Result := (Area.Col1 >= 0) and (Area.Row1 >= 0);
end;

function CellAreaToAreaStr(Area: TXLSCellArea): AxUCString;
begin
  if (Area.Col1 < 0) or (Area.Row1 < 0) then
    Result := ''
  else if (Area.Col2 < 0) or (Area.Row2 < 0) or ((Area.Col1 = Area.Col2) and (Area.Row1 = Area.Row2)) then
    Result := ColRowToRefStr(Area.Col1,Area.Row1)
  else
    Result := AreaToRefStr(Area.Col1,Area.Row1,Area.Col2,Area.Row2);
end;


function RefStrToColRow(const S: AxUCString; out Col,Row: integer; var AbsCol,AbsRow: boolean): boolean;
var
  i,j,m: integer;
begin
  AbsCol := False;
  AbsRow := False;
  Result := False;
  if Length(S) < 1 then
    Exit;
  Col := 0;
  m := 1;
  i := Length(S);
  while CharInSet(S[i],['0'..'9']) do
    Dec(i);
  if S[i] = '$' then begin
    AbsRow := True;
    Dec(i);
  end;
  j := i;
  repeat
    Col := Col + (Ord(S[j]) - Ord('A') + 1) * m;
    m := m * 26;
    Dec(j);
  until ((j = 0) or not (CharInSet(S[j],['A'..'Z'])));
  if (j > 0) and (S[j] = '$') then begin
    AbsCol := True;
    Inc(i);
    Dec(j);
  end;
  Result := j = 0;
  if not Result then
    Exit;
  Dec(Col);
  Row := StrToInt(Copy(S,i + 1,MAXINT)) - 1;
  Result := (Col >= 0) and (Col <= XLS_MAXCOL) and (Row >= 0) and (ROW <= XLS_MAXROW);
end;

function RefStrToColRow(const S: AxUCString; out Col,Row: integer): boolean;
var
  i,j,m: integer;
begin
  Result := False;
  if Length(S) < 1 then
    Exit;
  Col := 0;
  m := 1;
  i := Length(S);
  while CharInSet(S[i],['0'..'9']) do
    Dec(i);
  j := i;
  if (j > 0) and (S[j] = '$') then
    Dec(j);
  repeat
    Col := Col + (Ord(S[j]) - Ord('A') + 1) * m;
    m := m * 26;
    Dec(j);
  until ((j = 0) or not (CharInSet(S[j],['A'..'Z'])));
  if (j > 0) and (S[j] = '$') then
    Dec(j);
  Result := j = 0;
  if not Result then
    Exit;
  Dec(Col);
  Result := TryStrToInt(Copy(S,i + 1,MAXINT),Row);
  if Result then begin
    Dec(Row);
    Result := (Col >= 0) and (Col <= XLS_MAXCOL) and (Row >= 0) and (ROW <= XLS_MAXROW);
  end;
end;

function AreaStrToColRow(S: AxUCString; out Col1,Row1,Col2,Row2: integer): boolean;
var
  p: integer;
  S1,S2: AxUCString;
begin
  p := CPos('!',S);
  if p > 1 then
    S := Copy(S,p + 1,MAXINT);

  p := CPos(':',S);
  if p < 1 then begin
    Result := RefStrToColRow(S,Col1,Row1);
    Col2 := Col1;
    Row2 := Row1;
  end
  else begin
    S1 := Copy(S,1,p - 1);
    S2 := Copy(S,p + 1,MAXINT);
    Result := RefStrToColRow(S1,Col1,Row1);
    if Result then
      Result := RefStrToColRow(S2,Col2,Row2);
  end;
  if not Result then begin
    Result := ColStrToCol(S1,Col1);
    if Result then begin
      Result := ColStrToCol(S2,Col2) and (Col2 >= Col1);
      if Result then begin
        Row1 := 1;
        Row2 := XLS_MAXROW;
      end;
    end
    else begin
      Result := RowStrToRow(S1,Row1);
      if Result then begin
        Result := RowStrToRow(S2,Row2) and (Row2 >= Row1);
        if Result then begin
          Col1 := 1;
          Col2 := XLS_MAXCOL;
        end;
      end;
    end;
  end;
end;

function  IsAreaStr(S: AxUCString): boolean;
var
  C1,R1,C2,R2: integer;
begin
  Result := AreaStrToColRow(S,C1,R1,C2,R2);
end;

function AreaStrToColRow(const S: AxUCString; out Col1,Row1,Col2,Row2: integer; var AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean): boolean;
var
  p: integer;
begin
  Result := False;
  p := CPos(':',S);
  if p < 1 then
    Exit;
  Result := RefStrToColRow(Copy(S,1,p - 1),Col1,Row1,AbsCol1,AbsRow1);
  if Result then
    Result := RefStrToColRow(Copy(S,p + 1,MAXINT),Col2,Row2,AbsCol2,AbsRow2);
end;

function  ColStrToCol97(const S: AxUCString; out Col: integer): boolean;
var
  i,j: integer;
  m: integer;
begin
  Result := False;
  if Length(S) < 1 then
    Exit;
  j := 0;
  if S[1] = '$' then
    Inc(j);

  m := 1;
  Col := 0;
  i := Length(S);
  repeat
    Col := Col + (Ord(S[i]) - Ord('A') + 1) * m;
    m := m * 26;
    Dec(i);
  until ((i = j) or not (CharInSet(S[i],['A'..'Z'])));
  Dec(Col);
  Result := (i = j) and (Col <= XLS_MAXCOL_97);
end;

function  RowStrToRow97(const S: AxUCString; out Row: integer): boolean;
var
  i,j: integer;
begin
  Result := False;
  if Length(S) < 1 then
    Exit;
  i := 1;
  if S[1] = '$' then
    Inc(i);

  for j := i to Length(S) do begin
    if not CharInSet(S[j],['0'..'9']) then
      Exit;
  end;
  Result := TryStrToInt(Copy(S,i,MAXINT),Row);
  if not Result then
    Exit;
  Dec(Row);
  Result := Row <= XLS_MAXROW_97;
end;

function  RefStrToColRow97(const S: AxUCString; out Col,Row: integer): boolean;
var
  i,j,m: integer;
begin
  Result := False;
  if Length(S) < 1 then
    Exit;
  Col := 0;
  m := 1;
  i := Length(S);
  while CharInSet(S[i],['0'..'9']) do
    Dec(i);
  j := i;
  if (j > 0) and (S[j] = '$') then
    Dec(j);
  repeat
    Col := Col + (Ord(S[j]) - Ord('A') + 1) * m;
    m := m * 26;
    Dec(j);
  until ((j = 0) or not (CharInSet(S[j],['A'..'Z'])));
  if (j > 0) and (S[j] = '$') then
    Dec(j);
  Result := j = 0;
  if not Result then
    Exit;
  Dec(Col);
  Result := TryStrToInt(Copy(S,i + 1,MAXINT),Row);
  if Result then begin
    Dec(Row);
    Result := (Col >= 0) and (Col <= XLS_MAXCOL_97) and (Row >= 0) and (ROW <= XLS_MAXROW_97);
  end;
end;

function  AreaStrToColRow97(const S: AxUCString; out Col1,Row1,Col2,Row2: integer): boolean;
var
  p: integer;
  S1,S2: AxUCString;
begin
  p := CPos(':',S);
  if p < 1 then begin
    Result := RefStrToColRow97(S,Col1,Row1);
    Col2 := Col1;
    Row2 := Row1;
  end
  else begin
    S1 := Copy(S,1,p - 1);
    S2 := Copy(S,p + 1,MAXINT);
    Result := RefStrToColRow97(S1,Col1,Row1);
    if Result then
      Result := RefStrToColRow97(S2,Col2,Row2);
  end;
  if not Result then begin
    Result := ColStrToCol97(S1,Col1);
    if Result then begin
      Result := ColStrToCol97(S2,Col2) and (Col2 >= Col1);
      if Result then begin
        Row1 := 1;
        Row2 := XLS_MAXROW_97;
      end;
    end
    else begin
      Result := RowStrToRow97(S1,Row1);
      if Result then begin
        Result := RowStrToRow97(S2,Row2) and (Row2 >= Row1);
        if Result then begin
          Col1 := 1;
          Col2 := XLS_MAXCOL_97;
        end;
      end;
    end;
  end;
end;

function ColStrToCol(const S: AxUCString; out Col: integer): boolean;
var
  i,j: integer;
  m: integer;
begin
  Result := False;
  if Length(S) < 1 then
    Exit;
  j := 0;
  if S[1] = '$' then
    Inc(j);

  m := 1;
  Col := 0;
  i := Length(S);
  repeat
    Col := Col + (Ord(S[i]) - Ord('A') + 1) * m;
    m := m * 26;
    Dec(i);
  until ((i = j) or not (CharInSet(S[i],['A'..'Z'])));
  Dec(Col);
  Result := (i = j) and (Col <= XLS_MAXCOL);
end;

function RowStrToRow(const S: AxUCString; out Row: integer): boolean;
var
  i,j: integer;
begin
  Result := False;
  if Length(S) < 1 then
    Exit;
  i := 1;
  if S[1] = '$' then
    Inc(i);

  for j := i to Length(S) do begin
    if not CharInSet(S[j],['0'..'9']) then
      Exit;
  end;
  Result := TryStrToInt(Copy(S,i,MAXINT),Row);
  if not Result then
    Exit;
  Dec(Row);
  Result := Row <= XLS_MAXROW;
end;

function R1C1Index(const S: AxUCString; out AIndex: integer; out ARel: boolean): boolean;
var
  S2: AxUCString;
  i,j: integer;
begin
  S2 := S;

  Result := Length(S2) > 0;
  if not Result then
    Exit;

  ARel := (S2[1] = '[') and (S2[Length(S2)] = ']');
  if ARel then
    S2 := Copy(S2,2,Length(S2) - 2);

  Result := Length(S2) > 0;
  if not Result then
    Exit;

  j := 1;
  if CharInSet(S2[1],['+','-']) then
    Inc(j);

  Result := j <= Length(S2);
  if not Result then
    Exit;

  for i := j to Length(S2) do begin
    Result := CharInSet(S2[i],['0'..'9']);
    if not Result then
      Exit;
  end;

  AIndex := StrToInt(S2);
  if not ARel then
    Dec(AIndex);
end;

function R1C1ToColRow(const S: AxUCString; out Col,Row: integer; out RelC,RelR: boolean): boolean;
var
  S2: AxUCString;
  pC: integer;
begin
  Col := -1;
  Row := -1;
  Result := S <> '';
  if Result then begin
    S2 := Uppercase(S);
    pC := CPos('C',S2);
    if pC < 1 then begin
      Result := S2[1] = 'R';
      if Result then begin
        if S2 = 'R' then begin
          Row := 0;
          RelR := True;
        end
        else
          Result := R1C1Index(Copy(S2,2,MAXINT),Row,RelR);
      end;
    end
    else if pC = 1 then begin
      if S2 = 'C' then begin
        Col := 0;
        RelC := True;
      end
      else
        Result := R1C1Index(Copy(S2,2,MAXINT),Col,RelC);
    end
    else begin
      Result := S2[1] = 'R';
      if Result then begin
        if Copy(S2,1,pC - 1) = 'R' then begin
          Row := 0;
          RelR := True;
        end
        else
          Result := R1C1Index(Copy(S2,2,pC - 2),Row,RelR);
        if Result then begin
          if Copy(S2,pC,MAXINT) = 'C' then begin
            Col := 0;
            RelC := True;
          end
          else
            Result := R1C1Index(Copy(S2,pC + 1,MAXINT),Col,RelC);
        end;
      end;
    end;
  end;
end;

function R1C1ToArea(const S: AxUCString; out Col1,Row1,Col2,Row2: integer; out RelC1,RelR1,RelC2,RelR2: boolean): boolean;
var
  p: integer;
begin
  p := CPos(':',S);
  if p > 1 then begin
    Result := R1C1ToColRow(Copy(S,1,p - 1),Col1,Row1,RelC1,RelR1);
    if Result then
      Result := R1C1ToColRow(Copy(S,p + 1,MAXINT),Col2,Row2,RelC2,RelR2);
  end
  else begin
    Result := R1C1ToColRow(S,Col1,Row1,RelC1,RelR1);
    if Result then begin
      Col2 := Col1;
      Row2 := Row1;
      RelC2 := RelC1;
      RelR2 := RelR1;
    end;
  end;
  if Col1 = -1 then Col1 := 0;
  if Row1 = -1 then Row1 := 0;
  if Col2 = -1 then Col2 := XLS_MAXCOL;
  if Row2 = -1 then Row2 := XLS_MAXROW;
end;

function ErrorTextToCellError(const S: AxUCString): TXc12CellError;
var
  i: TXc12CellError;
begin
  for i := Low(TXc12CellError) to High(Xc12CellErrorNames) do begin
    if S = Xc12CellErrorNames[i] then begin
      Result := TXc12CellError(i);
      Exit;
    end;
  end;
  Result := errUnknown;
end;

function CellErrorToErrorText(AError: TXc12CellError): AxUCString;
begin
  if (Integer(AError) >= Integer(Low(TXc12CellError))) and (Integer(AError) <= Integer(High(TXc12CellError))) then
    Result := Xc12CellErrorNames[AError]
  else
    Result := Xc12CellErrorNames[errUnknown];
end;

procedure ClipAreaToExtent(var C1,R1,C2,R2: integer);
begin
  C1 := Max(C1,0);
  R1 := Max(R1,0);
  C2 := Min(C2,XLS_MAXCOL);
  R2 := Min(R2,XLS_MAXROW);
end;

function InsideExtent(const C,R: integer): boolean;
begin
  Result := (C >= 0) and (C <= XLS_MAXCOL) and (R >= 0) and (R <= XLS_MAXROW);
end;

function InsideExtent(const C1,R1,C2,R2: integer): boolean;
begin
  Result := (C1 <= XLS_MAXCOL) and (C1 >= 0) and (C2 <= XLS_MAXCOL) and (C2 >= 0) and (R1 <= XLS_MAXROW) and (R1 >= 0) and (R2 <= XLS_MAXROW) and (R2 >= 0);
end;

function ColRowToRefStr(ACol,ARow: integer): AxUCString;
var
  C: integer;
begin
  Result := '';
  if ACol < 26 then
    Result := Result + Char(Ord('A') + ACol)
  else if ACol < 702 then
    Result := Result + Char(Ord('@') + ACol div 26) + Char(Ord('A') + ACol mod 26)
  else begin
    C := Trunc((ACol - 702) / 676);
    Result := Result + Char(Ord('A') + C);
    Dec(ACol,702 + C * 676);
    C := Trunc(ACol / 26);
    Result := Result + Char(Ord('A') + C);
    Dec(ACol,26 * C);
    Result := Result + Char(Ord('A') + ACol);
  end;
  Result := Result + IntToStr(ARow + 1);
end;

function ColRowToRefStr(ACol,ARow: integer; AbsCol,AbsRow: boolean): AxUCString;
var
  C: integer;
begin
  if AbsCol then
    Result := '$'
  else
    Result := '';
  if ACol < 26 then
    Result := Result + Char(Ord('A') + ACol)
  else if ACol < 702 then
    Result := Result + Char(Ord('@') + ACol div 26) + Char(Ord('A') + ACol mod 26)
  else begin
    C := Trunc((ACol - 702) / 676);
    Result := Result + Char(Ord('A') + C);
    Dec(ACol,702 + C * 676);
    C := Trunc(ACol / 26);
    Result := Result + Char(Ord('A') + C);
    Dec(ACol,26 * C);
    Result := Result + Char(Ord('A') + ACol);
  end;
  if AbsRow then
    Result := Result + '$' + IntToStr(ARow + 1)
  else
    Result := Result + IntToStr(ARow + 1);
end;

function ColToRefStr(ACol: integer; const AAbsolute: boolean): AxUCString; overload;
var
  C: integer;
begin
  Result := '';
  if ACol < 26 then
    Result := Result + Char(Ord('A') + ACol)
  else if ACol < 702 then
    Result := Result + Char(Ord('@') + ACol div 26) + Char(Ord('A') + ACol mod 26)
  else begin
    C := Trunc((ACol - 702) / 676);
    Result := Result + Char(Ord('A') + C);
    Dec(ACol,702 + C * 676);
    C := Trunc(ACol / 26);
    Result := Result + Char(Ord('A') + C);
    Dec(ACol,26 * C);
    Result := Result + Char(Ord('A') + ACol);
  end;
  if AAbsolute then
    Result := '$' + Result;
end;

function RowToRefStr(ARow: integer; const AAbsolute: boolean): AxUCString;
begin
  if AAbsolute then
    Result := '$' + IntToStr(ARow + 1)
  else
    Result := IntToStr(ARow + 1);
end;

function ColsToRefStr(ACol1,ACol2: integer; const AAbsolute: boolean): AxUCString;
begin
  Result := ColToRefStr(ACol1,AAbsolute) + ':' + ColToRefStr(ACol2,AAbsolute);
end;

function RowsToRefStr(ARow1,ARow2: integer; const AAbsolute: boolean): AxUCString;
begin
  Result := RowToRefStr(ARow1,AAbsolute) + ':' + RowToRefStr(ARow2,AAbsolute);
end;

function ShortAreaToRefStr(Col1,Row1,Col2,Row2: integer; AAbsolute: boolean = False): AxUCString;
begin
  if (Col1 = Col2) and (Row1 = Row2) then
    Result := ColRowToRefStr(Col1,Row1,AAbsolute,AAbsolute)
  else
    Result := AreaToRefStr(Col1,Row1,Col2,Row2,AAbsolute,AAbsolute,AAbsolute,AAbsolute);
end;

function AreaToRefStr(Col1,Row1,Col2,Row2: integer): AxUCString;
begin
  Result := ColRowToRefStr(Col1,Row1) + ':' + ColRowToRefStr(Col2,Row2);
end;

function AreaToRefStr(Col1,Row1,Col2,Row2: integer; AbsCol1,AbsRow1,AbsCol2,AbsRow2: boolean): AxUCString;
begin
  Result := ColRowToRefStr(Col1,Row1,AbsCol1,AbsRow1) + ':' + ColRowToRefStr(Col2,Row2,AbsCol2,AbsRow2);
end;

function  AreaToRefStrAbs(Col1,Row1,Col2,Row2: integer): AxUCString;
begin
  Result := AreaToRefStr(Col1,Row1,Col2,Row2,True,True,True,True);
end;

function ColRowToRefStrEnc(ACol,ARow: integer): AxUCString;
begin
  Result := ColRowToRefStr(ACol and not COL_ABSFLAG,ARow and not ROW_ABSFLAG,(ACol and COL_ABSFLAG) = COL_ABSFLAG,(ARow and ROW_ABSFLAG) = ROW_ABSFLAG);
end;

function AreaToRefStrEnc(Col1,Row1,Col2,Row2: integer): AxUCString;
begin
  Result := AreaToRefStr(Col1 and not COL_ABSFLAG,Row1 and not ROW_ABSFLAG,Col2 and not COL_ABSFLAG,Row2 and not ROW_ABSFLAG,(Col1 and COL_ABSFLAG) = COL_ABSFLAG,(Row1 and ROW_ABSFLAG) = ROW_ABSFLAG,(Col2 and COL_ABSFLAG) = COL_ABSFLAG,(Row2 and ROW_ABSFLAG) = ROW_ABSFLAG);
end;

function IntersectCellArea(A1, A2: TXLSCellArea; out Dest: TXLSCellArea): Boolean;
begin
  if A1.Col1 >= A2.Col1 then Dest.Col1 := A1.Col1 else Dest.Col1 := A2.Col1;
  if A1.Col2 <= A2.Col2 then Dest.Col2 := A1.Col2 else Dest.Col2 := A2.Col2;
  if A1.Row1 >= A2.Row1 then Dest.Row1 := A1.Row1 else Dest.Row1 := A2.Row1;
  if A1.Row2 <= A2.Row2 then Dest.Row2 := A1.Row2 else Dest.Row2 := A2.Row2;
  Result := (Dest.Col2 >= Dest.Col1) and (Dest.Row2 >= Dest.Row1);
end;

procedure ExtendCellArea(A1, A2: TXLSCellArea; out Dest: TXLSCellArea);
begin
  Dest.Col1 := Min(A1.Col1,A2.Col1);
  Dest.Row1 := Min(A1.Row1,A2.Row1);
  Dest.Col2 := Max(A1.Col2,A2.Col2);
  Dest.Row2 := Max(A1.Row2,A2.Row2);
end;

procedure ExtendCellArea(var AArea: TXLSCellArea; C1,R1,C2,R2: integer);
begin
  AArea.Col1 := Min(AArea.Col1,C1);
  AArea.Row1 := Min(AArea.Row1,R1);
  AArea.Col2 := Max(AArea.Col2,C2);
  AArea.Row2 := Max(AArea.Row2,R2);
end;

function AreaIsRef(AArea: TXLSCellArea): boolean;
begin
  Result := (AArea.Col1 = AArea.Col2) and (AArea.Row1 = AArea.Row2);
end;

function RGBColorToXc12(const ARGB: longword): TXc12Color;
begin
  Result.ColorType := exctRgb;
  Result.OrigRGB := ARGB;
  Result.ARGB := ARGB;
  Result.Tint := 0;
end;

function  RGBColorToXc12(const ARGB: longword; Color: PXc12Color): TXc12Color;
begin
  Color.ColorType := exctRgb;
  Color.OrigRGB := ARGB;
  Color.ARGB := ARGB;
  Color.Tint := 0;
end;

function  ThemeColorToXc12(const AScheme: TXc12ClrSchemeColor; const ATint: double): TXc12Color;
begin
  Result.ColorType := exctTheme;
  Result.Theme := AScheme;
  Result.ARGB := Xc12DefColorScheme[AScheme];
  Result.Tint := Fork(ATint,-1,1);
end;

procedure ThemeColorToXc12(const AColor: PXc12Color; const AScheme: TXc12ClrSchemeColor; const ATint: double);
begin
  AColor.ColorType := exctTheme;
  AColor.Theme := AScheme;
  AColor.ARGB := Xc12DefColorScheme[AScheme];
  AColor.Tint := Fork(ATint,-1,1);
end;

function  ThemeColorToRGB(const AScheme: TXc12ClrSchemeColor; const ATint: double): longword;
var
  Xc12Cl: TXc12Color;
begin
  Xc12Cl.ColorType := exctTheme;
  Xc12Cl.Theme := AScheme;
  Xc12Cl.Tint := ATInt;
  Result := Xc12ColorToRGB(Xc12Cl);
end;

function  _Temp_Xc12IndexColorToRGB(XColor: longword; AutoColor: longword): longword;
begin
  if XColor = EX12_AUTO_COLOR then
    Result := AutoColor
  else
    Result := XColor;
  Result := ((Result and $FF) shl 16) + (Result and $00FF00) + (Result shr 16);
end;

procedure FillDefaultColors;
var
  i: integer;
  j: TXc12ClrSchemeColor;
begin
  for i := 9 to High(TXc12DefaultIndexColorPalette) do
    Xc12IndexColorPalette[i] := TXc12DefaultIndexColorPalette[i];

  for i := 9 to High(TXc12DefaultIndexColorPalette) do
    Xc12IndexColorPaletteRGB[i] := TXc12DefaultIndexColorPaletteRGB[i];

  for j := Low(TXc12ClrSchemeColor) to High(TXc12ClrSchemeColor) do
    Xc12DefColorScheme[j] := Xc12DefColorSchemeVals[j];

  for j := Low(TXc12ClrSchemeColor) to High(TXc12ClrSchemeColor) do
    Xc12DefColorSchemeRGB[j] := Xc12DefColorSchemeValsRGB[j];
end;

procedure NormalizeArea(var C1,R1,C2,R2: integer);
var
  T: integer;
begin
  if C1 > C2 then begin
    T := C1;
    C1 := C2;
    C2 := T;
  end;
  if R1 > R2 then begin
    T := R1;
    R1 := R2;
    R2 := T;
  end;
end;

function  AreaInsideSheet(const C1,R1,C2,R2: integer): boolean;
begin
  Result := (C1 >= 0) and (C1 <= XLS_MAXCOL) and (R1 >= 0) and (R1 <= XLS_MAXROW) and (C2 >= 0) and (C2 <= XLS_MAXCOL) and (R2 >= 0) and (R2 <= XLS_MAXROW);
end;

function ClipAreaToSheet(var C1,R1,C2,R2: integer): boolean;
begin
  if (C1 > XLS_MAXCOL) or (R1 > XLS_MAXROW) or (C2 < 0) or (R2 < 0) then
    Result := False
  else begin
    C1 := Max(C1,0);
    R1 := Max(R1,0);
    C2 := Min(C2,XLS_MAXCOL);
    R2 := Min(R2,XLS_MAXROW);
    Result := True;
  end;
end;

function TColorToClosestIndexColor(Color: TColor): TXc12IndexColor;
var
  i,j: integer;
  C: integer;
  R1,G1,B1: byte;
  R2,G2,B2: byte;
  V1,V2: double;
begin
  j := 8;
  R1 := Color and $FF;
  G1 := (Color and $FF00) shr 8;
  B1 := (Color and $FF0000) shr 16;
  V1 := $FFFFFF;
  for i := 8 to 63 do begin
    C := Xc12IndexColorPalette[i];
    R2 := C and $FF;
    G2 := (C and $FF00) shr 8;
    B2 := (C and $FF0000) shr 16;
    V2 := Abs(R1 - R2) + Abs(G1 - G2) + Abs(B1 - B2);
    if Abs(V2) < Abs(V1) then begin
      V1 := V2;
      j := i;
    end;
  end;
  Result := TXc12IndexColor(j);
end;

function GetBasePtgs(APtg: byte): byte; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
begin
  if APtg > $7F then
    Result := APtg
  else begin
    if (APtg and $40) <> 0 then
      Result := (APtg or $20) and $3F
    else
      Result := APtg and $3F;
  end;
end;

initialization
  G_StrTRUE := 'TRUE';
  G_StrFALSE := 'FALSE';

  FillDefaultColors;


end.
