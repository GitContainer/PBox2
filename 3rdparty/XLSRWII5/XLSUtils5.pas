unit XLSUtils5;

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

uses SysUtils, Classes, Contnrs,
{$ifdef DELPHI_XE5_OR_LATER}
     UITypes,
{$else}
     Graphics,
{$endif}
{$ifdef MSWINDOWS}
     Windows,
{$endif}
     Math;

{$ifdef AXLINUX}
type AnsiString = string;
type AnsiChar = char;
type PAnsiChar = PChar;
{$endif}


const HexCharTable: array[0..$7F] of byte = (
$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
$00,$01,$02,$03,$04,$05,$06,$07,$08,$09,$00,$00,$00,$00,$00,$00,
$00,$0A,$0B,$0C,$0D,$0E,$0F,$00,$00,$00,$00,$00,$00,$00,$00,$00,
$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,
$00,$0A,$0B,$0C,$0D,$0E,$0F,$00,$00,$00,$00,$00,$00,$00,$00,$00,
$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00);

const HexStrs255: array[0..255] of AnsiString = (
'00','01','02','03','04','05','06','07','08','09','0A','0B','0C','0D','0E','0F','10','11','12','13','14','15',
'16','17','18','19','1A','1B','1C','1D','1E','1F','20','21','22','23','24','25','26','27','28','29','2A','2B',
'2C','2D','2E','2F','30','31','32','33','34','35','36','37','38','39','3A','3B','3C','3D','3E','3F','40','41',
'42','43','44','45','46','47','48','49','4A','4B','4C','4D','4E','4F','50','51','52','53','54','55','56','57',
'58','59','5A','5B','5C','5D','5E','5F','60','61','62','63','64','65','66','67','68','69','6A','6B','6C','6D',
'6E','6F','70','71','72','73','74','75','76','77','78','79','7A','7B','7C','7D','7E','7F','80','81','82','83',
'84','85','86','87','88','89','8A','8B','8C','8D','8E','8F','90','91','92','93','94','95','96','97','98','99',
'9A','9B','9C','9D','9E','9F','A0','A1','A2','A3','A4','A5','A6','A7','A8','A9','AA','AB','AC','AD','AE','AF',
'B0','B1','B2','B3','B4','B5','B6','B7','B8','B9','BA','BB','BC','BD','BE','BF','C0','C1','C2','C3','C4','C5',
'C6','C7','C8','C9','CA','CB','CC','CD','CE','CF','D0','D1','D2','D3','D4','D5','D6','D7','D8','D9','DA','DB',
'DC','DD','DE','DF','E0','E1','E2','E3','E4','E5','E6','E7','E8','E9','EA','EB','EC','ED','EE','EF','F0','F1',
'F2','F3','F4','F5','F6','F7','F8','F9','FA','FB','FC','FD','FE','FF');

{$ifdef DELPHI_XE5_OR_LATER}
type TColor = UITypes.TColor;
{$else}
type TColor = Graphics.TColor;
{$endif}

{$ifdef BABOON}
const clDefault =  $20000000;
const clBlack   =  $00000000;
const clFuchsia =  $00FF00FF;
const clPurple  =  $00800080;
const clBlue    =  $000000FF;
const clGreen   =  $0000FF00;
const clRed     =  $00FF0000;
const clYellow  =  $00FFFF00;
const clWhite   =  $00FFFFFF;
{$endif}

const ComponentName = 'XLSReadWriteII';
const CurrentVersionNumber = '6.00.47';
// Max 4 digits
const CurrentBuildNumber = '6047';

// C++ Builder don't like the name MININT as it's already defined.
const XLS_MININT    = -(MAXINT - 1);

//// const MINDOUBLE     = -Infinity;
//const MINDOUBLE     = -(1.0 / 0.0);
////const MAXDOUBLE     = Infinity;
//const MAXDOUBLE     = 1.0 / 0.0;

const EMU_PER_PT = 12700;

const StrDigits = '0123456789';

const XLS_MAGIC_COMPOUND: array[0..5] of byte = ($D0,$CF,$11,$E0,$A1,$B1);
      XLS_MAGIC_ZIP     : array[0..3] of byte = ($50,$4B,$03,$04);
      XLS_MAGIC_RTF     : array[0..4] of byte = ($7B,$5C,$72,$74,$66);

type TXLSKnownFiletype = (xkftUnknown,xkftCompound,xkftZIP,xkftXLSX,xkftDOCX,xkftODT,xkftRTF,
                          xkftXLS5,xkftXLS97,xkftEncryptedXLSX,xkftEncryptedDOCX);

type XLSRWException = class(Exception);

type TDynDoubleArray  = array of double;
type TDoubleArray     = array [0..$FFFF] of double;
type PDoubleArray     = ^TDoubleArray;

type TDynSingleArray  = array of single;
type TSingleArray     = array [0..$FFFF] of single;
type PSingleArray     = ^TSingleArray;

type TDynIntegerArray = array of integer;

type PLongBool = ^LongBool;

{$ifdef DELPHI_2009_OR_LATER}
type XLS8String = RawByteString;
type XLS8Char = AnsiChar;
type XLS8PChar = PAnsiChar;
type AxUCChar = char;
type AxUCString = string;
type AxPUCChar = PChar;
{$else}
type XLS8String = string;
type XLS8Char = Char;
type XLS8PChar = PChar;
type AxUCChar = WideChar;
type AxUCString = WideString;
type AxPUCChar = PWideChar;

type NativeInt = integer;
{$endif}

type TDynStringArray  = array of AxUCString;

{$ifdef DELPHI_2009_OR_LATER}
type TXLSWideStringList = TStringList;
{$endif}

type TXLSPointF = record
     X,Y: double;
     end;

type TXLSRectXY = record
     X1,Y1,X2,Y2: integer;
     end;

type TStringEvent = procedure(ASender: TObject; const AText: AxUCString) of object;
type TIntegerEvent = procedure(ASender: TObject; const AValue: integer) of object;
type TTwoIntegerEvent = procedure(ASender: TObject; const AValue1,AValue2: integer) of object;

type TXLSProgressType  = (xptReadFile,xptWriteFile,xptCompileFormulas,xptCalculateFormulas,xptExport,xptSortCells);
type TXLSProgressState = (xpsBegin,xpsWork,xpsEnd);
type TXLSProgressEvent = procedure (AProgressType: TXLSProgressType; AProgressState: TXLSProgressState; AValue: double) of object;
type TXLSPasswordEvent = procedure(Sender: TObject; var Password: AxUCString) of object;

{$ifndef DELPHI_2009_OR_LATER}
type TSysCharSet = set of AnsiChar;
function  CharInSet(C: WideChar; const CharSet: TSysCharSet): Boolean;
{$endif}
function  CPos(const C: AxUCChar; const S: AxUCString): integer; overload;
function  CPos(const C: AnsiChar; const S: AxUCString): integer; overload;
function  XPos(const ASubStr,AStr: AxUCString; out AP: integer): boolean;
function  RCPos(const C: AxUCChar; const S: AxUCString): integer; overload;
function  RCPos(const C: AnsiChar; const S: AxUCString): integer; overload;
function  SplitAtChar(const C: AxUCChar; var S: AxUCString): AxUCString;
function  SplitAtCRLF(var S: AxUCString; out AFoundCR: boolean): AxUCString;
function  XLSCalcCRC32(const P: PByteArray; Len: integer): longword;
function  EmptyGUID(const AGUID: PGUID): boolean;
function  EmuToCm(const EMU: integer): double;
function  CmToEmu(const Cm: double): integer;
function  RoundToDecimal(x: Extended; const d: Integer): Extended;
function  TryCurrencyStrToFloat(const AValue: AxUCString; out AResult: double): boolean;
procedure SortDoubleArray(AArray: TDynDoubleArray);
function  BytesToHexStr(const ABytes: PByteArray; const ACount: integer): AxUCString;
function  UnicodeIs8Bit(const AString: AxUCString): boolean;
function  _XLSReplaceStr(aSourceString, aFindString, aReplaceString : AxUCString) : AxUCString;
procedure StripQuotes(var S: AxUCString);
function StripCRLF(const S: AxUCString): AxUCString;
function  SafeTryStrToFloat(S: AxUCString; out Value: double): boolean;
// APassword shall be AnsiString
function  MakePasswordHash(const APassword: AnsiString): word;
function  PasswordFromHash(const AHash: word): AxUCString;
function  Fork(const AVal,AMin,AMax: double): double; overload;
function  Fork(const AVal,AMin,AMax: integer): integer; overload;
function  TColorToRGB(const AValue: TColor): longword;
function  RGBToTColor(const ARGB: longword): TColor;
function  BoolAsWord(const ABool: boolean): word; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function  MakeIndent(const AIndent: integer): AxUCString;
{$ifndef DELPHI_7_OR_LATER}
function  PosEx(const FindText,WithinText: AxUCString; const StartIndex: integer): integer;
function IsNan(const AValue: Single): Boolean;
{$endif}
// function  AbsToRelPath(const AbsPath, BasePath: AxUCString): AxUCString;
function  FileTypeFromMagic(AStream: TStream): TXLSKnownFiletype;
procedure StreamToByteStrings(AStream: TStream; const RowLength: integer);
procedure GetAvailableFonts(AList: TStrings);
function  RevRGB(RGB: longword): longword; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function  HSLToRGB(const H, S, L: longword): longword; overload;
function  HSLToRGB(const AHSL: longword): longword; overload;
procedure RGBToHSL(const R, G, B: byte; var H, S, L: integer); overload;
procedure RGBToHSL(const RGB: longword; var H, S, L: integer); overload;
function  MakeGradientColor(const RGB1,RGB2: longword; APercent: double): longword;
function  XLSUTF8Encode(S: AxUCString): AxUCString;
function  _XLSUTF8Encode(S: AxUCString): AxUCString;
function  AdjustDeleteRowsCols(var ARC1,ARC2: integer; AMove1,AMove2: integer): boolean;

{$ifdef DELPHI_5}

type PByte = ^Byte;
type PWord = ^Word;
type PDouble = ^Double;
type PShortint = ^Shortint;
type PSmallint = ^Smallint;
type PSingle = ^Single;
type PInteger = ^Integer;
type PLongword = ^Longword;
type PBoolean = ^Boolean;

type IntegerArray  = array[0..$effffff] of Integer;
type PIntegerArray = ^IntegerArray;

const NaN =  0.0 / 0.0;

function  TryStrToFloat(const S: AxUCString; out AValue: double): boolean; overload;
function  TryStrToFloat(const S: AxUCString; out AValue: extended): boolean; overload;
function  StrToFloatDef(const S: AxUCString; const ADefault: extended): double; overload;
function  TryStrToInt(const S: AxUCString; out AValue: integer): boolean; overload;
function  TryStrToDateTime(const S: AxUCString; out AValue: TDateTime): boolean; overload;
function  TryStrToDate(const S: AxUCString; out AValue: TDateTime): boolean; overload;
function  TryStrToTime(const S: AxUCString; out AValue: TDateTime): boolean; overload;
function  SameValue(const A, B: Double; Epsilon: Double): Boolean;
function  BoolToStr(const B: boolean; const AUseBoolStrs: boolean): AxUCString;
function  StringsToDelimited(AList: TStrings; const ADelimiter: AxUCChar): AxUCString;
procedure DelimitedToStrings(AText: AxUCString; AList: TStrings; const ADelimiter: AxUCChar);
function  Sign(const AValue: Double): integer;
function  RightStr(const AText: AxUCString; const ACount: Integer): AxUCString; overload;
function  SimpleRoundTo(const AValue: Single; const ADigit: integer = -2): Single;
function  UTF8Encode(S: AxUCString): AxUCString;

type PPHashItem = ^PHashItem;
  PHashItem = ^THashItem;
  THashItem = record
    Next: PHashItem;
    Key: string;
    Value: Integer;
  end;

type TStringHash = class
  private
    Buckets: array of PHashItem;
  protected
    function Find(const Key: string): PPHashItem;
    function HashOf(const Key: string): Cardinal; virtual;
  public
    constructor Create(Size: Integer = 256);
    destructor Destroy; override;
    procedure Add(const Key: string; Value: Integer);
    procedure Clear;
    procedure Remove(const Key: string);
    function Modify(const Key: string; Value: Integer): Boolean;
    function ValueOf(const Key: string): Integer;
  end;
{$endif}

type TXLSErrorLevel = (xelMessage,xelHint,xelWarning,xelError,xelFatal,xelDebug);
     TXLSErrorLevels = set of TXLSErrorLevel;

type TXLSErrorEvent = procedure(ASender: TObject; ALevel: TXLSErrorLevel; AText: AxUCString) of object;

// Error numbers:
// 1000 - 1999: File read errors
// 2000 - 2999: File write errors
// 3000 - 3999: Runtime (Access, uknown name, etc).
// 4000 - 4999: Formula errors.
// 5000 - 5999: BIFF file errors.
// 6000 - 6999: General errors.

const XLSERR_FILEREAD_UNKNOWENCRYPT          = 1002;
const XLSERR_FILEREAD_PASSWORDMISSING        = 1003;
const XLSERR_FILEREAD_WRONGPASSWORD          = 1004;

const XLSWARN_FILEREAD_XMLPARSE              = 2001;

const XLSERR_DIRECTWR_CELL_NOT_INC           = 1005;

const XLSHINT_BAD_FORMULA_RESULT             = 7001;

// TODO 3001
const XLSWARN_CIRCULAR_FORMULA               = 3002;
const XLSWARN_USEDSHEETNAME                  = 3003;
const XLSWARN_INVALIDCELLREF                 = 3004;
const XLSWARN_INVALIDCELLVALUE               = 3005;
const XLSWARN_UNKNOWNIMAGE                   = 3006;
const XLSWARN_UNSUPPORTEDIMAGE               = 3007;
const XLSWARN_IMAGEERROR                     = 3008;
const XLSWARN_CAN_NOT_CHANGE_MERGED          = 3009;

const XLSWARN_SHAPE_EDIT_TEXTBOX             = 3101;
const XLSWARN_SHAPE_CAN_NOT_EDIT             = 3102;

const XLSERR_FMLA_BADERRORCONST              = 4001;
const XLSERR_FMLA_BADEXPONENT                = 4002;
const XLSERR_FMLA_BADFRACTIONAL              = 4003;
const XLSERR_FMLA_MISSINGQUOTE               = 4004;
const XLSERR_FMLA_BADARRAYCONST              = 4005;
const XLSERR_FMLA_FORMULA                    = 4006;
const XLSERR_FMLA_MISSINGRPAR                = 4007;
const XLSERR_FMLA_INVALIDUSEOF               = 4008;
const XLSERR_FMLA_MISSINGEXPR                = 4009;
const XLSERR_FMLA_BADCELLREF                 = 4010;
const XLSERR_FMLA_BADSHEETNAME               = 4011;
const XLSERR_FMLA_BADWORKBOOKNAME            = 4012;
const XLSERR_FMLA_MISSING                    = 4013;
const XLSERR_FMLA_BADR1C1REF                 = 4014;
const XLSERR_FMLA_MISSINGOPERAND             = 4015;
const XLSERR_FMLA_UNKNOWNSHEET               = 4016;
const XLSERR_FMLA_UNKNOWNXBOOK               = 4017;
const XLSERR_FMLA_BADUSEOFOP                 = 4018;
const XLSERR_FMLA_UNKNOWNNAME                = 4019;
const XLSERR_FMLA_MISSINGARG                 = 4020;
const XLSERR_FMLA_TOMANYGARGS                = 4021;
const XLSERR_FMLA_BADARRAYCOLS               = 4022;

const XLSERR_FMLA_BADTABLE                   = 4023;
const XLSERR_FMLA_UNKNOWNTABLE               = 4024;
const XLSERR_FMLA_UNKNOWNTABLECOL            = 4025;
const XLSERR_FMLA_UNKNOWNTBLSPEC             = 4026;
const XLSERR_FMLA_TABLESQBRACKET             = 4027;
const XLSERR_FMLA_BADUSEOFTABLEOP            = 4028;
const XLSERR_FMLA_BADUSEOFTBLSPEC            = 4029;
const XLSERR_FMLA_TBLSPECNOTFIRST            = 4030;

const XLSERR_NAME_EMPTY                      = 4031;
const XLSERR_NAME_DUPLICATE                  = 4032;

const XLSERR_HLINK_INVALID_CELLREF           = 4033;

const XLSERR_BIFF_BADBUILTIN_NAME            = 5001;
const XLSERR_BIFF_CANT_FIND_BUILTIN_NAME     = 5002;

const XLSERR_FILENAME_IS_MISSING             = 6001;
const XLSERR_NO_XLS_DEFINED                  = 6002;
const XLSWARN_WORKBOOK_NOT_EMPTY             = 6003;
const XLSWARN_MULTIPLE_SEL_AREAS             = 6004;
const XLSERR_INVALID_SIMPLE_TAG              = 6005;

type TXLSErrorManager = class(TPersistent)
private
     function GetErrorCount: integer;
     function GetFatalCount: integer;
     function GetHintCount: integer;
     function GetWarningCount: integer;
     procedure SetErrorList(const Value: TStrings);
protected
     FErrorEvent: TXLSErrorEvent;
     FFilter: TXLSErrorLevels;
     FSaveToList: boolean;
     FErrorList: TStrings;
     FOwnsErrorList: boolean;
     FLastError: integer;

     FHintCount: integer;
     FWarningCount: integer;
     FErrorCount: integer;
     FFatalCount: integer;

     FIgnoreErrors: boolean;
     FIgnoreXMLWarnings: boolean;

     function  GetErrorText(const AText: AxUcString; AErrorNo: integer): AxUcString;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure Message(const AText: AxUcString);
     procedure Hint(const AText: AxUcString; AErrorNo: integer);
     procedure Warning(const AText: AxUcString; AErrorNo: integer);
     procedure Error(const AText: AxUcString; AErrorNo: integer);
     procedure Fatal(const AText: AxUcString; AErrorNo: integer);
     procedure Debug(const AText: AxUcString);

     property ErrorList: TStrings read FErrorList write SetErrorList;
     // Usefull when testing things such if a formula is valid or not.
     // When True, only LastError is updated.
     property IgnoreErrors: boolean read FIgnoreErrors write FIgnoreErrors;
     property IgnoreXMLWarnings: boolean read FIgnoreXMLWarnings write FIgnoreXMLWarnings;
published
     property LastError: integer read FLastError write FLastError;
     property SaveToList: boolean read FSaveToList write FSaveToList;
     // Fatal errors can not be filtered.
     property Filter: TXLSErrorLevels read FFilter write FFilter;
     property HintCount: integer read GetHintCount;
     property WarningCount: integer read GetWarningCount;
     property ErrorCount: integer read GetErrorCount;
     property FatalCount: integer read GetFatalCount;

     property OnError: TXLSErrorEvent read FErrorEvent write FErrorEvent;
     end;

type TIntegerList = class(TList)
private
     function GetItems(Index: integer): integer;
protected
public
     procedure Add(AValue: integer);

     property Items[Index: integer]: integer read GetItems; default;
     end;

type TIntegerStack = class(TIntegerList)
protected
public
     procedure Push(const AValue: integer);
     function  Pop: integer;
     function  Peek: integer;
     end;

type TIndexObject = class(TObject)
protected
     FIndex: integer;
public
     function RequestObject(AClass: TClass): TObject; virtual;

     property Index: integer read FIndex;
     end;

type TIndexObjectList = class(TObjectList)
protected
     FIsClearing: boolean;

     procedure Notify(Ptr: Pointer; Action: TListNotification); override;
public
     // Use ReIndex when items are inserted. Normally Notify is fired with
     // Add notification when an item is inserted. As very few objects
     // inserts items (worksheets), and by using ReIndex a whole reindex
     // is not required when items that never is inserted such as style
     // items are added.
     procedure ReIndex;
     end;

type TPointerMemoryStream = class(TMemoryStream)
protected
public
     procedure SetStreamData(Ptr: Pointer; const ASize: NativeInt);
     end;

{$ifndef DELPHI_XE_OR_LATER}
type TFormatSettings = class(TObject)
private
     function  GetListSeparator: AxUCChar;
     procedure SetDecimalSeparator(const Value: AxUCChar);
     procedure SetListSeparator(const Value: AxUCChar);
     procedure SetThousandSeparator(const Value: AxUCChar);
     function  GetDecimalSeparator: AxUCChar;
     function  GetThousandSeparator: AxUCChar;protected
public
     CurrencyString: AxUCString;
     CurrencyFormat: Byte;
     CurrencyDecimals: Byte;
     DateSeparator: AxUCChar;
     TimeSeparator: AxUCChar;
     ShortDateFormat: AxUCString;
     LongDateFormat: AxUCString;
     TimeAMString: AxUCString;
     TimePMString: AxUCString;
     ShortTimeFormat: AxUCString;
     LongTimeFormat: AxUCString;
     ShortMonthNames: array[1..12] of AxUCString;
     LongMonthNames: array[1..12] of AxUCString;
     ShortDayNames: array[1..7] of AxUCString;
     LongDayNames: array[1..7] of AxUCString;
     TwoDigitYearCenturyWindow: Word;
     NegCurrFormat: Byte;

     property ListSeparator: AxUCChar read GetListSeparator write SetListSeparator;
     property ThousandSeparator: AxUCChar read GetThousandSeparator write SetThousandSeparator;
     property DecimalSeparator: AxUCChar read GetDecimalSeparator write SetDecimalSeparator;
     end;

var FormatSettings: TFormatSettings;
{$endif}

var G_Counter: integer;
var G_DirSepChar: AxUCChar;

implementation

{$ifdef BABOON}
function MulDiv(V1,V2,V3: integer): integer;
var
  R: int64;
begin
  R := V1 * V2;
  Result := R div V3;
end;
{$endif}

const
  CRC32_POLYNOMIAL = $EDB88320;
var
  Ccitt32Table: array[0..255] of longword;
{$ifndef DELPHI_XE_OR_LATER}
  init_i: integer;
{$endif}

procedure BuildCRCTable;
var i, j: longint;
    value: longword;
begin
  for i := 0 to 255 do begin
    value := i;
    for j := 8 downto 1 do
      if ((value and 1) <> 0) then
        value := (value shr 1) xor CRC32_POLYNOMIAL
      else
        value := value shr 1;
    Ccitt32Table[i] := value;
  end
end;

{$ifndef DELPHI_2009_OR_LATER}
function CharInSet(C: WideChar; const CharSet: TSysCharSet): Boolean;
begin
  Result := not ((Word(C) and $FF00) <> 0);
  if Result then
    Result := AnsiChar(C) in CharSet;
end;
{$endif}

function CPos(const C: AxUCChar; const S: AxUCString): integer;
begin
  for Result := 1 to Length(S) do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function  CPos(const C: AnsiChar; const S: AxUCString): integer; overload;
begin
  Result := CPos(AxUCChar(C),S);
end;

function  XPos(const ASubStr,AStr: AxUCString; out AP: integer): boolean;
begin
  AP := Pos(ASubStr,AStr);
  Result := AP >= 1;
end;

function RCPos(const C: AxUCChar; const S: AxUCString): integer;
begin
  for Result := Length(S) downto 1 do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function  RCPos(const C: AnsiChar; const S: AxUCString): integer;
begin
  Result := RCPos(AxUCChar(C),S);
end;

function SplitAtChar(const C: AxUCChar; var S: AxUCString): AxUCString;
var
  p: integer;
begin
  p := CPos(C,S);
  if p > 0 then begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

function  SplitAtCRLF(var S: AxUCString; out AFoundCR: boolean): AxUCString;
var
  p: integer;
begin
  for p := 1 to Length(S) do begin
    if (S[p] = #$000A) or (S[p] = #$000D) then begin
      Result := Copy(S,1,p - 1);
      if (p < Length(S)) and (S[p] = #$000D) and (S[p + 1] = #$000A) then
        S := Copy(S,p + 2,MAXINT)
      else
        S := Copy(S,p + 1,MAXINT);
      AFoundCR := True;
      Exit;
    end;
  end;
  AFoundCR := False;
  Result := S;
  S := '';
end;

function XLSCalcCRC32(const P: PByteArray; Len: integer): longword;
var j: integer;
begin
  Result := $FFFFFFFF;
  for j:=0 to Len - 1 do
    Result:= (((Result shr 8) and $00FFFFFF) xor (Ccitt32Table[(Result xor P[j]) and $FF]));
end;

function EmptyGUID(const AGUID: PGUID): boolean;
var
  i: integer;
  P: PByteArray;
begin
  P := PByteArray(AGUID);
  for i := 0 to SizeOf(TGUID) - 1 do begin
    if P[i] <> 0 then begin
      Result := False;
      Exit;
    end;
  end;
  Result := True;
end;

// EMU (English Metric Units) to Centimeters
function EmuToCm(const EMU: integer): double;
begin
  Result := Emu / 360000;
end;

function CmToEmu(const Cm: double): integer;
begin
  Result := Round(Cm * 360000);
end;

function RoundToDecimal(x: Extended; const d: Integer): Extended;
// RoundD(123.456, 0) = 123.00
// RoundD(123.456, 2) = 123.46
// RoundD(123456, -3) = 123000
var
  n: Extended;
begin
  n := IntPower(10, d);
  x := x * n;
  Result := (Int(x) + Int(Frac(x) * 2)) / n;
end;

function TryCurrencyStrToFloat(const AValue: AxUCString; out AResult: double): boolean;
var
  L: integer;
  S: AxUCString;
begin
  Result := False;
  S := Trim(AValue);
  L := Length(FormatSettings.CurrencyString);
  if FormatSettings.CurrencyFormat in [0,2] then begin
    if Copy(S,1,L) = FormatSettings.CurrencyString then
      Result := TryStrToFloat(Copy(S,L + 1,MAXINT),AResult);
  end
  else begin
    if Copy(S,1,Length(S) - L + 1) = FormatSettings.CurrencyString then
      Result := TryStrToFloat(Copy(S,1,Length(S) - L),AResult);
  end;
end;

procedure SortDoubleArray(AArray: TDynDoubleArray);

procedure QSortArray(iLo, iHi: Integer) ;
var
  Lo,Hi: Integer;
  Pivot,T: double;
begin
  Lo := iLo;
  Hi := iHi;
  Pivot := AArray[(Lo + Hi) div 2];
  repeat
    while AArray[Lo] < Pivot do Inc(Lo) ;
    while AArray[Hi] > Pivot do Dec(Hi) ;
     if Lo <= Hi then begin
       T := AArray[Lo];
       AArray[Lo] := AArray[Hi];
       AArray[Hi] := T;
       Inc(Lo) ;
       Dec(Hi) ;
     end;
   until Lo > Hi;
   if Hi > iLo then QSortArray(iLo,Hi) ;
   if Lo < iHi then QSortArray(Lo,iHi) ;
 end;

 begin
   if Length(AArray) > 1 then
     QSortArray(0,High(AArray));
 end;

function BytesToHexStr(const ABytes: PByteArray; const ACount: integer): AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to ACount - 1 do
    Result := Result + Format('%.2X,',[ABytes[i]]);
  Result := Copy(Result,1,Length(Result) - 1);
end;

function  UnicodeIs8Bit(const AString: AxUCString): boolean;
var
  sugga: integer;
begin
  for sugga := 1 to Length(AString) do begin
    if Word(AString[sugga]) > $00FF then begin
      Result := False;
      Exit;
    end;
  end;
  Result := True;
end;

function _XLSReplaceStr(aSourceString, aFindString, aReplaceString : AxUCString) : AxUCString;
var
  SearchStr, Patt, NewStr: AxUCString;
  Offset: Integer;
begin
  SearchStr := aSourceString;
  Patt := aFindString;
  NewStr := aSourceString;
  Result := '';
  while SearchStr <> '' do begin
    Offset := Pos(Patt, SearchStr);
    if Offset = 0 then begin
      Result := Result + NewStr;
      Break;
    end;
    Result := Result + Copy(NewStr, 1, Offset - 1) + aReplaceString;
    NewStr := Copy(NewStr, Offset + Length(aFindString), MAXINT);
    SearchStr := Copy(SearchStr, Offset + Length(Patt), MaxInt);
  end;
end;

procedure StripQuotes(var S: AxUCString);
var
  C1,C2: AxUCChar;
  L: integer;
begin
  L := Length(S);
  if L >= 2 then begin
    C1 := S[1];
    C2 := S[L];
    if ((C1 = '"') and (C2 = '"')) or ((C1 = '''') and (C2 = '''')) then
      S := Copy(S,2,L - 2);
  end;
end;

function StripCRLF(const S: AxUCString): AxUCString;
var
  i,j: integer;
begin
  j := 0;
  SetLength(Result,Length(S));
  for i := 1 to Length(S) do begin
    if S[i] = #10 then begin
      Inc(j);
      Result[i] := ' ';
    end
    else if S[i] <> #13 then begin
      Inc(j);
      Result[i] := S[i];
    end;
  end;
  SetLength(Result,j);
end;

function SafeTryStrToFloat(S: AxUCString; out Value: double): boolean;
var
  TempDS: AxUCChar;
begin
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
{$ifdef DELPHI_5}
    try
      StrToFloat(S);
      Result := True;
    except
      Result := False;
    end;
{$else}
    Result := TryStrToFloat(S,Value);
{$endif}
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

function MakePasswordHash(const APassword: AnsiString): word;
var
  C,T: word;
  S: AnsiString;
  i: integer;
begin
  Result := 0;
  S := Copy(APassword,1,15);
  for i := 1 to Length(S) do begin
    T := Byte(S[i]);
    C := ((T shl i) + (T shr (15 - i))) and $7FFF;
    Result := Result xor C;
  end;
  Result := Result xor Length(S) xor $CE4B;
end;

function PasswordFromHash(const AHash: word): AxUCString;
var
  Chars,S: AnsiString;
  i,j,k: integer;
begin
  Result := '';
  Chars := 'abcdefghijklmnopqrstuwvxyzABCDEFGHIJKLMNOPQRSTUWVXYZ0123456789_';
  Randomize;
  for k := 1 to 1000000 do begin
    i := 12; // Password length. Don't make it too short as that will fail.
    SetLength(S,i);
    for j := 1 to i do
      S[j] := Chars[Random(Length(Chars)) + 1];
    if AHash = MakePasswordHash(S) then begin
      Result := AxUCString(S);
      Exit;
    end;
  end;
end;

function  Fork(const AVal,AMin,AMax: double): double;
begin
  if AVal < AMin then
    Result := AMin
  else if AVal > AMax then
    Result := AMax
  else
    Result := AVal;
end;

function  Fork(const AVal,AMin,AMax: integer): integer; overload;
begin
  if AVal < AMin then
    Result := AMin
  else if AVal > AMax then
    Result := AMax
  else
    Result := AVal;
end;

function TColorToRGB(const AValue: TColor): longword;
begin
  Result := ((AValue and $000000FF) shl 16) + (AValue and $0000FF00) + ((AValue and $00FF0000) shr 16);
end;

function  RGBToTColor(const ARGB: longword): TColor;
begin
  Result := ((ARGB and $000000FF) shl 16) + (ARGB and $0000FF00) + ((ARGB and $00FF0000) shr 16);
end;

function  BoolAsWord(const ABool: boolean): word;
begin
  Result := Word(ABool);
end;

function MakeIndent(const AIndent: integer): AxUCString;
var
  i: integer;
begin
  SetLength(Result,AIndent * 2);
  for i := 1 to AIndent * 2 do
    Result[i] := ' ';
end;

{$ifndef DELPHI_7_OR_LATER}
function  PosEx(const FindText,WithinText: AxUCString; const StartIndex: integer): integer;
begin
  raise XLSRWException.Create('TODO D6');
end;

function IsNan(const AValue: Single): Boolean;
begin
  Result := ((PLongWord(@AValue)^ and $7F800000)  = $7F800000) and
            ((PLongWord(@AValue)^ and $007FFFFF) <> $00000000);
end;

{$endif}

{$ifdef DELPHI_5}
function  TryStrToFloat(const S: AxUCString; out AValue: double): boolean;
begin
  try
    AValue := StrToFloat(S);
    Result := True;
  except
    Result := False;
  end;
end;

function  TryStrToFloat(const S: AxUCString; out AValue: extended): boolean;
begin
  try
    AValue := StrToFloat(S);
    Result := True;
  except
    Result := False;
  end;
end;

function  StrToFloatDef(const S: AxUCString; const ADefault: extended): double; overload;
begin
  try
    Result := StrToFloat(S);
  except
    Result := ADefault;
  end;
end;

function  TryStrToInt(const S: AxUCString; out AValue: integer): boolean; overload;
begin
  if S = '' then
    Result := False
  else begin
    try
      AValue := StrToInt(S);
      Result := True;
    except
      Result := False;
    end;
  end;
end;

function  TryStrToDateTime(const S: AxUCString; out AValue: TDateTime): boolean; overload;
begin
  try
    AValue := StrToDateTime(S);
    Result := True;
  except
    Result := False;
  end;
end;

function  TryStrToDate(const S: AxUCString; out AValue: TDateTime): boolean; overload;
begin
  try
    AValue := StrToDate(S);
    Result := True;
  except
    Result := False;
  end;
end;

function  TryStrToTime(const S: AxUCString; out AValue: TDateTime): boolean; overload;
begin
  try
    AValue := StrToTime(S);
    Result := True;
  except
    Result := False;
  end;
end;

function SameValue(const A, B: Double; Epsilon: Double): Boolean;
begin
  if Epsilon = 0 then
    Epsilon := Max(Min(Abs(A), Abs(B)) * 1E-15 * 1000, 1E-15 * 1000);
  if A > B then
    Result := (A - B) <= Epsilon
  else
    Result := (B - A) <= Epsilon;
end;

function BoolToStr(const B: boolean; const AUseBoolStrs: boolean): AxUCString;
begin
  if B then
    Result := 'True'
  else
    Result := 'False';
end;

function StringsToDelimited(AList: TStrings; const ADelimiter: AxUCChar): AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to AList.Count - 1 do
    Result := Result + AList[i] + ADelimiter;

  if Result <> '' then
    Result := Copy(Result,1,Length(Result) - 1);
end;

procedure DelimitedToStrings(AText: AxUCString; AList: TStrings; const ADelimiter: AxUCChar);
var
  S: AxUCString;
begin
  while AText <> '' do begin
    S := SplitAtChar(ADelimiter,AText);
    if S <> '' then
      AList.Add(S);
  end;
end;

function Sign(const AValue: Double): integer;
begin
  if ((PInt64(@AValue)^ and $7FFFFFFFFFFFFFFF) = $0000000000000000) then
    Result := 0
  else if ((PInt64(@AValue)^ and $8000000000000000) = $8000000000000000) then
    Result := -1
  else
    Result := 1;
end;

function  RightStr(const AText: AxUCString; const ACount: Integer): AxUCString; overload;
begin
  Result := Copy(AText, Length(AText) + 1 - ACount, ACount);
end;

function  SimpleRoundTo(const AValue: Single; const ADigit: integer = -2): Single;
var
  LFactor: Extended;
begin
  LFactor := IntPower(10.0, ADigit);
  if AValue < 0 then
    Result := Int((AValue / LFactor) - 0.5) * LFactor
  else
    Result := Int((AValue / LFactor) + 0.5) * LFactor;
end;

function  UTF8Encode(S: AxUCString): AxUCString;

function UnicodeToUtf8(Dest: PAnsiChar; MaxDestBytes: Cardinal; Source: PWideChar; SourceChars: Cardinal): Cardinal;
begin
  Result := 0;
  if Source = nil then Exit;
  if Dest <> nil then
  begin
    Result := Cardinal(WideCharToMultiByte(CP_UTF8, 0, Source, Integer(SourceChars), Dest, Integer(MaxDestBytes), nil, nil));
    if (Result > 0) and (Result <= MaxDestBytes) and (Dest[Result - 1] <> #0) then
    begin
      if Result = MaxDestBytes then
      begin
        while (Result > 1) and (Byte(Dest[Result - 1]) > $7F) and (Byte(Dest[Result - 1]) and $80 <> 0) and (Byte(Dest[Result - 1]) and $C0 <> $C0) do
          Dec(Result);
      end else
        Inc(Result);
      Dest[Result - 1] := #0;
    end;
  end else
    Result := Cardinal(WideCharToMultiByte(CP_UTF8, 0, Source, Integer(SourceChars), nil, 0, nil, nil));
end;

function _Utf8Encode(const WS: WideString): string;
var
  L: Integer;
  Temp: string;
begin
  Result := '';
  if WS = '' then Exit;
  L := Length(WS);
  SetLength(Temp, L * 3); // SetLength includes space for null terminator

  L := UnicodeToUtf8(PAnsiChar(Temp), Length(Temp) + 1, PWideChar(WS), L);
  if L > 0 then
    SetLength(Temp, L - 1)
  else
    Temp := '';
  Result := Temp;
//  if Result <> '' then
//    PStrRec(Integer(Result) - SizeOf(StrRec)).codePage := CP_UTF8;
end;

begin
  Result := _Utf8Encode(S);
end;

{ TStringHash }

procedure TStringHash.Add(const Key: string; Value: Integer);
var
  Hash: Integer;
  Bucket: PHashItem;
begin
  Hash := HashOf(Key) mod Cardinal(Length(Buckets));
  New(Bucket);
  Bucket^.Key := Key;
  Bucket^.Value := Value;
  Bucket^.Next := Buckets[Hash];
  Buckets[Hash] := Bucket;
end;

procedure TStringHash.Clear;
var
  I: Integer;
  P, N: PHashItem;
begin
  for I := 0 to Length(Buckets) - 1 do
  begin
    P := Buckets[I];
    while P <> nil do
    begin
      N := P^.Next;
      Dispose(P);
      P := N;
    end;
    Buckets[I] := nil;
  end;
end;

constructor TStringHash.Create(Size: Integer);
begin
  inherited Create;
  SetLength(Buckets, Size);
end;

destructor TStringHash.Destroy;
begin
  Clear;
  inherited;
end;

function TStringHash.Find(const Key: string): PPHashItem;
var
  Hash: Integer;
begin
  Hash := HashOf(Key) mod Cardinal(Length(Buckets));
  Result := @Buckets[Hash];
  while Result^ <> nil do
  begin
    if Result^.Key = Key then
      Exit
    else
      Result := @Result^.Next;
  end;
end;

function TStringHash.HashOf(const Key: string): Cardinal;
var
  I: Integer;
begin
  Result := 0;
  for I := 1 to Length(Key) do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor
      Ord(Key[I]);
end;

function TStringHash.Modify(const Key: string; Value: Integer): Boolean;
var
  P: PHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
  begin
    Result := True;
    P^.Value := Value;
  end
  else
    Result := False;
end;

procedure TStringHash.Remove(const Key: string);
var
  P: PHashItem;
  Prev: PPHashItem;
begin
  Prev := Find(Key);
  P := Prev^;
  if P <> nil then
  begin
    Prev^ := P^.Next;
    Dispose(P);
  end;
end;

function TStringHash.ValueOf(const Key: string): Integer;
var
  P: PHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
    Result := P^.Value else
    Result := -1;
end;
{$endif}

//function  AbsToRelPath(const AbsPath, BasePath: AxUCString): AxUCString;
//var
//  Path: array[0..MAX_PATH-1] of char;
//begin
//  PathRelativePathTo(@Path[0], PChar(BasePath), FILE_ATTRIBUTE_DIRECTORY, PChar(AbsPath), 0);
//  result := Path;
//end;

function TestMagic(AStream: TStream; AMagic: array of byte): boolean;
var
  i  : integer;
  Buf: array[0..15] of byte;
begin
  Result := Length(AMagic) <= Length(Buf);
  if Result then begin
    AStream.Read(Buf[0],Length(AMagic));
    for i := 0 to High(AMagic) do begin
      if AMagic[i] <> Buf[i] then begin
        AStream.Seek(0,soBeginning);
        Result := False;
        Exit;
      end;
    end;
  end;

  AStream.Seek(0,soBeginning);
end;

function FileTypeFromMagic(AStream: TStream): TXLSKnownFiletype;
begin
       if TestMagic(AStream,XLS_MAGIC_COMPOUND) then Result := xkftCompound
  else if TestMagic(AStream,XLS_MAGIC_ZIP)      then Result := xkftZIP
  else if TestMagic(AStream,XLS_MAGIC_RTF)      then Result := xkftRTF
  else                                               Result := xkftUnknown;
end;

procedure StreamToByteStrings(AStream: TStream; const RowLength: integer);
var
  b: byte;
  S,S2: AxUCString;
begin
  S := '';
  S2 := '';
  while AStream.Read(b,1) = 1 do begin
    S2 := S2 + Format('$%.2X,',[b]);
    if Length(S2) > RowLength then begin
      S := S + S2 + #13 + #10;
      S2 := '';
    end;
  end;
end;

{$ifdef BABOON}
procedure GetAvailableFonts(AList: TStrings);
begin
  AList.Add('Times');
  AList.Add('Helvetica');
  AList.Add('Courier');
end;
{$else}
function EnumFontsProc(var LogFont: TLogFont; var TextMetric: TTextMetric; FontType: Integer; Data: Pointer): Integer; stdcall;
var
  S: TStrings;
//  P: ^ENUMLOGFONTEX;
  Temp: AxUCString;
begin
  S := TStrings(Data);
//  P := @LogFont;
//  Temp := P.elfStyle;
  Temp := LogFont.lfFaceName;
  // @ = Exclude vertical fonts.
  if ((S.Count = 0) or (AnsiCompareText(S[S.Count - 1], Temp) <> 0)) and (Temp[1] <> '@') then
    S.AddObject(Temp,TObject(FontType));
  Result := 1;
end;

procedure GetAvailableFonts(AList: TStrings);
var
  DC: HDC;
  LF: TLogFont;
begin
  DC := GetDC(0);
  FillChar(LF,SizeOf(TLogFont),0);
  LF.lfCharset := DEFAULT_CHARSET;
  EnumFontFamiliesEx(DC, LF, @EnumFontsProc, Windows.LPARAM(AList), 0);
  ReleaseDC(0, DC);
end;
{$endif}

function RevRGB(RGB: longword): longword; {$ifdef D2006PLUS} inline; {$endif}
begin
  Result := ((RGB and $FF0000) shr 16) + (RGB and $00FF00) + ((RGB and $0000FF) shl 16);
end;

function HSLToRGB(const H, S, L: longword): longword;
var
 M1,M2: double;
 R,G,B: byte;

function HueToColorValue(Hue: double): byte;
var
  V : double;
begin
 if Hue < 0 then
   Hue := Hue + 1
 else if Hue > 1 then
   Hue := Hue - 1;
 if 6 * Hue < 1 then
   V := M1 + (M2 - M1) * Hue * 6
 else begin
   if 2 * Hue < 1 then
     V := M2
   else begin
     if 3 * Hue < 2 then
       V := M1 + (M2 - M1) * (2/3 - Hue) * 6
     else
       V := M1;
   end;
 end;
 Result := Round(255 * V);
end;

begin
 if S = 0 then begin
   R := Round($FF * L);
   G := R;
   B := R;
 end
 else begin
   if L <= 0.5 then
    M2 := L * (1 + S)
   else
    M2 := L + S - L * S;
   M1 := 2 * L - M2;
   R := HueToColorValue (H + 1/3);
   G := HueToColorValue (H);
   B := HueToColorValue (H - 1/3)
  end;
  Result := (R shl 16) + (G shl 8) + B;
end;

function HSLToRGB(const AHSL: longword): longword;
var
 H,S,L: byte;
begin
  H := (AHSL and $00FF0000) shr 16;
  S := (AHSL and $0000FF00) shr 8;
  L := AHSL and $000000FF;
  Result := HSLToRGB(H,S,L);
end;

procedure RGBToHSL(const R, G, B: byte; var H, S, L: integer);

function RGBMaxValue(const R, G, B: byte): byte;
begin
  Result := R;
  if (Result < G) then Result := G;
  if (Result < B) then Result := B;
end;

function RGBMinValue(const R, G, B: byte) : byte;
begin
  Result := R;
  if (Result > G) then Result := G;
  if (Result > B) then Result := B;
end;

var
 Delta, Min: byte;
begin
 L := RGBMaxValue(R,G,B);
 Min := RGBMinValue(R,G,B);
 Delta := L - Min;
 if L = Min then begin
   H := 0;
   S := 0;
 end
 else begin
   S := MulDiv(Delta, 255, L);
   if (R = L) then
     H := MulDiv(60, G - B, Delta)
   else begin
    if (G = L) then
      H := MulDiv(60, B - R, Delta) + 120
    else
     if (B = L) then
       H := MulDiv(60, R - G, Delta) + 240;
   end;
   if (H < 0) then
     H := H + 360;
  end;
end;

procedure RGBToHSL(const RGB: longword; var H, S, L: integer);
var
  R,G,B: byte;
begin
  R := (RGB and $00FF0000) shr 16;
  G := (RGB and $0000FF00) shr 8;
  B := RGB and $000000FF;
  RGBToHSL(R,G,B,H,S,L);
end;

function MakeGradientColor(const RGB1,RGB2: longword; APercent: double): longword;
var
  R,R1,R2: integer;
  G,G1,G2: integer;
  B,B1,B2: integer;
begin
  R1 := (RGB1 and $00FF0000) shr 16;
  G1 := (RGB1 and $0000FF00) shr 8;
  B1 :=  RGB1 and $000000FF;

  R2 := (RGB2 and $00FF0000) shr 16;
  G2 := (RGB2 and $0000FF00) shr 8;
  B2 :=  RGB2 and $000000FF;

  R := Round(R1 + (R2 - R1) * APercent);
  G := Round(G1 + (G2 - G1) * APercent);
  B := Round(B1 + (B2 - B1) * APercent);

  Result := (R shl 16) + (G shl 8) + B;
end;


function XLS_LocaleCharsFromUnicode(CodePage, Flags: Cardinal;  UnicodeStr: PWideChar; UnicodeStrLen: Integer; LocaleStr: PAnsiChar;
         LocaleStrLen: Integer; DefaultChar: PAnsiChar; UsedDefaultChar: PLongBool): Integer;
begin
  Result := WideCharToMultiByte(CodePage, Flags, UnicodeStr, UnicodeStrLen, LocaleStr,
    LocaleStrLen, DefaultChar, PBOOL(UsedDefaultChar));
end;


function  XLSUTF8Encode(S: AxUCString): AxUCString;
var
  L: Integer;
  Temp: PAnsiChar;

function MyUnicodeToUtf8(Dest: PAnsiChar; MaxDestBytes: Cardinal; Source: PWideChar; SourceChars: Cardinal): Cardinal;
begin
  Result := 0;
  if Source = nil then Exit;
  if Dest <> nil then begin
{$ifdef DELPHI_XE_OR_LATER}
    Result := Cardinal(LocaleCharsFromUnicode(CP_UTF8, 0, Source, Integer(SourceChars), Dest, Integer(MaxDestBytes), nil, nil));
{$else}
    Result := Cardinal(XLS_LocaleCharsFromUnicode(CP_UTF8, 0, Source, Integer(SourceChars), Dest, Integer(MaxDestBytes), nil, nil));
{$endif}
    if (Result > 0) and (Result <= MaxDestBytes) then begin
      if (SourceChars = Cardinal(-1)) and (Dest[Result -1] = #0) then Exit;

      if Result = MaxDestBytes then begin
        while (Result > 1) and (Byte(Dest[Result - 1]) > $7F) and (Byte(Dest[Result - 1]) and $80 <> 0) and (Byte(Dest[Result - 1]) and $C0 <> $C0) do
          Dec(Result);
      end
      else
        Inc(Result);
      Dest[Result - 1] := #0;
    end;
  end else
{$ifdef DELPHI_XE_OR_LATER}
    Result := Cardinal(LocaleCharsFromUnicode(CP_UTF8, 0, Source, Integer(SourceChars), nil, 0, nil, nil));
{$else}
    Result := Cardinal(XLS_LocaleCharsFromUnicode(CP_UTF8, 0, Source, Integer(SourceChars), nil, 0, nil, nil));
{$endif}
end;

begin
  Result := '';
  if S = '' then
    Exit;

  L := Length(S);
  GetMem(Temp,L * 4 + 2);
  try
    MyUnicodeToUtf8(Temp,L * 4 + 2, PWideChar(S), L);

    Result := AxUCString(Temp);
  finally
    FreeMem(Temp);
  end;
end;

// Not complete. See: http://en.wikipedia.org/wiki/UTF-8
function  _XLSUTF8Encode(S: AxUCString): AxUCString;
var
  i: integer;
  w: word;
  b1,b2: byte;
begin
  Result := '';
  for i := 1 to Length(S) do begin
    w := Word(S[i]);
    if w <= $007F then
      Result := Result + Char(S[i])
    else if w <= $07FF then begin
      b1 := (w and $07C0) shr 6;
      b2 := w and $003F;
      Result := Result + Char($C0 or b1);
      Result := Result + Char($80 or b2);
    end;
  end;
end;

function AdjustDeleteRowsCols(var ARC1,ARC2: integer; AMove1,AMove2: integer): boolean;
var
  D: integer;
begin
  Result := True;

  D := AMove2 - AMove1 + 1;

  // AMove1 and AMove2 are above.
  if AMove2 < ARC1 then begin
    ARC1 := ARC1 - D;
    ARC2 := ARC2 - D;
  end
  // AMove1 is above and AMove2 is inside.
  else if (AMove1 < ARC1) and (AMove2 >= ARC1) and (AMove2 < ARC2) then begin
    ARC1 := AMove1;
    ARC2 := ARC2 - D;
  end
  // AMove1 is above or equal and AMove2 is below or equal. Delete.
  else if (AMove1 <= ARC1) and (AMove2 >= ARC2) then begin
    Result := False;
  end
  // AMove1 is below and AMove2 is inside.
  else if (AMove1 > ARC1) and (AMove2 <= ARC2) then begin
    ARC2 := ARC2 - D;
  end
  // AMove1 is inside and AMove2 is below.
  else if (AMove1 > ARC1) and (AMove1 <= ARC2) and (AMove2 > ARC2) then begin
    ARC2 := AMove1 - 1;
  end;
end;

{ TIntegerList }

procedure TIntegerList.Add(AValue: integer);
begin
  inherited Add(Pointer(AValue));
end;

function TIntegerList.GetItems(Index: integer): integer;
begin
  Result := Integer(inherited Items[Index]);
end;

{ TXLSErrorManager }

procedure TXLSErrorManager.Clear;
begin
  FErrorList.Clear;
  FHintCount := 0;
  FWarningCount := 0;
  FErrorCount := 0;
  FFatalCount := 0;
end;

constructor TXLSErrorManager.Create;
begin
  FFilter := [xelDebug];
  FErrorList := TStringList.Create;
  FOwnsErrorList := True;
  FIgnoreXMLWarnings := True;
end;

procedure TXLSErrorManager.Debug(const AText: AxUcString);
begin
  if FSaveToList then
    FErrorList.Add(AText);
  if Assigned(FErrorEvent) and not (xelDebug in FFilter) then
    FErrorEvent(Self,xelDebug,AText);
end;

destructor TXLSErrorManager.Destroy;
begin
  if FOwnsErrorList then
    FErrorList.Free;
  inherited;
end;

procedure TXLSErrorManager.Error(const AText: AxUcString; AErrorNo: integer);
var
  S: AxUCString;
begin
  if not FIgnoreErrors then begin
    FLastError := AErrorNo;

    Inc(FErrorCount);
    S := 'E' + GetErrorText(AText,AErrorNo);
    if FSaveToList then
      FErrorList.AddObject(S,TObject(AErrorNo));
    if Assigned(FErrorEvent) and not (xelError in FFilter) then
      FErrorEvent(Self,xelError,S)
    else if not FSaveToList then
      raise XLSRWException.Create(S);
  end;
end;

procedure TXLSErrorManager.Fatal(const AText: AxUcString; AErrorNo: integer);
var
  S: AxUCString;
begin
  Inc(FFatalCount);
  S := 'F' + GetErrorText(AText,AErrorNo);
  if FSaveToList then
    FErrorList.AddObject(S,TObject(AErrorNo));
  if Assigned(FErrorEvent) then
    FErrorEvent(Self,xelFatal,S)
  else
    raise XLSRWException.Create(AText);
end;

function TXLSErrorManager.GetErrorCount: integer;
begin
  Result := FErrorCount + FFatalCount;
end;

function TXLSErrorManager.GetErrorText(const AText: AxUcString; AErrorNo: integer): AxUcString;
begin
  case AErrorNo of
    XLSERR_FILEREAD_UNKNOWENCRYPT      : Result := 'Unknown type of encryption';
    XLSERR_FILEREAD_PASSWORDMISSING    : Result := 'File is encrypted, password missing. If you think the file is unencrypted, try "VelvetSweatshop"';
    XLSERR_FILEREAD_WRONGPASSWORD      : Result := 'File is encrypted, password is wrong.';

    XLSWARN_FILEREAD_XMLPARSE           : Result := Format('XML read error: "%s"',[AText]);

    XLSERR_DIRECTWR_CELL_NOT_INC       : Result := 'Direct write: Cell column and row can not be the same as previous';

    3001                               : Result := Format('Can not find Defined Name "%s"',[AText]);
    XLSWARN_CIRCULAR_FORMULA           : Result := Format('Circular formula "%s"',[AText]);
    XLSWARN_USEDSHEETNAME              : Result := Format('There is already a sheet with name "%s"',[AText]);
    XLSWARN_INVALIDCELLREF             : Result := Format('Invalid cell reference "%s"',[AText]);
    XLSWARN_INVALIDCELLVALUE           : Result := 'Invalid cell value';
    XLSWARN_UNKNOWNIMAGE               : Result := Format('Unknown image type "%s"',[AText]);
    XLSWARN_UNSUPPORTEDIMAGE           : Result := Format('Unsupported image type "%s" Use JPEG or PNG images',[AText]);
    XLSWARN_IMAGEERROR                 : Result := 'Image error';
    XLSWARN_CAN_NOT_CHANGE_MERGED      : Result := 'Can not change part of a merged cell';

    XLSWARN_SHAPE_EDIT_TEXTBOX         : Result := 'Can not edit this shape as Text box';
    XLSWARN_SHAPE_CAN_NOT_EDIT         : Result := 'Can not edit this shape';

    XLSERR_FMLA_BADERRORCONST          : Result := Format('Unknown error constant "%s"',[AText]);
    XLSERR_FMLA_BADEXPONENT            : Result := 'Invalid exponent value';
    XLSERR_FMLA_BADFRACTIONAL          : Result := 'Invalid fractional number';
    XLSERR_FMLA_MISSINGQUOTE           : Result := 'Missing end quote character';
    XLSERR_FMLA_BADARRAYCONST          : Result := 'Invalid array constant';
    XLSERR_FMLA_FORMULA                : Result := 'Error in formula';

    XLSERR_FMLA_MISSINGRPAR            : Result := Format('Missing right parenthesis "%s"',[AText]);
    XLSERR_FMLA_INVALIDUSEOF           : Result := Format('Invalid use of "%s"',[AText]);
    XLSERR_FMLA_MISSINGEXPR            : Result := Format('Missing expression after "%s"',[AText]);
    XLSERR_FMLA_BADCELLREF             : Result := 'Invalid cell reference';
    XLSERR_FMLA_BADSHEETNAME           : Result := 'Invalid sheet name';
    XLSERR_FMLA_BADWORKBOOKNAME        : Result := 'Invalid workbook name';
    XLSERR_FMLA_MISSING                : Result := Format('Missing "%s"',[AText]);
    XLSERR_FMLA_BADR1C1REF             : Result := 'Invalid R1C1 reference';
    XLSERR_FMLA_MISSINGOPERAND         : Result := 'Missing operand';
    XLSERR_FMLA_UNKNOWNSHEET           : Result := Format('Unknown sheet name "%s"',[AText]);
    XLSERR_FMLA_UNKNOWNXBOOK           : Result := Format('Unknown external sheet "%s"',[AText]);
    XLSERR_FMLA_BADUSEOFOP             : Result := 'Invalid use of operator';
    XLSERR_FMLA_UNKNOWNNAME            : Result := Format('Unknown name "%s"',[AText]);
    XLSERR_FMLA_MISSINGARG             : Result := Format('Missing argument for function "%s"',[AText]);
    XLSERR_FMLA_TOMANYGARGS            : Result := Format('To many arguments for function "%s"',[AText]);
    XLSERR_FMLA_BADARRAYCOLS           : Result := 'Array constant columns not the same on each row';
    XLSERR_FMLA_BADUSEOFTABLEOP        : Result := 'Invalid use of operator in table';

    XLSERR_FMLA_BADTABLE               : Result := Format('Error in table "%s"',[AText]);
    XLSERR_FMLA_UNKNOWNTABLE           : Result := Format('Unknown table "%s"',[AText]);
    XLSERR_FMLA_UNKNOWNTABLECOL        : Result := Format('Unknown table columns "%s"',[AText]);
    XLSERR_FMLA_UNKNOWNTBLSPEC         : Result := Format('Unknown table special "%s"',[AText]);
    XLSERR_FMLA_TABLESQBRACKET         : Result := Format('Table "%s" is missing end bracket"',[AText]);
    XLSERR_FMLA_BADUSEOFTBLSPEC        : Result := 'Invalid combination of table specials';
    XLSERR_FMLA_TBLSPECNOTFIRST        : Result := 'Table specials must be first in expression';

    XLSERR_NAME_EMPTY                  : Result := 'Name text is empty';
    XLSERR_NAME_DUPLICATE              : Result := Format('Name "%s" is already defined',[AText]);

    XLSERR_HLINK_INVALID_CELLREF       : Result := Format('Invalid reference "%s" in hyperlink',[AText]);

    XLSERR_BIFF_BADBUILTIN_NAME        : Result := Format('Unknown built in name "%s"',[AText]);
    XLSERR_BIFF_CANT_FIND_BUILTIN_NAME : Result := Format('Can not find built in name "%s"',[AText]);

    XLSHINT_BAD_FORMULA_RESULT         : Result := Format('Formula result in cell "%s" do not match.',[AText]);


    XLSERR_FILENAME_IS_MISSING         : Result := Format('Filename is missing in %s',[AText]);
    XLSERR_NO_XLS_DEFINED              : Result := 'No ' + ComponentName + ' defined';
    XLSWARN_WORKBOOK_NOT_EMPTY         : Result := 'Workbook not empty. Data will be discarded when changing filetype.';
    XLSWARN_MULTIPLE_SEL_AREAS         : Result := AText + ' can not be performed on multiple selected areas';
    XLSERR_INVALID_SIMPLE_TAG          : Result := Format('Invalid simple tag "%s"',[AText]);
  end;
  Result := IntToStr(AErrorNo) + ': ' + Result;
end;

function TXLSErrorManager.GetFatalCount: integer;
begin
  Result := FFatalCount;
end;

function TXLSErrorManager.GetHintCount: integer;
begin
  Result := FHintCount + FWarningCount + FErrorCount + FFatalCount;
end;

function TXLSErrorManager.GetWarningCount: integer;
begin
  Result := FWarningCount + FErrorCount + FFatalCount;
end;

procedure TXLSErrorManager.Hint(const AText: AxUcString; AErrorNo: integer);
var
  S: AxUCString;
begin
  Inc(FHintCount);
  S := 'H' + GetErrorText(AText,AErrorNo);
  if FSaveToList then
    FErrorList.AddObject(S,TObject(AErrorNo));
  if Assigned(FErrorEvent) and not (xelHint in FFilter) then
    FErrorEvent(Self,xelHint,S);
end;

procedure TXLSErrorManager.Message(const AText: AxUcString);
begin
  if FSaveToList then
    FErrorList.Add(AText);
  if Assigned(FErrorEvent) and not (xelMessage in FFilter) then
    FErrorEvent(Self,xelMessage,AText);
end;

procedure TXLSErrorManager.SetErrorList(const Value: TStrings);
begin
  if FOwnsErrorList then
    FErrorList.Free;
  if Value = Nil then
    FErrorList := TStringList.Create
  else
    FErrorList := Value;
  FOwnsErrorList := Value = Nil;
  FSaveToList := Value <> Nil;
end;

procedure TXLSErrorManager.Warning(const AText: AxUcString; AErrorNo: integer);
var
  S: AxUCString;
begin
  if FIgnoreXMLWarnings and (AErrorNo = XLSWARN_FILEREAD_XMLPARSE) then
    Exit;
  Inc(FWarningCount);
  S := 'W' + GetErrorText(AText,AErrorNo);
  if FSaveToList then
    FErrorList.AddObject(S,TObject(AErrorNo));
  if Assigned(FErrorEvent) and not (xelWarning in FFilter) then
    FErrorEvent(Self,xelWarning,S);
end;

{ TIndexObjectList }

procedure TIndexObjectList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  inherited;
  case Action of
    lnAdded    : TIndexObject(Ptr).FIndex := Count - 1;
    lnExtracted,
    lnDeleted  : ReIndex;
  end;
end;

procedure TIndexObjectList.ReIndex;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    TIndexObject(Items[i]).FIndex := i
end;

{ TIntegerStack }

function TIntegerStack.Peek: integer;
begin
  Result := Items[Count - 1];
end;

function TIntegerStack.Pop: integer;
begin
  Result := Items[Count - 1];
  Delete(Count - 1);
end;

procedure TIntegerStack.Push(const AValue: integer);
begin
  Add(AValue);
end;

{ TPoinerMemoryStream }

procedure TPointerMemoryStream.SetStreamData(Ptr: Pointer; const ASize: NativeInt);
begin
  SetPointer(Ptr,ASize);
end;

{ TFormatSettings }

{$ifndef DELPHI_XE_OR_LATER}
function TFormatSettings.GetDecimalSeparator: AxUCChar;
begin
  Result := AxUCChar(SysUtils.DecimalSeparator);
end;

function TFormatSettings.GetListSeparator: AxUCChar;
begin
  Result := AxUCChar(SysUtils.ListSeparator);
end;

function TFormatSettings.GetThousandSeparator: AxUCChar;
begin
  Result := AxUCChar(SysUtils.ThousandSeparator);
end;

procedure TFormatSettings.SetDecimalSeparator(const Value: AxUCChar);
begin
{$ifdef DELPHI_D2009_OR_LATER}
  SysUtils.DecimalSeparator := Value;
{$else}
  SysUtils.DecimalSeparator := Char(Value);
{$endif}
end;

procedure TFormatSettings.SetListSeparator(const Value: AxUCChar);
begin
{$ifdef DELPHI_D2009_OR_LATER}
  SysUtils.ListSeparator := Value;
{$else}
  SysUtils.ListSeparator := Char(Value);
{$endif}
end;

procedure TFormatSettings.SetThousandSeparator(const Value: AxUCChar);
begin
{$ifdef DELPHI_D2009_OR_LATER}
  SysUtils.ThousandSeparator := Value;
{$else}
  SysUtils.ThousandSeparator := Char(Value);
{$endif}
end;
{$endif}

{ TIndexObject }

function TIndexObject.RequestObject(AClass: TClass): TObject;
begin
  Result := Nil;
end;

initialization
  BuildCRCTable;
{$ifndef DELPHI_XE_OR_LATER}

  FormatSettings := TFormatSettings.Create;

  FormatSettings.CurrencyString := CurrencyString;
  FormatSettings.CurrencyFormat := CurrencyFormat;
  FormatSettings.CurrencyDecimals := CurrencyDecimals;
  FormatSettings.DateSeparator := AxUCChar(DateSeparator);
  FormatSettings.TimeSeparator := AxUCChar(TimeSeparator);
  FormatSettings.ListSeparator := AxUCChar(ListSeparator);
  FormatSettings.ShortDateFormat := ShortDateFormat;
  FormatSettings.LongDateFormat := LongDateFormat;
  FormatSettings.TimeAMString := TimeAMString;
  FormatSettings.TimePMString := TimePMString;
  FormatSettings.ShortTimeFormat := ShortTimeFormat;
  FormatSettings.LongTimeFormat := LongTimeFormat;
  for init_i := Low(ShortMonthNames) to High(ShortMonthNames) do
    FormatSettings.ShortMonthNames[init_i] := ShortMonthNames[init_i];
  for init_i := Low(LongMonthNames) to High(LongMonthNames) do
    FormatSettings.LongMonthNames[init_i] := LongMonthNames[init_i];
  for init_i := Low(ShortDayNames) to High(ShortDayNames) do
    FormatSettings.ShortDayNames[init_i] := ShortDayNames[init_i];
  for init_i := Low(LongDayNames) to High(LongDayNames) do
    FormatSettings.LongDayNames[init_i] := LongDayNames[init_i];
  FormatSettings.ThousandSeparator := AxUCChar(ThousandSeparator);
  FormatSettings.DecimalSeparator := AxUCChar(DecimalSeparator);
  FormatSettings.TwoDigitYearCenturyWindow := TwoDigitYearCenturyWindow;
  FormatSettings.NegCurrFormat := NegCurrFormat;
{$endif}

G_Counter := 0;
G_DirSepChar := '\';

finalization
{$ifndef DELPHI_XE_OR_LATER}
  FormatSettings.Free;
{$endif}

end.
