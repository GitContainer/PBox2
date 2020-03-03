unit XLSFormulaTypes5;

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

uses Classes, SysUtils, IniFiles,
     XLSUtils5;

const  XLSFMLA_MAXDATE = 2958465;  // 9999-12-31

// _IMPORTANT_! When adding new Ptgs, update FillPtgsSizes in this unit.

const xptgNone        = $00;
const xptgOpAdd       = $03;
const xptgOpSub       = $04;
const xptgOpMult      = $05;
const xptgOpDiv       = $06;
const xptgOpPower     = $07;
const xptgOpConcat    = $08;
const xptgOpLT        = $09;
const xptgOpLE        = $0A;
const xptgOpEQ        = $0B;
const xptgOpGE        = $0C;
const xptgOpGT        = $0D;
const xptgOpNE        = $0E;

const xptgOpIsect     = $0F;
const xptgOpUnion     = $10;
const xptgOpRange     = $11;

const xptgOpUPlus     = $12;
const xptgOpUMinus    = $13;

const xptgOpPercent   = $14;
const xptgLPar        = $15;

const xptgErr         = $1C;
const xptgBool        = $1D;

const xptgNum         = $1F;

// Excel 97 ptgs

const xptgExp97       = $01;

const xptgStr97       = $17;
const xptgAttr97      = $19;
const xptgInt97       = $1E;

const xptgMissArg97   = $16;
const xptgExtend97    = $18;

const xptgArray97     = $20;
const xptgArrayV97    = $40;
const xptgArrayA97    = $60;

const xptgName97      = $23;
const xptgNameV97     = $43;
const xptgNameA97     = $63;

const xptgRef97       = $24;
const xptgRefV97      = $44;
const xptgRefA97      = $64;

const xptgRefN97      = $2C;
const xptgRefNV97     = $4C;
const xptgRefNA97     = $6C;

const xptgArea97      = $25;
const xptgAreaV97     = $45;
const xptgAreaA97     = $65;

const xptgRef3d97     = $3A;
const xptgRef3dV97    = $5A;
const xptgRef3dA97    = $7A;

const xptgRefErr3d97  = $3C;
const xptgRefErr3dV97 = $5C;
const xptgRefErr3dA97 = $7C;

const xptgArea3d97    = $3B;
const xptgArea3dV97   = $5B;
const xptgArea3dA97   = $7B;

const xptgAreaErr3d97 = $3D;
const xptgAreaErr3dV97= $5D;
const xptgAreaErr3dA97= $7D;

const xptgAreaN97     = $2D;
const xptgAreaNV97    = $4D;
const xptgAreaNA97    = $6D;

const xptgRefErr97    = $2A;
const xptgRefErrV97   = $4A;
const xptgRefErrA97   = $6A;

const xptgAreaErr97   = $2B;
const xptgAreaErrV97  = $4B;
const xptgAreaErrA97  = $6B;

const xptgFunc97      = $21;
const xptgFuncV97     = $41;
const xptgFuncA97     = $61;

const xptgFuncVar97   = $22;
const xptgFuncVarV97  = $42;
const xptgFuncVarA97  = $62;

const xptgNameX97     = $39;
const xptgNameXV97    = $59;
const xptgNameXA97    = $79;

const xptgMemFunc97   = $29;
const xptgMemFuncV97  = $49;
const xptgMemFuncA97  = $69;

const xptgMemArea97   = $26;
const xptgMemAreaV97  = $46;
const xptgMemAreaA97  = $66;

const xptgMemErr97    = $27;
const xptgMemErrV97   = $47;
const xptgMemErrA97   = $67;

// Same ptgs as excel, but not exactly the same meaning.
const xptgArray       = $80;

const xptgFunc        = $81;
const xptgFuncVar     = $82;
const xptgName        = $83;
const xptgRef         = $84;
const xptgArea        = $86;
// Same size as xptgRef/xptArea but the cells are outside the sheet due to a copy or similar operation
const xptgRefErr      = $87;
const xptgAreaErr     = $88;

const xptgInt         = $89;
const xptgStr         = $8A;


// Non-excel ptgs
const xptgWS          = $A0;
const xptgRPar        = $A1;
const xptgRef1d       = $A2; // Sheet1!A1
const xptgRef3d       = $A3; // Sheet1:Sheet3!A1
const xptgXRef1d      = $A4; // [T1.xlsx]Sheet1!A1
const xptgXRef3d      = $A5; // [T1.xlsx]Sheet1:Sheet3!A1
const xptgArea1d      = $A6; // Sheet1!A1:B2
const xptgArea3d      = $A7; // Sheet1:Sheet3!A1:B2
const xptgXArea1d     = $A8; // [T1.xlsx]Sheet1!A1:B2
const xptgXArea3d     = $A9; // [T1.xlsx]Sheet1:Sheet3!A1:B2

const xptgRef1dErr    = $B0; // Sheet1!#REF!
const xptgRef3dErr    = $B1; // Sheet1!#REF!
const xptgArea1dErr   = $B2; // Sheet1!#REF!
const xptgArea3dErr   = $B3; // Sheet1!#REF!

const xptgArrayFmlaChild     = $B8;
const xptgArrayFmlaChild97   = $B9;
const xptgDataTableFmla      = $BA;
const xptgDataTableFmlaChild = $BB;

const xptgTable              = $BC;
const xptgTableCol           = $BD;
const xptgTableSpecial       = $BE;

const xptgFuncArg     = $F0;
const xptgMissingArg  = $F1;
const xptgUserFunc    = $F2;

// If this is the first Ptgs, the rest of the ptgs is Excel 97.
const xptg_EXCEL_97 = $FF;
// There are array constants in the expression. These are stored after the
// normal ptgs. The size of the array consts is the word following this ptgs.
// There is no xptg_EXCEL_97 in the expression as this is the first ptgs.
const xptg_ARRAYCONSTS_97 = $FE;

const XPtgsLParType    = [xptgLPar,xptgArray,xptgFunc,xptgFuncVar,xptgUserFunc];
const XPtgsRefAreaType = [xptgRef,xptgArea,xptgRefErr,xptgAreaErr,xptgRef1d,xptgRef1dErr,xptgRef3d,xptgRef3dErr,xptgXRef1d,xptgXRef3d,xptgArea1d,xptgArea1dErr,xptgArea3d,xptgArea3dErr,xptgXArea1d,xptgXArea3d];
const XPtgsFuncType    = [xptgFunc,xptgFuncvar,xptgUserFunc];
const XPtgsVarFuncType = [xptgFuncVar,xptgUserFunc];

type PXLSPtgs = ^TXLSPtgs;
     TXLSPtgs = packed record
     Id: byte;
     end;

type TXLSArrayPtgs = array[0..$FFFF] of PXLSPtgs;
     PXLSArrayPtgs = ^TXLSArrayPtgs;

type PXLSPtgsErr = ^TXLSPtgsErr;
     TXLSPtgsErr = packed record
     Id: byte;
     Value: byte;
     end;

type PXLSPtgsBool = ^TXLSPtgsBool;
     TXLSPtgsBool = packed record
     Id: byte;
     Value: byte;
     end;

type PXLSPtgsWS = ^TXLSPtgsWS;
     TXLSPtgsWS = packed record
     Id: byte;
     Count: byte;
     end;

type PXLSPtgsISect = ^TXLSPtgsISect;
     TXLSPtgsISect = packed record
     Id: byte;
     SpacesCount: byte; // There can be more than one space in an intersection expression: (A1:A10   A5:D5)
     end;

type PXLSPtgsRef = ^TXLSPtgsRef;
     TXLSPtgsRef = packed record
     Id: byte;
     Row: longword;
     Col: word;
     end;

type PXLSPtgsArea = ^TXLSPtgsArea;
     TXLSPtgsArea = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     end;

type PXLSPtgsRef1d = ^TXLSPtgsRef1d;
     TXLSPtgsRef1d = packed record
     Id: byte;
     Row: longword;
     Col: word;
     Sheet: byte;
     end;

type PXLSPtgsRef1dError = ^TXLSPtgsRef1dError;
     TXLSPtgsRef1dError = packed record
     Id: byte;
     Row: longword;
     Col: word;
     Error: byte;
     end;

type PXLSPtgsRef3d = ^TXLSPtgsRef3d;
     TXLSPtgsRef3d = packed record
     Id: byte;
     Row: longword;
     Col: word;
     Sheet1: byte;
     Sheet2: byte;
     end;

type PXLSPtgsRef3dError = ^TXLSPtgsRef3dError;
     TXLSPtgsRef3dError = packed record
     Id: byte;
     Row: longword;
     Col: word;
     Error1: byte; // Sheet1
     Error2: byte; // Sheet2
     end;

type PXLSPtgsXRef1d = ^TXLSPtgsXRef1d;
     TXLSPtgsXRef1d = packed record
     Id: byte;
     Row: longword;
     Col: word;
     XBook: word;
     Sheet: byte;
     end;

type PXLSPtgsXRef3d = ^TXLSPtgsXRef3d;
     TXLSPtgsXRef3d = packed record
     Id: byte;
     Row: longword;
     Col: word;
     XBook: word;
     Sheet1: byte;
     Sheet2: byte;
     end;

type PXLSPtgsArea1d = ^TXLSPtgsArea1d;
     TXLSPtgsArea1d = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     Sheet: byte;
     end;

type PXLSPtgsArea1dError = ^TXLSPtgsArea1dError;
     TXLSPtgsArea1dError = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     Error: byte;
     end;

type PXLSPtgsArea3d = ^TXLSPtgsArea3d;
     TXLSPtgsArea3d = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     Sheet1: byte;
     Sheet2: byte;
     end;

type PXLSPtgsArea3dError = ^TXLSPtgsArea3dError;
     TXLSPtgsArea3dError = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     Error1: byte; // Sheet1
     Error2: byte; // Sheet2
     end;

type PXLSPtgsXArea1d = ^TXLSPtgsXArea1d;
     TXLSPtgsXArea1d = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     XBook: word;
     Sheet: byte;
     end;

type PXLSPtgsXArea3d = ^TXLSPtgsXArea3d;
     TXLSPtgsXArea3d = packed record
     Id: byte;
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     XBook: word;
     Sheet1: byte;
     Sheet2: byte;
     end;

type PXLSPtgsNum = ^TXLSPtgsNum;
     TXLSPtgsNum = packed record
     Id: byte;
     Value: double;
     end;

type PXLSPtgsInt = ^TXLSPtgsInt;
     TXLSPtgsInt = packed record
     Id: byte;
     Value: integer;
     end;

const FixedSzPtgsStr = 3;
type PXLSPtgsStr = ^TXLSPtgsStr;
     TXLSPtgsStr = packed record
     Id: byte;
     Len: word;
     Str: array[0..$7FFF] of AxUCChar;
     end;

// TODO If there is a formula with an unknown name, the name shall be saved
// in the Ptgs, "ptgsUnknownName", but this is not the case now. When writing
// compiled formulas, the original unknown name is lost.
const XLS_NAME_UNKNOWN = $FFFF;
type PXLSPtgsName = ^TXLSPtgsName;
     TXLSPtgsName = packed record
     Id: byte;
     NameId: word;
     end;

type PXLSPtgsArray = ^TXLSPtgsArray;
     TXLSPtgsArray = packed record
     Id: byte;
     Cols: word;
     Rows: longword;
     end;

type PXLSPtgsFunc = ^TXLSPtgsFunc;
     TXLSPtgsFunc = packed record
     Id: byte;
     FuncId: word;
     end;

type PXLSPtgsFuncVar = ^TXLSPtgsFuncVar;
     TXLSPtgsFuncVar = packed record
     Id: byte;
     ArgCount: byte;
     FuncId: word;
     end;

const FixedSzPtgsUserFunc = 4;
type PXLSPtgsUserFunc = ^TXLSPtgsUserFunc;
     TXLSPtgsUserFunc = packed record
     Id: byte;
     ArgCount: byte;
     Len: word;
     Name: array[0..$7FFF] of AxUCChar;
     end;

type PXLSPtgsArrayChildFmla = ^TXLSPtgsArrayChildFmla;
     TXLSPtgsArrayChildFmla = packed record
     Id: byte;
     ParentCol: word;
     ParentRow: longword;
     end;

type PXLSPtgsArrayChildFmla97 = ^TXLSPtgsArrayChildFmla97;
     TXLSPtgsArrayChildFmla97 = packed record
     Id: byte;
     ParentCol: word;
     ParentRow: longword;
     end;

type PXLSPtgsDataTableFmla = ^TXLSPtgsDataTableFmla;
     TXLSPtgsDataTableFmla = packed record
     Id: byte;
     Options: byte;
     R1Col: word;
     R1Row: longword;
     R2Col: word;
     R2Row: longword;
     end;

type PXLSPtgsDataTableChildFmla = ^TXLSPtgsDataTableChildFmla;
     TXLSPtgsDataTableChildFmla = packed record
     Id: byte;
     ParentCol: word;
     ParentRow: longword;
     end;

type PXLSPtgsTable = ^TXLSPtgsTable;
     TXLSPtgsTable = packed record
     Id: byte;
     TableId: word;
     // Set if the table not is on the same sheet as the formula, else $FFFF
     Sheet: word;
     ArgCount: byte;

     AreaPtg: byte;
     // The rest must match TXLSPtgsArea
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     end;

type PXLSPtgsTableCol = ^TXLSPtgsTableCol;
     TXLSPtgsTableCol = packed record
     Id: byte;
     TableId: word;
     ColId: word;

     AreaPtg: byte;
     // The rest must match TXLSPtgsArea
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     end;

type TXLSTableSpecialSpecifier = (xtssNone,xtssAll,xtssData,xtssHeaders,xtssTotals,xtssThisRow);

type PXLSPtgsTableSpecial = ^TXLSPtgsTableSpecial;
     TXLSPtgsTableSpecial = packed record
     Id: byte;
     SpecialId: byte;  // = TXLSTableSpecialSpecifier
     Error: byte;

     AreaId: byte;
     // The rest must match TXLSPtgsArea
     Row1: longword;
     Row2: longword;
     Col1: word;
     Col2: word;
     end;

type PPTGArray97 = ^TPTGArray97;
     TPTGArray97 = packed record
     Cols: byte;
     Rows: word;
     Data: record end;
     end;

type PArrayFloat97 = ^TArrayFloat97;
     TArrayFloat97 = packed record
     ID: byte;
     Value: double;
     end;

type PArrayString97 = ^TArrayString97;
     TArrayString97 = packed record
     ID: byte;
     Len: word;
     Data: record end;
     end;


type PXLSPtgsRef97 = ^TXLSPtgsRef97;
     TXLSPtgsRef97 = packed record
     Id: byte;
     Row: word;
     Col: word;
     end;

type PXLSPtgsArea97 = ^TXLSPtgsArea97;
     TXLSPtgsArea97 = packed record
     Id: byte;
     Row1: word;
     Row2: word;
     Col1: word;
     Col2: word;
     end;

type PXLSPtgsArea3d97 = ^TXLSPtgsArea3d97;
     TXLSPtgsArea3d97 = packed record
     Id: byte;
     ExtSheetIndex: word;
     Row1,Row2: word;
     Col1,Col2: word;
     end;

type PXLSPtgsRef3d97 = ^TXLSPtgsRef3d97;
     TXLSPtgsRef3d97 = packed record
     Id: byte;
     ExtSheetIndex: word;
     Row: word;
     Col: word;
     end;

type PXLSPtgsInt97 = ^TXLSPtgsInt97;
     TXLSPtgsInt97 = packed record
     Id: byte;
     Value: word;
     end;

type PXLSPtgsName97 = ^TXLSPtgsName97;
     TXLSPtgsName97 = packed record
     Id: byte;
     NameId: word;
     Unused: word;
     end;

type PXLSPtgsNameX97 = ^TXLSPtgsNameX97;
     TXLSPtgsNameX97 = packed record
     Id: byte;
     ExtSheetIndex: word;
     NameIndex: word;
     Unused: word;
     end;

type TXLSExcelFuncType = (xeftNormal,xeftArray);
type TXLSExcelFuncArgType = (xefatNone,xefatRef,xefatArea,xefatRange,xefatValue);

type PXLSExcelFuncArgs = ^TXLSExcelFuncArgs;
     TXLSExcelFuncArgs = record
     FuncId: integer;
     ArgCount: integer;
     Result: TXLSExcelFuncArgType;
     Args: array[0..9] of TXLSExcelFuncArgType;
     end;

type PXc12FuncData = ^TXc12FuncData;
     TXc12FuncData = record
     Type_: TXLSExcelFuncType;
     Args: PXLSExcelFuncArgs;
     Id: integer;
     Min,Max: integer;
     Name: AxUCString;
     end;

const XLSFUNCID_LOOKUP  = 028;
const XLSFUNCID_HLOOKUP = 101;
const XLSFUNCID_VLOOKUP = 102;

type TXLSExcelFuncNames = class(TObject)
private
     function GetItems(Index: integer): PXc12FuncData;
protected
{$ifdef DELPHI_5}
     FList: TStringList;
{$else}
     FList: THashedStringList;
{$endif}
public
     constructor Create;
     destructor Destroy; override;

     function Count: integer;

     function  FindId(const AName: AxUCString): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  FindName(AId: integer): AxUCString; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  ArgCount(AId: integer; out AMin,AMax: integer): boolean; overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  ArgCount(AId: integer): integer; overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  FixedArgs(AId: integer): boolean;

     property Items[Index: integer]: PXc12FuncData read GetItems; default;
     end;

var
  G_XLSExcelFuncNames: TXLSExcelFuncNames;
  G_XLSPtgsSize: array[0..$FF] of integer;

implementation

type TXc12FunctionsArray = array[0..327] of TXc12FuncData;

const XLSFMLA_MAXARGS = 254;

// Min and Max = $FF; illegal function name.
const Xc12Functions: TXc12FunctionsArray = (
(Type_: xeftNormal; Args: Nil; Id: 000; Min: 01; Max: 254; Name: 'COUNT'),
(Type_: xeftNormal; Args: Nil; Id: 001; Min: 02; Max: 003; Name: 'IF'),
(Type_: xeftNormal; Args: Nil; Id: 002; Min: 01; Max: 001; Name: 'ISNA'),
(Type_: xeftNormal; Args: Nil; Id: 003; Min: 01; Max: 001; Name: 'ISERROR'),
(Type_: xeftNormal; Args: Nil; Id: 004; Min: 01; Max: 254; Name: 'SUM'),
(Type_: xeftNormal; Args: Nil; Id: 005; Min: 01; Max: 254; Name: 'AVERAGE'),
(Type_: xeftNormal; Args: Nil; Id: 006; Min: 01; Max: 254; Name: 'MIN'),
(Type_: xeftNormal; Args: Nil; Id: 007; Min: 01; Max: 254; Name: 'MAX'),
(Type_: xeftNormal; Args: Nil; Id: 008; Min: 00; Max: 001; Name: 'ROW'),
(Type_: xeftNormal; Args: Nil; Id: 009; Min: 00; Max: 001; Name: 'COLUMN'),
(Type_: xeftNormal; Args: Nil; Id: 010; Min: 00; Max: 000; Name: 'NA'),
(Type_: xeftNormal; Args: Nil; Id: 011; Min: 02; Max: 254; Name: 'NPV'),
(Type_: xeftNormal; Args: Nil; Id: 012; Min: 01; Max: 254; Name: 'STDEV'),
(Type_: xeftNormal; Args: Nil; Id: 013; Min: 01; Max: 002; Name: 'DOLLAR'),
(Type_: xeftNormal; Args: Nil; Id: 014; Min: 01; Max: 003; Name: 'FIXED'),
(Type_: xeftNormal; Args: Nil; Id: 015; Min: 01; Max: 001; Name: 'SIN'),
(Type_: xeftNormal; Args: Nil; Id: 016; Min: 01; Max: 001; Name: 'COS'),
(Type_: xeftNormal; Args: Nil; Id: 017; Min: 01; Max: 001; Name: 'TAN'),
(Type_: xeftNormal; Args: Nil; Id: 018; Min: 01; Max: 001; Name: 'ATAN'),
(Type_: xeftNormal; Args: Nil; Id: 019; Min: 00; Max: 000; Name: 'PI'),
(Type_: xeftNormal; Args: Nil; Id: 020; Min: 01; Max: 001; Name: 'SQRT'),
(Type_: xeftNormal; Args: Nil; Id: 021; Min: 01; Max: 001; Name: 'EXP'),
(Type_: xeftNormal; Args: Nil; Id: 022; Min: 01; Max: 001; Name: 'LN'),
(Type_: xeftNormal; Args: Nil; Id: 023; Min: 01; Max: 001; Name: 'LOG10'),
(Type_: xeftNormal; Args: Nil; Id: 024; Min: 01; Max: 001; Name: 'ABS'),
(Type_: xeftNormal; Args: Nil; Id: 025; Min: 01; Max: 001; Name: 'INT'),
(Type_: xeftNormal; Args: Nil; Id: 026; Min: 01; Max: 001; Name: 'SIGN'),
(Type_: xeftNormal; Args: Nil; Id: 027; Min: 02; Max: 002; Name: 'ROUND'),
(Type_: xeftNormal; Args: Nil; Id: 028; Min: 02; Max: 003; Name: 'LOOKUP'),
(Type_: xeftNormal; Args: Nil; Id: 029; Min: 02; Max: 004; Name: 'INDEX'),
(Type_: xeftNormal; Args: Nil; Id: 030; Min: 02; Max: 002; Name: 'REPT'),
(Type_: xeftNormal; Args: Nil; Id: 031; Min: 03; Max: 003; Name: 'MID'),
(Type_: xeftNormal; Args: Nil; Id: 032; Min: 01; Max: 001; Name: 'LEN'),
(Type_: xeftNormal; Args: Nil; Id: 033; Min: 01; Max: 001; Name: 'VALUE'),
(Type_: xeftNormal; Args: Nil; Id: 034; Min: 00; Max: 000; Name: 'TRUE'),
(Type_: xeftNormal; Args: Nil; Id: 035; Min: 00; Max: 000; Name: 'FALSE'),
(Type_: xeftNormal; Args: Nil; Id: 036; Min: 01; Max: 254; Name: 'AND'),
(Type_: xeftNormal; Args: Nil; Id: 037; Min: 01; Max: 254; Name: 'OR'),
(Type_: xeftNormal; Args: Nil; Id: 038; Min: 01; Max: 001; Name: 'NOT'),
(Type_: xeftNormal; Args: Nil; Id: 039; Min: 02; Max: 002; Name: 'MOD'),
(Type_: xeftNormal; Args: Nil; Id: 040; Min: 03; Max: 003; Name: 'DCOUNT'),
(Type_: xeftNormal; Args: Nil; Id: 041; Min: 03; Max: 003; Name: 'DSUM'),
(Type_: xeftNormal; Args: Nil; Id: 042; Min: 03; Max: 003; Name: 'DAVERAGE'),
(Type_: xeftNormal; Args: Nil; Id: 043; Min: 03; Max: 003; Name: 'DMIN'),
(Type_: xeftNormal; Args: Nil; Id: 044; Min: 03; Max: 003; Name: 'DMAX'),
(Type_: xeftNormal; Args: Nil; Id: 045; Min: 03; Max: 003; Name: 'DSTDEV'),
(Type_: xeftNormal; Args: Nil; Id: 046; Min: 01; Max: 254; Name: 'VAR'),
(Type_: xeftNormal; Args: Nil; Id: 047; Min: 03; Max: 003; Name: 'DVAR'),
(Type_: xeftNormal; Args: Nil; Id: 048; Min: 02; Max: 002; Name: 'TEXT'),
(Type_: xeftNormal; Args: Nil; Id: 049; Min: 01; Max: 004; Name: 'LINEST'),
(Type_: xeftNormal; Args: Nil; Id: 050; Min: 01; Max: 004; Name: 'TREND'),
(Type_: xeftNormal; Args: Nil; Id: 051; Min: 01; Max: 004; Name: 'LOGEST'),
(Type_: xeftNormal; Args: Nil; Id: 052; Min: 01; Max: 004; Name: 'GROWTH'),
(Type_: xeftNormal; Args: Nil; Id: 056; Min: 02; Max: 005; Name: 'PV'),
(Type_: xeftNormal; Args: Nil; Id: 057; Min: 02; Max: 005; Name: 'FV'),
(Type_: xeftNormal; Args: Nil; Id: 058; Min: 02; Max: 005; Name: 'NPER'),
(Type_: xeftNormal; Args: Nil; Id: 059; Min: 03; Max: 005; Name: 'PMT'),
(Type_: xeftNormal; Args: Nil; Id: 060; Min: 03; Max: 005; Name: 'RATE'),
(Type_: xeftNormal; Args: Nil; Id: 061; Min: 03; Max: 003; Name: 'MIRR'),
(Type_: xeftNormal; Args: Nil; Id: 062; Min: 01; Max: 002; Name: 'IRR'),
(Type_: xeftNormal; Args: Nil; Id: 063; Min: 00; Max: 000; Name: 'RAND'),
(Type_: xeftNormal; Args: Nil; Id: 064; Min: 02; Max: 003; Name: 'MATCH'),
(Type_: xeftNormal; Args: Nil; Id: 065; Min: 03; Max: 003; Name: 'DATE'),
(Type_: xeftNormal; Args: Nil; Id: 066; Min: 03; Max: 003; Name: 'TIME'),
(Type_: xeftNormal; Args: Nil; Id: 067; Min: 01; Max: 001; Name: 'DAY'),
(Type_: xeftNormal; Args: Nil; Id: 068; Min: 01; Max: 001; Name: 'MONTH'),
(Type_: xeftNormal; Args: Nil; Id: 069; Min: 01; Max: 001; Name: 'YEAR'),
(Type_: xeftNormal; Args: Nil; Id: 070; Min: 01; Max: 002; Name: 'WEEKDAY'),
(Type_: xeftNormal; Args: Nil; Id: 071; Min: 01; Max: 001; Name: 'HOUR'),
(Type_: xeftNormal; Args: Nil; Id: 072; Min: 01; Max: 001; Name: 'MINUTE'),
(Type_: xeftNormal; Args: Nil; Id: 073; Min: 01; Max: 001; Name: 'SECOND'),
(Type_: xeftNormal; Args: Nil; Id: 074; Min: 00; Max: 000; Name: 'NOW'),
(Type_: xeftNormal; Args: Nil; Id: 075; Min: 01; Max: 001; Name: 'AREAS'),
(Type_: xeftNormal; Args: Nil; Id: 076; Min: 01; Max: 001; Name: 'ROWS'),
(Type_: xeftNormal; Args: Nil; Id: 077; Min: 01; Max: 001; Name: 'COLUMNS'),
(Type_: xeftNormal; Args: Nil; Id: 078; Min: 02; Max: 005; Name: 'OFFSET'),
(Type_: xeftNormal; Args: Nil; Id: 082; Min: 02; Max: 003; Name: 'SEARCH'),
(Type_: xeftNormal; Args: Nil; Id: 083; Min: 01; Max: 001; Name: 'TRANSPOSE'),
(Type_: xeftNormal; Args: Nil; Id: 086; Min: 01; Max: 001; Name: 'TYPE'),
(Type_: xeftNormal; Args: Nil; Id: 097; Min: 02; Max: 002; Name: 'ATAN2'),
(Type_: xeftNormal; Args: Nil; Id: 098; Min: 01; Max: 001; Name: 'ASIN'),
(Type_: xeftNormal; Args: Nil; Id: 099; Min: 01; Max: 001; Name: 'ACOS'),
(Type_: xeftNormal; Args: Nil; Id: 100; Min: 02; Max: 254; Name: 'CHOOSE'),
(Type_: xeftNormal; Args: Nil; Id: 101; Min: 03; Max: 004; Name: 'HLOOKUP'),
(Type_: xeftNormal; Args: Nil; Id: 102; Min: 03; Max: 004; Name: 'VLOOKUP'),
(Type_: xeftNormal; Args: Nil; Id: 105; Min: 01; Max: 001; Name: 'ISREF'),
(Type_: xeftNormal; Args: Nil; Id: 109; Min: 01; Max: 002; Name: 'LOG'),
(Type_: xeftNormal; Args: Nil; Id: 111; Min: 01; Max: 001; Name: 'CHAR'),
(Type_: xeftNormal; Args: Nil; Id: 112; Min: 01; Max: 001; Name: 'LOWER'),
(Type_: xeftNormal; Args: Nil; Id: 113; Min: 01; Max: 001; Name: 'UPPER'),
(Type_: xeftNormal; Args: Nil; Id: 114; Min: 01; Max: 001; Name: 'PROPER'),
(Type_: xeftNormal; Args: Nil; Id: 115; Min: 01; Max: 002; Name: 'LEFT'),
(Type_: xeftNormal; Args: Nil; Id: 116; Min: 01; Max: 002; Name: 'RIGHT'),
(Type_: xeftNormal; Args: Nil; Id: 117; Min: 02; Max: 002; Name: 'EXACT'),
(Type_: xeftNormal; Args: Nil; Id: 118; Min: 01; Max: 001; Name: 'TRIM'),
(Type_: xeftNormal; Args: Nil; Id: 119; Min: 04; Max: 004; Name: 'REPLACE'),
(Type_: xeftNormal; Args: Nil; Id: 120; Min: 03; Max: 004; Name: 'SUBSTITUTE'),
(Type_: xeftNormal; Args: Nil; Id: 121; Min: 01; Max: 001; Name: 'CODE'),
(Type_: xeftNormal; Args: Nil; Id: 124; Min: 02; Max: 003; Name: 'FIND'),
(Type_: xeftNormal; Args: Nil; Id: 125; Min: 01; Max: 002; Name: 'CELL'),
(Type_: xeftNormal; Args: Nil; Id: 126; Min: 01; Max: 001; Name: 'ISERR'),
(Type_: xeftNormal; Args: Nil; Id: 127; Min: 01; Max: 001; Name: 'ISTEXT'),
(Type_: xeftNormal; Args: Nil; Id: 128; Min: 01; Max: 001; Name: 'ISNUMBER'),
(Type_: xeftNormal; Args: Nil; Id: 129; Min: 01; Max: 001; Name: 'ISBLANK'),
(Type_: xeftNormal; Args: Nil; Id: 130; Min: 01; Max: 001; Name: 'T'),
(Type_: xeftNormal; Args: Nil; Id: 131; Min: 01; Max: 001; Name: 'N'),
(Type_: xeftNormal; Args: Nil; Id: 140; Min: 01; Max: 001; Name: 'DATEVALUE'),
(Type_: xeftNormal; Args: Nil; Id: 141; Min: 01; Max: 001; Name: 'TIMEVALUE'),
(Type_: xeftNormal; Args: Nil; Id: 142; Min: 03; Max: 003; Name: 'SLN'),
(Type_: xeftNormal; Args: Nil; Id: 143; Min: 04; Max: 004; Name: 'SYD'),
(Type_: xeftNormal; Args: Nil; Id: 144; Min: 04; Max: 005; Name: 'DDB'),
(Type_: xeftNormal; Args: Nil; Id: 148; Min: 01; Max: 002; Name: 'INDIRECT'),
(Type_: xeftNormal; Args: Nil; Id: 162; Min: 01; Max: 001; Name: 'CLEAN'),
(Type_: xeftNormal; Args: Nil; Id: 163; Min: 01; Max: 001; Name: 'MDETERM'),
(Type_: xeftArray;  Args: Nil; Id: 164; Min: 01; Max: 001; Name: 'MINVERSE'),
(Type_: xeftNormal; Args: Nil; Id: 165; Min: 02; Max: 002; Name: 'MMULT'),
(Type_: xeftNormal; Args: Nil; Id: 167; Min: 04; Max: 006; Name: 'IPMT'),
(Type_: xeftNormal; Args: Nil; Id: 168; Min: 01; Max: 254; Name: 'PPMT'),
(Type_: xeftNormal; Args: Nil; Id: 169; Min: 01; Max: 254; Name: 'COUNTA'),
(Type_: xeftNormal; Args: Nil; Id: 183; Min: 01; Max: 254; Name: 'PRODUCT'),
(Type_: xeftNormal; Args: Nil; Id: 184; Min: 01; Max: 001; Name: 'FACT'),
(Type_: xeftNormal; Args: Nil; Id: 189; Min: 03; Max: 003; Name: 'DPRODUCT'),
(Type_: xeftNormal; Args: Nil; Id: 190; Min: 01; Max: 001; Name: 'ISNONTEXT'),
(Type_: xeftNormal; Args: Nil; Id: 193; Min: 01; Max: 254; Name: 'STDEVP'),
(Type_: xeftNormal; Args: Nil; Id: 194; Min: 01; Max: 254; Name: 'VARP'),
(Type_: xeftNormal; Args: Nil; Id: 195; Min: 03; Max: 003; Name: 'DSTDEVP'),
(Type_: xeftNormal; Args: Nil; Id: 196; Min: 03; Max: 003; Name: 'DVARP'),
(Type_: xeftNormal; Args: Nil; Id: 197; Min: 01; Max: 002; Name: 'TRUNC'),
(Type_: xeftNormal; Args: Nil; Id: 198; Min: 01; Max: 001; Name: 'ISLOGICAL'),
(Type_: xeftNormal; Args: Nil; Id: 199; Min: 03; Max: 003; Name: 'DCOUNTA'),
(Type_: xeftNormal; Args: Nil; Id: 205; Min: 02; Max: 003; Name: 'FINDB'),
(Type_: xeftNormal; Args: Nil; Id: 206; Min: 02; Max: 003; Name: 'SEARCHB'),
(Type_: xeftNormal; Args: Nil; Id: 207; Min: 04; Max: 004; Name: 'REPLACEB'),
(Type_: xeftNormal; Args: Nil; Id: 208; Min: 01; Max: 002; Name: 'LEFTB'),
(Type_: xeftNormal; Args: Nil; Id: 209; Min: 01; Max: 002; Name: 'RIGHTB'),
(Type_: xeftNormal; Args: Nil; Id: 210; Min: 03; Max: 003; Name: 'MIDB'),
(Type_: xeftNormal; Args: Nil; Id: 211; Min: 01; Max: 001; Name: 'LENB'),
(Type_: xeftNormal; Args: Nil; Id: 212; Min: 02; Max: 002; Name: 'ROUNDUP'),
(Type_: xeftNormal; Args: Nil; Id: 213; Min: 02; Max: 002; Name: 'ROUNDDOWN'),
(Type_: xeftNormal; Args: Nil; Id: 216; Min: 02; Max: 003; Name: 'RANK'),
(Type_: xeftNormal; Args: Nil; Id: 219; Min: 02; Max: 005; Name: 'ADDRESS'),
(Type_: xeftNormal; Args: Nil; Id: 220; Min: 01; Max: 254; Name: 'DAYS360'),
(Type_: xeftNormal; Args: Nil; Id: 221; Min: 00; Max: 000; Name: 'TODAY'),
(Type_: xeftNormal; Args: Nil; Id: 222; Min: 05; Max: 007; Name: 'VDB'),
(Type_: xeftNormal; Args: Nil; Id: 227; Min: 01; Max: 254; Name: 'MEDIAN'),
(Type_: xeftNormal; Args: Nil; Id: 228; Min: 01; Max: 254; Name: 'SUMPRODUCT'),
(Type_: xeftNormal; Args: Nil; Id: 229; Min: 01; Max: 001; Name: 'SINH'),
(Type_: xeftNormal; Args: Nil; Id: 230; Min: 01; Max: 001; Name: 'COSH'),
(Type_: xeftNormal; Args: Nil; Id: 231; Min: 01; Max: 001; Name: 'TANH'),
(Type_: xeftNormal; Args: Nil; Id: 232; Min: 01; Max: 001; Name: 'ASINH'),
(Type_: xeftNormal; Args: Nil; Id: 233; Min: 01; Max: 001; Name: 'ACOSH'),
(Type_: xeftNormal; Args: Nil; Id: 234; Min: 01; Max: 001; Name: 'ATANH'),
(Type_: xeftNormal; Args: Nil; Id: 235; Min: 03; Max: 003; Name: 'DGET'),
(Type_: xeftNormal; Args: Nil; Id: 244; Min: 01; Max: 001; Name: 'INFO'),
(Type_: xeftNormal; Args: Nil; Id: 247; Min: 01; Max: 254; Name: 'DB'),
(Type_: xeftNormal; Args: Nil; Id: 252; Min: 02; Max: 002; Name: 'FREQUENCY'),
(Type_: xeftNormal; Args: Nil; Id: 261; Min: 01; Max: 001; Name: 'ERROR.TYPE'),
(Type_: xeftNormal; Args: Nil; Id: 269; Min: 01; Max: 254; Name: 'AVEDEV'),
(Type_: xeftNormal; Args: Nil; Id: 270; Min: 03; Max: 005; Name: 'BETADIST'),
(Type_: xeftNormal; Args: Nil; Id: 271; Min: 01; Max: 001; Name: 'GAMMALN'),
(Type_: xeftNormal; Args: Nil; Id: 272; Min: 03; Max: 005; Name: 'BETAINV'),
(Type_: xeftNormal; Args: Nil; Id: 273; Min: 04; Max: 004; Name: 'BINOMDIST'),
(Type_: xeftNormal; Args: Nil; Id: 274; Min: 02; Max: 002; Name: 'CHIDIST'),
(Type_: xeftNormal; Args: Nil; Id: 275; Min: 02; Max: 002; Name: 'CHIINV'),
(Type_: xeftNormal; Args: Nil; Id: 276; Min: 02; Max: 002; Name: 'COMBIN'),
(Type_: xeftNormal; Args: Nil; Id: 277; Min: 03; Max: 003; Name: 'CONFIDENCE'),
(Type_: xeftNormal; Args: Nil; Id: 278; Min: 03; Max: 003; Name: 'CRITBINOM'),
(Type_: xeftNormal; Args: Nil; Id: 279; Min: 01; Max: 001; Name: 'EVEN'),
(Type_: xeftNormal; Args: Nil; Id: 280; Min: 03; Max: 003; Name: 'EXPONDIST'),
(Type_: xeftNormal; Args: Nil; Id: 281; Min: 03; Max: 003; Name: 'FDIST'),
(Type_: xeftNormal; Args: Nil; Id: 282; Min: 03; Max: 003; Name: 'FINV'),
(Type_: xeftNormal; Args: Nil; Id: 283; Min: 01; Max: 001; Name: 'FISHER'),
(Type_: xeftNormal; Args: Nil; Id: 284; Min: 01; Max: 001; Name: 'FISHERINV'),
(Type_: xeftNormal; Args: Nil; Id: 285; Min: 02; Max: 002; Name: 'FLOOR'),
(Type_: xeftNormal; Args: Nil; Id: 286; Min: 04; Max: 004; Name: 'GAMMADIST'),
(Type_: xeftNormal; Args: Nil; Id: 287; Min: 03; Max: 003; Name: 'GAMMAINV'),
(Type_: xeftNormal; Args: Nil; Id: 288; Min: 02; Max: 002; Name: 'CEILING'),
(Type_: xeftNormal; Args: Nil; Id: 289; Min: 04; Max: 004; Name: 'HYPGEOMDIST'),
(Type_: xeftNormal; Args: Nil; Id: 290; Min: 03; Max: 003; Name: 'LOGNORMDIST'),
(Type_: xeftNormal; Args: Nil; Id: 291; Min: 03; Max: 003; Name: 'LOGINV'),
(Type_: xeftNormal; Args: Nil; Id: 292; Min: 03; Max: 003; Name: 'NEGBINOMDIST'),
(Type_: xeftNormal; Args: Nil; Id: 293; Min: 04; Max: 004; Name: 'NORMDIST'),
(Type_: xeftNormal; Args: Nil; Id: 294; Min: 01; Max: 001; Name: 'NORMSDIST'),
(Type_: xeftNormal; Args: Nil; Id: 295; Min: 03; Max: 003; Name: 'NORMINV'),
(Type_: xeftNormal; Args: Nil; Id: 296; Min: 01; Max: 001; Name: 'NORMSINV'),
(Type_: xeftNormal; Args: Nil; Id: 297; Min: 03; Max: 003; Name: 'STANDARDIZE'),
(Type_: xeftNormal; Args: Nil; Id: 298; Min: 01; Max: 001; Name: 'ODD'),
(Type_: xeftNormal; Args: Nil; Id: 299; Min: 02; Max: 002; Name: 'PERMUT'),
(Type_: xeftNormal; Args: Nil; Id: 300; Min: 03; Max: 003; Name: 'POISSON'),
(Type_: xeftNormal; Args: Nil; Id: 301; Min: 03; Max: 003; Name: 'TDIST'),
(Type_: xeftNormal; Args: Nil; Id: 302; Min: 04; Max: 004; Name: 'WEIBULL'),
(Type_: xeftNormal; Args: Nil; Id: 303; Min: 02; Max: 002; Name: 'SUMXMY2'),
(Type_: xeftNormal; Args: Nil; Id: 304; Min: 02; Max: 002; Name: 'SUMX2MY2'),
(Type_: xeftNormal; Args: Nil; Id: 305; Min: 02; Max: 002; Name: 'SUMX2PY2'),
(Type_: xeftNormal; Args: Nil; Id: 306; Min: 02; Max: 002; Name: 'CHITEST'),
(Type_: xeftNormal; Args: Nil; Id: 307; Min: 02; Max: 002; Name: 'CORREL'),
(Type_: xeftNormal; Args: Nil; Id: 308; Min: 02; Max: 002; Name: 'COVAR'),
(Type_: xeftNormal; Args: Nil; Id: 309; Min: 03; Max: 003; Name: 'FORECAST'),
(Type_: xeftNormal; Args: Nil; Id: 310; Min: 02; Max: 002; Name: 'FTEST'),
(Type_: xeftNormal; Args: Nil; Id: 311; Min: 02; Max: 002; Name: 'INTERCEPT'),
(Type_: xeftNormal; Args: Nil; Id: 312; Min: 02; Max: 002; Name: 'PEARSON'),
(Type_: xeftNormal; Args: Nil; Id: 313; Min: 02; Max: 002; Name: 'RSQ'),
(Type_: xeftNormal; Args: Nil; Id: 314; Min: 02; Max: 002; Name: 'STEYX'),
(Type_: xeftNormal; Args: Nil; Id: 315; Min: 02; Max: 002; Name: 'SLOPE'),
(Type_: xeftNormal; Args: Nil; Id: 316; Min: 04; Max: 004; Name: 'TTEST'),
(Type_: xeftNormal; Args: Nil; Id: 317; Min: 03; Max: 004; Name: 'PROB'),
(Type_: xeftNormal; Args: Nil; Id: 318; Min: 01; Max: 254; Name: 'DEVSQ'),
(Type_: xeftNormal; Args: Nil; Id: 319; Min: 01; Max: 254; Name: 'GEOMEAN'),
(Type_: xeftNormal; Args: Nil; Id: 320; Min: 01; Max: 254; Name: 'HARMEAN'),
(Type_: xeftNormal; Args: Nil; Id: 321; Min: 01; Max: 254; Name: 'SUMSQ'),
(Type_: xeftNormal; Args: Nil; Id: 322; Min: 01; Max: 254; Name: 'KURT'),
(Type_: xeftNormal; Args: Nil; Id: 323; Min: 01; Max: 254; Name: 'SKEW'),
(Type_: xeftNormal; Args: Nil; Id: 324; Min: 02; Max: 003; Name: 'ZTEST'),
(Type_: xeftNormal; Args: Nil; Id: 325; Min: 02; Max: 002; Name: 'LARGE'),
(Type_: xeftNormal; Args: Nil; Id: 326; Min: 02; Max: 002; Name: 'SMALL'),
(Type_: xeftNormal; Args: Nil; Id: 327; Min: 02; Max: 002; Name: 'QUARTILE'),
(Type_: xeftNormal; Args: Nil; Id: 328; Min: 02; Max: 002; Name: 'PERCENTILE'),
(Type_: xeftNormal; Args: Nil; Id: 329; Min: 02; Max: 003; Name: 'PERCENTRANK'),
(Type_: xeftNormal; Args: Nil; Id: 330; Min: 01; Max: 254; Name: 'MODE'),
(Type_: xeftNormal; Args: Nil; Id: 331; Min: 02; Max: 002; Name: 'TRIMMEAN'),
(Type_: xeftNormal; Args: Nil; Id: 332; Min: 02; Max: 002; Name: 'TINV'),
(Type_: xeftNormal; Args: Nil; Id: 336; Min: 01; Max: 254; Name: 'CONCATENATE'),
(Type_: xeftNormal; Args: Nil; Id: 337; Min: 02; Max: 002; Name: 'POWER'),
(Type_: xeftNormal; Args: Nil; Id: 342; Min: 01; Max: 001; Name: 'RADIANS'),
(Type_: xeftNormal; Args: Nil; Id: 343; Min: 01; Max: 001; Name: 'DEGREES'),
(Type_: xeftNormal; Args: Nil; Id: 344; Min: 02; Max: 254; Name: 'SUBTOTAL'),
(Type_: xeftNormal; Args: Nil; Id: 345; Min: 02; Max: 003; Name: 'SUMIF'),
(Type_: xeftNormal; Args: Nil; Id: 346; Min: 02; Max: 002; Name: 'COUNTIF'),
(Type_: xeftNormal; Args: Nil; Id: 347; Min: 01; Max: 001; Name: 'COUNTBLANK'),
(Type_: xeftNormal; Args: Nil; Id: 350; Min: 04; Max: 004; Name: 'ISPMT'),
(Type_: xeftNormal; Args: Nil; Id: 354; Min: 01; Max: 002; Name: 'ROMAN'),
(Type_: xeftNormal; Args: Nil; Id: 359; Min: 01; Max: 002; Name: 'HYPERLINK'),
(Type_: xeftNormal; Args: Nil; Id: 360; Min: 01; Max: 001; Name: 'PHONETIC'),
(Type_: xeftNormal; Args: Nil; Id: 361; Min: 01; Max: 254; Name: 'AVERAGEA'),
(Type_: xeftNormal; Args: Nil; Id: 362; Min: 01; Max: 254; Name: 'MAXA'),
(Type_: xeftNormal; Args: Nil; Id: 363; Min: 01; Max: 254; Name: 'MINA'),
(Type_: xeftNormal; Args: Nil; Id: 364; Min: 01; Max: 254; Name: 'STDEVPA'),
(Type_: xeftNormal; Args: Nil; Id: 365; Min: 01; Max: 254; Name: 'VARPA'),
(Type_: xeftNormal; Args: Nil; Id: 366; Min: 01; Max: 254; Name: 'STDEVA'),
(Type_: xeftNormal; Args: Nil; Id: 367; Min: 01; Max: 254; Name: 'VARA'),
(Type_: xeftNormal; Args: Nil; Id: 368; Min: 01; Max: 001; Name: 'BAHTTEXT'),
(Type_: xeftNormal; Args: Nil; Id: 369; Min: 01; Max: 001; Name: 'THAIDAYOFWEEK'),
(Type_: xeftNormal; Args: Nil; Id: 370; Min: 01; Max: 001; Name: 'THAIDIGIT'),
(Type_: xeftNormal; Args: Nil; Id: 371; Min: 01; Max: 001; Name: 'THAIMONTHOFYEAR'),
(Type_: xeftNormal; Args: Nil; Id: 372; Min: 01; Max: 001; Name: 'THAINUMSOUND'),
(Type_: xeftNormal; Args: Nil; Id: 373; Min: 01; Max: 001; Name: 'THAINUMSTRING'),
(Type_: xeftNormal; Args: Nil; Id: 374; Min: 01; Max: 001; Name: 'THAISTRINGLENGTH'),
(Type_: xeftNormal; Args: Nil; Id: 375; Min: 01; Max: 001; Name: 'ISTHAIDIGIT'),
(Type_: xeftNormal; Args: Nil; Id: 376; Min: 01; Max: 001; Name: 'ROUNDBAHTDOWN'),
(Type_: xeftNormal; Args: Nil; Id: 377; Min: 01; Max: 001; Name: 'ROUNDBAHTUP'),
(Type_: xeftNormal; Args: Nil; Id: 378; Min: 01; Max: 001; Name: 'THAIYEAR'),
(Type_: xeftNormal; Args: Nil; Id: 379; Min: 01; Max: 254; Name: 'RTD'),
(Type_: xeftNormal; Args: Nil; Id: 380; Min: 01; Max: 001; Name: 'ISHYPERLINK'),
(Type_: xeftNormal; Args: Nil; Id: 381; Min: 02; Max: 002; Name: 'QUOTIENT'),

// Excel 2010 functions.
(Type_: xeftNormal; Args: Nil; Id: 800; Min: 02; Max: 002; NAME: 'T.DIST.2T'),
(Type_: xeftNormal; Args: Nil; Id: 801; Min: 02; Max: 002; NAME: 'T.DIST.RT'),
(Type_: xeftNormal; Args: Nil; Id: 802; Min: 02; Max: 002; NAME: 'T.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 803; Min: 02; Max: 002; NAME: 'T.INV'),
(Type_: xeftNormal; Args: Nil; Id: 804; Min: 02; Max: 002; NAME: 'T.INV.2T'),
(Type_: xeftNormal; Args: Nil; Id: 805; Min: 02; Max: 002; NAME: 'CHISQ.INV'),
(Type_: xeftNormal; Args: Nil; Id: 806; Min: 02; Max: 002; NAME: 'CHI.INV.RT'),
(Type_: xeftNormal; Args: Nil; Id: 807; Min: 03; Max: 003; NAME: 'GAMMA.INV'),
(Type_: xeftNormal; Args: Nil; Id: 808; Min: 03; Max: 003; NAME: 'LOGNORM.INV'),
(Type_: xeftNormal; Args: Nil; Id: 809; Min: 03; Max: 003; NAME: 'NORM.INV'),
(Type_: xeftNormal; Args: Nil; Id: 810; Min: 01; Max: 001; NAME: 'NORM.S.INV'),
(Type_: xeftNormal; Args: Nil; Id: 811; Min: 03; Max: 003; NAME: 'F.INV'),
(Type_: xeftNormal; Args: Nil; Id: 812; Min: 03; Max: 003; NAME: 'F.INV.RT'),
(Type_: xeftNormal; Args: Nil; Id: 813; Min: 03; Max: 003; NAME: 'BINOM.INV'),
(Type_: xeftNormal; Args: Nil; Id: 814; Min: 03; Max: 003; NAME: 'LOGNORM.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 815; Min: 02; Max: 002; NAME: 'F.TEST'),
(Type_: xeftNormal; Args: Nil; Id: 816; Min: 01; Max: 001; NAME: 'ERFC'),
(Type_: xeftNormal; Args: Nil; Id: 817; Min: 01; Max: 001; NAME: 'ERFC.PRECISE'),
(Type_: xeftNormal; Args: Nil; Id: 818; Min: 03; Max: 254; NAME: 'AGGREGATE'),
(Type_: xeftNormal; Args: Nil; Id: 819; Min: 01; Max: 254; NAME: 'STDEV.S'),
(Type_: xeftNormal; Args: Nil; Id: 820; Min: 01; Max: 254; NAME: 'STDEV.P'),
(Type_: xeftNormal; Args: Nil; Id: 821; Min: 01; Max: 254; NAME: 'VAR.S'),
(Type_: xeftNormal; Args: Nil; Id: 822; Min: 01; Max: 254; NAME: 'VAR.P'),
(Type_: xeftNormal; Args: Nil; Id: 823; Min: 01; Max: 254; NAME: 'MODE.SNGL'),
(Type_: xeftNormal; Args: Nil; Id: 824; Min: 02; Max: 003; NAME: 'PERCENTILE.INC'),
(Type_: xeftNormal; Args: Nil; Id: 825; Min: 02; Max: 003; NAME: 'QUARTILE.INC'),
(Type_: xeftNormal; Args: Nil; Id: 826; Min: 02; Max: 002; NAME: 'PERCENTILE.EXC'),
(Type_: xeftNormal; Args: Nil; Id: 827; Min: 02; Max: 003; NAME: 'QUARTILE.EXC'),
(Type_: xeftNormal; Args: Nil; Id: 828; Min: 02; Max: 003; Name: 'AVERAGEIF'),
(Type_: xeftNormal; Args: Nil; Id: 829; Min: 02; Max: 254; Name: 'AVERAGEIFS'),
(Type_: xeftNormal; Args: Nil; Id: 830; Min: 01; Max: 002; Name: 'CEILING.PRECISE'),
(Type_: xeftNormal; Args: Nil; Id: 831; Min: 03; Max: 003; Name: 'CHISQ.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 832; Min: 03; Max: 003; Name: 'CONFIDENCE.T'),
(Type_: xeftNormal; Args: Nil; Id: 833; Min: 02; Max: 254; Name: 'COUNTIFS'),
(Type_: xeftNormal; Args: Nil; Id: 834; Min: 02; Max: 002; Name: 'COVARIANCE.S'),
(Type_: xeftNormal; Args: Nil; Id: 835; Min: 01; Max: 001; Name: 'ERF.PRECISE'),
(Type_: xeftNormal; Args: Nil; Id: 836; Min: 04; Max: 004; Name: 'F.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 837; Min: 01; Max: 002; Name: 'FLOOR.PRECISE'),
(Type_: xeftNormal; Args: Nil; Id: 838; Min: 01; Max: 001; Name: 'GAMMALN.PRECISE'),
(Type_: xeftNormal; Args: Nil; Id: 839; Min: 02; Max: 002; Name: 'IFERROR'),
(Type_: xeftNormal; Args: Nil; Id: 840; Min: 01; Max: 002; Name: 'ISO.CEILING'),
(Type_: xeftNormal; Args: Nil; Id: 841; Min: 01; Max: 254; Name: 'MODE.MULT'),
(Type_: xeftNormal; Args: Nil; Id: 842; Min: 02; Max: 004; Name: 'NETWORKDAYS.INTL'),
(Type_: xeftNormal; Args: Nil; Id: 843; Min: 02; Max: 003; Name: 'PERCENTRANK.EXC'),
(Type_: xeftNormal; Args: Nil; Id: 844; Min: 03; Max: 254; Name: 'SUMIFS'),
(Type_: xeftNormal; Args: Nil; Id: 845; Min: 02; Max: 004; Name: 'WORKDAY.INTL'),
(Type_: xeftNormal; Args: Nil; Id: 846; Min: 04; Max: 006; Name: 'BETA.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 847; Min: 03; Max: 005; Name: 'BETA.INV'),
(Type_: xeftNormal; Args: Nil; Id: 848; Min: 04; Max: 004; Name: 'BINOM.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 849; Min: 02; Max: 002; Name: 'CHISQ.DIST.RT'),
(Type_: xeftNormal; Args: Nil; Id: 850; Min: 02; Max: 002; Name: 'CHISQ.TEST'),
(Type_: xeftNormal; Args: Nil; Id: 851; Min: 03; Max: 003; Name: 'CONFIDENCE.NORM'),
(Type_: xeftNormal; Args: Nil; Id: 852; Min: 02; Max: 002; Name: 'COVARIANCE.P'),
(Type_: xeftNormal; Args: Nil; Id: 853; Min: 03; Max: 003; Name: 'EXPON.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 854; Min: 03; Max: 003; Name: 'F.DIST.RT'),
(Type_: xeftNormal; Args: Nil; Id: 855; Min: 04; Max: 004; Name: 'GAMMA.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 856; Min: 05; Max: 005; Name: 'HYPGEOM.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 857; Min: 04; Max: 004; Name: 'NEGBINOM.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 858; Min: 04; Max: 004; Name: 'NORM.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 859; Min: 02; Max: 002; Name: 'NORM.S.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 860; Min: 02; Max: 003; Name: 'PERCENTRANK.INC'),
(Type_: xeftNormal; Args: Nil; Id: 861; Min: 03; Max: 003; Name: 'POISSON.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 862; Min: 02; Max: 003; Name: 'RANK.EQ'),
(Type_: xeftNormal; Args: Nil; Id: 863; Min: 04; Max: 004; Name: 'T.TEST'),
(Type_: xeftNormal; Args: Nil; Id: 864; Min: 04; Max: 004; Name: 'WEIBULL.DIST'),
(Type_: xeftNormal; Args: Nil; Id: 865; Min: 02; Max: 003; Name: 'Z.TEST'),
(Type_: xeftNormal; Args: Nil; Id: 866; Min: 02; Max: 003; Name: 'WORKDAY'),
(Type_: xeftNormal; Args: Nil; Id: 867; Min: 02; Max: 003; Name: 'NETWORKDAYS'),
(Type_: xeftNormal; Args: Nil; Id: 868; Min: 04; Max: 007; Name: 'ACCRINT'),
// Not sure when these where introduced
(Type_: xeftNormal; Args: Nil; Id: 869; Min: 02; Max: 002; Name: 'EOMONTH'),
(Type_: xeftNormal; Args: Nil; Id: 870; Min: 02; Max: 003; Name: 'YEARFRAC'),
(Type_: xeftNormal; Args: Nil; Id: 871; Min: 02; Max: 003; Name: 'XIRR'),
(Type_: xeftNormal; Args: Nil; Id: 872; Min: 03; Max: 003; Name: 'XNPV'),
(Type_: xeftNormal; Args: Nil; Id: 873; Min: 02; Max: 002; Name: 'EDATE')
);

// -BINOM.INV          (CRITBINOM)
// -CHI.INV.RT         (CHIINV)
// -F.INV.RT           (FINV);
// -F.TEST             (FTEST)
// -GAMMA.INV          (GAMMAINV)
// -LOGNORM.DIST       (LOGNORMDIST)
// -LOGNORM.INV        (LOGINV)
// -MODE.SNGL          (MODE)
// -NORM.INV           (NORMINV)
// -NORM.S.INV         (NORMSINV)
// -PERCENTILE.INC     (PERCENTILE)
// -QUARTILE.INC       (QUARTILE)
// -STDEV.P            (STDEVP)
// -STDEV.S            (STDEV)
// -T.DIST.2T          (TDIST)
// -T.DIST.RT          (TDIST)
// -T.INV.2T           (TINV)
// -VAR.P              (VARP)
// -VAR.S              (VAR)

// New in Excel 2010 (26)
// BETA.DIST          (BETADIST)
// BETA.INV           (BETAINV)
// BINOM.DIST         (BINOMDIST)
// CHISQ.DIST.RT      (CHIDIST)
// CHISQ.TEST         (CHITEST)
// CONFIDENCE.NORM    (CONFIDENCE)
// COVARIANCE.P       (COVAR)
// EXPON.DIST         (EXPONDIST)
// F.DIST.RT          (FDIST)
// GAMMA.DIST         (GAMMADIST)
// HYPGEOM.DIST       (HYPGEOMDIST)
// NEGBINOM.DIST      (NEGBINOMDIST)
// NORM.DIST          (NORMDIST)
// NORM.S.DIST        (NORMSDIST)
// PERCENTRANK.INC    (PERCENTRANK)
// POISSON.DIST       (POISSON)
// RANK.EQ            (RANK)
// T.TEST             (TTEST)
// WEIBULL.DIST       (WEIBULL)
// Z.TEST             (ZTEST)

// -AGGREGATE
// -AVERAGEIF
// -AVERAGEIFS
// -CEILING.PRECISE
// -CHISQ.DIST
// -CHISQ.INV
// CONFIDENCE.T
// -COUNTIFS
// -COVARIANCE.S
// * ERF.PRECISE
// -ERFC.PRECISE
// -F.DIST
// -F.INV
// -FLOOR.PRECISE
// -GAMMALN.PRECISE
// -IFERROR
// -ISO.CEILING
// -MODE.MULT
// *NETWORKDAYS.INTL
// *PERCENTILE.EXC
// *PERCENTRANK.EXC
// *QUARTILE.EXC
// -SUMIFS
// -T.DIST
// -T.INV
// *WORKDAY.INTL

const Xc12FunctionArgs: array[0..260] of TXLSExcelFuncArgs = (
(FuncId:   0; ArgCount : 255; Result: xefatValue),
(FuncId:   1; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatRef,xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:   2; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:   3; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:   4; ArgCount : 255; Result: xefatValue),
(FuncId:   5; ArgCount : 255; Result: xefatValue),
(FuncId:   6; ArgCount : 255; Result: xefatValue),
(FuncId:   7; ArgCount : 255; Result: xefatValue),
(FuncId:   8; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:   9; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  10; ArgCount :   0; Result: xefatValue),
(FuncId:  11; ArgCount : 255; Result: xefatValue),
(FuncId:  12; ArgCount : 255; Result: xefatValue),
(FuncId:  13; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  14; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  15; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  16; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  17; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  18; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  19; ArgCount :   0; Result: xefatValue),
(FuncId:  20; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  21; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  22; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  23; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  24; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  25; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  26; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  27; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  28; ArgCount :   3; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  29; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  30; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  31; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  32; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  33; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  34; ArgCount :   0; Result: xefatValue),
(FuncId:  35; ArgCount :   0; Result: xefatValue),
(FuncId:  36; ArgCount : 255; Result: xefatValue),
(FuncId:  37; ArgCount : 255; Result: xefatValue),
(FuncId:  38; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  39; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  40; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  41; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  42; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  43; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  44; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  45; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  46; ArgCount : 255; Result: xefatValue),
(FuncId:  47; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  48; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  49; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  50; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  51; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  52; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  56; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  57; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  58; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  59; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  60; ArgCount :   6; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  61; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  62; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  63; ArgCount :   0; Result: xefatValue),
(FuncId:  64; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  65; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  66; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  67; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  68; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  69; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  70; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  71; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  72; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  73; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  74; ArgCount :   0; Result: xefatValue),
(FuncId:  75; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  76; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  77; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  78; ArgCount :   5; Result: xefatValue; Args: (xefatRef,xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  82; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  83; ArgCount :   1; Result: xefatArea; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  86; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  97; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  98; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId:  99; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 100; ArgCount : 255; Result: xefatArea),
(FuncId: 101; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 102; ArgCount :   4; Result: xefatArea; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 105; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 109; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 111; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 112; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 113; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 114; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 115; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 116; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 117; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 118; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 119; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 120; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 121; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 124; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 125; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 126; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 127; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 128; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 129; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 130; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 131; ArgCount :   1; Result: xefatValue; Args: (xefatRef,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 140; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 141; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 142; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 143; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 144; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 148; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 162; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 163; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 164; ArgCount :   1; Result: xefatArea; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 165; ArgCount :   2; Result: xefatArea; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 167; ArgCount :   6; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 168; ArgCount :   6; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 169; ArgCount : 255; Result: xefatValue),
(FuncId: 183; ArgCount : 255; Result: xefatValue),
(FuncId: 184; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 189; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 190; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 193; ArgCount : 255; Result: xefatValue),
(FuncId: 194; ArgCount : 255; Result: xefatValue),
(FuncId: 195; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 196; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 197; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 198; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 199; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 205; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 206; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 207; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 208; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 209; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 210; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 211; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 212; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 213; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 216; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatRange,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 219; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 220; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 221; ArgCount :   0; Result: xefatValue),
(FuncId: 222; ArgCount :   7; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone)),
(FuncId: 227; ArgCount : 255; Result: xefatValue),
(FuncId: 228; ArgCount : 255; Result: xefatValue),
(FuncId: 229; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 230; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 231; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 232; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 233; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 234; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 235; ArgCount :   3; Result: xefatArea; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 244; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 247; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 252; ArgCount :   2; Result: xefatArea; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 261; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 269; ArgCount : 255; Result: xefatValue),
(FuncId: 270; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 271; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 272; ArgCount :   5; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 273; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 274; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 275; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 276; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 277; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 278; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 279; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 280; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 281; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 282; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 283; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 284; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 285; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 286; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 287; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 288; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 289; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 290; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 291; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 292; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 293; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 294; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 295; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 296; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 297; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 298; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 299; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 300; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 301; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 302; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 303; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 304; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 305; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 306; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 307; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 308; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 309; ArgCount :   3; Result: xefatValue; Args: (xefatValue,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 310; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 311; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 312; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 313; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 314; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 315; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 316; ArgCount :   4; Result: xefatValue; Args: (xefatArea,xefatArea,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 317; ArgCount :   4; Result: xefatValue; Args: (xefatArea,xefatArea,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 318; ArgCount : 255; Result: xefatValue),
(FuncId: 319; ArgCount : 255; Result: xefatValue),
(FuncId: 320; ArgCount : 255; Result: xefatValue),
(FuncId: 321; ArgCount : 255; Result: xefatValue),
(FuncId: 322; ArgCount : 255; Result: xefatValue),
(FuncId: 323; ArgCount : 255; Result: xefatValue),
(FuncId: 324; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 325; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 326; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 327; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 328; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 329; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 330; ArgCount : 255; Result: xefatValue),
(FuncId: 331; ArgCount :   2; Result: xefatValue; Args: (xefatArea,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 332; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 336; ArgCount : 255; Result: xefatValue),
(FuncId: 337; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 342; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 343; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 344; ArgCount : 255; Result: xefatValue),
(FuncId: 345; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 346; ArgCount :   2; Result: xefatValue; Args: (xefatRange,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 347; ArgCount :   1; Result: xefatValue; Args: (xefatRange,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 350; ArgCount :   4; Result: xefatValue; Args: (xefatValue,xefatValue,xefatValue,xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 354; ArgCount :   2; Result: xefatValue; Args: (xefatValue,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 360; ArgCount :   1; Result: xefatValue; Args: (xefatRange,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 361; ArgCount : 255; Result: xefatValue),
(FuncId: 362; ArgCount : 255; Result: xefatValue),
(FuncId: 363; ArgCount : 255; Result: xefatValue),
(FuncId: 364; ArgCount : 255; Result: xefatValue),
(FuncId: 365; ArgCount : 255; Result: xefatValue),
(FuncId: 366; ArgCount : 255; Result: xefatValue),
(FuncId: 367; ArgCount : 255; Result: xefatValue),
(FuncId: 368; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 369; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 370; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 371; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 372; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 373; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 374; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 375; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 376; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 377; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 378; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 379; ArgCount : 255; Result: xefatArea),
(FuncId: 380; ArgCount :   1; Result: xefatValue; Args: (xefatValue,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 816; ArgCount :   1; Result: xefatValue; Args: (xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 828; ArgCount :   3; Result: xefatValue; Args: (xefatRange,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 829; ArgCount : 255; Result: xefatValue),
(FuncId: 833; ArgCount : 255; Result: xefatValue),
(FuncId: 839; ArgCount :   2; Result: xefatArea; Args: (xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 844; ArgCount : 255; Result: xefatValue),
(FuncId: 866; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 867; ArgCount :   3; Result: xefatValue; Args: (xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone,xefatNone)),
(FuncId: 868; ArgCount :   7; Result: xefatValue; Args: (xefatArea,xefatArea,xefatArea,xefatArea,xefatArea,xefatArea,xefatArea,xefatNone,xefatNone,xefatNone)));
//Not defined: T.DIST.2T
//Not defined: T.DIST.RT
//Not defined: T.DIST
//Not defined: T.INV
//Not defined: T.INV.2T
//Not defined: CHISQ.INV
//Not defined: CHI.INV.RT
//Not defined: GAMMA.INV
//Not defined: LOGNORM.INV
//Not defined: NORM.INV
//Not defined: NORM.S.INV
//Not defined: F.INV
//Not defined: F.INV.RT
//Not defined: BINOM.INV
//Not defined: LOGNORM.DIST
//Not defined: F.TEST
//Not defined: ERFC.PRECISE
//Not defined: AGGREGATE
//Not defined: STDEV.S
//Not defined: STDEV.P
//Not defined: VAR.S
//Not defined: VAR.P
//Not defined: MODE.SNGL
//Not defined: PERCENTILE.INC
//Not defined: QUARTILE.INC
//Not defined: PERCENTILE.EXC
//Not defined: QUARTILE.EXC
//Not defined: CEILING.PRECISE
//Not defined: CHISQ.DIST
//Not defined: CONFIDENCE.T
//Not defined: COVARIANCE.S
//Not defined: ERF.PRECISE
//Not defined: F.DIST
//Not defined: FLOOR.PRECISE
//Not defined: GAMMALN.PRECISE
//Not defined: ISO.CEILING
//Not defined: MODE.MULT
//Not defined: NETWORKDAYS.INTL
//Not defined: PERCENTRANK.EXC
//Not defined: WORKDAY.INTL
//Not defined: BETA.DIST
//Not defined: BETA.INV
//Not defined: BINOM.DIST
//Not defined: CHISQ.DIST.RT
//Not defined: CHISQ.TEST
//Not defined: CONFIDENCE.NORM
//Not defined: COVARIANCE.P
//Not defined: EXPON.DIST
//Not defined: F.DIST.RT
//Not defined: GAMMA.DIST
//Not defined: HYPGEOM.DIST
//Not defined: NEGBINOM.DIST
//Not defined: NORM.DIST
//Not defined: NORM.S.DIST
//Not defined: PERCENTRANK.INC
//Not defined: POISSON.DIST
//Not defined: RANK.EQ
//Not defined: T.TEST
//Not defined: WEIBULL.DIST
//Not defined: Z.TEST

{ TXLSExcelFuncNames }
function TXLSExcelFuncNames.ArgCount(AId: integer; out AMin, AMax: integer): boolean;
var
  Data: PXc12FuncData;
begin
  Data := PXc12FuncData(FList.Objects[AId]);
  if Data = Nil then
    raise XLSRWException.CreateFmt('Unknown function id "%d"',[AId]);
  AMin := Data.Min;
  AMax := Data.Max;
  Result := Data.Min = Data.Max;
end;

function TXLSExcelFuncNames.ArgCount(AId: integer): integer;
var
  Data: PXc12FuncData;
begin
  Data := PXc12FuncData(FList.Objects[AId]);
  if Data = Nil then
    raise XLSRWException.CreateFmt('Unknown function id "%d"',[AId]);
  Result := Data.Min;
end;

function TXLSExcelFuncNames.Count: integer;
begin
  Result := FList.Count;
end;

constructor TXLSExcelFuncNames.Create;
var
  i,n: integer;
begin
{$ifdef DELPHI_5}
  FList := TStringList.Create;
{$else}
  FList := THashedStringList.Create;
{$endif}

  n := Xc12Functions[High(Xc12Functions)].Id;
  for i := 0 to n do
    FList.AddObject('',Nil);

  for i := 0 to High(Xc12Functions) do begin
    FList[Xc12Functions[i].Id] := Xc12Functions[i].Name;
    FList.Objects[Xc12Functions[i].Id] := @Xc12Functions[i];
  end;

  for i := 0 to High(Xc12FunctionArgs) do
    PXc12FuncData(FList.Objects[Xc12FunctionArgs[i].FuncId]).Args := @Xc12FunctionArgs[i];
end;

destructor TXLSExcelFuncNames.Destroy;
begin
  FList.Free;
  inherited;
end;

function TXLSExcelFuncNames.FindId(const AName: AxUCString): integer;
begin
  Result := FList.IndexOf(AName);
end;

function TXLSExcelFuncNames.FindName(AId: integer): AxUCString;
begin
  Result := FList[AId];
end;

function TXLSExcelFuncNames.FixedArgs(AId: integer): boolean;
var
  Data: PXc12FuncData;
begin
  Data := PXc12FuncData(FList.Objects[AId]);
  if Data = Nil then
    raise XLSRWException.CreateFmt('Unknown function id "%d"',[AId]);
  Result := Data.Min = Data.Max;
end;

function TXLSExcelFuncNames.GetItems(Index: integer): PXc12FuncData;
begin
  Result := PXc12FuncData(FList.Objects[Index]);
end;

procedure FillPtgsSizes;
var
  i: integer;
begin
  for i := 0 to High(G_XLSPtgsSize) do
    G_XLSPtgsSize[i] := -1;

  G_XLSPtgsSize[xptg_EXCEL_97]    := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptg_ARRAYCONSTS_97]  := SizeOf(TXLSPtgs);

  G_XLSPtgsSize[xptgOpAdd]        := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpSub]        := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpMult]       := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpDiv]        := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpPower]      := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpConcat]     := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpLT]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpLE]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpEQ]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpGE]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpGT]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpNE]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpIsect]      := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpUnion]      := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpRange]      := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpUPlus]      := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpUMinus]     := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgOpPercent]    := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgLPar]         := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgLPar]         := SizeOf(TXlsPtgs);
  G_XLSPtgsSize[xptgArray]        := SizeOf(TXlsPtgsArray);
  G_XLSPtgsSize[xptgErr]          := SizeOf(TXlsPtgsErr);
  G_XLSPtgsSize[xptgBool]         := SizeOf(TXlsPtgsBool);
  G_XLSPtgsSize[xptgStr]          := -1; // Variable size
  G_XLSPtgsSize[xptgOpIsect]      := SizeOf(TXlsPtgsISect);

  G_XLSPtgsSize[xptgRef]          := SizeOf(TXlsPtgsRef);
  G_XLSPtgsSize[xptgArea]         := SizeOf(TXlsPtgsArea);
  G_XLSPtgsSize[xptgRef1d]        := SizeOf(TXlsPtgsRef1d);
  G_XLSPtgsSize[xptgArea1d]       := SizeOf(TXlsPtgsArea1d);
  G_XLSPtgsSize[xptgRef3d]        := SizeOf(TXlsPtgsRef3d);
  G_XLSPtgsSize[xptgArea3d]       := SizeOf(TXlsPtgsArea3d);
  G_XLSPtgsSize[xptgXRef1d]       := SizeOf(TXlsPtgsXRef1d);
  G_XLSPtgsSize[xptgXRef3d]       := SizeOf(TXlsPtgsXRef3d);
  G_XLSPtgsSize[xptgXArea1d]      := SizeOf(TXlsPtgsXArea1d);
  G_XLSPtgsSize[xptgXArea3d]      := SizeOf(TXlsPtgsXArea3d);

  G_XLSPtgsSize[xptgRef1dErr] := SizeOf(TXLSPtgsRef1dError);
  G_XLSPtgsSize[xptgRef3dErr] := SizeOf(TXLSPtgsRef3dError);
  G_XLSPtgsSize[xptgArea1dErr] := SizeOf(TXLSPtgsArea1dError);
  G_XLSPtgsSize[xptgArea3dErr] := SizeOf(TXLSPtgsArea3dError);

  G_XLSPtgsSize[xptgInt]          := SizeOf(TXlsPtgsInt);
  G_XLSPtgsSize[xptgNum]          := SizeOf(TXlsPtgsNum);
  G_XLSPtgsSize[xptgName]         := SizeOf(TXlsPtgsName);
  G_XLSPtgsSize[xptgFunc]         := SizeOf(TXlsPtgsFunc);
  G_XLSPtgsSize[xptgFuncVar]      := SizeOf(TXlsPtgsFuncVar);
  G_XLSPtgsSize[xptgUserFunc]     := -1; // Variable size
  G_XLSPtgsSize[xptgFuncArg]      := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgMissingArg]   := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgWS]           := SizeOf(TXlsPtgsWS);

  G_XLSPtgsSize[xptgArrayFmlaChild] := SizeOf(TXLSPtgsArrayChildFmla);
  G_XLSPtgsSize[xptgDataTableFmlaChild] := SizeOf(TXLSPtgsDataTableChildFmla);
  G_XLSPtgsSize[xptgDataTableFmla] := SizeOf(TXLSPtgsDataTableFmla);

  G_XLSPtgsSize[xptgTable] := SizeOf(TXLSPtgsTable);
  G_XLSPtgsSize[xptgTableCol] := SizeOf(TXLSPtgsTableCol);
  G_XLSPtgsSize[xptgTableSpecial] := SizeOf(TXLSPtgsTableSpecial);

  G_XLSPtgsSize[xptgMissArg97]    := SizeOf(TXLSPtgs);
  G_XLSPtgsSize[xptgExtend97]     := 2;

  G_XLSPtgsSize[xptgStr97]        := -1;
  G_XLSPtgsSize[xptgInt97]        := SizeOf(TXLSPtgsInt97);

  G_XLSPtgsSize[xptgArray97]      := 8; // See docs
  G_XLSPtgsSize[xptgArrayV97]     := 8;
  G_XLSPtgsSize[xptgArrayA97]     := 8;

  G_XLSPtgsSize[xptgName97]       := SizeOf(TXLSPtgsName97);
  G_XLSPtgsSize[xptgNameV97]      := SizeOf(TXLSPtgsName97);
  G_XLSPtgsSize[xptgNameA97]      := SizeOf(TXLSPtgsName97);

  G_XLSPtgsSize[xptgNameX97]      := SizeOf(TXLSPtgsNameX97);
  G_XLSPtgsSize[xptgNameXV97]     := SizeOf(TXLSPtgsNameX97);
  G_XLSPtgsSize[xptgNameXA97]     := SizeOf(TXLSPtgsNameX97);

  G_XLSPtgsSize[xptgAttr97]       := -1;

  G_XLSPtgsSize[xptgRef97]        := SizeOf(TXLSPtgsRef97);
  G_XLSPtgsSize[xptgRefV97]       := SizeOf(TXLSPtgsRef97);
  G_XLSPtgsSize[xptgRefA97]       := SizeOf(TXLSPtgsRef97);

  G_XLSPtgsSize[xptgRefN97]       := SizeOf(TXLSPtgsRef97);
  G_XLSPtgsSize[xptgRefNV97]      := SizeOf(TXLSPtgsRef97);
  G_XLSPtgsSize[xptgRefNA97]      := SizeOf(TXLSPtgsRef97);

  G_XLSPtgsSize[xptgRef3d97]      := SizeOf(TXLSPtgsRef3d97);
  G_XLSPtgsSize[xptgRef3dV97]     := SizeOf(TXLSPtgsRef3d97);
  G_XLSPtgsSize[xptgRef3dA97]     := SizeOf(TXLSPtgsRef3d97);

  G_XLSPtgsSize[xptgRefErr3d97]      := SizeOf(TXLSPtgsRef3d97);
  G_XLSPtgsSize[xptgRefErr3dV97]     := SizeOf(TXLSPtgsRef3d97);
  G_XLSPtgsSize[xptgRefErr3dA97]     := SizeOf(TXLSPtgsRef3d97);

  G_XLSPtgsSize[xptgArea97]       := SizeOf(TXLSPtgsArea97);
  G_XLSPtgsSize[xptgAreaV97]      := SizeOf(TXLSPtgsArea97);
  G_XLSPtgsSize[xptgAreaA97]      := SizeOf(TXLSPtgsArea97);

  G_XLSPtgsSize[xptgAreaN97]      := SizeOf(TXLSPtgsArea97);
  G_XLSPtgsSize[xptgAreaNV97]     := SizeOf(TXLSPtgsArea97);
  G_XLSPtgsSize[xptgAreaNA97]     := SizeOf(TXLSPtgsArea97);

  G_XLSPtgsSize[xptgArea3d97]     := SizeOf(TXLSPtgsArea3d97);
  G_XLSPtgsSize[xptgArea3dV97]    := SizeOf(TXLSPtgsArea3d97);
  G_XLSPtgsSize[xptgArea3dA97]    := SizeOf(TXLSPtgsArea3d97);

  G_XLSPtgsSize[xptgAreaErr3d97]  := SizeOf(TXLSPtgsArea3d97);
  G_XLSPtgsSize[xptgAreaErr3dV97] := SizeOf(TXLSPtgsArea3d97);
  G_XLSPtgsSize[xptgAreaErr3dA97] := SizeOf(TXLSPtgsArea3d97);

  G_XLSPtgsSize[xptgRefErr97]     := SizeOf(TXLSPtgsRef97);
  G_XLSPtgsSize[xptgRefErrV97]    := SizeOf(TXLSPtgsRef97);
  G_XLSPtgsSize[xptgRefErrA97]    := SizeOf(TXLSPtgsRef97);

  G_XLSPtgsSize[xptgAreaErr97]    := SizeOf(TXLSPtgsArea97);
  G_XLSPtgsSize[xptgAreaErrV97]   := SizeOf(TXLSPtgsArea97);
  G_XLSPtgsSize[xptgAreaErrA97]   := SizeOf(TXLSPtgsArea97);

  G_XLSPtgsSize[xptgFunc97]       := SizeOf(TXLSPtgsFunc);
  G_XLSPtgsSize[xptgFuncV97]      := SizeOf(TXLSPtgsFunc);
  G_XLSPtgsSize[xptgFuncA97]      := SizeOf(TXLSPtgsFunc);
  G_XLSPtgsSize[xptgFuncVar97]    := SizeOf(TXLSPtgsFuncVar);
  G_XLSPtgsSize[xptgFuncVarV97]   := SizeOf(TXLSPtgsFuncVar);
  G_XLSPtgsSize[xptgFuncVarA97]   := SizeOf(TXLSPtgsFuncVar);
end;

initialization
  G_XLSExcelFuncNames := TXLSExcelFuncNames.Create;
  FillPtgsSizes;

finalization
  G_XLSExcelFuncNames.Free;

end.
