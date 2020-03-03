unit XLSCellMMU5;

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

uses Classes, SysUtils, Math,
     Xc12Utils5, Xc12DataSST5, Xc12DataStyleSheet5,
     XLSUtils5, Xc12Common5, XLSMMU5, XLSFormulaTypes5;

{.$define MMU_DEBUG}

// Formula memory storage
// Formula type and options are stored in TXLSCellItem.Flags
//   word     : Size of the entire formula data.
//
//   byte     : Flags. Lower 5 bits are formula type, TXLSCellFormulaType
//              Bit 7: Formula is array entered formula.
//              Bit 8: Formula is TABLE() formula.
//
//  [Style]   : Word. Style index if formula has a style assigned,
//              CELLXL_FLAG_STYLE in XLate byte.
//
//              Result, read XLate byte for result type.
// ** Numeric rersult **
//   double   : formula value
// ** string result **
//   word     : string chars count
//   string   : unicode chars
// ** boolean result **
//   byte     : boolean value
// ** error result **
//   byte     : error value

//   word     : Formula size. Ptgs size or formula length in chars.
//   Formula  : Chars of formula expression, or if first byte is $FF, formula ptgs.
//              If formula is stored as chars, and first byte is $01, then all
//              chars are 8-bit stored in ISO/IEC 8859-1, Latin-1 codepage.
//              If first byte is $00, then all chars are 16-bit unicode.
//
//   Ref      : TXLSFormulaArea. Result area for array entered formulas,
//              only if bit 7 in Flags is set.
//
//   TableOpts: byte. Options for TABLE() formulas.
//              Only when bit 8 in TXLSCellItem.Flags is set.
//   r1,r2    : Two TXLSFormulaRef structures. Used by TABLE() function. See doc for more details.
//              Only when bit 8 in TXLSCellItem.Flags is set.

const CELLXL_FLAG_STYLE   = $40;

const CELLXL_BLANK        =  1;

const CELLXL_NUM_8        =  2;
const CELLXL_NUM_16       =  3;
const CELLXL_NUM_32       =  4;
const CELLXL_NUM_2DEC     =  5;
const CELLXL_NUM_4DEC     =  6;
const CELLXL_NUM_SINGLE   =  7;
const CELLXL_NUM_DOUBLE   =  8;

const CELLXL_STR_16       =  9;
const CELLXL_STR_32       = 10;
const CELLXL_BOOLEAN      = 11;
const CELLXL_ERROR        = 12;

const CELLXL_FMLA_FIRST   = 20;
const CELLXL_FMLA_FLOAT   = 20;
const CELLXL_FMLA_STR     = 21;
const CELLXL_FMLA_BOOLEAN = 22;
const CELLXL_FMLA_ERROR   = 23;
const CELLXL_FMLA_LAST    = 23;

const CELLMMU_STYLE_SZ    = 2;
const CELLMMU_FMLA_FLAG_SZ= 1;

const CELLMMU_FMLA_FLAG_ARRAY = $40;
const CELLMMU_FMLA_FLAG_TABLE = $80;
const CELLMMU_FMLA_FLAG_MASK  = $1F;

const CELLMMU_FMLA_IS_PTGS    = $FF;
const CELLMMU_FMLA_IS_STR_8   = $01;
const CELLMMU_FMLA_IS_STR_UC  = $00;

type TXLSRowOption = (xroHidden,xroCustomHeight,xroCollapsed,xroThickTop,xroThickBottom,xroPhonetic);
     TXLSRowOptions = set of TXLSRowOption;

type TXLSCellFormulaType = (xcftNormal,xcftArray,xcftDataTable,xcftArrayChild,xcftArrayChild97,xcftDataTableChild);

type PXLSMMURowHeader = ^TXLSMMURowHeader;
     TXLSMMURowHeader = packed record
     Options     : TXLSRowOptions;
     Style       : word;
     OutlineLevel: byte;
     Height      : word;
     end;

type PXLSMMUFormulaHeader = ^TXLSMMUFormulaHeader;
     TXLSMMUFormulaHeader = packed record
     AllocSize: word;
     FmlaType: TXLSCellFormulaType;
     end;

type PXLSCellItem = ^TXLSCellItem;
     TXLSCellItem = record
     XLate: integer;
     Data: TMMUPtr;
     end;

type PXLSFormulaArea = ^TXLSFormulaArea;
     TXLSFormulaArea = packed record
     Col1: word;
     Row1: longword;
     Col2: word;
     Row2: longword;
     end;

type PXLSFormulaRef = ^TXLSFormulaRef;
     TXLSFormulaRef = packed record
     Col: word;
     Row: longword;
     end;

type TXLSCellMMUIterate = record
     Col: integer;
     Row: integer;
     LastRow: integer;

     RowHeader: PXLSMMURowHeader;
     Cell: TXLSCellItem;
     end;

type TXLSFormulaHelper = class(TObject)
private
     function  GetRef: AxUCString;
     procedure SetRef(const Value: AxUCString); {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetAsFloat: double;
     procedure SetAsFloat(const Value: double); {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetAsBoolean: boolean;
     procedure SetAsBoolean(const Value: boolean); {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetAsString: AxUCString;
     procedure SetAsString(const Value: AxUCString); {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetAsErrorStr: AxUCString;
     procedure SetAsErrorStr(const Value: AxUCString); {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetStrTargetRef: AxUCString;
     procedure SetStrTragetRef(const Value: AxUCString);
     function  GetStrR1: AxUCString;
     function  GetStrR2: AxUCString;
     procedure SetStrR1(const Value: AxUCString);
     procedure SetStrR2(const Value: AxUCString);
     function  GetAsError: TXc12CellError;
     procedure SetAsError(const Value: TXc12CellError);
     procedure SetPtgs(const Value: PXLSPtgs);
{$ifdef DELPHI_XE_OR_LATER} strict {$endif} private
     FCol         : integer;
     FRow         : integer;
     // For array childs
     FParentCol   : integer;
     FParentRow   : integer;

     FStyle       : integer;
     FFormula     : AxUCString;
     FFmlaIs8Bit  : boolean;
     FPtgs        : PXLSPtgs;
     FPtgsSize    : integer;
     FOwnsPtgs    : boolean;
     FCellType    : TXLSCellType;
     FFormulaType : TXLSCellFormulaType;
     FTargetRef   : TXLSFormulaArea;
     FHasApply    : boolean;
     FResult      : Int64;
     FStrResult   : AxUCString;
     FR1          : TXLSFormulaRef;
     FR2          : TXLSFormulaRef;
     FIsTABLE     : boolean;
     FOptions     : byte;
     FTableOptions: byte;
public
     destructor Destroy; override;

     procedure Clear;

     function  IsCompiled: boolean;
     function  AllocSize: integer;
     procedure SetR1(const ACol,ARow: integer);
     procedure SetR2(const ACol,ARow: integer);
     procedure AllocPtgs(const APtgs: PXLSPtgs; const APtgsSize: integer);
     procedure AllocPtgs97(const APtgs: PXLSPtgs; const APtgsSize: integer);
     procedure AllocPtgsArrayConsts97(const APtgs: PXLSPtgs; const APtgsSize,AArrayConstsSize: integer);
     procedure SetTargetRef(const ACol1,ARow1,ACol2,ARow2: integer);

     property Col         : integer read FCol write FCol;
     property Row         : integer read FRow write FRow;
     property ParentCol   : integer read FParentCol write FParentCol;
     property ParentRow   : integer read FParentRow write FParentRow;
     property Ref         : AxUCString read GetRef write SetRef;
     property Style       : integer read FStyle write FStyle;
     property Formula     : AxUCString read FFormula write FFormula;
     property FmlaIs8Bit  : boolean read FFmlaIs8Bit;
     property Ptgs        : PXLSPtgs read FPtgs write SetPtgs;
     property PtgsSize    : integer read FPtgsSize write FPtgsSize;

     property AsFloat     : double read GetAsFloat write SetAsFloat;
     property AsString    : AxUCString read GetAsString write SetAsString;
     property AsBoolean   : boolean read GetAsBoolean write SetAsBoolean;
     property AsError     : TXc12CellError read GetAsError write SetAsError;
     property AsErrorStr  : AxUCString read GetAsErrorStr write SetAsErrorStr;

     property FormulaType : TXLSCellFormulaType read FFormulaType write FFormulaType;
     property CellType    : TXLSCellType read FCellType write FCellType;
     property StrTargetRef: AxUCString read GetStrTargetRef write SetStrTragetRef;
     property TargetRef   : TXLSFormulaArea read FTargetRef write FTargetRef;
     property HasApply    : boolean read FHasApply write FHasApply;
     property IsTABLE     : boolean read FIsTABLE write FIsTABLE;
     property Options     : byte read FOptions write FOptions;
     property TableOptions: byte read FTableOptions write FTableOptions;
     property StrR1       : AxUCString read GetStrR1 write SetStrR1;
     property StrR2       : AxUCString read GetStrR2 write SetStrR2;
     property R1          : TXLSFormulaRef read FR1 write FR1;
     property R2          : TXLSFormulaRef read FR2 write FR2;
     end;

type TXLSCellMMU = class;

     TXLSCellMMUDebugData = class(TObject)
protected
     FOwner: TXLSCellMMU;
     FList: TStringList;
public
     constructor Create(AOwner: TXLSCellMMU);
     destructor Destroy; override;

     procedure Clear;

     procedure Save(AFilename: string);
     procedure Load(AFilename: string);

     procedure Delete(const ACol,ARow: integer);
     procedure AddBlank(const ACol,ARow,AStyle: integer);
     procedure AddBoolean(const ACol,ARow,AStyle: integer; const AValue: boolean);
     procedure AddError(const ACol,ARow,AStyle: integer; const AValue: TXc12CellError);
     procedure AddFloat(const ACol,ARow,AStyle: integer; const AValue: double);
     procedure AddString(const ACol,ARow,AStyle: integer; const AValue: integer);
     procedure AddFormula(const ACol,ARow,AStyle: integer; const AFormula: AxUCString);
     procedure AddFormulaFloat(const ACol,ARow,AStyle: integer; const AValue: double; const APtgs: PXLSPtgs; APtgsSize: integer);

     procedure AddRow(const ARow,AStyle: integer);
     end;

     TXLSCellMMU = class(TObject)
private
     procedure SetDebugSave(const Value: boolean);
protected
     FSST          : TXc12DataSST;
     FStyles       : TXc12DataStyleSheet;
     FDefaultRowHdr: TXLSMMURowHeader;

     FBlocks       : array of TXLSMMUBlock;
     FBlockManager : TXLSMMUBlockManager;
     FIsUpdating   : boolean;
     FiCurrBlock   : integer;

     FDimension    : TXLSCellArea;

     FFormulaHelper: TXLSFormulaHelper;

     FIterate      : TXLSCellMMUIterate;

     FSheetIndex   : integer;

     FDebugLoad    : boolean;
     FDebugSave    : boolean;
     FDebugList    : TStrings;
     FDebugMem     : integer;
     FDebugCnt     : integer;
     FDebugData    : TXLSCellMMUDebugData;

     function  CheckIfRowHeaderAssigned(AHeader: TMMUPtr): boolean;

     procedure InitBlock(const AIndex: integer);
     procedure FreeBlock(const AIndex: integer);
     function  EncodeNumber(const AValue: double; out AResult: double): integer;
     procedure AddBlocks(const ALastIndex: integer);

     function  SelectBlock(const ARow: integer): boolean;  {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  ExtendBlocks(ARow: integer): boolean;  {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}

     function  ItemSize(const AXLate: integer; const APtr: TMMUPtr): integer;
     function  XLateIsFormula(const AXLate: integer): boolean; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  XLateToCellType(const AXLate: integer): TXLSCellType;
     procedure TrimEndBlocks;

     function  GetValue(const ACell: PXLSCellItem): TMMUPtr; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}

     function  GetFormulaFmla(const AFormula: TMMUPtr; const AXLate: integer; out ASize: integer): TMMUPtr;
     function  GetFormulaPtgs(const AFormula: TMMUPtr; const AXLate: integer; out ASize: integer): TMMUPtr; overload;
     function  GetFormulaStr(const AFormula: TMMUPtr; const AXLate: integer): AxUCString;
     function  GetFormulaValueSize(const AFormulaValue: TMMUPtr; const AXLate: integer): integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  AdjustFormulaMem(const ACell: PXLSCellItem; const ACol,ARow,ANewSize: integer; out AMem: TMMUPtr; out ASize: integer): TMMUPtr;

     // If ARow is outside the last block, the last block is returned.
     function  FindBlock(const ARow: integer): integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetLastRow: integer;

     function  CheckCol(const ACol: integer): boolean; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  CheckRow(const ARow: integer): boolean; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}

     procedure DEBUG_AddFormulaFloat(const ACol,ARow,AStyle: integer; const AValue: double; const APtgs: PXLSPtgs; APtgsSize: integer);
public
     constructor Create(const ASheetIndex: integer; const ADebugList: TStrings; const ASST: TXc12DataSST; const AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure Clear;

     function  IsEmpty: boolean;
     function  CalcDimensions: boolean;
     procedure CalcRowDimensions(ARow: integer; out ACol1,ACol2: integer);

     function  GetStyle(ACell: PXLSCellItem): integer; overload; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  GetStyle(const ACol,ARow: integer): integer; overload; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     procedure SetStyle(const ACol,ARow: integer; AStyle: integer); overload;
     procedure SetStyle(const ACol,ARow: integer; ACell: PXLSCellItem; AStyle: integer); overload;

     procedure SetRowStyle(const ARow,AStyle: integer);
     function  GetRowStyle(const ARow: integer): integer;

     // Allocates memory for a TXLSCellItem
     function  CopyCell(const ACol,ARow: integer): TXLSCellItem;
     procedure InsertCell(const ACol,ARow: integer; const ACell: PXLSCellItem);
     procedure MoveCell(const ASrcCol,ASrcRow,ADestCol,ADestRow: integer);
     procedure ExchangeCells(const ACol1,ARow1,ACol2,ARow2: integer);
     // Releases the memory for ACell
     procedure FreeCell(const ACell: PXLSCellItem);

     // Format is not deleted.
     procedure ClearCell(const ACol,ARow: integer);
     procedure DeleteCell(const ACol,ARow: integer);
//{$ifdef _AXOLOT_DEBUG}
//{$message warn '_TODO_ Check that XF is deleted when cells are deleted'}
//{$endif}
     procedure DeleteColumns(const ACol1,ACol2: integer);
     procedure DeleteAndShiftColumns(const ACol1,ACol2: integer);
     procedure DeleteRows(const ARow1,ARow2: integer);
     procedure DeleteAndShiftRows(const ARow1,ARow2: integer);

     procedure InsertColumns(const ACol,ACount: integer);
     procedure InsertRows(const ARow,ACount: integer);
     procedure InsertRows_ORIG(const ARow,ACount: integer);

     function  FindRow(const ARow: integer): PXLSMMURowHeader;
     function  AddRow(const ARow,AStyle: integer): PXLSMMURowHeader; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}

     procedure ClearRowHeaders(const ARow1,ARow2: integer);
     procedure CopyRowHeaders(const ARow1,ARow2,ADestRow: integer); overload;
     procedure CopyRowHeaders(ASrcMMU: TXLSCellMMU; const ARow1,ARow2,ADestRow: integer); overload;

     function  CellType(const ACell: PXLSCellItem): TXLSCellType; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}

     // Add values to memory. Any existing values are overwritten.
     // Styles are reference counted here.
     procedure StoreBlank(const ACol,ARow,AStyle: integer; const AAcceptDefStyle: boolean = False);
     procedure StoreBoolean(const ACol,ARow,AStyle: integer; const AValue: boolean);
     procedure StoreFloat(const ACol,ARow,AStyle: integer; const AValue: double);
     procedure StoreError(const ACol,ARow,AStyle: integer; const AValue: TXc12CellError);
     procedure StoreString(const ACol,ARow,AStyle: integer; const AValue: integer); overload;
     procedure StoreString(const ACol,ARow,AStyle: integer; const AValue: AxUCString); overload;

     function  EmptyCell(const ACol, ARow: integer): boolean;

     function  FindCell(const ACol, ARow: integer): TXLSCellItem; overload;
     function  FindCell(const ACol, ARow: integer; out ACell: TXLSCellItem): boolean; overload;

     function  GetBlank(const ACol,ARow: integer): boolean; overload;
     function  GetBlank(const ACell: PXLSCellItem): boolean; overload;

     // True if the cell is blank or there is no cell.
     function  GetEmpty(const ACol,ARow: integer): boolean; overload;

     function  GetBoolean(const ACol,ARow: integer; out AValue: boolean): boolean; overload;
     function  GetBoolean(const ACell: PXLSCellItem): boolean; overload;

     function  GetFloat(const ACol,ARow: integer): double; overload;
     function  GetFloat(const ACell: PXLSCellItem): double; overload;

     function  GetError(const ACol,ARow: integer; out AValue: TXc12CellError): boolean; overload;
     function  GetError(const ACell: PXLSCellItem): TXc12CellError; overload;

     function  GetString(const ACol,ARow: integer; out AValue: AxUCString): boolean; overload;
     function  GetString(const ACell: PXLSCellItem): AxUCString; overload;
     function  GetString(const ACell: PXLSCellItem; out AFontRuns: PXc12FontRunArray; out AFontRunsCount: integer): AxUCString; overload;
     function  GetStringSST(const ACell: PXLSCellItem): integer;

     function  AsString(const ACol, ARow: integer; var AResult: AxUCString): boolean;
     function  AsFloat(ACol, ARow: integer; var Aresult: double; out AEmptyCell: boolean): TXc12CellError;

     // Preserves excisting format if the cell is formatted. Otherwhise Format is used.
     // The intention is to set Format to the Row or Column format, and use that
     // if the cell not is formatted.
     procedure UpdateBoolean(ACol,ARow: integer; AValue: boolean; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
     procedure UpdateFloat(const ACol,ARow: integer; const AValue: double; const AStyle: integer = XLS_STYLE_DEFAULT_XF); overload;
     procedure UpdateError(const ACol,ARow: integer; const AValue: TXc12CellError; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
     procedure UpdateString(const ACol,ARow: integer; const AValue: AxUCString; const AStyle: integer = XLS_STYLE_DEFAULT_XF);

     // Overwrite Formula. If the cell contains a formula, the formula is overwritten and replaced with the value.
     // The above methods, UpdateXXX updates the result of a formula if there is one in the cell.
     procedure UpdateBlankOF(ACol,ARow: integer; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
     procedure UpdateBooleanOF(ACol,ARow: integer; AValue: boolean; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
     procedure UpdateFloatOF(const ACol,ARow: integer; const AValue: double; const AStyle: integer = XLS_STYLE_DEFAULT_XF); overload;
     procedure UpdateErrorOF(const ACol,ARow: integer; const AValue: TXc12CellError; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
     procedure UpdateStringOF(const ACol,ARow: integer; const AValue: AxUCString; const AStyle: integer = XLS_STYLE_DEFAULT_XF);

     procedure ClearFormulaKeepValue(ACol,ARow: integer);

     procedure AddFormulaVal(const ACell: PXLSCellItem; const ACol,ARow: integer; const AValue: boolean); overload;
     procedure AddFormulaVal(const ACol,ARow: integer; const AValue: boolean); overload;

     procedure AddFormulaVal(const ACell: PXLSCellItem; const ACol,ARow: integer; const AValue: double); overload;
     procedure AddFormulaVal(const ACol,ARow: integer; const AValue: double); overload;

     procedure AddFormulaVal(const ACell: PXLSCellItem; const ACol,ARow: integer; const AValue: TXc12CellError); overload;
     procedure AddFormulaVal(const ACol,ARow: integer; const AValue: TXc12CellError); overload;

     procedure AddFormulaVal(const ACell: PXLSCellItem; const ACol,ARow: integer; const AValue: AxUCString); overload;
     procedure AddFormulaVal(const ACol,ARow: integer; const AValue: AxUCString); overload;

     function  IsFormula(const ACell: PXLSCellItem): boolean; overload;
     function  IsFormula(const ACol,ARow: integer): boolean; overload;
     function  FormulaType(const ACell: PXLSCellItem): TXLSCellFormulaType; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IsFormulaCompiled(const ACell: PXLSCellItem): boolean; overload;
     function  IsFormulaCompiled(const ACol,ARow: integer): boolean; overload;
     function  GetFormulaPtgs(const ACell: PXLSCellItem; out APtgs: PXLSPtgs): integer; overload;
     procedure SetFormulaPtgs(const ACol,ARow: integer; const ACell: PXLSCellItem; const APtgs: PXLSPtgs; const APtgsSize: integer);
     function  GetFormulaTargetArea(const ACell: PXLSCellItem): PXLSFormulaArea;

     procedure AddFormula(const ACol,ARow: integer; const AFormula: AxUCString; const AValue: double); overload;

     procedure AddFormula(const ACol,ARow,AStyle: integer; const AFormula: AxUCString; const AValue: boolean); overload;
     procedure AddFormula(const ACol,ARow,AStyle: integer; const AFormula: AxUCString; const AValue: double); overload;
     procedure AddFormula(const ACol,ARow,AStyle: integer; const AFormula: AxUCString; const AValue: AxUCString); overload;
     procedure AddFormula(const ACol,ARow,AStyle: integer; const AFormula: AxUCString; const AValue: TXc12CellError); overload;

     procedure AddFormula(const ACol,ARow,AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: boolean); overload;
     procedure AddFormula(const ACol,ARow,AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: double); overload;
     procedure AddFormula(const ACol,ARow,AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: AxUCString); overload;
     procedure AddFormula(const ACol,ARow,AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: TXc12CellError); overload;

     procedure AddFormula97(const ACol,ARow: integer; APtgs: Pointer; const APtgsSz: integer; const AStyle: integer);

     function  GetFormula(const ACol,ARow: integer): AxUCString; overload;
     function  GetFormula(const ACell: PXLSCellItem): AxUCString; overload;

     function  GetFormulaValBoolean(const ACol,ARow: integer; out AValue: boolean): boolean; overload;
     function  GetFormulaValBoolean(const ACell: PXLSCellItem): boolean; overload;

     function  GetFormulaValFloat(const ACol,ARow: integer; out AValue: double): boolean; overload;
     function  GetFormulaValFloat(const ACell: PXLSCellItem): double; overload;

     function  GetFormulaValError(const ACol,ARow: integer; out AValue: TXc12CellError): boolean; overload;
     function  GetFormulaValError(const ACell: PXLSCellItem): TXc12CellError; overload;

     function  GetFormulaValString(const ACol,ARow: integer; out AValue: AxUCString): boolean; overload;
     function  GetFormulaValString(const ACell: PXLSCellItem): AxUCString; overload;

     procedure BeginIterate(const ARow: integer = 0);
     function  IterateNext: boolean;

     // ARow = start row - 1.
     procedure BeginIterateRow(const ARow: integer = -1);
     function  IterateNextRow: boolean;
     procedure BeginIterateCol;
     function  IterateNextCol: boolean;

     function  IterCellType: TXLSCellType; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterFormulaType: TXLSCellFormulaType; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterCellCol: integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterCellRow: integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
//     function  IterRef: TXLSRefItem; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterRow: PXLSMMURowHeader; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterCell: PXLSCellItem; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterRefStr: AxUCString;
     function  IterGetStyleIndex: integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterGetBoolean: boolean; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterGetFloat: double; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterGetError: TXc12CellError; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     function  IterGetStringIndex: integer; {$ifdef DELPHI_XE_OR_LATER} inline; {$endif}
     procedure IterGetFormula;
     function  IterGetFormulaPtgs(var APtgs: PXLSPtgs): integer;

     procedure StoreFormula;
     procedure RetriveFormula(const ACell: PXLSCellItem);

     // *********** DEBUG ***********
     procedure DebugTest(const ATest: AxUCString; AList: TStrings);

     function  CheckIntegrity(const AList: TStrings): boolean;

     procedure SaveDebug(const AFilename: string);
     procedure LoadDebug(const AFilename: string);

     procedure BeginDebug;
     procedure EndDebug;
     procedure DebugAddEmpty(const ACount: integer);
     procedure DebugAddFloat(const AValue: double); overload;
     procedure DebugAddFloat(const AValues: array of double); overload;
     procedure DebugAddFormulaFloat(const AValue: double; const ASize: integer); overload;
     procedure Test;

     property DebugData: TXLSCellMMUDebugData read FDebugData;

     property DebugSave: boolean read FDebugSave write SetDebugSave;

     property Dimension    : TXLSCellArea read FDimension write FDimension;
     property FormulaHelper: TXLSFormulaHelper read FFormulaHelper;
     end;

implementation

{ TXLSCellMMU }

procedure TXLSCellMMU.SaveDebug(const AFilename: string);
begin
  FDebugData.Save(AFilename);
end;

function TXLSCellMMU.SelectBlock(const ARow: integer): boolean;
var
  iBlock: integer;
begin
  iBlock := ARow shr XBLOCK_VECTOR_COUNT_BITS;
  Result := iBlock <= High(FBlocks);
  if Result then
    FBlockManager.Block := @FBlocks[ARow shr XBLOCK_VECTOR_COUNT_BITS];
end;

procedure TXLSCellMMU.SetDebugSave(const Value: boolean);
begin
  FDebugData.Clear;
  FDebugSave := Value;
end;

procedure TXLSCellMMU.SetFormulaPtgs(const ACol,ARow: integer; const ACell: PXLSCellItem; const APtgs: PXLSPtgs; const APtgsSize: integer);
//var
//  C: TXLSCellItem;
begin
  if (ACell.XLate = 0) or (ACell.Data = Nil) then
    raise XLSRWException.Create('Invalid formula');
  RetriveFormula(ACell);
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Ptgs := APtgs;
  FFormulaHelper.PtgsSize := APtgsSize;
  StoreFormula;

//  C := FindCell(ACol,ARow);
end;

procedure TXLSCellMMU.SetRowStyle(const ARow, AStyle: integer);
begin
  AddRow(ARow,AStyle);
end;

procedure TXLSCellMMU.SetStyle(const ACol, ARow: integer; AStyle: integer);
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    SetStyle(ACol,ARow,@Cell,AStyle)
  else if AStyle <> XLS_STYLE_DEFAULT_XF then
    StoreBlank(ACol,ARow,AStyle);
end;

procedure TXLSCellMMU.SetStyle(const ACol,ARow: integer; ACell: PXLSCellItem; AStyle: integer);
var
  P: TMMUPtr;
  OldStyle: integer;
  vBool: boolean;
  vError: TXc12CellError;
  vFloat: double;
  vString: integer;
begin
  OldStyle := GetStyle(ACell);
  if OldStyle = AStyle then
    Exit;
  if (OldStyle <> XLS_STYLE_DEFAULT_XF) and (AStyle <> XLS_STYLE_DEFAULT_XF) then begin
    if (ACell.XLate and XVECT_XLATE_MASK) in [CELLXL_FMLA_FIRST..CELLXL_FMLA_LAST] then begin
      P := ACell.Data;
      Inc(P,SZ_HDR_EX + CELLMMU_FMLA_FLAG_SZ);
      PWord(P)^ := AStyle;
    end
    else
      PWord(ACell.Data)^ := AStyle;
    FStyles.XFEditor.UseStyle(AStyle);
  end
  else begin
    case CellType(ACell) of
      xctBlank         : StoreBlank(ACol,ARow,AStyle);
      xctBoolean       : begin
        vBool := GetBoolean(ACell);
        StoreBoolean(ACol,ARow,AStyle,vBool);
      end;
      xctError         : begin
        vError := GetError(ACell);
        StoreError(ACol,ARow,AStyle,vError);
      end;
      xctString        : begin
        vString := GetStringSST(ACell);
        StoreString(ACol,ARow,AStyle,vString);
      end;
      xctFloat         : begin
        vFloat := GetFloat(ACell);
        StoreFloat(ACol,ARow,AStyle,vFloat);
      end;
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : begin
        RetriveFormula(ACell);
        FFormulaHelper.Col := ACol;
        FFormulaHelper.Row := ARow;
        FFormulaHelper.Style := AStyle;
        StoreFormula;
      end;
    end;
  end;
  FStyles.XFEditor.FreeStyle(OldStyle);
end;

procedure TXLSCellMMU.StoreFormula;
var
  i: integer;
  P: TMMUPtr;
  Sz: integer;
  XL: byte;
begin
  case FFormulaHelper.CellType of
    xctFloatFormula  : XL := CELLXL_FMLA_FLOAT;
    xctStringFormula : XL := CELLXL_FMLA_STR;
    xctBooleanFormula: XL := CELLXL_FMLA_BOOLEAN;
    xctErrorFormula  : XL := CELLXL_FMLA_ERROR;
    else               raise XLSRWException.Create('Invalid formula cell type');
  end;

  if FFormulaHelper.Style <> XLS_STYLE_DEFAULT_XF then
    XL := XL or CELLXL_FLAG_STYLE;

  Sz := FFormulaHelper.AllocSize;

  if ExtendBlocks(FFormulaHelper.Row) then
    FStyles.XFEditor.UseDefault;

  P := FBlockManager.AllocMemEx(FFormulaHelper.Row and XBLOCK_VECTOR_MASK,FFormulaHelper.Col,XL,Sz);
  Inc(P,SizeOf(word));

  PByte(P)^ := Byte(FFormulaHelper.FormulaType);
  if FFormulaHelper.HasApply then
    PByte(P)^ := PByte(P)^ or CELLMMU_FMLA_FLAG_ARRAY;
  if FFormulaHelper.IsTABLE then
    PByte(P)^ := PByte(P)^ or CELLMMU_FMLA_FLAG_TABLE;
  Inc(P,SizeOf(byte));

  if FFormulaHelper.Style <> XLS_STYLE_DEFAULT_XF then begin
    PWord(P)^ := FFormulaHelper.Style;
    Inc(P,SizeOf(word));
  end;
  FStyles.XFEditor.UseStyle(FFormulaHelper.Style);

  case FFormulaHelper.CellType of
    xctFloatFormula  : begin
      PDouble(P)^ := FFormulaHelper.AsFloat;
      Inc(P,SizeOf(double));
    end;
    xctStringFormula : begin
      PWord(P)^ := Length(FFormulaHelper.AsString);
      Inc(P,SizeOf(word));
      System.Move(FFormulaHelper.AsString[1],P^,Length(FFormulaHelper.AsString) * 2);
      Inc(P,Length(FFormulaHelper.AsString) * 2);
    end;
    xctBooleanFormula: begin
      if FFormulaHelper.AsBoolean then
        PByte(P)^ := 1
      else
        PByte(P)^ := 0;
      Inc(P,SizeOf(byte));
    end;
    xctErrorFormula  : begin
      PByte(P)^ := Byte(FFormulaHelper.AsError);
      Inc(P,SizeOf(byte));
    end;
  end;

  case FFormulaHelper.FormulaType of
    xcftNormal,
    xcftArray: begin
      if FFormulaHelper.PtgsSize > 0 then begin
        PWord(P)^ := FFormulaHelper.PtgsSize;
        Inc(P,SizeOf(word));
        PByte(P)^ := CELLMMU_FMLA_IS_PTGS;
        Inc(P,SizeOf(byte));
        System.Move(FFormulaHelper.Ptgs^,P^,FFormulaHelper.PtgsSize);
        Inc(P,FFormulaHelper.PtgsSize);
      end
      else begin
        PWord(P)^ := Length(FFormulaHelper.Formula);
        Inc(P,SizeOf(word));
        if UnicodeIs8Bit(FFormulaHelper.Formula) then begin
          PByte(P)^ := CELLMMU_FMLA_IS_STR_8;
          Inc(P,SizeOf(byte));
          for i := 1 to Length(FFormulaHelper.Formula) do
            PByteArray(P)[i - 1] := Byte(FFormulaHelper.Formula[i]);
          Inc(P,Length(FFormulaHelper.Formula));
        end
        else begin
          PByte(P)^ := CELLMMU_FMLA_IS_STR_UC;
          Inc(P,SizeOf(byte));
          System.Move(FFormulaHelper.Formula[1],Pointer(P)^,Length(FFormulaHelper.Formula) * 2);
          Inc(P,Length(FFormulaHelper.Formula) * 2);
        end;
      end;
    end;
    xcftDataTable: begin
      PWord(P)^ := SizeOf(TXLSPtgsDataTableFmla);
      Inc(P,SizeOf(word));
      PByte(P)^ := CELLMMU_FMLA_IS_PTGS;
      Inc(P,SizeOf(byte));
      PXLSPtgsDataTableFmla(P).Id := xptgDataTableFmla;
      PXLSPtgsDataTableFmla(P).Options := FFormulaHelper.TableOptions;
      PXLSPtgsDataTableFmla(P).R1Col := FFormulaHelper.R1.Col;
      PXLSPtgsDataTableFmla(P).R1Row := FFormulaHelper.R1.Row;
      PXLSPtgsDataTableFmla(P).R2Col := FFormulaHelper.R2.Col;
      PXLSPtgsDataTableFmla(P).R2Row := FFormulaHelper.R2.Row;
      Inc(P,SizeOf(TXLSPtgsDataTableFmla));
    end;
    xcftArrayChild: begin
      PWord(P)^ := SizeOf(TXLSPtgsArrayChildFmla);
      Inc(P,SizeOf(word));
      PByte(P)^ := CELLMMU_FMLA_IS_PTGS;
      Inc(P,SizeOf(byte));
      PXLSPtgsArrayChildFmla(P).Id := xptgArrayFmlaChild;
      PXLSPtgsArrayChildFmla(P).ParentCol := FFormulaHelper.ParentCol;
      PXLSPtgsArrayChildFmla(P).ParentRow := FFormulaHelper.ParentRow;
      Inc(P,SizeOf(TXLSPtgsArrayChildFmla));
    end;
    xcftArrayChild97: begin
      PWord(P)^ := SizeOf(TXLSPtgsArrayChildFmla97);
      Inc(P,SizeOf(word));
      PByte(P)^ := CELLMMU_FMLA_IS_PTGS;
      Inc(P,SizeOf(byte));
      PXLSPtgsArrayChildFmla97(P).Id := xptgArrayFmlaChild97;
      PXLSPtgsArrayChildFmla97(P).ParentCol := FFormulaHelper.ParentCol;
      PXLSPtgsArrayChildFmla97(P).ParentRow := FFormulaHelper.ParentRow;
      Inc(P,SizeOf(TXLSPtgsArrayChildFmla));
    end;
    xcftDataTableChild: begin
      PWord(P)^ := SizeOf(TXLSPtgsDataTableChildFmla);
      Inc(P,SizeOf(word));
      PByte(P)^ := CELLMMU_FMLA_IS_PTGS;
      Inc(P,SizeOf(byte));
      PXLSPtgsArrayChildFmla(P).Id := xptgDataTableFmlaChild;
      PXLSPtgsArrayChildFmla(P).ParentCol := FFormulaHelper.ParentCol;
      PXLSPtgsArrayChildFmla(P).ParentRow := FFormulaHelper.ParentRow;
      Inc(P,SizeOf(TXLSPtgsDataTableChildFmla));
    end;
  end;

  if FFormulaHelper.HasApply then begin
    PXLSFormulaArea(P)^ := FFormulaHelper.TargetRef;
    Inc(P,SizeOf(TXLSFormulaArea));
  end;

  if FFormulaHelper.IsTABLE then begin
    PByte(P)^ := FFormulaHelper.TableOptions;
    Inc(P,SizeOf(Byte));
    PXLSFormulaRef(P)^ := FFormulaHelper.R1;
    Inc(P,SizeOf(TXLSFormulaRef));
    PXLSFormulaRef(P)^ := FFormulaHelper.R2;
  end;
end;

procedure TXLSCellMMU.Test;
begin
  SelectBlock(0);

end;

procedure TXLSCellMMU.TrimEndBlocks;
var
  i: integer;
  n: integer;
begin
  n := 0;
  for i := High(FBlocks) downto 0 do begin
    if FBlocks[i].Memory = Nil then
      Inc(n)
    else
      Break;
  end;
  SetLength(FBlocks,Length(FBlocks) - n);
end;

procedure TXLSCellMMU.UpdateBlankOF(ACol, ARow: integer; const AStyle: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    Style := GetStyle(@Cell);
    if Style <> XLS_STYLE_DEFAULT_XF then
      FStyles.XFEditor.DecStyle(Style)
    else
      Style := AStyle;

    StoreBlank(ACol,ARow,Style);
  end
  else
    StoreBlank(ACol,ARow,AStyle);
end;

procedure TXLSCellMMU.UpdateBoolean(ACol, ARow: integer; AValue: boolean; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    case XLateToCellType(Cell.XLate) of
      xctBoolean       : PByte(GetValue(@Cell))^ := Byte(AValue);
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : AddFormulaVal(@Cell,ACol,ARow,AValue);
      else begin
        Style := GetStyle(@Cell);
        if Style <> XLS_STYLE_DEFAULT_XF then
          FStyles.XFEditor.DecStyle(Style)
        else
          Style := AStyle;

        StoreBoolean(ACol,ARow,Style,AValue);
      end;
    end;
  end
  else
    StoreBoolean(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateBooleanOF(ACol, ARow: integer; AValue: boolean; const AStyle: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    Style := GetStyle(@Cell);
    if Style <> XLS_STYLE_DEFAULT_XF then
      FStyles.XFEditor.DecStyle(Style)
    else
      Style := AStyle;

    StoreBoolean(ACol,ARow,Style,AValue);
  end
  else
    StoreBoolean(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateError(const ACol, ARow: integer; const AValue: TXc12CellError; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    case XLateToCellType(Cell.XLate) of
      xctError         : PByte(GetValue(@Cell))^ := Byte(AValue);
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : AddFormulaVal(@Cell,ACol,ARow,AValue);
      else begin
        Style := GetStyle(@Cell);
        if Style <> XLS_STYLE_DEFAULT_XF then
          FStyles.XFEditor.DecStyle(Style)
        else
          Style := AStyle;

        StoreError(ACol,ARow,Style,AValue);
      end;
    end;
  end
  else
    StoreError(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateErrorOF(const ACol, ARow: integer; const AValue: TXc12CellError; const AStyle: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    Style := GetStyle(@Cell);
    if Style <> XLS_STYLE_DEFAULT_XF then
      FStyles.XFEditor.DecStyle(Style)
    else
      Style := AStyle;

    StoreError(ACol,ARow,Style,AValue);
  end
  else
    StoreError(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateFloat(const ACol, ARow: integer; const AValue: double; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    case XLateToCellType(Cell.XLate) of
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : AddFormulaVal(@Cell,ACol,ARow,AValue);
      else begin
        Style := GetStyle(@Cell);
        if Style <> XLS_STYLE_DEFAULT_XF then
          FStyles.XFEditor.DecStyle(Style)
        else
          Style := AStyle;

        StoreFloat(ACol,ARow,Style,AValue);
      end;
    end;
  end
  else
    StoreFloat(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateFloatOF(const ACol, ARow: integer; const AValue: double; const AStyle: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    Style := GetStyle(@Cell);
    if Style <> XLS_STYLE_DEFAULT_XF then
      FStyles.XFEditor.DecStyle(Style)
    else
      Style := AStyle;

    StoreFloat(ACol,ARow,Style,AValue);
  end
  else
    StoreFloat(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateString(const ACol, ARow: integer; const AValue: AxUCString; const AStyle: integer = XLS_STYLE_DEFAULT_XF);
var
  i: integer;
  XL: byte;
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    case XLateToCellType(Cell.XLate) of
      xctString        : begin
        if (Cell.XLate and XVECT_XLATE_MASK) = CELLXL_STR_16 then
          FSST.ReleaseString(PWord(GetValue(@Cell))^)
        else
          FSST.ReleaseString(PLongword(GetValue(@Cell))^);

        i := FSST.AddString(AValue);
        if i <= $FFFF then
          XL := CELLXL_STR_16
        else
          XL := CELLXL_STR_32;

        if (Cell.XLate and XVECT_XLATE_MASK) = XL then begin
          FSST.UsesString(i);
          if XL = CELLXL_STR_16 then
            PWord(GetValue(@Cell))^ := i
          else
            PLongword(GetValue(@Cell))^ := i;
        end
        else begin
          Style := GetStyle(@Cell);
          FStyles.XFEditor.DecStyle(Style);
          StoreString(ACol,ARow,Style,i);
        end;
      end;
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : AddFormulaVal(@Cell,ACol,ARow,AValue);
      else begin
        Style := GetStyle(@Cell);
        if Style <> XLS_STYLE_DEFAULT_XF then
          FStyles.XFEditor.DecStyle(Style)
        else
          Style := AStyle;

        StoreString(ACol,ARow,Style,AValue);
      end;
    end;
  end
  else
    StoreString(ACol,ARow,AStyle,AValue);
end;

procedure TXLSCellMMU.UpdateStringOF(const ACol, ARow: integer; const AValue: AxUCString; const AStyle: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then begin
    Style := GetStyle(@Cell);
    if Style <> XLS_STYLE_DEFAULT_XF then
      FStyles.XFEditor.DecStyle(Style)
    else
      Style := AStyle;

    StoreString(ACol,ARow,Style,AValue);
  end
  else
    StoreString(ACol,ARow,AStyle,AValue);
end;

function TXLSCellMMU.XLateIsFormula(const AXLate: integer): boolean;
var
  XL: integer;
begin
  XL := AXLate and XVECT_XLATE_MASK;
  Result := (XL >= CELLXL_FMLA_FIRST) and (XL <= CELLXL_FMLA_LAST);
end;

function TXLSCellMMU.XLateToCellType(const AXLate: integer): TXLSCellType;
begin
  case AXLate and XVECT_XLATE_MASK of
    CELLXL_BLANK       : Result := xctBlank;
    CELLXL_NUM_8       : Result := xctFloat;
    CELLXL_NUM_16      : Result := xctFloat;
    CELLXL_NUM_32      : Result := xctFloat;
    CELLXL_NUM_2DEC    : Result := xctFloat;
    CELLXL_NUM_4DEC    : Result := xctFloat;
    CELLXL_NUM_SINGLE  : Result := xctFloat;
    CELLXL_NUM_DOUBLE  : Result := xctFloat;
    CELLXL_STR_16      : Result := xctString;
    CELLXL_STR_32      : Result := xctString;
    CELLXL_BOOLEAN     : Result := xctBoolean;
    CELLXL_ERROR       : Result := xctError;

    CELLXL_FMLA_FLOAT  : Result := xctFloatFormula;
    CELLXL_FMLA_STR    : Result := xctStringFormula;
    CELLXL_FMLA_BOOLEAN: Result := xctBooleanFormula;
    CELLXL_FMLA_ERROR  : Result := xctErrorFormula;
    else                 raise XLSRWException.Create('Invalid XLate code');
  end;
end;

//procedure TXLSCellMMU.StoreBlank(const ACol, ARow, AStyle: integer; const AAcceptDefStyle: boolean = False);
//var
//  P: TMMUPtr;
//  XLate: byte;
//  Style: integer;
//  Cell: TXLSCellItem;
//begin
//{$ifdef MMU_DEBUG}
//  if FDebugSave then
//    FDebugData.AddBlank(ACol,ARow,AStyle)
//  else begin
//{$endif}
//   if FindCell(ACol,ARow,Cell) then
//     Style := GetStyle(@Cell)
//   else
//     Style := AStyle;
//
//   XLate := CELLXL_BLANK or CELLXL_FLAG_STYLE;
//
//   if ExtendBlocks(ARow) then
//     FStyles.XFEditor.UseDefault;
//
//   P := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);
//
//   PWord(P)^ := Style;
//
//   FStyles.XFEditor.UseStyle(AStyle);
//
//{$ifdef MMU_DEBUG}
//  end;
//{$endif}
//end;

procedure TXLSCellMMU.StoreBlank(const ACol, ARow, AStyle: integer; const AAcceptDefStyle: boolean = False);
var
  P: TMMUPtr;
  XLate: byte;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.AddBlank(ACol,ARow,AStyle)
  else begin
{$endif}
//    if AAcceptDefStyle or (AStyle <> XLS_STYLE_DEFAULT_XF) then begin
      XLate := CELLXL_BLANK or CELLXL_FLAG_STYLE;

      if ExtendBlocks(ARow) then
        FStyles.XFEditor.UseDefault;

      P := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);
      PWord(P)^ := AStyle;

      FStyles.XFEditor.UseStyle(AStyle);
//    end;
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.AddBlocks(const ALastIndex: integer);
var
  i,j: integer;
begin
  j := High(FBlocks) + 1;
  SetLength(FBlocks,ALastIndex + 1);
  for i := j to High(FBlocks) do
    InitBlock(i);
end;

procedure TXLSCellMMU.StoreBoolean(const ACol, ARow, AStyle: integer; const AValue: boolean);
var
  P: TMMUPtr;
  XLate: byte;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.AddBoolean(ACol,ARow,AStyle,AValue)
  else begin
{$endif}
    XLate := CELLXL_BOOLEAN;

    if AStyle <> XLS_STYLE_DEFAULT_XF then
      XLate := XLate or CELLXL_FLAG_STYLE;

    if ExtendBlocks(ARow) then
      FStyles.XFEditor.UseDefault;

    P := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);

    if AStyle <> XLS_STYLE_DEFAULT_XF then begin
      PWord(P)^ := AStyle;
      Inc(P,SizeOf(word));
    end;
    if AValue then
      PByte(P)^ := 1
    else
      PByte(P)^ := 0;

    FStyles.XFEditor.UseStyle(AStyle);
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.StoreError(const ACol, ARow, AStyle: integer; const AValue: TXc12CellError);
var
  P: TMMUPtr;
  XLate: byte;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.AddError(ACol,ARow,AStyle,AValue)
  else begin
{$endif}
    XLate := CELLXL_ERROR;

    if AStyle <> XLS_STYLE_DEFAULT_XF then
      XLate := XLate or CELLXL_FLAG_STYLE;

    if ExtendBlocks(ARow) then
      FStyles.XFEditor.UseDefault;

    P := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);

    if AStyle <> XLS_STYLE_DEFAULT_XF then begin
      PWord(P)^ := AStyle;
      Inc(P,SizeOf(word));
    end;

    PByte(P)^ := Integer(AValue);

    FStyles.XFEditor.UseStyle(AStyle);
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.StoreFloat(const ACol, ARow, AStyle: integer; const AValue: double);
var
  P: TMMUPtr;
  Res: double;
  XLate: byte;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.AddFloat(ACol,ARow,AStyle,AValue)
  else begin
{$endif}
    XLate := EncodeNumber(AValue,Res);

    if AStyle <> XLS_STYLE_DEFAULT_XF then
      XLate := XLate or CELLXL_FLAG_STYLE;

    if ExtendBlocks(ARow) then
      FStyles.XFEditor.UseDefault;

    P := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);

    if AStyle <> XLS_STYLE_DEFAULT_XF then begin
      PWord(P)^ := AStyle;
      Inc(P,SizeOf(word));
    end;

    case XLate and XVECT_XLATE_MASK of
      CELLXL_NUM_8     : PShortint(P)^ := PShortint(@Res)^;
      CELLXL_NUM_16    : PSmallint(P)^ := PSmallint(@Res)^;
      CELLXL_NUM_32,
      CELLXL_NUM_2DEC,
      CELLXL_NUM_4DEC  : PInteger(P)^ := PInteger(@Res)^;
      CELLXL_NUM_SINGLE: PSingle(P)^ := PSingle(@Res)^;
      CELLXL_NUM_DOUBLE: PDouble(P)^ := Res;
      else raise XLSRWException.Create('Invalid XLate');
    end;

    FStyles.XFEditor.UseStyle(AStyle);
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow: integer; const AFormula: AxUCString; const AValue: double);
begin
  FFormulaHelper.Clear;
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Formula := AFormula;
  FFormulaHelper.AsFloat := AValue;
  FFormulaHelper.Style := XLS_STYLE_DEFAULT_XF;

  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const AFormula: AxUCString; const AValue: double);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Formula := AFormula;
  FFormulaHelper.AsFloat := AValue;

  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: double);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;

  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Ptgs := APtgs;
  FFormulaHelper.PtgsSize := APtgsSize;

  FFormulaHelper.AsFloat := AValue;
  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const AFormula: AxUCString; const AValue: boolean);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Formula := AFormula;
  FFormulaHelper.AsBoolean := AValue;

  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const AFormula, AValue: AxUCString);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Formula := AFormula;
  FFormulaHelper.AsString := AValue;

  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const AFormula: AxUCString; const AValue: TXc12CellError);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Formula := AFormula;
  FFormulaHelper.AsError := AValue;

  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula97(const ACol, ARow: integer; APtgs: Pointer; const APtgsSz: integer; const AStyle: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;
  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.AllocPtgs97(APtgs, APtgsSz);
  FFormulaHelper.AsFloat := 0;

  StoreFormula;
end;

procedure TXLSCellMMU.DEBUG_AddFormulaFloat(const ACol, ARow, AStyle: integer; const AValue: double; const APtgs: PXLSPtgs; APtgsSize: integer);
var
  i: integer;
  P: TMMUPtr;
  Sz: integer;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.AddFormulaFloat(ACol,ARow,AStyle,AValue,APtgs,APtgsSize)
  else begin
{$endif}
    Sz := SizeOf(double) + APtgsSize;
    if AStyle <> 0 then
      Inc(Sz,CELLMMU_STYLE_SZ);

    ExtendBlocks(ARow);

    P := FBlockManager.AllocMemEx(ARow and XBLOCK_VECTOR_MASK,ACol,CELLXL_FMLA_FLOAT,Sz);

    Inc(P,SZ_HDR_EX);

    if AStyle <> XLS_STYLE_DEFAULT_XF then begin
      PWord(P)^ := AStyle;
      P := Pointer(NativeInt(P) + SizeOf(word));
    end;

    PDouble(P)^ := AValue;
    Inc(P,SizeOf(double));

    for i := 0 to APtgsSize - 1 do
      PByteArray(P)[i] := $DC;
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.AddFormulaVal(const ACell: PXLSCellItem; const ACol, ARow: integer; const AValue: boolean);
var
  P: TMMUPtr;
begin
  if (ACell.Data <> Nil) and ((ACell.XLate and XVECT_XLATE_EX) <> 0) then begin
    if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_BOOLEAN then begin
      P := ACell.Data;
      Inc(P,SZ_HDR_EX + SizeOf(byte));
      if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
        Inc(P,CELLMMU_STYLE_SZ);
      if AValue then
        PByte(P)^ := 1
      else
        PByte(P)^ := 0;
    end
    else begin
      RetriveFormula(ACell);
      FFormulaHelper.Col := ACol;
      FFormulaHelper.Row := ARow;
      FFormulaHelper.AsBoolean := AValue;
      StoreFormula;
    end;
  end
  else
    raise XLSRWException.Create('Cell is not formula');
end;

procedure TXLSCellMMU.AddFormulaVal(const ACell: PXLSCellItem; const ACol, ARow: integer; const AValue: TXc12CellError);
var
  P: TMMUPtr;
begin
  if (ACell.Data <> Nil) and ((ACell.XLate and XVECT_XLATE_EX) <> 0) then begin
    if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_ERROR then begin
      P := ACell.Data;
      Inc(P,SZ_HDR_EX + SizeOf(byte));
      if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
        Inc(P,CELLMMU_STYLE_SZ);
      PByte(P)^ := Byte(AValue);
    end
    else begin
      RetriveFormula(ACell);
      FFormulaHelper.Col := ACol;
      FFormulaHelper.Row := ARow;
      FFormulaHelper.AsError := AValue;
      StoreFormula;
    end;
  end
  else
    raise XLSRWException.Create('Cell is not formula');
end;

procedure TXLSCellMMU.AddFormulaVal(const ACell: PXLSCellItem; const ACol, ARow: integer; const AValue: double);
var
  P: TMMUPtr;
begin
  if (ACell.Data <> Nil) and ((ACell.XLate and XVECT_XLATE_EX) <> 0) then begin
    if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_FLOAT then begin
      P := ACell.Data;
      Inc(P,SZ_HDR_EX + SizeOf(byte));
      if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
        Inc(P,CELLMMU_STYLE_SZ);
      PDouble(P)^ := AValue;
    end
    else begin
      RetriveFormula(ACell);
      FFormulaHelper.Col := ACol;
      FFormulaHelper.Row := ARow;
      FFormulaHelper.AsFloat := AValue;
      StoreFormula;
    end;
  end
  else
    raise XLSRWException.Create('Cell is not formula');
end;

procedure TXLSCellMMU.AddFormulaVal(const ACell: PXLSCellItem; const ACol, ARow: integer; const AValue: AxUCString);
var
  S: AxUCString;
begin
  if (ACell.Data <> Nil) and ((ACell.XLate and XVECT_XLATE_EX) <> 0) then begin
    S := GetFormulaValString(ACell);
    if AValue <> S then begin
      RetriveFormula(ACell);
      FFormulaHelper.Col := ACol;
      FFormulaHelper.Row := ARow;
      FFormulaHelper.AsString := AValue;
      StoreFormula;
    end;
  end
  else
    raise XLSRWException.Create('Cell is not formula');
end;

function TXLSCellMMU.AddRow(const ARow, AStyle: integer): PXLSMMURowHeader;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then begin
    FDebugData.AddRow(ARow,0);
    Result := Nil;
  end
  else begin
{$endif}
    ExtendBlocks(ARow);

    if (FBLockManager.Block.VectorsOffs[ARow and XBLOCK_VECTOR_MASK] <> Integer(XBLOCK_VECTOR_UNUSED)) then begin
      Result := PXLSMMURowHeader(FBlockManager.GetItemsHeader(ARow and XBLOCK_VECTOR_MASK));
      FStyles.XFEditor.FreeStyle(Result.Style);
    end;

    Result := PXLSMMURowHeader(FBLockManager.NewVector(ARow and XBLOCK_VECTOR_MASK));
    System.Move(FDefaultRowHdr,Result^,SizeOf(TXLSMMURowHeader));
    Result.Style := AStyle;
    FStyles.XFEditor.UseStyle(Result.Style)
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.StoreString(const ACol, ARow, AStyle: integer; const AValue: AxUCString);
var
  i: integer;
begin
  i := FSST.AddString(AValue);
  StoreString(ACol,ARow,AStyle,i);
end;

function TXLSCellMMU.AdjustFormulaMem(const ACell: PXLSCellItem; const ACol,ARow,ANewSize: integer; out AMem: TMMUPtr; out ASize: integer): TMMUPtr;
var
  P,P2: TMMUPtr;
  Delta: integer;
begin
  P := ACell.Data;

  ASize := PWord(ACell.Data)^;
  GetMem(AMem,ASize);

  Inc(P,SZ_HDR_EX);
  System.Move(P^,AMem^,ASize);
  P := AMem;

  FBlockManager.FreeMem(ARow and XBLOCK_VECTOR_MASK,ACol);

  Inc(P,SizeOf(byte));
  if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
    Inc(P,CELLMMU_STYLE_SZ);

  Result := P;

  Delta := ANewSize - GetFormulaValueSize(P,ACell.XLate);
  if Delta <> 0 then begin
    P2 := P;
    Inc(P2,SizeOf(byte));

    Inc(ASize,Delta);

    if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
      System.Move(P2^,P^,ASize - SizeOf(byte) - CELLMMU_STYLE_SZ)
    else
      System.Move(P2^,P^,ASize - SizeOf(byte));
  end;
end;

procedure TXLSCellMMU.StoreString(const ACol, ARow, AStyle, AValue: integer);
var
  P: TMMUPtr;
  XLate: byte;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.AddString(ACol,ARow,AStyle,AValue)
  else begin
{$endif}
    if AValue <= $FFFF then
      XLate := CELLXL_STR_16
    else
      XLate := CELLXL_STR_32;

    if AStyle <> XLS_STYLE_DEFAULT_XF then
      XLate := XLate or CELLXL_FLAG_STYLE;

    if ExtendBlocks(ARow) then
      FStyles.XFEditor.UseDefault;

    P := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);

    if AStyle <> XLS_STYLE_DEFAULT_XF then begin
      PWord(P)^ := AStyle;
      Inc(P,SizeOf(word));
    end;

    if (XLate and XVECT_XLATE_MASK) = CELLXL_STR_16 then
      PWord(P)^ := AValue
    else
      PInteger(P)^ := AValue;

    FSST.UsesString(AValue);

    FStyles.XFEditor.UseStyle(AStyle);
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

function TXLSCellMMU.AsFloat(ACol, ARow: integer; var Aresult: double; out AEmptyCell: boolean): TXc12CellError;
var
  Cell: TXLSCellItem;
begin
  Result := errUnknown;
  AResult := 0;
  AEmptyCell := not FindCell(ACol,ARow,Cell);
  if not AEmptyCell then begin
    case CellType(@Cell) of
      xctNone,
      xctBlank         : AEmptyCell := True;
      xctError         : Result := GetError(@Cell);
      xctString        : AEmptyCell := True;
      xctFloat         : AResult := GetFloat(@Cell);
//      xctCurrency      : AResult := GetFloat(@Cell);
      xctBoolean       : AEmptyCell := True;
      xctFloatFormula  : AResult := GetFormulaValFloat(@Cell);
      xctStringFormula : AEmptyCell := True;
      xctBooleanFormula: AEmptyCell := True;
      xctErrorFormula  : Result := GetFormulaValError(@Cell);
    end;
  end;
end;

function TXLSCellMMU.AsString(const ACol, ARow: integer; var Aresult: AxUCString): boolean;
var
  Cell: TXLSCellItem;
begin
  AResult := '';
  Result := FindCell(ACol,ARow,Cell);
  if Result then begin
    case CellType(@Cell) of
      xctNone,
      xctBlank         : Result := False;
      xctError         : AResult := Xc12CellErrorNames[GetError(@Cell)];
      xctString        : AResult := GetString(@Cell);
      xctFloat         : AResult := FloatToStr(GetFloat(@Cell));
      xctBoolean       : if GetBoolean(@Cell) then AResult := G_StrTRUE else AResult := G_StrFALSE;
      xctFloatFormula  : AResult := FloatToStr(GetFormulaValFloat(@Cell));
      xctStringFormula : AResult := GetFormulaValString(@Cell);
      xctBooleanFormula: if GetFormulaValBoolean(@Cell) then AResult := G_StrTRUE else AResult := G_StrFALSE;
      xctErrorFormula  : AResult := Xc12CellErrorNames[GetFormulaValError(@Cell)];
    end;
  end;
end;

procedure TXLSCellMMU.BeginDebug;
begin
  FDebugMem := SizeOf(TXVectHeader) + SizeOf(TXLSMMURowHeader);
  FDebugList.Add(Format('[Headers = %d bytes]',[FDebugMem]));
end;

procedure TXLSCellMMU.BeginIterate(const ARow: integer = 0);
begin
  BeginIterateRow(ARow - 1);
  if IterateNextRow then
    BeginIterateCol;
end;

procedure TXLSCellMMU.BeginIterateCol;
begin
  FIterate.Col := -1;
end;

procedure TXLSCellMMU.BeginIterateRow(const ARow: integer = -1);
begin
  FIterate.Row := ARow;
  FIterate.LastRow := GetLastRow;
end;

function TXLSCellMMU.CalcDimensions: boolean;
var
  i: integer;
  C1,C2: integer;
  R1,R2: integer;
  BlockOffs: integer;
begin
  Result := False;

  BlockOffs := 0;
  ClearCellArea(FDimension);
  if Length(FBlocks) > 0 then begin
    FDimension.Col1 := MAXINT;
    FDimension.Col2 := -(MAXINT - 1);
    FDimension.Row1 := MAXINT;
    FDimension.Row2 := -(MAXINT - 1);
    for i := 0 to High(FBlocks) do begin
      if FBlocks[i].Memory <> Nil then begin
        FBlockManager.Block := @FBlocks[i];
        FBlockManager.CalcDimensions(R1,R2,C1,C2);

        Inc(R1,BlockOffs);
        Inc(R2,BlockOffs);

        if (C1 >= 0) and (C1 < FDimension.Col1) then
          FDimension.Col1 := C1;
        if C2 > FDimension.Col2 then
          FDimension.Col2 := C2;
        if (R1 >= 0) and (R1 < FDimension.Row1) then
          FDimension.Row1 := R1;
        if R2 > FDimension.Row2 then
          FDimension.Row2 := R2;

        Inc(BlockOffs,XBLOCK_VECTOR_COUNT);
      end;
    end;
  end;
  if (FDimension.Col1 > XLS_MAXCOLS) or (FDimension.Col2 > XLS_MAXCOLS) or (FDimension.Row1 > XLS_MAXROWS) or (FDimension.Row2 > XLS_MAXROWS) then begin
    FDimension.Col1 := 0;
    FDimension.Col2 := 0;
    FDimension.Row1 := 0;
    FDimension.Row2 := 0;

    Result := True;
  end;
end;

procedure TXLSCellMMU.CalcRowDimensions(ARow: integer; out ACol1, ACol2: integer);
var
  iBlock: integer;
begin
//  if High(FBlocks) < 0 then
//    ARow := 1000000;

  iBlock := ARow shr XBLOCK_VECTOR_COUNT_BITS;
  if iBlock <= High(FBlocks) then begin
    FBlockManager.Block := @FBlocks[iBlock];
    FBlockManager.CalcDimensions(ARow,ARow,ACol1,ACol2);
  end
  else begin
    ACol1 := -1;
    ACol2 := -1;
  end;
end;

function TXLSCellMMU.CellType(const ACell: PXLSCellItem): TXLSCellType;
begin
  Result := XLateToCellType(ACell.XLate);
end;

function TXLSCellMMU.CheckCol(const ACol: integer): boolean;
begin
  Result := (ACol >= 0) and (ACol <= XLS_MAXCOL);
end;

function TXLSCellMMU.CheckIfRowHeaderAssigned(AHeader: TMMUPtr): boolean;
begin
  Result := False;
end;

function TXLSCellMMU.CheckIntegrity(const AList: TStrings): boolean;
var
  i: integer;
  Res: boolean;
begin
  Result := True;
  for i := 0 to High(FBlocks) do begin
    if FBlocks[i].Memory <> Nil then begin
      FBlockManager.Block := @FBlocks[i];
      Res := FBlockManager.WalkVectors(AList,i * XBLOCK_VECTOR_COUNT);
      if not Res then
        Result := Res;
    end;
  end;
end;

function TXLSCellMMU.CheckRow(const ARow: integer): boolean;
begin
  Result := (ARow >= 0) and (ARow <= XLS_MAXROW);
end;

procedure TXLSCellMMU.Clear;
var
  i: integer;
begin
  for i := 0 to High(FBlocks) do
    System.FreeMem(FBlocks[i].Memory);
  SetLength(FBlocks,0);
  ClearCellArea(FDimension);

  ExtendBlocks(0);
end;

procedure TXLSCellMMU.ClearCell(const ACol, ARow: integer);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) and (XLateToCellType(Cell.XLate) <> xctBlank) then begin
    Style := GetStyle(@Cell);
    if Style <> XLS_STYLE_DEFAULT_XF then
      StoreBlank(ACol,ARow,Style)
    else
      DeleteCell(ACol,ARow);
  end;
end;

procedure TXLSCellMMU.ClearFormulaKeepValue(ACol,ARow: integer);
var
  Cell: TXLSCellItem;
  S   : integer;
  vF  : double;
  vS  : AxUCString;
  vB  : boolean;
  vE  : TXc12CellError;
begin
  if FindCell(ACol,ARow,Cell) then begin
    S := GetStyle(@Cell);
    case Cell.XLate and XVECT_XLATE_MASK of
      CELLXL_FMLA_FLOAT  : begin
        vF := GetFormulaValFloat(@Cell);
        StoreFloat(ACol,ARow,S,vF);
      end;
      CELLXL_FMLA_STR    : begin
        vS := GetFormulaValString(@Cell);
        StoreString(ACol,ARow,S,vS);
      end;
      CELLXL_FMLA_BOOLEAN: begin
        vB := GetFormulaValBoolean(@Cell);
        StoreBoolean(ACol,ARow,S,vB);
      end;
      CELLXL_FMLA_ERROR  : begin
        vE := GetFormulaValError(@Cell);
        StoreError(ACol,ARow,S,vE);
      end;
    end;
  end;
end;

procedure TXLSCellMMU.ClearRowHeaders(const ARow1, ARow2: integer);
var
  R: integer;
  Row: PXLSMMURowHeader;
begin
  for R := ARow1 to ARow2 do begin
    Row := FindRow(R);
    if Row <> Nil then begin
      FStyles.XFEditor.FreeStyle(Row.Style);
      if FBlockManager._VectorItemsCount > 0 then
        System.Move(FDefaultRowHdr,Row^,SizeOf(TXLSMMURowHeader))
      else
        FBlockManager.FreeVector(R and XBLOCK_VECTOR_MASK);
    end;
  end;
end;

// $1B $00 $00 $00 $00 $00 $00 $00 $00 $00 $00 $0F $00 $FF $84 $06 $00 $00 $00 $06 $00 $84 $06 $00 $00 $00 $07 $00 $03
function TXLSCellMMU.CopyCell(const ACol, ARow: integer): TXLSCellItem;
var
  P: TMMUPtr;
  Sz: integer;
  XLate: integer;
begin
  Result.Data := Nil;
  if SelectBlock(ARow) then begin
    P := FBlockManager.GetMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate);
    if P <> Nil then begin
      Sz := ItemSize(XLate,P);
      System.GetMem(Result.Data,Sz);
      System.Move(P^,Result.Data^,Sz);
      Result.XLate := XLate;
      if CellType(@Result) = xctString then
        FSST.UsesString(GetStringSST(@Result));
    end;
  end;
end;

procedure TXLSCellMMU.CopyRowHeaders(const ARow1, ARow2, ADestRow: integer);
begin
  CopyRowHeaders(Self,ARow1,ARow2,ADestRow);
end;

procedure TXLSCellMMU.CopyRowHeaders(ASrcMMU: TXLSCellMMU; const ARow1, ARow2, ADestRow: integer);
var
  R: integer;
  RowSrc,
  RowDest: PXLSMMURowHeader;
begin
  for R := ARow1 to ARow2 do begin
    RowSrc := ASrcMMU.FindRow(R);
    RowDest := FindRow(ADestRow + (R - ARow1));
    if (RowSrc <> Nil) and (RowDest <> Nil) then begin
      FStyles.XFEditor.FreeStyle(RowDest.Style);
      FStyles.XFEditor.UseStyle(RowSrc.Style);
      System.Move(RowSrc^,RowDest^,SizeOf(TXLSMMURowHeader));
    end
    else if RowSrc <> Nil then begin
      RowDest := AddRow(ADestRow + (R - ARow1),XLS_STYLE_DEFAULT_XF);

      FStyles.XFEditor.FreeStyle(RowDest.Style);
      FStyles.XFEditor.UseStyle(RowSrc.Style);
      System.Move(RowSrc^,RowDest^,SizeOf(TXLSMMURowHeader));
    end
    else if RowDest <> Nil then begin
      System.Move(FDefaultRowHdr,RowDest^,SizeOf(TXLSMMURowHeader));
      FStyles.XFEditor.UseStyle(RowDest.Style);
    end;
  end;
end;

constructor TXLSCellMMU.Create(const ASheetIndex: integer; const ADebugList: TStrings; const ASST: TXc12DataSST; const AStyles: TXc12DataStyleSheet);
begin
  FSheetIndex := ASheetIndex;
  FDebugList := ADebugList;
  FSST := ASST;
  FStyles := AStyles;
  FFormulaHelper := TXLSFormulaHelper.Create;


  ClearCellArea(FDimension);

  FDebugData := TXLSCellMMUDebugData.Create(Self);

  FBlockManager := TXLSMMUBlockManager.Create;

  G_XLSCellMMU5_KeepItemsHeaderEvent := CheckIfRowHeaderAssigned;

  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_8,1);
  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_16,2);
  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_32,4);
  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_2DEC,4);
  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_4DEC,4);
  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_SINGLE,4);
  TXLSMMUVectorManager.SetXLate(CELLXL_NUM_DOUBLE,8);
  TXLSMMUVectorManager.SetXLate(CELLXL_STR_16,2);
  TXLSMMUVectorManager.SetXLate(CELLXL_STR_32,4);
  TXLSMMUVectorManager.SetXLate(CELLXL_BOOLEAN,1);
  TXLSMMUVectorManager.SetXLate(CELLXL_ERROR,1);

  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_BLANK,0 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_8,1 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_16,2 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_32,4 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_2DEC,4 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_4DEC,4 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_SINGLE,4 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_NUM_DOUBLE,8 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_STR_16,2 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_STR_32,4 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_BOOLEAN,1 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_ERROR,1 + 2);

  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_FMLA_FLOAT,8 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_FMLA_STR,2 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_FMLA_BOOLEAN,1 + 2);
  TXLSMMUVectorManager.SetXLate(CELLXL_FLAG_STYLE or CELLXL_FMLA_ERROR,1 + 2);

  FDefaultRowHdr.Options := [];
  FDefaultRowHdr.Style := XLS_STYLE_DEFAULT_XF;
  FDefaultRowHdr.OutlineLevel := 0;
  FDefaultRowHdr.Height := XLS_DEFAULT_ROWHEIGHT;

  TXLSMMUVectorManager.SetItemsHeaderSize(SizeOf(TXLSMMURowHeader),@FDefaultRowHdr,SizeOf(TXLSMMURowHeader));

  ExtendBlocks(0);
end;

procedure TXLSCellMMU.DebugAddFloat(const AValue: double);
var
  XLate: byte;
  S: string;
  Sz: integer;
  Bytes: PByteArray;
  Res: double;
begin
  XLate := EncodeNumber(AValue,Res);
  S := Format('%.4X %.2X,%.2X                    Item# %d',[FDebugMem,1,XLate,FDebugCnt]);
  FDebugList.Add(S);
  Inc(FDebugMem,2);
  Inc(FDebugCnt);

  Bytes := PByteArray(@Res);
  case XLate of
    CELLXL_NUM_8     : Sz := 1;
    CELLXL_NUM_16    : Sz := 2;
    CELLXL_NUM_32    : Sz := 4;
    CELLXL_NUM_2DEC  : Sz := 4;
    CELLXL_NUM_4DEC  : Sz := 4;
    CELLXL_NUM_SINGLE: Sz := 4;
    CELLXL_NUM_DOUBLE: Sz := 8;
    else raise XLSRWException.Create('Invalid XLate code');
  end;

  S := Format('%.4X %s',[FDebugMem,BytesToHexStr(Bytes,Sz)]);
  FDebugList.Add(S);
  Inc(FDebugMem,Sz);
end;

procedure TXLSCellMMU.DebugAddEmpty(const ACount: integer);
var
  i: integer;
  n: integer;
  S: string;
begin
  n := ACount div XVECT_MAX_EQCOUNT;
  for i := 0 to n - 1 do begin
    S := Format('%.4X %.2X                       Item# %d..%d Empty',[FDebugMem,XVECT_MAX_EQCOUNT or XVECT_EQCOUNT_EMPTY,FDebugCnt,FDebugCnt + XVECT_MAX_EQCOUNT - 1]);
    FDebugList.Add(S);
    Inc(FDebugCnt,XVECT_MAX_EQCOUNT);
    Inc(FDebugMem);
  end;
  n := ACount - (n * XVECT_MAX_EQCOUNT);
  if n > 0 then begin
    if n > 1 then
      S := Format('%.4X %.2X                       Item# %d..%d Empty',[FDebugMem,n or XVECT_EQCOUNT_EMPTY,FDebugCnt,FDebugCnt + n - 1])
    else
      S := Format('%.4X %.2X                       Item# %d Empty',[FDebugMem,n or XVECT_EQCOUNT_EMPTY,FDebugCnt,FDebugCnt + n]);
    FDebugList.Add(S);
    Inc(FDebugCnt,n);
    Inc(FDebugMem);
  end;
end;

procedure TXLSCellMMU.DebugAddFloat(const AValues: array of double);
var
  XLate: byte;
  S: string;
  i: integer;
  n: integer;
  Sz: integer;
  Bytes: PByteArray;
  Res: double;
begin
  n := Length(AValues);
  XLate := EncodeNumber(AValues[0],Res);
  S := Format('%.4X %.2X,%.2X                    Item# %d..%d',[FDebugMem,n,XLate,FDebugCnt,FDebugCnt + n - 1]);
  FDebugList.Add(S);
  Inc(FDebugMem,2);
  Inc(FDebugCnt,n);

  for i := 0 to n - 1 do begin
    Bytes := @AValues[i];
    case XLate of
      CELLXL_NUM_8     : Sz := 1;
      CELLXL_NUM_16    : Sz := 2;
      CELLXL_NUM_32    : Sz := 4;
      CELLXL_NUM_2DEC  : Sz := 4;
      CELLXL_NUM_4DEC  : Sz := 4;
      CELLXL_NUM_SINGLE: Sz := 4;
      CELLXL_NUM_DOUBLE: Sz := 8;
      else raise XLSRWException.Create('Invalid XLate code');
    end;

    S := Format('%.4X %s',[FDebugMem,BytesToHexStr(Bytes,Sz)]);
    FDebugList.Add(S);
    Inc(FDebugMem,Sz);
  end;
end;

procedure TXLSCellMMU.DebugAddFormulaFloat(const AValue: double; const ASize: integer);
var
  S: string;
  Bytes: PByteArray;
begin
  S := Format('%.4X %.2X,%.2X %.4X               Item# %d [Formula]',[FDebugMem,1,XVECT_XLATE_EX,ASize,FDebugCnt]);
  FDebugList.Add(S);
  Inc(FDebugMem,4);
  Inc(FDebugCnt,1);

  Bytes := PByteArray(@AValue);
  S := Format('%.4X %s',[FDebugMem,BytesToHexStr(Bytes,8)]);
  FDebugList.Add(S);
  Inc(FDebugMem,8);

  S := Format('%.4X [Ptgs: %d]',[FDebugMem,ASize]);
  FDebugList.Add(S);
  Inc(FDebugMem,ASize);
end;

procedure TXLSCellMMU.DebugTest(const ATest: AxUCString; AList: TStrings);
begin
  if Lowercase(ATest) = 'walkvectors' then
    FBlockManager.WalkVectors(AList)
  else
    AList.Add('Unkown debug test: ' + ATest);
end;

procedure TXLSCellMMU.DeleteAndShiftColumns(const ACol1, ACol2: integer);
var
  i: integer;
begin
  if not CheckCol(ACol1) or not CheckCol(ACol2) or (ACol1 > ACol2) then
    Exit;

//  for i := 0 to High(FBlocks) do begin
//    if (FBlocks[i].Memory <> nil) and SelectBlock(i) then
//      FBlockManager.MoveMemAll(ACol2 + 1,ACol1);
//  end;


//  FBlockManager.MoveMem(2,ACol2 + 1,ACol1 + 1);
//  Exit;

// Changed 2013-06-10
  for i := 0 to High(FBlocks) do begin
    if FBlocks[i].Memory <> nil then begin


      FBlockManager.Block := @FBlocks[i];
      FBlockManager.MoveMemAll(ACol2 + 1,ACol1);
    end;
  end;
end;

procedure TXLSCellMMU.DeleteAndShiftRows(const ARow1, ARow2: integer);
var
  i1,i2: integer;
  R1,R2: integer;
  MaxRow: integer;
  P: TMMUPtr;
begin
  if not CheckRow(ARow1) or not CheckRow(ARow2) or (ARow1 > ARow2) or (High(FBlocks) < 0) then
    Exit;

  DeleteRows(ARow1,ARow2);

  R1 := ARow1;
  R2 := ARow2 + 1;
  MaxRow := XBLOCK_VECTOR_COUNT * Length(FBlocks) - 1;

  while R2 <= MaxRow do begin
    i1 := FindBlock(R1);
    i2 := FindBlock(R2);

    if i1 = i2 then begin
      FBlockManager.Block := @FBlocks[i1];
      FBlockManager.MoveVector(R2 and XBLOCK_VECTOR_MASK,R1 and XBLOCK_VECTOR_MASK);
    end
    else begin
      FBlockManager.Block := @FBlocks[i2];
      P := FBlockManager.CopyVector(R2 and XBLOCK_VECTOR_MASK);
      FBlockManager.Block := @FBlocks[i1];
      if P <> Nil then begin
        FBlockManager.InsertVector(R1 and XBLOCK_VECTOR_MASK,P);
        FreeMem(P);
      end
      else
        FBlockManager.FreeVector(R1 and XBLOCK_VECTOR_MASK);
    end;

    Inc(R1);
    Inc(R2);
  end;

  TrimEndBlocks;
end;

procedure TXLSCellMMU.DeleteCell(const ACol, ARow: integer);
var
  Style: integer;
  RowStyle: integer;
  Cell: TXLSCellItem;
  RowHeader: PXLSMMURowHeader;
begin
{$ifdef MMU_DEBUG}
  if FDebugSave then
    FDebugData.Delete(ACol,ARow)
  else begin
{$endif}

    if FindCell(ACol,ARow,Cell) then begin
      Style := GetStyle(@Cell);
      FStyles.XFEditor.FreeStyle(Style);

      if CellType(@Cell) = xctString then
        FSST.ReleaseString(GetStringSST(@Cell));
    end;

    if SelectBlock(ARow) then begin
      RowHeader := PXLSMMURowHeader(FBlockManager.GetItemsHeader(ARow and XBLOCK_VECTOR_MASK));
      if RowHeader <> Nil then
        RowStyle := RowHeader.Style
      else
        RowStyle := -1;

      FBlockManager.FreeMem(ARow and XBLOCK_VECTOR_MASK,ACol);

      // Free row style if the entire row was deleted.
      if (RowStyle >= 0) and (FBlockManager._VectorItemsCount = 0) then begin
        FStyles.XFEditor.FreeStyle(RowStyle);
        FBlockManager.FreeVector(ARow and XBLOCK_VECTOR_MASK);
      end;

//      if (RowStyle >= 0) and (FBlockManager.GetItemsHeader(ARow and XBLOCK_VECTOR_MASK) = Nil) then
//        FStyles.XFEditor.FreeStyle(RowStyle);
    end;
{$ifdef MMU_DEBUG}
  end;
{$endif}
end;

procedure TXLSCellMMU.DeleteColumns(const ACol1, ACol2: integer);
var
  i: integer;
begin
  if not CheckCol(ACol1) or not CheckCol(ACol2) or (ACol1 > ACol2) then
    Exit;

  for i := 0 to High(FBlocks) do begin
    if (FBlocks[i].Memory <> nil) and SelectBlock(i) then
      FBlockManager.FreeMemAll(ACol1,ACol2);
  end;
end;

procedure TXLSCellMMU.DeleteRows(const ARow1, ARow2: integer);
var
  i: integer;
  i1,i2: integer;
  RowHdr: PXLSMMURowHeader;
begin
  if not CheckRow(ARow1) or not CheckRow(ARow2) or (ARow1 > ARow2) then
    Exit;

  i1 := FindBlock(ARow1);
  i2 := FindBlock(ARow2);

  if i1 < 0 then
    Exit;

  i := -1;
  BeginIterate(ARow1);
  while IterateNext and (IterCellRow <= ARow2) do begin
    FStyles.XFEditor.FreeStyle(IterGetStyleIndex);
    if IterCellRow <> i then begin
      RowHdr := IterRow;
      if RowHdr <> Nil then
        FStyles.XFEditor.FreeStyle(RowHdr.Style);
      i := IterCellRow;
    end;
  end;

  if i1 = i2 then begin
    FBlockManager.Block := @FBlocks[i1];
    FBlockManager.FreeVectors(ARow1 and XBLOCK_VECTOR_MASK,ARow2 and XBLOCK_VECTOR_MASK);
  end
  else begin
    FBlockManager.Block := @FBlocks[i1];
    FBlockManager.FreeVectors(ARow1 and XBLOCK_VECTOR_MASK,XBLOCK_VECTOR_COUNT - 1);

    for i := i1 + 1 to i2 - 1 do
      FreeBlock(i);

    FBlockManager.Block := @FBlocks[i2];
    FBlockManager.FreeVectors(0,ARow2 and XBLOCK_VECTOR_MASK);
  end;

  TrimEndBlocks;
end;

destructor TXLSCellMMU.Destroy;
begin
  Clear;

  FBlockManager.Free;
  FDebugData.Free;
  FFormulaHelper.Free;
  inherited;
end;

function TXLSCellMMU.EmptyCell(const ACol, ARow: integer): boolean;
var
  XLate: integer;
begin
  if (ARow >= 0) and SelectBlock(ARow) then
    Result := FBlockManager.GetMem(ARow and XBLOCK_VECTOR_MASK,ACol,XLate) = Nil
  else
    Result := True;
end;

function TXLSCellMMU.EncodeNumber(const AValue: double; out AResult: double): integer;
const
  XLS_CELL_EPSILON = 0.00000001;
var
  V: double;
  i64: Int64;
  vCurr: Currency;
  Delta: double;
begin
  if (Int(AValue) = 0) and (Frac(AValue) < XLS_CELL_EPSILON) then begin
    Result := CELLXL_NUM_DOUBLE;
    AResult := AValue;
    Exit;
  end;

{$ifdef DELPHI_2009_OR_LATER}
  {$ifdef CPUX64}
    V := AValue;
  {$else}
    V := RoundTo(AValue,-12);
  {$endif}
{$else}
// Don't works with D7 and earlier.
//  V := RoundTo(AValue,-12);
  V := AValue;
{$endif}
  if (AValue = MAXDOUBLE) or (AValue = MINDOUBLE) then begin
    Result := CELLXL_NUM_DOUBLE;
    AResult := V;
  end
  else if Frac(V) = 0 then begin
    if (V >= -128) and (V <= 127) then begin
      Result := CELLXL_NUM_8;
      PShortint(@AResult)^ := Round(V);
    end
    else if (V >= -32768) and (V <= 32767) then begin
      Result := CELLXL_NUM_16;
      PSmallint(@AResult)^ := Round(V);
    end
    else if (V >= XLS_MININT) and (V <= MAXINT) then begin
      Result := CELLXL_NUM_32;
      PInteger(@AResult)^ := Round(V);
    end
    else begin
      Result := CELLXL_NUM_DOUBLE;
      AResult := V;
    end;
  end
  else begin
{$ifdef DELPHI_5}
    Delta := 1;
{$else}
    Delta := RoundTo(AValue * 10000,0) - AValue * 10000;
{$endif}
    if (Delta > -XLS_CELL_EPSILON) and (Delta < XLS_CELL_EPSILON) then begin
// Changed 140923
      if (AValue > MAXINT) or (AValue < XLS_MININT) then begin
        Result := CELLXL_NUM_DOUBLE;
        AResult := V;
      end
      else begin
        vCurr := AValue;
        i64 := PInt64(@vCurr)^;
        if Frac(vCurr) = 0 then begin
          i64 := i64 div 10000;
          PInteger(@AResult)^ := i64;
          Result := CELLXL_NUM_32;
        end
        else if (Frac(vCurr * 100) = 0) and ((vCurr * 100) <= MAXINT) and ((vCurr * 100) >= XLS_MININT) then begin
          i64 := i64 div 100;
          PInteger(@AResult)^ := i64;
          Result := CELLXL_NUM_2DEC;
        end
        else if (i64 >= -MAXINT) and (i64 <= MAXINT) then begin
          PInteger(@AResult)^ := i64;
          Result := CELLXL_NUM_4DEC;
        end
        else begin
          Result := CELLXL_NUM_DOUBLE;
          AResult := V;
        end;
      end;
// End change

//        vCurr := AValue;
//        i64 := PInt64(@vCurr)^;
//        if (i64 > MAXINT) or (i64 < XLS_MININT) then begin
//          Result := CELLXL_NUM_DOUBLE;
//          AResult := V;
//        end
//        else if Frac(vCurr) = 0 then begin
//          i64 := i64 div 10000;
//          PInteger(@AResult)^ := i64;
//          Result := CELLXL_NUM_32;
//        end
//        else if Frac(vCurr * 100) = 0 then begin
//          i64 := i64 div 100;
//          PInteger(@AResult)^ := i64;
//          Result := CELLXL_NUM_2DEC;
//        end
//        else if (i64 >= -MAXINT) and (i64 <= MAXINT) then begin
//          PInteger(@AResult)^ := i64;
//          Result := CELLXL_NUM_4DEC;
//        end
//        else begin
//          Result := CELLXL_NUM_DOUBLE;
//          AResult := V;
//        end;
    end
    else begin
      Result := CELLXL_NUM_DOUBLE;
      AResult := V;
    end;
  end;
end;

procedure TXLSCellMMU.EndDebug;
begin
  FDebugList.Add('Memory used: ' + IntToStr(FDebugMem));
end;

procedure TXLSCellMMU.ExchangeCells(const ACol1, ARow1, ACol2, ARow2: integer);
var
  Ref: TXLSCellItem;
begin
  Ref := CopyCell(ACol1,ARow1);
  MoveCell(ACol2,ARow2,ACol1,ARow1);
  if Ref.Data <> Nil then begin
    InsertCell(ACol2,ARow2,@Ref);
    System.FreeMem(Ref.Data);
  end;
end;

function TXLSCellMMU.ExtendBlocks(ARow: integer): boolean;
var
  iBlock: integer;
begin
//  if High(FBlocks) < 0 then
//    ARow := 1000000;

  iBlock := ARow shr XBLOCK_VECTOR_COUNT_BITS;
  if iBlock > High(FBlocks) then
    AddBlocks(iBlock);

  FBlockManager.Block := @FBlocks[iBlock];

  Result := FBlockManager.Block.VectorsOffs[ARow and XBLOCK_VECTOR_MASK] = Integer(XBLOCK_VECTOR_UNUSED);
end;

function TXLSCellMMU.FindBlock(const ARow: integer): integer;
begin
  Result := ARow shr XBLOCK_VECTOR_COUNT_BITS;
  if Result > High(FBlocks) then
    Result := High(FBlocks);
end;

function TXLSCellMMU.FindCell(const ACol, ARow: integer): TXLSCellItem;
begin
  if (ARow >= 0) and SelectBlock(ARow) then
    Result.Data := FBlockManager.GetMem(ARow and XBLOCK_VECTOR_MASK,ACol,Result.XLate)
  else
    Result.Data := Nil;
end;

function TXLSCellMMU.FindCell(const ACol, ARow: integer; out ACell: TXLSCellItem): boolean;
begin
  Result := (ARow >= 0) and SelectBlock(ARow);
  if Result then begin
    ACell.Data := FBlockManager.GetMem(ARow and XBLOCK_VECTOR_MASK,ACol,ACell.XLate);
    Result := ACell.Data <> Nil;

//    if G_DoWrite then begin
//      G_DebugStream.Write(FSheetIndex,4);
//      G_DebugStream.Write(ACol,4);
//      G_DebugStream.Write(ARow,4);
//    end;
  end
  else
    ACell.Data := Nil;
end;

function TXLSCellMMU.FindRow(const ARow: integer): PXLSMMURowHeader;
begin
  if SelectBlock(ARow) then
    Result := PXLSMMURowHeader(FBlockManager.GetItemsHeader(ARow and XBLOCK_VECTOR_MASK))
  else
    Result := Nil;
end;

function TXLSCellMMU.FormulaType(const ACell: PXLSCellItem): TXLSCellFormulaType;
var
  P: TMMUPtr;
begin
  if (ACell.XLate and XVECT_XLATE_EX) <> 0 then begin
    P := ACell.Data;
    Inc(P,SZ_HDR_EX);
    Result := TXLSCellFormulaType(CELLMMU_FMLA_FLAG_MASK and PByte(P)^);
  end
  else
    raise XLSRWException.Create('Cell is not formula');
end;

procedure TXLSCellMMU.FreeBlock(const AIndex: integer);
begin
  if FBlocks[AIndex].Memory <> Nil then begin
    FreeMem(FBlocks[AIndex].Memory);
    InitBlock(AIndex);
  end;
end;

procedure TXLSCellMMU.FreeCell(const ACell: PXLSCellItem);
begin
  System.FreeMem(ACell.Data);
end;

function TXLSCellMMU.GetBoolean(const ACol, ARow: integer; out AValue: boolean): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell) and ((Cell.XLate and XVECT_XLATE_MASK) = CELLXL_BOOLEAN);
  if Result then
    AValue := PByte(GetValue(@Cell))^ = 1
  else
    AValue := False;
end;

function TXLSCellMMU.GetBlank(const ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell) and ((Cell.XLate and XVECT_XLATE_MASK) = CELLXL_BLANK);
end;

function TXLSCellMMU.GetBlank(const ACell: PXLSCellItem): boolean;
begin
  Result := (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_BLANK;
end;

function TXLSCellMMU.GetBoolean(const ACell: PXLSCellItem): boolean;
begin
  if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_BOOLEAN then
    Result := PByte(GetValue(ACell))^ = 1
  else
    Result := False;
end;

function TXLSCellMMU.GetError(const ACol, ARow: integer; out AValue: TXc12CellError): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell) and ((Cell.XLate and XVECT_XLATE_MASK) = CELLXL_ERROR);
  if Result then
    AValue := TXc12CellError(PByte(GetValue(@Cell))^)
  else
    AValue := errUnknown;
end;

function TXLSCellMMU.GetEmpty(const ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := not FindCell(ACol,ARow,Cell);
  if not Result then
    Result := (Cell.XLate and XVECT_XLATE_MASK) = CELLXL_BLANK;
end;

function TXLSCellMMU.GetError(const ACell: PXLSCellItem): TXc12CellError;
begin
  if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_ERROR then
    Result := TXc12CellError(PByte(GetValue(ACell))^)
  else
    Result := errUnknown;
end;

function TXLSCellMMU.GetFloat(const ACol, ARow: integer): double;
var
  Cell: TXLSCellItem;
  P: TMMUPtr;
begin
  if FindCell(ACol,ARow,Cell) then begin
    P := GetValue(@Cell);
    case Cell.XLate and XVECT_XLATE_MASK of
      CELLXL_NUM_8     : Result := PShortint(P)^;
      CELLXL_NUM_16    : Result := PSmallint(P)^;
      CELLXL_NUM_32    : Result := PInteger(P)^;
      CELLXL_NUM_2DEC  : Result := PInteger(P)^ / 100;
      CELLXL_NUM_4DEC  : Result := PInteger(P)^ / 10000;
      CELLXL_NUM_SINGLE: Result := PSingle(P)^;
      CELLXL_NUM_DOUBLE: Result := PDouble(P)^;
      CELLXL_FMLA_FLOAT: Result := PDouble(P)^;
      else               Result := 0;
    end;
  end
  else
    Result := 0;
end;

function TXLSCellMMU.GetFloat(const ACell: PXLSCellItem): double;
var
  P: TMMUPtr;
begin
  P := GetValue(ACell);
  case ACell.XLate and XVECT_XLATE_MASK of
    CELLXL_NUM_8     : Result := PShortint(P)^;
    CELLXL_NUM_16    : Result := PSmallint(P)^;
    CELLXL_NUM_32    : Result := PInteger(P)^;
    CELLXL_NUM_2DEC  : Result := PInteger(P)^ / 100;
    CELLXL_NUM_4DEC  : Result := PInteger(P)^ / 10000;
    CELLXL_NUM_SINGLE: Result := PSingle(P)^;
    CELLXL_NUM_DOUBLE: Result := PDouble(P)^;
    CELLXL_FMLA_FLOAT: Result := PDouble(P)^;
    else               Result := 0;
  end;
end;

function TXLSCellMMU.GetFormula(const ACol, ARow: integer): AxUCString;
var
  Cell: TXLSCellItem;
begin
  Cell := FindCell(ACol,ARow);
  if Cell.Data <> Nil then
    Result := GetFormulaStr(Cell.Data,Cell.XLate)
  else
    Result := '';
end;

function TXLSCellMMU.GetStyle(ACell: PXLSCellItem): integer;
var
  P: TMMUPtr;
begin
  if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then begin
    if (ACell.XLate and XVECT_XLATE_MASK) in [CELLXL_FMLA_FIRST..CELLXL_FMLA_LAST] then begin
      P := ACell.Data;
      Inc(P,SZ_HDR_EX + CELLMMU_FMLA_FLAG_SZ);
      Result := PWord(P)^
    end
    else
      Result := PWord(ACell.Data)^
  end
  else
    Result := XLS_STYLE_DEFAULT_XF;
end;

function TXLSCellMMU.GetFormula(const ACell: PXLSCellItem): AxUCString;
begin
  if ACell.Data <> Nil then
    Result := GetFormulaStr(ACell.Data,ACell.XLate)
  else
    Result := '';
end;

function TXLSCellMMU.GetFormulaFmla(const AFormula: TMMUPtr; const AXLate: integer; out ASize: integer): TMMUPtr;
var
  Sz: integer;
begin
  Result := AFormula;

  Inc(Result,SizeOf(word) + SizeOf(byte));

  if (AXLate and CELLXL_FLAG_STYLE) <> 0 then
    Inc(Result,SizeOf(word));

  case AXLate and XVECT_XLATE_MASK of
    CELLXL_FMLA_FLOAT  : Inc(Result,SizeOf(double));
    CELLXL_FMLA_STR    : begin
      Sz := PWord(Result)^;
      Inc(Result,SizeOf(word) + Sz * 2);
    end;
    CELLXL_FMLA_BOOLEAN: Inc(Result);
    CELLXL_FMLA_ERROR  : Inc(Result);
    else raise XLSRWException.Create('Invalid XLate code');
  end;

  ASize := PWord(Result)^;
  Inc(Result,SizeOf(word));
end;

function TXLSCellMMU.GetFormulaPtgs(const ACell: PXLSCellItem; out APtgs: PXLSPtgs): integer;
begin
  APtgs := PXLSPtgs(GetFormulaPtgs(ACell.Data,ACell.XLate,Result));
end;

function TXLSCellMMU.GetFormulaPtgs(const AFormula: TMMUPtr; const AXLate: integer; out ASize: integer): TMMUPtr;
begin
  Result := GetFormulaFmla(AFormula,AXLate,ASize);

  if PByte(Result)^ = CELLMMU_FMLA_IS_PTGS then
    Inc(Result)
  else begin
    raise XLSRWException.Create('Formula is not compiled');
  end;
end;

function TXLSCellMMU.GetFormulaStr(const AFormula: TMMUPtr; const AXLate: integer): AxUCString;
var
  i: integer;
  Sz: integer;
  P: TMMUPtr;
begin
  P := GetFormulaFmla(AFormula,AXLate,Sz);

  case PByte(P)^ of
    CELLMMU_FMLA_IS_PTGS  : raise XLSRWException.Create('Formula is compiled');
    CELLMMU_FMLA_IS_STR_8 : begin
      Inc(P,SizeOf(byte));
      SetLength(Result,Sz);
      for i := 1 to Sz do
        Result[i] := AxUCChar(PByteArray(P)[i - 1]);
    end;
    CELLMMU_FMLA_IS_STR_UC: begin
      Inc(P,SizeOf(byte));
      SetLength(Result,Sz);
      System.Move(P^,Pointer(Result)^,Sz * 2);
    end;
    else
      raise XLSRWException.Create('Invalid formula');
  end;
end;

function TXLSCellMMU.GetFormulaTargetArea(const ACell: PXLSCellItem): PXLSFormulaArea;
var
  P: TMMUPtr;
  Sz: integer;
begin
  Result := Nil;

  P := ACell.Data;

  Inc(P,SizeOf(word));

  if (PByte(P)^ and CELLMMU_FMLA_FLAG_ARRAY) <> 0 then begin
    Inc(P,SizeOf(byte));

    if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
      Inc(P,SizeOf(word));

    case ACell.XLate and XVECT_XLATE_MASK of
      CELLXL_FMLA_FLOAT  : Inc(P,SizeOf(double));
      CELLXL_FMLA_STR    : begin
        Sz := PWord(P)^;
        Inc(P,SizeOf(word) + Sz * 2);
      end;
      CELLXL_FMLA_BOOLEAN: Inc(P);
      CELLXL_FMLA_ERROR  : Inc(P);
      else raise XLSRWException.Create('Invalid XLate code');
    end;

    Sz := PWord(P)^;
    Inc(P,SizeOf(word) + SizeOf(byte) + Sz);

    Result := PXLSFormulaArea(P);
  end;
end;

function TXLSCellMMU.GetFormulaValBoolean(const ACol, ARow: integer; out AValue: boolean): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell);
  if Result then
    AValue := GetFormulaValBoolean(@Cell);
end;

function TXLSCellMMU.GetFormulaValBoolean(const ACell: PXLSCellItem): boolean;
var
  P: TMMUPtr;
begin
  if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_BOOLEAN then begin
    P := ACell.Data;
    Inc(P,SZ_HDR_EX + SizeOf(byte));
    if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
      Inc(P,CELLMMU_STYLE_SZ);
    Result := PByte(P)^ = 1;
  end
  else
    Result := False;
end;

function TXLSCellMMU.GetFormulaValError(const ACol, ARow: integer; out AValue: TXc12CellError): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell);
  if Result then
    AValue := GetFormulaValError(@Cell);
end;

function TXLSCellMMU.GetFormulaValError(const ACell: PXLSCellItem): TXc12CellError;
var
  P: TMMUPtr;
begin
  if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_ERROR then begin
    P := ACell.Data;
    Inc(P,SZ_HDR_EX + SizeOf(byte));
    if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
      Inc(P,CELLMMU_STYLE_SZ);
    Result := TXc12CellError(PByte(P)^);
  end
  else
    Result := errUnknown;
end;

function TXLSCellMMU.GetFormulaValFloat(const ACol, ARow: integer; out AValue: double): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell);
  if Result then
    AValue := GetFormulaValFloat(@Cell);
end;

function TXLSCellMMU.GetFormulaValFloat(const ACell: PXLSCellItem): double;
var
  P: TMMUPtr;
begin
  if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_FLOAT then begin
    P := ACell.Data;
    Inc(P,SZ_HDR_EX + SizeOf(byte));
    if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
      Inc(P,CELLMMU_STYLE_SZ);
    Result := PDouble(P)^;
  end
  else
    Result := 0;
end;

function TXLSCellMMU.GetFormulaValString(const ACol, ARow: integer; out AValue: AxUCString): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell);
  if Result then
    AValue := GetFormulaValString(@Cell);
end;

function TXLSCellMMU.GetFormulaValString(const ACell: PXLSCellItem): AxUCString;
var
  Sz: integer;
  P: TMMUPtr;
begin
  if (ACell.XLate and XVECT_XLATE_MASK) = CELLXL_FMLA_STR then begin
    P := ACell.Data;
    Inc(P,SZ_HDR_EX + SizeOf(byte));
    if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
      Inc(P,CELLMMU_STYLE_SZ);
    Sz := PWord(P)^;
    SetLength(Result,Sz);
    Inc(P,SizeOf(word));
    System.Move(P^,Pointer(Result)^,Sz * 2);
  end
  else
    Result := '';
end;

function TXLSCellMMU.GetFormulaValueSize(const AFormulaValue: TMMUPtr; const AXLate: integer): integer;
begin
  case AXLate and XVECT_XLATE_MASK of
    CELLXL_FMLA_FLOAT  : Result := 8;
    CELLXL_FMLA_STR    : Result := PWord(AFormulaValue)^;
    CELLXL_FMLA_BOOLEAN: Result := 1;
    CELLXL_FMLA_ERROR  : Result := 1;
    else raise XLSRWException.Create('Invalid XLate code');
  end;
end;

procedure TXLSCellMMU.InitBlock(const AIndex: integer);
var
  i: integer;
begin
  FBlocks[AIndex].MemSize := 0;
  FBlocks[AIndex].Memory := Nil;
  for i := 0 to XBLOCK_VECTOR_COUNT - 1 do
    FBlocks[AIndex].VectorsOffs[i] := Integer(XBLOCK_VECTOR_UNUSED);
end;

procedure TXLSCellMMU.InsertCell(const ACol, ARow: integer; const ACell: PXLSCellItem);
var
  Sz: integer;
  Style: integer;
  pDest: TMMUPtr;
begin
  if ExtendBlocks(ARow) then
    FStyles.XFEditor.UseDefault;

  if XLateIsFormula(ACell.XLate) then begin
    RetriveFormula(ACell);
    FFormulaHelper.Col := ACol;
    FFormulaHelper.Row := ARow;
    StoreFormula;

    // This don't works.
//    Sz := PWord(ACell.Data)^;
//    pDest := FBlockManager.AllocMemEx(ARow and XBLOCK_VECTOR_MASK,ACol,ACell.XLate,Sz);
//    System.Move(ACell.Data^,pDest^,Sz);
  end
  else begin
    Sz := ItemSize(ACell.XLate,ACell.Data);
    pDest := FBlockManager.AllocMem(ARow and XBLOCK_VECTOR_MASK,ACol,ACell.XLate);
    System.Move(ACell.Data^,pDest^,Sz);
  end;

  if CellType(ACell) = xctString then
    FSST.UsesString(GetStringSST(ACell));

  Style := GetStyle(ACell);
  FStyles.XFEditor.UseStyle(Style);
end;

procedure TXLSCellMMU.InsertColumns(const ACol, ACount: integer);
var
  i: integer;
begin
  if not CheckCol(ACol) then
    Exit;

  for i := 0 to High(FBlocks) do begin
    if FBlocks[i].Memory <> nil then begin
      FBlockManager.Block := @FBlocks[i];
//      SelectBlock(i * $FF + 1);
      FBlockManager.InsertEmptyAll(ACol,ACount);
    end;
  end;
end;

procedure TXLSCellMMU.InsertRows(const ARow, ACount: integer);
var
  i: integer;
  LastRow: integer;
  R1,R2: integer;
  i1,i2: integer;
  P: TMMUPtr;
begin
  if not CheckRow(ARow) or (ACount < 1) then
    Exit;

  LastRow := GetLastRow;
  if ARow > LastRow then
    Exit;

  ExtendBlocks(LastRow + ACount);

  R1 := LastRow;
  R2 := LastRow + ACount;

  while R1 > XLS_MAXROW do begin
    i1 := FindBlock(R1);
    FBlockManager.Block := @FBlocks[i1];
    FBlockManager.FreeVector(R1 and XBLOCK_VECTOR_MASK);
    Dec(R1);
    Dec(R2);
  end;

  while R1 >= ARow do begin
    i1 := FindBlock(R1);
    i2 := FindBlock(R2);

    if i1 = i2 then begin
      FBlockManager.Block := @FBlocks[i1];
      FBlockManager.MoveVector(R1 and XBLOCK_VECTOR_MASK,R2 and XBLOCK_VECTOR_MASK);
    end
    else begin
      FBlockManager.Block := @FBlocks[i1];
      P := FBlockManager.CopyVector(R1 and XBLOCK_VECTOR_MASK);
      FBlockManager.Block := @FBlocks[i2];
      if P <> Nil then begin
        FBlockManager.InsertVector(R2 and XBLOCK_VECTOR_MASK,P);
        FreeMem(P);
      end
      else
        FBlockManager.FreeVector(R2 and XBLOCK_VECTOR_MASK);
    end;
    Dec(R1);
    Dec(R2);
  end;
  for i := Arow to ARow + ACount - 1 do begin
    i1 := FindBlock(i);
    FBlockManager.Block := @FBlocks[i1];
    FBlockManager.FreeVector(i and XBLOCK_VECTOR_MASK);
  end;
end;

// 2015-02-26
procedure TXLSCellMMU.InsertRows_ORIG(const ARow, ACount: integer);
var
  i: integer;
  LastRow: integer;
  R1,R2: integer;
  i1,i2: integer;
  P: TMMUPtr;
begin
  if not CheckRow(ARow) or (ACount < 1) then
    Exit;

  LastRow := GetLastRow;
  if ARow > LastRow then
    Exit;

  ExtendBlocks(LastRow + ACount);

  R1 := LastRow;
  R2 := LastRow + ACount;

  while R1 > XLS_MAXROW do begin
    i1 := FindBlock(R1);
    FBlockManager.Block := @FBlocks[i1];
    FBlockManager.FreeVector(R1 and XBLOCK_VECTOR_MASK);
    Dec(R1);
    Dec(R2);
  end;

  while R1 >= ARow do begin
    i1 := FindBlock(R1);
    i2 := FindBlock(R2);

    if i1 = i2 then begin
      FBlockManager.Block := @FBlocks[i1];
      FBlockManager.MoveVector(R1 and XBLOCK_VECTOR_MASK,R2 and XBLOCK_VECTOR_MASK);
    end
    else begin
      FBlockManager.Block := @FBlocks[i1];
      P := FBlockManager.CopyVector(R1 and XBLOCK_VECTOR_MASK);
      FBlockManager.Block := @FBlocks[i2];
      if P <> Nil then begin
        FBlockManager.InsertVector(R2 and XBLOCK_VECTOR_MASK,P);
        FreeMem(P);
      end
      else
        FBlockManager.FreeVector(R1 and XBLOCK_VECTOR_MASK);
    end;
    Dec(R1);
    Dec(R2);
  end;
  for i := Arow to ARow + ACount - 1 do begin
    i1 := FindBlock(i);
    FBlockManager.Block := @FBlocks[i1];
    FBlockManager.FreeVector(i and XBLOCK_VECTOR_MASK);
  end;
end;

function TXLSCellMMU.IsEmpty: boolean;
begin
  Result := not CellAreaAssigned(FDimension);
end;

function TXLSCellMMU.IsFormula(const ACell: PXLSCellItem): boolean;
begin
  Result := CellType(ACell) in XLSCellTypeFormulas;
end;

function TXLSCellMMU.IsFormula(const ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Result := IsFormula(@Cell)
  else
    Result := False;
end;

function TXLSCellMMU.IsFormulaCompiled(const ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Result := IsFormulaCompiled(@Cell)
  else
    Result := False;
end;

function TXLSCellMMU.IsFormulaCompiled(const ACell: PXLSCellItem): boolean;
var
  P: TMMUPtr;
  Sz: integer;
begin
  P := GetFormulaFmla(ACell.Data,ACell.XLate,Sz);
  Result := PByte(P)^ = CELLMMU_FMLA_IS_PTGS;
end;

function TXLSCellMMU.ItemSize(const AXLate: integer; const APtr: TMMUPtr): integer;
begin
  case AXLate and XVECT_XLATE_MASK of
    CELLXL_BLANK       : Result := 0;
    CELLXL_NUM_8       : Result := 1;
    CELLXL_NUM_16      : Result := 2;
    CELLXL_NUM_32      : Result := 4;
    CELLXL_NUM_2DEC    : Result := 4;
    CELLXL_NUM_4DEC    : Result := 4;
    CELLXL_NUM_SINGLE  : Result := 8;
    CELLXL_NUM_DOUBLE  : Result := 8;
    CELLXL_STR_16      : Result := 2;
    CELLXL_STR_32      : Result := 4;
    CELLXL_BOOLEAN     : Result := 1;
    CELLXL_ERROR       : Result := 1;

    CELLXL_FMLA_FLOAT  ,
    CELLXL_FMLA_STR    ,
    CELLXL_FMLA_BOOLEAN,           // Formula size word + formula size.
    CELLXL_FMLA_ERROR  : Result := SizeOf(word) + PWord(APtr)^;
    else raise XLSRWException.CreateFmt('Invalid XLate value "%d"',[AXLate]);
  end;
  if (AXLate and CELLXL_FLAG_STYLE) <> 0 then
    Inc(Result,CELLMMU_STYLE_SZ);
end;

function TXLSCellMMU.IterateNext: boolean;
begin
  Result := IterateNextCol;
  while not Result and IterateNextRow do
    Result := IterateNextCol;
end;

function TXLSCellMMU.IterateNextCol: boolean;
begin
  // Changed 130101
  Result := FIterate.Row >= 0;
  if Result then begin
  // End changed 130101
    Inc(FIterate.Col);
    FIterate.Cell.Data := FBlockManager.GetNextItem(FIterate.Col,FIterate.Cell.XLate);
    Result := FIterate.Cell.Data <> Nil;
  // Changed 130101
  end;
  // End changed 130101
end;

function TXLSCellMMU.IterateNextRow: boolean;
var
  i: integer;
  R: integer;
  P: TMMUPtr;
begin
  Result := FBlocks <> Nil;
  if Result then begin
    while FIterate.Row < FIterate.LastRow do begin
      Inc(FIterate.Row);
      i := FIterate.Row shr XBLOCK_VECTOR_COUNT_BITS;
      Result := FBlocks[i].Memory <> Nil;
      if Result then begin
        FBlockManager.Block := @FBlocks[i];
        R := FIterate.Row and XBLOCK_VECTOR_MASK;
        P := FBlockManager.GetNextVector(R);
        FIterate.Row := (i shl XBLOCK_VECTOR_COUNT_BITS) + R;
        Result := P <> Nil;

        if Result then begin
          Inc(P,SizeOf(TXVectHeader));
          FIterate.RowHeader := PXLSMMURowHeader(P);
          BeginIterateCol;
          Exit;
        end
        else
          FIterate.Row := XBLOCK_VECTOR_COUNT * (i + 1) - 1;
      end
      else
        FIterate.Row := (i + 1) * XBLOCK_VECTOR_COUNT - 1;
    end;
    Result := False;
  end;
// Changed 140121
  FIterate.Row := -1;
end;

function TXLSCellMMU.IterCell: PXLSCellItem;
begin
  Result := @FIterate.Cell;
end;

function TXLSCellMMU.IterCellCol: integer;
begin
  Result := FIterate.Col;
end;

function TXLSCellMMU.IterCellRow: integer;
begin
  Result := FIterate.Row;
end;

function TXLSCellMMU.IterCellType: TXLSCellType;
begin
  Result := XLateToCellType(FIterate.Cell.XLate);
end;

function TXLSCellMMU.IterFormulaType: TXLSCellFormulaType;
begin
  Result := PXLSMMUFormulaHeader(FIterate.Cell.Data).FmlaType;
end;

function TXLSCellMMU.IterGetBoolean: boolean;
var
  P: TMMUPtr;
begin
  if (FIterate.Cell.XLate and CELLXL_FLAG_STYLE) <> 0 then begin
    P := FIterate.Cell.Data;
    Inc(P,CELLMMU_STYLE_SZ);
    Result := PByte(P)^ <> 0;
  end
  else
    Result := PByte(FIterate.Cell.Data)^ <> 0;
end;

function TXLSCellMMU.IterGetError: TXc12CellError;
var
  P: TMMUPtr;
begin
  if (FIterate.Cell.XLate and CELLXL_FLAG_STYLE) <> 0 then begin
    P := FIterate.Cell.Data;
    Inc(P,CELLMMU_STYLE_SZ);
    Result := TXc12CellError(PByte(P)^);
  end
  else
    Result := TXc12CellError(PByte(FIterate.Cell.Data)^);
end;

function TXLSCellMMU.IterGetFloat: double;
var
  P: TMMUPtr;
begin
  P := FIterate.Cell.Data;
  if (FIterate.Cell.XLate and CELLXL_FLAG_STYLE) <> 0 then
    Inc(P,CELLMMU_STYLE_SZ);
  case FIterate.Cell.XLate and XVECT_XLATE_MASK of
    CELLXL_NUM_8       : Result := PShortInt(P)^;
    CELLXL_NUM_16      : Result := PSmallint(P)^;
    CELLXL_NUM_32      : Result := PInteger(P)^;
    CELLXL_NUM_2DEC    : Result := PInteger(P)^ / 100;
    CELLXL_NUM_4DEC    : Result := PInteger(P)^ / 10000;
    CELLXL_NUM_SINGLE  : Result := PSingle(P)^;
    CELLXL_NUM_DOUBLE  : Result := PDouble(P)^;
    else                  raise XLSRWException.Create('Invalid float XLate');
  end;
end;

procedure TXLSCellMMU.IterGetFormula;
begin
  RetriveFormula(@FIterate.Cell);
end;

function TXLSCellMMU.IterGetFormulaPtgs(var APtgs: PXLSPtgs): integer;
begin
  APtgs := PXLSPtgs(GetFormulaPtgs(FIterate.Cell.Data,FIterate.Cell.XLate,Result));
end;

function TXLSCellMMU.IterGetStringIndex: integer;
var
  P: TMMUPtr;
begin
  P := FIterate.Cell.Data;
  if (FIterate.Cell.XLate and CELLXL_FLAG_STYLE) <> 0 then
    Inc(P,CELLMMU_STYLE_SZ);
  case FIterate.Cell.XLate and XVECT_XLATE_MASK of
    CELLXL_STR_16: Result := PWord(P)^;
    CELLXL_STR_32: Result := PInteger(P)^;
    else            raise XLSRWException.Create('Invalid string XLate');
  end;
end;

function TXLSCellMMU.IterGetStyleIndex: integer;
begin
  Result := GetStyle(@FIterate.Cell);
end;

function TXLSCellMMU.IterRefStr: AxUCString;
begin
  Result := ColRowToRefStr(FIterate.Col,FIterate.Row);
end;

function TXLSCellMMU.IterRow: PXLSMMURowHeader;
begin
  Result := FIterate.RowHeader;
end;

function TXLSCellMMU.GetLastRow: integer;
var
  i: integer;
  B: TXLSMMUBlock;
begin
  if Length(FBlocks) > 0 then begin
    B := FBlocks[High(FBlocks)];
    Result := 0;
    for i := XBLOCK_VECTOR_COUNT - 1 downto 0 do begin
      if B.VectorsOffs[i] <> Integer(XBLOCK_VECTOR_UNUSED) then
        Break;
      Inc(Result);
    end;
    Result := Length(FBlocks) * XBLOCK_VECTOR_COUNT - Result - 1;
  end
  else
    Result := -1;
end;

function TXLSCellMMU.GetRowStyle(const ARow: integer): integer;
var
  RH: PXLSMMURowHeader;
begin
  RH := FindRow(ARow);
  if RH <> Nil then
    Result := RH.Style
  else
    Result := XLS_STYLE_DEFAULT_XF;
end;

function TXLSCellMMU.GetString(const ACol, ARow: integer; out AValue: AxUCString): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := FindCell(ACol,ARow,Cell);
  if Result then begin
    case Cell.XLate and XVECT_XLATE_MASK of
      CELLXL_STR_16: AValue := FSST.ItemText[PWord(GetValue(@Cell))^];
      CELLXL_STR_32: AValue := FSST.ItemText[PLongword(GetValue(@Cell))^];
      else begin
        AValue := '';
      end;
    end;
  end
  else
    AValue := '';
end;

function TXLSCellMMU.GetString(const ACell: PXLSCellItem): AxUCString;
begin
  case ACell.XLate and XVECT_XLATE_MASK of
    CELLXL_STR_16: Result := FSST.ItemText[PWord(GetValue(ACell))^];
    CELLXL_STR_32: Result := FSST.ItemText[PLongword(GetValue(ACell))^];
    else           Result := '';
  end;
end;

function TXLSCellMMU.GetString(const ACell: PXLSCellItem; out AFontRuns: PXc12FontRunArray; out AFontRunsCount: integer): AxUCString;
var
  i: integer;
  pS: PXLSString;
begin
  case ACell.XLate and XVECT_XLATE_MASK of
    CELLXL_STR_16: i := PWord(GetValue(ACell))^;
    CELLXL_STR_32: i := PLongword(GetValue(ACell))^;
    else begin
      Result := '';
      AFontRuns := Nil;
      Exit;
    end;
  end;

  pS := FSST.Items[i];
  Result := FSST.GetText(pS);
  AFontRunsCount := FSST.GetFontRunsCount(pS);
  if AFontRunsCount > 0 then
    AFontRuns := FSST.GetFontRuns(pS);
end;

function TXLSCellMMU.GetStringSST(const ACell: PXLSCellItem): integer;
begin
  case ACell.XLate and XVECT_XLATE_MASK of
    CELLXL_STR_16: Result := PWord(GetValue(ACell))^;
    CELLXL_STR_32: Result := PLongword(GetValue(ACell))^;
    else           Result := -1;
  end;
end;

function TXLSCellMMU.GetStyle(const ACol, ARow: integer): integer;
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Result := GetStyle(@Cell)
  else
    Result := XLS_STYLE_DEFAULT_XF;
end;

function TXLSCellMMU.GetValue(const ACell: PXLSCellItem): TMMUPtr;
begin
  Result := ACell.Data;

  if (ACell.XLate and XVECT_XLATE_EX) <> 0 then
    Inc(Result,SZ_HDR_EX + SizeOf(byte));

  if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then
    Inc(Result,CELLMMU_STYLE_SZ);
end;

procedure TXLSCellMMU.LoadDebug(const AFilename: string);
begin
  FDebugLoad := True;
  FDebugData.Load(AFilename);
end;

procedure TXLSCellMMU.MoveCell(const ASrcCol, ASrcRow, ADestCol, ADestRow: integer);
var
  Cell: TXLSCellItem;
begin
  if not ((ASrcCol = ADestCol) and (ASrcRow = ADestRow)) then begin
    Cell := CopyCell(ASrcCol,ASrcRow);
    if Cell.Data <> Nil then begin
      InsertCell(ADestCol,ADestRow,@Cell);
      FreeMem(Cell.Data);
    end;
    DeleteCell(ASrcCol,ASrcRow);
  end;
end;

procedure TXLSCellMMU.RetriveFormula(const ACell: PXLSCellItem);
var
  i: integer;
  P: TMMUPtr;
  Sz: integer;
  S: AxUCString;
  Flags: byte;
begin
  FFormulaHelper.Clear;

  P := ACell.Data;
  Inc(P,SizeOf(word));

  FFormulaHelper.CellType := XLateToCellType(ACell.XLate);

  Flags := PByte(P)^;
  FFormulaHelper.FormulaType := TXLSCellFormulaType(CELLMMU_FMLA_FLAG_MASK and Flags);
  Inc(P,SizeOf(byte));

  if (ACell.XLate and CELLXL_FLAG_STYLE) <> 0 then begin
    FFormulaHelper.Style := PWord(P)^;
    Inc(P,SizeOf(word));
  end
  else
    FFormulaHelper.Style := XLS_STYLE_DEFAULT_XF;

  case FFormulaHelper.CellType of
    xctFloatFormula  : begin
      FFormulaHelper.AsFloat := PDouble(P)^;
      Inc(P,SizeOf(double));
    end;
    xctStringFormula : begin
      SetLength(S,PWord(P)^);
      Inc(P,SizeOf(word));
      Move(P^,Pointer(S)^,Length(S) * 2);
      Inc(P,Length(S) * 2);
      FFormulaHelper.AsString := S;
    end;
    xctBooleanFormula: begin
      FFormulaHelper.AsBoolean := PByte(P)^ <> 0;
      Inc(P,SizeOf(byte));
    end;
    xctErrorFormula  : begin
      FFormulaHelper.AsError := TXc12CellError(PByte(P)^);
      Inc(P,SizeOf(byte));
    end;
  end;

  Sz := PWord(P)^;
  Inc(P,SizeOf(word));
  case FFormulaHelper.FormulaType of
    xcftNormal,
    xcftArray         : begin
      case PByte(P)^ of
        CELLMMU_FMLA_IS_PTGS: begin
          Inc(P,SizeOf(byte));
          FFormulaHelper.AllocPtgs(PXLSPtgs(P),Sz);
          Inc(P,FFormulaHelper.PtgsSize);
        end;
        CELLMMU_FMLA_IS_STR_8: begin
          Inc(P,SizeOf(byte));
          SetLength(S,Sz);
          for i := 1 to Sz do
            S[i] := AxUCChar(PByteArray(P)[i - 1]);
          FFormulaHelper.Formula := S;
          Inc(P,Sz);
        end;
        CELLMMU_FMLA_IS_STR_UC: begin
          Inc(P,SizeOf(byte));
          SetLength(S,Sz);
          System.Move(P^,Pointer(S)^,Sz * 2);
          FFormulaHelper.Formula := S;
          Inc(P,Sz);
        end;
        else raise XLSRWException.Create('Invalid stored formula');
      end;
    end;
    xcftDataTable: begin
      Inc(P,SizeOf(byte));
      FFormulaHelper.TableOptions := PXLSPtgsDataTableFmla(P).Options;
      FFormulaHelper.SetR1(PXLSPtgsDataTableFmla(P).R1Col,PXLSPtgsDataTableFmla(P).R1Row);
      FFormulaHelper.SetR2(PXLSPtgsDataTableFmla(P).R2Col,PXLSPtgsDataTableFmla(P).R2Row);
      Inc(P,SizeOf(TXLSPtgsDataTableFmla));
    end;
    xcftArrayChild: begin
      Inc(P,SizeOf(byte));
      FFormulaHelper.ParentCol := PXLSPtgsArrayChildFmla(P).ParentCol;
      FFormulaHelper.ParentRow := PXLSPtgsArrayChildFmla(P).ParentRow;
      Inc(P,SizeOf(TXLSPtgsArrayChildFmla));
    end;
    xcftArrayChild97: begin
      Inc(P,SizeOf(byte));
      FFormulaHelper.ParentCol := PXLSPtgsArrayChildFmla97(P).ParentCol;
      FFormulaHelper.ParentRow := PXLSPtgsArrayChildFmla97(P).ParentRow;
      Inc(P,SizeOf(TXLSPtgsArrayChildFmla));
    end;
    xcftDataTableChild: begin
      Inc(P,SizeOf(byte));
      FFormulaHelper.ParentCol := PXLSPtgsArrayChildFmla(P).ParentCol;
      FFormulaHelper.ParentRow := PXLSPtgsArrayChildFmla(P).ParentRow;
      Inc(P,SizeOf(TXLSPtgsDataTableChildFmla));
    end;
  end;

  FFormulaHelper.HasApply := (Flags and CELLMMU_FMLA_FLAG_ARRAY) <> 0;
  if FFormulaHelper.HasApply then begin
    FFormulaHelper.TargetRef := PXLSFormulaArea(P)^;
    Inc(P,SizeOf(TXLSFormulaArea));
  end;

  FFormulaHelper.IsTABLE := (Flags and CELLMMU_FMLA_FLAG_TABLE) <> 0;
  if FFormulaHelper.IsTABLE then begin
    FFormulaHelper.TableOptions := PByte(P)^;
    Inc(P,SizeOf(Byte));
    FFormulaHelper.R1 := PXLSFormulaRef(P)^;
    Inc(P,SizeOf(TXLSFormulaRef));
    FFormulaHelper.R2 := PXLSFormulaRef(P)^;
  end;
end;

procedure TXLSCellMMU.AddFormulaVal(const ACol, ARow: integer; const AValue: double);
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    AddFormulaVal(@Cell,ACol,ARow,AValue);
end;

procedure TXLSCellMMU.AddFormulaVal(const ACol,ARow: integer; const AValue: boolean);
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    AddFormulaVal(@Cell,ACol,ARow,AValue);
end;

procedure TXLSCellMMU.AddFormulaVal(const ACol, ARow: integer; const AValue: AxUCString);
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    AddFormulaVal(@Cell,ACol,ARow,AValue);
end;

procedure TXLSCellMMU.AddFormulaVal(const ACol, ARow: integer; const AValue: TXc12CellError);
var
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    AddFormulaVal(@Cell,ACol,ARow,AValue);
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: boolean);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;

  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Ptgs := APtgs;
  FFormulaHelper.PtgsSize := APtgsSize;

  FFormulaHelper.AsBoolean := AValue;
  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: AxUCString);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;

  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Ptgs := APtgs;
  FFormulaHelper.PtgsSize := APtgsSize;

  FFormulaHelper.AsString := AValue;
  StoreFormula;
end;

procedure TXLSCellMMU.AddFormula(const ACol, ARow, AStyle: integer; const APtgs: PXLSPtgs; const APtgsSize: integer; const AValue: TXc12CellError);
var
  Style: integer;
  Cell: TXLSCellItem;
begin
  if FindCell(ACol,ARow,Cell) then
    Style := GetStyle(@Cell)
  else
    Style := AStyle;

  FFormulaHelper.Clear;

  FFormulaHelper.Col := ACol;
  FFormulaHelper.Row := ARow;
  FFormulaHelper.Style := Style;
  FFormulaHelper.Ptgs := APtgs;
  FFormulaHelper.PtgsSize := APtgsSize;

  FFormulaHelper.AsError := AValue;
  StoreFormula;
end;

{ TXLSCellMMUDebugData }

procedure TXLSCellMMUDebugData.AddBlank(const ACol, ARow, AStyle: integer);
var
  S: string;
begin
  S := Format('Blank:%d:%d:%d',[ACol,ARow,AStyle]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddBoolean(const ACol, ARow, AStyle: integer; const AValue: boolean);
var
  S: string;
begin
  if AValue then
    S := Format('Boolean:%d:%d:%d:True',[ACol,ARow,AStyle])
  else
    S := Format('Boolean:%d:%d:%d:False',[ACol,ARow,AStyle]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddError(const ACol, ARow, AStyle: integer; const AValue: TXc12CellError);
var
  S: string;
begin
  S := Format('Error:%d:%d:%d:%d',[ACol,ARow,AStyle,Integer(AValue)]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddFloat(const ACol, ARow, AStyle: integer; const AValue: double);
var
  S: string;
begin
  S := Format('Float:%d:%d:%d:%f',[ACol,ARow,AStyle,AValue]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddFormula(const ACol, ARow, AStyle: integer; const AFormula: AxUCString);
var
  S: string;
begin
  // Ptgs not saved.
  S := Format('Formula:%d:%d:%d:%s',[ACol,ARow,AStyle,AFormula]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddFormulaFloat(const ACol, ARow, AStyle: integer; const AValue: double; const APtgs: PXLSPtgs; APtgsSize: integer);
var
  S: string;
begin
  // Ptgs not saved.
  S := Format('FormulaFloat:%d:%d:%d:%f:%d',[ACol,ARow,AStyle,AValue,APtgsSize]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddRow(const ARow, AStyle: integer);
var
  S: string;
begin
  S := Format('Row:%d:%d',[ARow,AStyle]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.AddString(const ACol, ARow, AStyle, AValue: integer);
var
  S: string;
begin
  S := Format('String:%d:%d:%d:%d',[ACol,ARow,AStyle,AValue]);
  FList.Add(S);
end;

procedure TXLSCellMMUDebugData.Clear;
begin
  FList.Clear;
end;

constructor TXLSCellMMUDebugData.Create(AOwner: TXLSCellMMU);
begin
  FOwner := AOwner;
  FList := TStringList.Create;
end;

procedure TXLSCellMMUDebugData.Delete(const ACol, ARow: integer);
var
  S: string;
begin
  S := Format('Delete:%d:%d',[ACol,ARow]);
  FList.Add(S);
end;

destructor TXLSCellMMUDebugData.Destroy;
begin
  FList.Free;
  inherited;
end;

procedure TXLSCellMMUDebugData.Load(AFilename: string);
var
  i,j: integer;
  S,S2: AxUCString;
  Col,Col2,Row: integer;
  Format: integer;
  ValFloat: double;
  ValStr: AxUCString;
  PtgsSz: integer;
begin
  FOwner.DebugSave := False;
  FList.LoadFromFile(AFilename);

  for i := 0 to FList.Count - 1 do begin
    S := Trim(FList[i]);
    if (S = '') or (Copy(S,1,1) = ';') then
      Continue;

    S2 := Uppercase(SplitAtChar(':',S));

    if S2 = 'DELETE' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      FOwner.DeleteCell(Col,Row);
    end
    else if S2 = 'BLANK' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      FOwner.StoreBlank(Col,Row,Format);
    end
    else if S2 = 'MULBLANK' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Col2 := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      for j := Col to Col2 do
        FOwner.StoreBlank(j,Row,Format);
    end
    else if S2 = 'BOOLEAN' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      ValStr := SplitAtChar(':',S);
//      FOwner.StoreBoolean(Col,Row,Format,Uppercase(ValStr) = 'TRUE');
      FOwner.UpdateBoolean(Col,Row,Uppercase(ValStr) = 'TRUE',Format);
    end
    else if S2 = 'FLOAT' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      ValFloat := StrToFloat(SplitAtChar(':',S));
//      FOwner.StoreFloat(Col,Row,Format,ValFloat);
      FOwner.UpdateFloat(Col,Row,ValFloat,Format);
    end
    else if S2 = 'FORMULA' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      ValStr := SplitAtChar(':',S);
      FOwner.AddFormula(Col,Row,Format,ValStr,0);
    end
    else if S2 = 'FORMULAFLOAT' then begin
      Col := StrToInt(SplitAtChar(':',S));
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      ValFloat := StrToFloat(SplitAtChar(':',S));
      PtgsSz := StrToInt(SplitAtChar(':',S));
      FOwner.DEBUG_AddFormulaFloat(Col,Row,Format,ValFloat,Nil,PtgsSz);
    end
    else if S2 = 'ROW' then begin
      Row := StrToInt(SplitAtChar(':',S));
      Format := StrToInt(SplitAtChar(':',S));
      FOwner.AddRow(Row,Format);
    end
    else
      raise XLSRWException.Create('Invalid debug data: ' + S2);
  end;
end;

procedure TXLSCellMMUDebugData.Save(AFilename: string);
begin
  FList.SaveToFile(AFilename);
end;

{ TXLSFormulaHelper }

procedure TXLSFormulaHelper.AllocPtgs(const APtgs: PXLSPtgs; const APtgsSize: integer);
begin
  if FOwnsPtgs then
    FreeMem(FPtgs);
  FOwnsPtgs := True;
  FPtgsSize := APtgsSize;
  GetMem(FPtgs,FPtgsSize);
  System.Move(APtgs^,FPtgs^,FPtgsSize);
end;

procedure TXLSFormulaHelper.AllocPtgs97(const APtgs: PXLSPtgs; const APtgsSize: integer);
var
  P: PXLSPtgs;
begin
  if FOwnsPtgs then
    FreeMem(FPtgs);
  FOwnsPtgs := True;
  FPtgsSize := APtgsSize + SizeOf(TXLSPtgs);
  GetMem(FPtgs,FPtgsSize + SizeOf(TXLSPtgs));
  P := FPtgs;
  P.Id := xptg_EXCEL_97;
  P := PXLSPtgs(NativeInt(P) + SizeOf(TXLSPtgs));
  System.Move(APtgs^,P^,FPtgsSize - SizeOf(TXLSPtgs));
end;

procedure TXLSFormulaHelper.AllocPtgsArrayConsts97(const APtgs: PXLSPtgs; const APtgsSize, AArrayConstsSize: integer);
var
  P: PXLSPtgs;
begin
  if FOwnsPtgs then
    FreeMem(FPtgs);
  FOwnsPtgs := True;
  FPtgsSize := APtgsSize + SizeOf(TXLSPtgs) + SizeOf(word) + AArrayConstsSize;
  GetMem(FPtgs,FPtgsSize);
  P := FPtgs;
  P.Id := xptg_ARRAYCONSTS_97;
  P := PXLSPtgs(NativeInt(P) + SizeOf(TXLSPtgs));
  PWord(P)^ := AArrayConstsSize;
  P := PXLSPtgs(NativeInt(P) + SizeOf(word));
  System.Move(APtgs^,P^,FPtgsSize - SizeOf(TXLSPtgs) - SizeOf(word));
end;

function TXLSFormulaHelper.AllocSize: integer;
begin
  Result := {SizeOf(word) + }CELLMMU_FMLA_FLAG_SZ;

  if FStyle <> XLS_STYLE_DEFAULT_XF then
    Inc(Result,CELLMMU_STYLE_SZ);

  case FCellType of
    xctFloatFormula  : Inc(Result,SizeOf(double));
    xctStringFormula : Inc(Result,SizeOf(word) + Length(FStrResult) * 2);
    xctBooleanFormula: Inc(Result,SizeOf(byte));
    xctErrorFormula  : Inc(Result,SizeOf(byte));
  end;

  Inc(Result,SizeOf(word));
  if FFormulaType = xcftArrayChild then
    Inc(Result,SizeOf(byte) + SizeOf(TXLSPtgsArrayChildFmla))
  else if FFormulaType = xcftArrayChild97 then
    Inc(Result,SizeOf(byte) + SizeOf(TXLSPtgsArrayChildFmla97))
  else if FFormulaType = xcftDataTable then
    Inc(Result,SizeOf(byte) + SizeOf(TXLSPtgsDataTableFmla))
  else if FFormulaType = xcftDataTableChild then
    Inc(Result,SizeOf(byte) + SizeOf(TXLSPtgsDataTableChildFmla))
  else if FPtgsSize > 0 then
    Inc(Result,SizeOf(byte) + FPtgsSize)
  else begin
    FFmlaIs8Bit := UnicodeIs8Bit(FFormula);
    if FFmlaIs8Bit then
      Inc(Result,SizeOf(byte) + Length(FFormula))
    else
      Inc(Result,SizeOf(byte) + Length(FFormula) * 2);
  end;
  if FHasApply then
    Inc(Result,SizeOf(TXLSFormulaArea));
  if FIsTABLE then
    Inc(Result,SizeOf(byte) + SizeOf(TXLSFormulaRef) * 2);
end;

procedure TXLSFormulaHelper.Clear;
begin
  if FOwnsPtgs then begin
    FreeMem(FPtgs);
    FOwnsPtgs := False;
  end;

  FCol        := -1;
  FRow        := -1;
  FStyle      := CELLXL_FLAG_STYLE;
  FFormula    := '';
  FFmlaIs8Bit := False;
  FPtgs       := Nil;
  FPtgsSize   := -1;
  FCellType   := xctNone;
  FFormulaType:= xcftNormal;
  FHasApply   := False;
  FResult     := 0;
  FStrResult  := '';
  FR1.Col     := $FFFF;
  FR2.Col     := $FFFF;
  FIsTABLE    := False;
  FOptions    := 0;
end;

destructor TXLSFormulaHelper.Destroy;
begin
  if FOwnsPtgs then
    FreeMem(FPtgs);
  inherited;
end;

function TXLSFormulaHelper.GetStrTargetRef: AxUCString;
begin
  Result := AreaToRefStr(FTargetRef.Col1,FTargetRef.Row1,FTargetRef.Col2,FTargetRef.Row2);
end;

function TXLSFormulaHelper.GetAsBoolean: boolean;
begin
  Result := PBoolean(@FResult)^;
end;

function TXLSFormulaHelper.GetAsError: TXc12CellError;
begin
  Result := TXc12CellError(FResult);
end;

function TXLSFormulaHelper.GetAsErrorStr: AxUCString;
begin
  Result := Xc12CellErrorNames[TXc12CellError(FResult)];
end;

function TXLSFormulaHelper.GetAsFloat: double;
begin
  Result := PDouble(@FResult)^;
end;

function TXLSFormulaHelper.GetAsString: AxUCString;
begin
  Result := FStrResult;
end;

function TXLSFormulaHelper.GetStrR1: AxUCString;
begin
  if FR1.Col <> $FFFF then
    Result := ColRowToRefStr(FR1.Col,FR1.Row)
  else
    Result := '';
end;

function TXLSFormulaHelper.GetStrR2: AxUCString;
begin
  if FR2.Col <> $FFFF then
    Result := ColRowToRefStr(FR2.Col,FR2.Row)
  else
    Result := '';
end;

function TXLSFormulaHelper.GetRef: AxUCString;
begin
  Result := ColRowToRefStr(FCol,FRow);
end;

function TXLSFormulaHelper.IsCompiled: boolean;
begin
  Result := FPtgsSize > 0;
end;

procedure TXLSFormulaHelper.SetStrTragetRef(const Value: AxUCString);
var
  C1,R1,C2,R2: integer;
begin
  FHasApply := Value <> '';
  if FHasApply then begin
    AreaStrToColRow(Value,C1,R1,C2,R2);
    FTargetRef.Col1 := C1;
    FTargetRef.Row1 := R1;
    FTargetRef.Col2 := C2;
    FTargetRef.Row2 := R2;
  end;
end;

procedure TXLSFormulaHelper.SetTargetRef(const ACol1, ARow1, ACol2, ARow2: integer);
begin
  FTargetRef.Col1 := ACol1;
  FTargetRef.Row1 := ARow1;
  FTargetRef.Col2 := ACol2;
  FTargetRef.Row2 := ARow2;
end;

procedure TXLSFormulaHelper.SetAsBoolean(const Value: boolean);
begin
  FResult := PInt64(@Value)^;
  FCellType := xctBooleanFormula;
end;

procedure TXLSFormulaHelper.SetAsError(const Value: TXc12CellError);
begin
  FResult := Int64(Value);
  if TXc12CellError(FResult) = errUnknown then
    raise XLSRWException.Create('Invalid error value');
  FCellType := xctErrorFormula;
end;

procedure TXLSFormulaHelper.SetAsErrorStr(const Value: AxUCString);
begin
  FResult := Int64(ErrorTextToCellError(Value));
  if TXc12CellError(FResult) = errUnknown then
    raise XLSRWException.Create('Invalid error value');
  FCellType := xctErrorFormula;
end;

procedure TXLSFormulaHelper.SetAsFloat(const Value: double);
begin
  FResult := PInt64(@Value)^;
  FCellType := xctFloatFormula;
end;

procedure TXLSFormulaHelper.SetAsString(const Value: AxUCString);
begin
  FStrResult := Value;
  FCellType := xctStringFormula;
end;

procedure TXLSFormulaHelper.SetPtgs(const Value: PXLSPtgs);
begin
  if FOwnsPtgs then begin
    FreeMem(FPtgs);
    FOwnsPtgs := False;
  end;
  FPtgs := Value;
end;

procedure TXLSFormulaHelper.SetStrR1(const Value: AxUCString);
var
  C,R: integer;
begin
  RefStrToColRow(Value,C,R);
  FR1.Col := C;
  FR1.Row := R;
  FIsTABLE := True;
end;

procedure TXLSFormulaHelper.SetStrR2(const Value: AxUCString);
var
  C,R: integer;
begin
  RefStrToColRow(Value,C,R);
  FR2.Col := C;
  FR2.Row := R;
end;

procedure TXLSFormulaHelper.SetRef(const Value: AxUCString);
begin
  RefStrToColRow(Value,FCol,FRow);
end;

procedure TXLSFormulaHelper.SetR1(const ACol,ARow: integer);
begin
  FR1.Col := ACol;
  FR1.Row := ARow;
end;

procedure TXLSFormulaHelper.SetR2(const ACol,ARow: integer);
begin
  FR2.Col := ACol;
  FR2.Row := ARow;
end;

// initialization
//  G_DebugStream := TFileStream.Create('d:\xtemp\CellFind.bin',fmCreate);
//  G_DoWrite := False;
//
//finalization
//  G_DebugStream.Free;

end.
