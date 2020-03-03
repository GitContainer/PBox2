unit XLSEvaluateFmla5;

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

// CHOOSE CONCATENATE

// Copyright
// PERCENTRANK, MODE, LINEST: from TurboPower SysTools

// http://www.excelfunctions.net/

// TODO
// TRUE() is decoded to (TRUE)
// Formulas with empty args, such as IF(TRUE;1;)
// Leaving a function without popping all args from stack (like when an error is detected) may cause problems.

// COMBIN only works with arguments <= 170, as it uses factorial.
// DoDatabase: Relative col number is also ok, e.g. "4"

// Missing functions
// PROPER, MODE.SNGL
//
// CONFIDENCE.T, NETWORKDAYS.INTL, PERCENTILE.EXC, PERCENTRANK.EXC, QUARTILE.EXC
// WORKDAY.INTL
//
// HYPERLINK is faked, result is correct, but it's no hyperlink.

// Wrong result
// VDB(10000,1000,5,4,5)

// Known issues
// ROMAN only works with mode 0 (Classic form)
// LINEST,LOGEST,GROWTH,TREND only partial implemented
// TTEST only works with type = 2

// Info
// ceil function in C++ don't gives the same result as Ceil in Delphi.
// C++: ceil(55) = 56, Delphi: ceil(55) = 55. Delphi answer seems to be correct.

uses Classes, SysUtils, Contnrs, IniFiles, Math, 
{$ifdef DELPHI_6_OR_LATER}
     Masks, DateUtils, StrUtils,
{$endif}
{$ifdef XLS_BIFF}
     BIFF_Utils5,
{$endif}
     Xc12Utils5, Xc12Manager5, Xc12DataXLinks5, Xc12DataWorkbook5, Xc12DataWorksheet5,
     XLSUtils5, XLSFormulaTypes5, XLSCellMMU5, XLSCellAreas5, XLSMask5,
     XLSEvaluate5, XLSMatrix5, XLSMath5, XLSMathData5, XLSFmlaDebugData5;

type TXLSDateDiffMethod = (xddmUS30_360,xddmActaul,xddm360,xddm365,xddmEU30_360);

type TLinEstData = record
     B0, B1     : double;   {model coefficients}
     seB0, seB1 : double;   {standard error of model coefficients}
     R2         : double;   {coefficient of determination}
     sigma      : double;   {standard error of regression}
     SSr, SSe   : double;   {elements for ANOVA table}
     F0         : double;   {F-statistic to test B1=0}
     df         : integer;  {denominator degrees of freedom for F-statistic}
     end;

type TCriteriaData = record
     CritVal: TXLSVarValue;
     Op: TXLSDbCondOperator;
     end;

type TDynCriteriaDataArray = array of TCriteriaData;

type TFVIterateData = record
     Source : TObject;
     FV     : TXLSFormulaValue;
     Col,Row: integer;
     Index  : integer;
     Result : TXLSFormulaValue;
     Done   : boolean;
     end;

type TXLSUserFuncData = class(TObject)
private
     function  GetAsFloat(Index: integer): double;
     function  GetCols(Index: integer): integer;
     function  GetRefType(Index: integer): TXLSFormulaRefType;
     function  GetRows(Index: integer): integer;
     function  GetAreaAsFloat(Index, ACol, ARow: integer): double;
     procedure SetResultAsFloat(const Value: double);
     function  GetAreaAsBoolean(Index, ACol, ARow: integer): boolean;
     function  GetAreaAsError(Index, ACol, ARow: integer): TXc12CellError;
     function  GetAreaAsString(Index, ACol, ARow: integer): AxUCString;
     function  GetAsBoolean(Index: integer): boolean;
     function  GetAsError(Index: integer): TXc12CellError;
     function  GetAsString(Index: integer): AxUCString;
     procedure SetResultAsBoolean(const Value: boolean);
     procedure SetResultAsError(const Value: TXc12CellError);
     procedure SetResultAsString(const Value: AxUCString);
     function  GetArgType(Index: integer): TXLSFormulaValueType;
protected
     FArgs: array of TXLSFormulaValue;
     FResult: TXLSFormulaValue;
public
     //* Number of arguments to the function.
     function  ArgCount: integer;

     //* The dimension of a cell reference argument.
     //* AIndex is the argument index.
     //* ACol1 = First column.
     //* ARow1 = First row.
     //* ACol2 = Last column. Same as ACol1 if the argument is a single cell.
     //* ARow2 = Last row. Same as ARow1 if the argument is a single cell.
     procedure Dimensions(const AIndex: integer; out ACol1,ARow1,ACol2,ARow2: integer);

     //* Type of cell reference
     //* AIndex is the argument index.
     //* TXLSFormulaRefType can be:
     //* xfrtNone = Argument is not a reference. Use the AsXXX properties to read the value.
     //* xfrtRef = Argument is a single cell reference. Use the AsXXX properties to read the value.
     //*           The Dimension method can be used to get the cell position.
     //* xfrtArea = Argument is an area cell reference. Use the AreaAsXXX properties to read the value.
     //*           The Dimension method can be used to get the area.
     //* xfrtXArea = Argument is an external reference (worksheet).
     property RefType[Index: integer]: TXLSFormulaRefType read GetRefType;

     //* Type of argument (cell tyle). This property has only meaning if the
     //* argument is a single cell.
     //* AIndex is the argument index.
     //* TXLSFormulaValueType can be:
     //* xfvtFloat = Cell is a floating point value cell.
     //* xfvtBoolean = Cell is a boolean value cell.
     //* xfvtError = Cell is an error value cell.
     //* xfvtString = Cell is a string value cell.
     property ArgType[Index: integer]: TXLSFormulaValueType read GetArgType;

     //* Number of columns in the argument. Most usefull if the argument is an array constant.
     property Cols[Index: integer]: integer read GetCols;
     //* Number of rows in the argument. Most usefull if the argument is an array constant.
     property Rows[Index: integer]: integer read GetRows;

     //* Returns the argument as a boolean value.
     //* AIndex is the argument index.
     //* If the argument not is a value or a single cell, there will be an exception.
     //* Is the value not a boolean cell or value, False is returned.
     property AsBoolean[Index: integer]: boolean read GetAsBoolean;

     //* Returns the argument as a floating point value.
     //* AIndex is the argument index.
     //* If the argument not is a value or a single cell, there will be an exception.
     //* Is the value not a float cell or value, zero is returned.
     property AsFloat[Index: integer]: double read GetAsFloat;

     //* Returns the argument as a string value.
     //* AIndex is the argument index.
     //* If the argument not is a value or a single cell, there will be an exception.
     //* Is the value not a string cell or value, a empty string is returned.
     property AsString[Index: integer]: AxUCString read GetAsString;

     //* Returns the argument as an error value.
     //* AIndex is the argument index.
     //* If the argument not is a value or a single cell, there will be an exception.
     //* Is the value not an error cell, errNA is returned.
     property AsError[Index: integer]: TXc12CellError read GetAsError;

     //* Returns the value of an area or array constant argument.
     //* AIndex is the argument index.
     //* ACol is the column.
     //* ARow is the row.
     //* If the argument not is an area or an array constant, there will be an exception.
     //* Is the value not a boolean value, False is returned.
     property AreaAsBoolean[Index,ACol,ARow: integer]: boolean read GetAreaAsBoolean;

     //* Returns the value of an area or array constant argument.
     //* AIndex is the argument index.
     //* ACol is the column.
     //* ARow is the row.
     //* If the argument not is an area or an array constant, there will be an exception.
     //* Is the value not a floating point value, zero is returned.
     property AreaAsFloat[Index,ACol,ARow: integer]: double read GetAreaAsFloat;

     //* Returns the value of an area or array constant argument.
     //* AIndex is the argument index.
     //* ACol is the column.
     //* ARow is the row.
     //* If the argument not is an area or an array constant, there will be an exception.
     //* Is the value not a string value, A empty string is returned.
     property AreaAsString[Index,ACol,ARow: integer]: AxUCString read GetAreaAsString;

     //* Returns the value of an area or array constant argument.
     //* AIndex is the argument index.
     //* ACol is the column.
     //* ARow is the row.
     //* If the argument not is an area or an array constant, there will be an exception.
     //* Is the value not an error value, errNA is returned.
     property AreaAsError[Index,ACol,ARow: integer]: TXc12CellError read GetAreaAsError;

     //* Sets the result as a boolean value.
     property ResultAsBoolean: boolean write SetResultAsBoolean;

     //* Sets the result as a floating point value.
     property ResultAsFloat: double write SetResultAsFloat;

     //* Sets the result as a string value.
     property ResultAsString: AxUCString write SetResultAsString;

     //* Sets the result as an error value.
     property ResultAsError: TXc12CellError write SetResultAsError;
     end;

type TUserFunctionEvent = procedure(const AName: AxUCString; AData: TXLSUserFuncData) of object;

type TValueStack = class;

     TValStackIterator = class(TObject)
protected
     FOwner       : TValueStack;
     FDimension   : PXLSCellArea;

     FIgnoreHidden: boolean;
     FIgnoreError : boolean;
     FSource      : TObject;
     FFV          : TXLSFormulaValue;
     FCol,FRow    : integer;
     FIndex       : integer;
     FResult      : TXLSFormulaValue;
     FDone        : boolean;
     FHaltOnError : boolean;

     FLinked      : TValStackIterator;

     function  DoNext: boolean;
     function  ResIsFloat: boolean;
     function  ResIsFloatA: boolean;
     function  ResIsError(out AError: TXc12CellError): boolean;
     function  ResIsNotEmpty: boolean;
public
     constructor Create(AOwner: TValueStack);

     destructor Destroy; override;

     procedure BeginIterate(AValue: TXLSFormulaValue; AHaltOnError: boolean); overload;
     procedure BeginIterate(AValue1,AValue2: TXLSFormulaValue; AHaltOnError: boolean); overload;
     function  Next: boolean;
     function  NextNotEmpty: boolean;
     function  NextFloat: boolean;
     // Used by "A" functions, AVERAGEA, MINA, MAXA, etc.
     function  NextFloatA: boolean;
     function  AsFloatA: double;

     procedure AddLinked;
     procedure Clear;

     property Result: TXLSFormulaValue read FResult;
     property Linked: TValStackIterator read FLinked;
     property IgnoreHidden: boolean read FIgnoreHidden write FIgnoreHidden;
     property IgnoreError: boolean read FIgnoreError write FIgnoreError;

     property Col: integer read FCol;
     property Row: integer read FRow;
     end;

     TValueStack = class(TObject)
{$ifdef DELPHI_2006_OR_LATER} strict {$endif} private
     FStackPtr  : integer;
     FResultPtr : integer;
protected
     FManager   : TXc12Manager;
     FIsArrayRes: boolean;
     FCol,FRow  : integer;
     FCells     : TXLSCellMMU;
     FSheetIndex: integer;

     FGarbage   : TObjectList;

     FStack     : array of TXLSFormulaValue;

     FIterator  : TValStackIterator;

     procedure IncStackPtr; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DecStackPtr; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  HandleError(AError: TXc12CellError): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     function  AreaCellCount(FV: PXLSFormulaValue): integer;
     function  AreasEqualSize(FV1,FV2: PXLSFormulaValue): boolean;
     function  IsVector(AFormulaVal: PXLSFormulaValue): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure MakeVector(AFormulaVal: PXLSFormulaValue; ALeftTop: boolean); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure GetCellValue(Cells: TXLSCellMMU; ACol,ARow: integer; AFormulaVal: PXLSFormulaValue);
     function  GetCellBlank(Cells: TXLSCellMMU; ACol,ARow: integer): boolean;
     function  GetCellType(Cells: TXLSCellMMU; ACol,ARow: integer): TXLSFormulaValueType ;
     function  GetCellValueFloat(Cells: TXLSCellMMU; ACol,ARow: integer; var AResult: double; out AEmptyCell: boolean): TXc12CellError; overload;
     function  GetCellValueFloat(Cells: TXLSCellMMU; ACol,ARow: integer; AAcceptText: boolean; var Aresult: double; out AEmptyCell: boolean): TXc12CellError; overload;

     function  GetAsString(AValue: TXLSFormulaValue; out AResult: AxUCString): TXc12CellError;
     function  GetAsString2(AValue: TXLSFormulaValue; out AResult: AxUCString): boolean;
     function  GetAsFloat(AValue: TXLSFormulaValue; out AResult: double; out AError: TXc12CellError): boolean;
     function  GetAsFloatNum(AValue: TXLSFormulaValue; out AResult: double; out AError: TXc12CellError): boolean;
     function  GetAsFloatNumOnly(AValue: TXLSFormulaValue; out AResult: double; out  AError: TXc12CellError): boolean;
     function  GetAsBoolean(AValue: TXLSFormulaValue; out AResult: boolean): TXc12CellError;
     function  GetAsError(AValue: TXLSFormulaValue): TXc12CellError;
     function  GetAsValueType(AValue: TXLSFormulaValue; out AResult: TXLSFormulaValueType): TXc12CellError;

     procedure FillArrayItem(AItem: TXLSArrayItem; AValue:  TXLSFormulaValue);
     procedure FillArrayItemFloat(AItem: TXLSArrayItem; AValue:  TXLSFormulaValue);

     procedure SetError(AError: TXc12CellError);
     function  Intersect(AValue: PXLSFormulaValue; var ACol,ARow: integer): boolean;
     // If intersects, AValue is changed from xfrtArea to xfrtRef
     function  MakeIntersect(AValue: PXLSFormulaValue): boolean;

     function  CoupDaysYear(ADate: TDateTime; ADateMethod: TXLSDateDiffMethod): integer;
     function  CoupDaysBetween(ADateFrom,ADateTo: TDateTime; ADateMethod: TXLSDateDiffMethod): integer;
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     procedure Clear(AIsArrayRes: boolean; ASheetIndex,ACol,ARow: integer);
     // Empty (unknown) value.
     procedure Push; overload;
     procedure Push(const AValue: TXLSFormulaValue); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const AValue: double); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const AValue: boolean); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const AValue: TXc12CellError); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     // if AFVValue is an error value, it has higher priority.
     procedure Push(const AValue: TXc12CellError; const AFVValue: TXLSFormulaValue); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const AValue: AxUCString); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const ACol,ARow: integer); overload;
     procedure Push(const ACol1,ARow1,ACol2,ARow2: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const ASheetIndex,ACol,ARow: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const ASheetIndex,ACol1,ARow1,ACol2,ARow2: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const ASheetIndex1,ASheetIndex2,ACol1,ARow1,ACol2,ARow2: integer); overload;
     procedure Push(const ASheetIndex: integer; const AError: TXc12CellError); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const ACells: TXLSCellMMU; const ACol,ARow: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const ACells: TXLSCellMMU; const ACol1,ARow1,ACol2,ARow2: integer); overload; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Push(const AArray: TXLSArrayItem); overload;
     procedure Push(const AMatrix: TXLSMatrix); overload;
     procedure Push(const ATableSpecial: TXLSTableSpecialSpecifier; const ACol1,ARow1,ACol2,ARow2: integer); overload;
//     procedure Push(ATableSpecial: TXLSTableSpecialSpecifier; ACol1,ARow1,ACol2,ARow2: integer); overload;

     function  Pop: TXLSFormulaValue; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  PopResult(out AValue: TXLSFormulaValue): boolean;
     procedure PopArray(out AArr: TXLSArrayItem; AFloat: boolean); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure PopArrays(out AArr1,AArr2: TXLSArrayItem; AFloat: boolean);

     function  Peek(ALevel: integer): TXLSFormulaValue; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  PeekFloat(ALevel: integer; out AValue: double): boolean;

     procedure SetResult(const ACount: integer);
     function  PopRes: TXLSFormulaValue; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  PopResFloat(out AValue: double): boolean;
     function  PopResInt(out AValue: integer): boolean;
     function  PopResPosInt(out AValue: integer; AOkNum: integer = 0): boolean;
     function  PopResBoolean(out AValue: boolean): boolean;
     function  PopResStr(out AValue: AxUCString; AEmptyOk: boolean = True): boolean;
     function  PopResDateTime(out AValue: double): boolean;
     function  PopResCellsSource(out AValue: TXLSCellsSource): boolean;
     procedure PopResArray(out AArr: TXLSArrayItem; AFloat: boolean); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure PopResArrays(out AArr1,AArr2: TXLSArrayItem; AFloat: boolean);
     function  PeekRes(ALevel: integer): TXLSFormulaValue; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure ClearStack; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  StackSize: integer;

     procedure OpFloat(AOperator: integer);
     procedure OpBoolean(AOperator: integer);
     procedure OpConcat;
     procedure OpArea(AOperator: integer);
     procedure OpUnary(AOperator: integer);
     procedure OpArrayFloat(AOperator: integer);
     procedure OpArrayBoolean(AOperator: integer);
     procedure OpArrayConcat;
     procedure OpArrayUnary(AOperator: integer);


     function  NPV(ARate: double; AValues: TDynDoubleArray): double;
     function  DDB(const Cost, Salvage, Life, Period, Factor: double): double;
     function  VDB(const Cost, Salvage, Life, StartPeriod, EndPeriod, Factor: double; NoSwitch: boolean): double;
     function  DB(Cost, Salvage: double; Life, Period, Month: integer) : double;

     function  PopulateMatrix(AMatrix: TXLSMatrix; AValue: PXLSFormulaValue): boolean;
     procedure WriteMatrix(AMatrix: TXLSMatrix; ACol,ARow: integer);

     procedure SetupIfs(AArgCount: integer; var AData: TDynCriteriaDataArray);

     function  DoCollectValues(var AResult: TDynDoubleArray): boolean;
     function  DoCollectValuesA(var AResult: TDynDoubleArray): boolean;
     function  DoCollectAllValues(AArgCount: integer; var AResult: TDynDoubleArray): boolean;
     function  DoCollectAllValuesA(AArgCount: integer; var AResult: TDynDoubleArray): boolean;

     function  Sum2Arrays(AFV1,AFV2: TXLSFormulaValue; out ASum1,ASum2: double): integer;
     function  Percentile(APercent: double; out AResult: double): boolean;
     function  TDist(X : double; DegreesFreedom : Integer; TwoTails : Boolean) : double;
     function  FDist(x: double; df1, df2: integer): double;
     function  TTest(Arr1,Arr2: TDynDoubleArray; Tails,Type_: integer): double;
     function  TTest2(Arr1,Arr2: TDynDoubleArray; Tails,Type_: integer): double;
     function  ZTest(Arr: TDynDoubleArray; x,sigma: double): double;
     function  LinEst(const KnownY,KnownX: TDynDoubleArray; var LF : TLinEstData; ErrorStats : Boolean): boolean;
     procedure LogEst(const KnownY,KnownX: TDynDoubleArray; var LF : TLinEstData; ErrorStats : Boolean);
     function  ForecastExponential(X : Double; const KnownY,KnownX: TDynDoubleArray) : Double;

     function  SolveFInvEvent(AData: PDoubleArray; AValue: double): double;
     function  SolveChiInvEvent(AData: PDoubleArray; AValue: double): double;

     procedure DoSum(AArgCount: integer);
     procedure DoAverage(AArgCount: integer);
     function  DoCount(AArgCount: integer): integer;
     procedure DoMin(AArgCount: integer);
     procedure DoMax(AArgCount: integer);
     function  DoNPV(AArgCount: integer; out AResult: double): boolean;
     procedure DoStDev(AArgCount: integer);
     procedure DoLookup(AArgCount: integer);
     procedure DoIndex(AArgCount: integer);
     procedure DoValue;
     procedure DoAND(AArgCount: integer);
     procedure DoOR(AArgCount: integer);
     procedure DoDatabase(AFuncId: integer);
     procedure DoVar(AArgCount: integer);
     procedure DoLinest(AArgCount: integer);
     procedure DoTrend(AArgCount: integer);
     procedure DoLogEst(AArgCount: integer);
     procedure DoGrowth(AArgCount: integer);
     procedure DoText;
     procedure DoAnnuity(AFuncId: integer; AArgCount: integer);
     procedure DoMIRR;
     procedure DoIRR(AArgCount: integer);
     procedure DoMatch(AArgCount: integer);
     procedure DoDate;
     procedure DoTime;
     procedure DoExtractDateTime(AFuncId: integer; AArgCount: integer);
     procedure DoAreas;
     procedure DoRowsColumns(AFuncId: integer);
     procedure DoOffset(AArgCount: integer);
     procedure DoSearch(AArgCount: integer; ACaseSensitive: boolean);
     procedure DoTranspose;
     procedure DoType;
     procedure DoChoose(AArgCount: integer);
     procedure DoHLookup(AArgCount: integer);
     procedure DoVLookup(AArgCount: integer);
     procedure DoReplace;
     procedure DoSubstitute(AArgCount: integer);
     procedure DoCell(AArgCount: integer);
     procedure DoSLN;
     procedure DoSYD;
     procedure DoDDB(AArgCount: integer);
     procedure DoClean;
     procedure DoMDeterm;
     procedure DoMInverse;
     procedure DoMMult;
     procedure DoIPMT(AArgCount: integer);
     procedure DoPPMT(AArgCount: integer);
     procedure DoCounta(AArgCount: integer);
     procedure DoProduct(AArgCount: integer);
     procedure DoFact;
     procedure DoStdevp(AArgCount: integer);
     procedure DoVarp(AArgCount: integer);
     procedure DoTrunc(AArgCount: integer);
     procedure DoQuotient;
     procedure DoRoundUp;
     procedure DoRank(AArgCount: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoAddress(AArgCount: integer);
     procedure DoDays360(AArgCount: integer);
     procedure DoVDB(AArgCount: integer);
     procedure DoMedian(AArgCount: integer);
     procedure DoSumProduct(AArgCount: integer);
     procedure DoInfo;
     procedure DoDB(AArgCount: integer);
     procedure DoFrequency;
     procedure DoError_Type;
     procedure DoAveDev(AArgCount: integer);
     procedure DoGammaLn;
     procedure DoBetaDist(AArgCount: integer; AVer2010: boolean);
     procedure DoBetaInv(AArgCount: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoBinomdist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoChiDist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoChiInv;
     procedure DoCombin;
     procedure DoConfidence; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoCritbinom;
     procedure DoBinom_Inv;
     procedure DoExpondist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoFDist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoFInv;
     procedure DoF_Inv;
     procedure DoF_Inv_RT;
     procedure DoFisher;
     procedure DoFisherInv;
     procedure DoFloor;
     procedure DoGammadist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoGammaInv;
     procedure DoGamma_Inv;
     procedure DoCeiling;
     procedure DoLognormdist;
     procedure DoLognorm_dist;
     procedure DoLognorm_Inv;
     procedure DoLoginv;
     procedure DoNegbinomdist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoNormdist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoNormsdist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoNormInv;
     procedure DoNorm_Inv;
     procedure DoNormsInv;
     procedure DoNorm_S_Inv;
     procedure DoStandardize;
     procedure DoOdd;
     procedure DoPermut;
     procedure DoPoisson; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoTDist;
     procedure DoWeibull; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoSumXmY2;
     procedure DoSumX2mY2;
     procedure DoSumX2pY2;
     procedure DoChiTest; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoCorrel;
     procedure DoCovar; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoForecast;
     procedure DoFTest;
     procedure DoF_Test;
     procedure DoIntercept;
     procedure DoRsq;
     procedure DoSteyx;
     procedure DoSlope;
     procedure DoTTest; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoLarge;
     procedure DoSmall;
     procedure DoQuartile;
     procedure DoPercentile;
     procedure DoTrimMean;
     procedure DoTInv;
     procedure DoPower;
     procedure DoRadians;
     procedure DoDegrees;
     procedure DoCountif;
     procedure DoCountblank;
     procedure DoISPMT;
     procedure DoPhonetic;

     procedure DoT_Dist_2T;
     procedure DoT_Dist_RT;
     procedure DoT_Dist;
     procedure DoT_Inv;
     procedure DoT_Inv_2T;
     procedure DoChiSq_Inv;
     procedure DoChi_Inv_RT;
     procedure DoErfc;
     procedure DoErfc_Precise;
     procedure DoAggregate(AArgCount: integer);
     procedure DoStdev_S(AArgCount: integer);
     procedure DoStdev_P(AArgCount: integer);
     procedure DoVar_S(AArgCount: integer);
     procedure DoVar_P(AArgCount: integer);
     procedure DoMode_Sngl(AArgCount: integer);
     procedure DoPercentile_Inc;
     procedure DoQuartile_Inc;
     procedure DoPercentile_Exc;
     procedure DoQuartile_Exc;

     procedure DoRoman(AArgCount: integer);
     procedure DoProb(AArgCount: integer);
     procedure DoDevSq(AArgCount: integer);
     procedure DoGeomean(AArgCount: integer);
     procedure DoHarmean(AArgCount: integer);
     procedure DoSumsq(AArgCount: integer);
     procedure DoKurt(AArgCount: integer);
     procedure DoSkew(AArgCount: integer);
     procedure DoZtest(AArgCount: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoPercentRank(AArgCount: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DoMode(AArgCount: integer);
     procedure DoConcatenate(AArgCount: integer);
     procedure DoSubtotal(AArgCount: integer);
     procedure DoSumif(AArgCount: integer);
     procedure DoAveragea(AArgCount: integer);
     procedure DoMaxa(AArgCount: integer);
     procedure DoMina(AArgCount: integer);
     procedure DoStdevpa(AArgCount: integer);
     procedure DoVarpa(AArgCount: integer);
     procedure DoStdeva(AArgCount: integer);
     procedure DoVara(AArgCount: integer);
     procedure DoHyperlink(AArgCount: integer);

     procedure DoChisq_Dist;
     procedure DoConfidence_T;
     procedure DoCovariance_S;
     procedure DoErf_Precise;
     procedure DoF_Dist;
     procedure DoGammaln_Precise;
     procedure DoIferror;
     procedure DoAverageIf(AArgCount: integer);
     procedure DoAverageIfs(AArgCount: integer);
     procedure DoCeiling_Precise(AArgCount: integer);
     procedure DoCountIfs(AArgCount: integer);
     procedure DoFloor_Precise(AArgCount: integer);
     procedure DoIso_Ceiling(AArgCount: integer);
     procedure DoMode_Mult(AArgCount: integer);
     procedure DoNetWorkdays(AArgCount: integer);
     procedure DoNetWorkdays_Intl(AArgCount: integer);
     procedure DoPercentrank_Exc(AArgCount: integer);
     procedure DoSumifs(AArgCount: integer);
     procedure DoWorkday(AArgCount: integer);
     procedure DoWorkday_Int(AArgCount: integer);

     procedure DoBeta_Inv(AArgCount: integer);
     procedure DoZ_Test(AArgCount: integer);

     procedure DoBinom_Dist;
     procedure DoChisq_Dist_rt;
     procedure DoChisq_Test;
     procedure DoConfidence_Norm;
     procedure DoCovariance_P;
     procedure DoExpon_Dist;
     procedure DoF_dist_RT;
     procedure DoGamma_Dist;
     procedure DoHypgeom_Dist;
     procedure DoNegbinom_Dist;
     procedure DoNorm_Dist;
     procedure DoNorm_S_Dist;
     procedure DoPercentrank_Inc(AArgCount: integer);
     procedure DoPoisson_Dist;
     procedure DoRank_Eq(AArgCount: integer);
     procedure DoT_Test;
     procedure DoWeibull_Dist;
     procedure DoEDate;
     procedure DoEOMonth;
     procedure DoYearfrac(AArgCount: integer);
     procedure DoXIRR(AArgCount: integer);
     procedure DoXNVP;

     procedure DoHypGeomDist; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure DoAccrInt(AArgCount: integer);
     end;

type TXLSFormulaEvaluator = class(TObject)
protected
     FManager      : TXc12Manager;
     FSheetIndex   : integer;
     FCol,FRow     : integer;
     FTargetArea   : PXLSFormulaArea;
     FUserFuncEvent: TUserFunctionEvent;

     FDebugData    : TFmlaDebugItems;
     FDebugSteps   : integer;

     FStack        : TValueStack;

     FVolatile     : boolean;

     FIsExcel97    : boolean;

     procedure FillTargetArea(FV: TXLSFormulaValue; AHasParentFormula: boolean);
     procedure DoArray(APtgs: PXLSPtgs);
{$ifdef XLS_BIFF}
     procedure DoArray97(var APtgs: PXLSPtgs);
{$endif}
     function  DoDataTableFormula(APtgs: PXLSPtgs; APtgsSz: integer; AIdRef1,AIdRef2,ASrcRef1,ASrcRef2: PXLSPtgsRef): TXLSFormulaValue;
     procedure DoDataTable(APtgs: PXLSPtgsDataTableFmla);
     // INDIRECT must execute here as a name will need to call DoEvaluate.
     procedure DoIndirect(AArgCount: integer);
     procedure DoName(const AName: AxUCString; const ASheetIndex: integer);
     procedure DoFunction(AId: integer); overload;
     procedure DoFunction(AId: integer; AArgCount: integer); overload;
     procedure DoFunction(const AName: AxUCString; AArgCount: integer); overload;
     procedure DoArrayFunction(APtgs: PXLSPtgs; ATargetArea: PXLSFormulaArea);
     procedure DoTable(ATable: PXLSPtgsTable);
     function  DoEvaluate(APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea; AIsName: boolean = False): TXLSFormulaValue;
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     function  Evaluate(ASheetIndex,ACol,ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea): TXLSFormulaValue; // inline;
     function  EvaluateStr(ASheetIndex,ACol,ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea): AxUCString;
     procedure DebugEvaluate(ADebugData : TFmlaDebugItems; ASheetIndex,ACol,ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea);

     property Stack: TValueStack read FStack;

     property Volatile: boolean read FVolatile;

     property OnUserFunction: TUserFunctionEvent read FUserFuncEvent write FUserFuncEvent;
     end;

implementation

const
  Array10: array[0..15] of double = (1,10,100,1000,10000,100000,1000000,10000000,1000000000,10000000000,100000000000,1000000000000,10000000000000,100000000000000,1000000000000000,10000000000000000);

{ TValueStack }

function TValueStack.HandleError(AError: TXc12CellError): boolean;
begin
  Result := AError = errUnknown;
  if not Result then
    SetError(AError);
end;

function TValueStack.AreaCellCount(FV: PXLSFormulaValue): integer;
begin
  case FV.RefType of
    xfrtNone,
    xfrtRef     : Result := 1;
    xfrtArea,
    xfrtXArea   : Result := (FV.Col2 - FV.Col1 + 1) * (FV.Row2 - FV.Row1 + 1);
    xfrtAreaList: Result := TCellAreas(FV.vSource).CellCount;
    xfrtArray   : Result := TXLSArrayItem(FV.vSource).Width * TXLSArrayItem(FV.vSource).Height;
    else          Result := 0;
  end;
end;

function TValueStack.AreasEqualSize(FV1, FV2: PXLSFormulaValue): boolean;
begin
  Result := (FV1.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) and (FV2.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]);
  if Result then begin
    Result := AreaCellCount(FV1) = AreaCellCount(FV2);
  end;
end;

procedure TValueStack.Clear(AIsArrayRes: boolean; ASheetIndex,ACol,ARow: integer);
begin
  FIsArrayRes := AIsArrayRes;
  FSheetIndex := ASheetIndex;
  FCells := FManager.Worksheets[FSheetIndex].Cells;
  FCol := ACol;
  FRow := ARow;
  FStackPtr := -1;
  FGarbage.Clear;
end;

procedure TValueStack.ClearStack;
begin
  FStackPtr := 0;
end;

function TValueStack.CoupDaysBetween(ADateFrom, ADateTo: TDateTime; ADateMethod: TXLSDateDiffMethod): integer;
var
  Y1,M1,D1: word;
  Y2,M2,D2: word;
  Months: integer;
  Days: integer;
begin
  case ADateMethod of
    xddmUS30_360: begin
      DecodeDate(ADateFrom,Y1,M1,D1);
      DecodeDate(ADateTo,Y2,M2,D2);
      Days := D2 - D1;
      Months := (M2 - M1) + (Y2 - Y1) * 12;
      Result := Days + Months * 30;
      if (M1 = 2) and (M2 <> 2) and (Y1 = Y2) then begin
        if IsLeapYear(Y1) then
          Dec(Result)
        else
          Dec(Result,2)
      end;
    end;
    xddmActaul,
    xddm360,
    xddm365     : Result := Round(ADateTo - ADateFrom);
    xddmEU30_360: begin
      DecodeDate(ADateFrom,Y1,M1,D1);
      DecodeDate(ADateTo,Y2,M2,D2);
      Days := D2 - D1;
      Months := (M2 - M1) + (Y2 - Y1) * 12;
      Result := Days + Months * 30;
    end;
    else Result := 0;
  end;
end;

function TValueStack.CoupDaysYear(ADate: TDateTime; ADateMethod: TXLSDateDiffMethod): integer;
var
  Y1,M1,D1: word;
begin
  case ADateMethod of
    xddmUS30_360: Result := 360;
    xddmActaul  : begin
      DecodeDate(ADate,Y1,M1,D1);
      if IsLeapYear(Y1) then
        Result := 366
      else
        Result := 365;
    end;
    xddm360     : Result := 360;
    xddm365     : Result := 365;
    xddmEU30_360: Result := 360;
    else Result := 0;
  end;
end;

procedure TValueStack.OpConcat;
var
  E1,E2: TXc12CellError;
  S1,S2: AxUCString;
begin
  DecStackPtr;
  E1 := GetAsString(FStack[FStackPtr],S1);
  E2 := GetAsString(FStack[FStackPtr + 1],S2);
  if E1 <> errUnknown then
    SetError(E1)
  else if E2 <> errUnknown then
    SetError(E2)
  else begin
    FStack[FStackPtr].vStr := S1 + S2;
    FStack[FStackPtr].ValType := xfvtString;
    FStack[FStackPtr].RefType := xfrtNone;
  end;
end;

constructor TValueStack.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FGarbage := TObjectList.Create;
  FIterator := TValStackIterator.Create(Self);
  FStackPtr := -1;
end;

function TValueStack.DB(Cost, Salvage: double; Life, Period, Month: integer): double;
var
  Rate: Extended;
  DPv : Extended;
  TDPv: Extended;
  I   : Integer;
begin
  DPv := 0.0;
  TDPv := 0.0;
  if Salvage = 0 then
    Salvage := 0.001;
  if Month = 0 then
    Month := 12;
  Rate := RoundToDecimal(1.0 - Power(Salvage / Cost, 1.0 / Life), 3);
  for I := 1 to Period do begin
    if (I = 1) then
      DPv := (Cost * Rate * Month) / 12.0
    else if (I = (Life + 1)) then
      DPv := (Cost - TDPv) * Rate * (12.0 - Month) / 12.0
    else
      DPv := (Cost - TDPv) * Rate;
    TDpv := TDpv + Dpv
  end;
  Result := Dpv;
end;

function TValueStack.DDB(const Cost, Salvage, Life, Period, Factor: double): double;
var
  DepreciatedVal: double;
  Fac: double;
begin
  if Life <= 2 then begin
    if Period = 1 then
      Result := Cost - Salvage
    else
      Result := 0;
    Exit;
  end;
  Fac := Factor / Life;

  DepreciatedVal := Cost * Power((1.0 - Fac), Period - 1);

  Result := Fac * DepreciatedVal;

  if Result > DepreciatedVal - Salvage then
    Result := DepreciatedVal - Salvage;

  if Result < 0.0 then
    Result := 0.0;
end;

procedure TValueStack.DecStackPtr;
begin
  Dec(FStackPtr);
  if FStackPtr < -1 then
    raise XLSRWException.Create('Missing operand in expression in cell ' + ColRowToRefStr(FCol,FRow));
end;

destructor TValueStack.Destroy;
begin
  FGarbage.Free;
  FIterator.Free;
  inherited;
end;

procedure TValueStack.DoAccrInt(AArgCount: integer);
var
  Issue: double;
  FirstInterest: double;
  Settlement: double;
  Rate: double;
  Par: double;
  Frequency: integer;
  Basis: integer;
  A,D: double;
begin
  Basis := 0;
  Frequency := 1;
  Par := 1000;
  if AArgCount = 7 then
    if not PopResPosInt(Basis) then Exit;
  if AArgCount >= 6 then
    if not PopResPosInt(Frequency) then Exit;
  if AArgCount >= 5 then
    if not PopResFloat(Par) then Exit;
  if not PopResFloat(Rate) then Exit;
  if not PopResFloat(Settlement) then Exit;
  if not PopResFloat(FirstInterest) then Exit;
  if not PopResFloat(Issue) then Exit;

  if (Rate <= 0) or (Par <= 0) or not (Frequency in [1,2,4]) or (Basis > 4) or (Issue >= Settlement) then
    Push(errNum)
  else begin
//    if FirstInterest >= Settlement then
      A := CoupDaysBetween(Issue,Settlement,TXLSDateDiffMethod(Basis));
//    else
//      A := CoupDaysBetween(Issue,FirstInterest,TXLSDateDiffMethod(Basis));
    D := CoupDaysYear(Settlement,TXLSDateDiffMethod(Basis));
    Push((Par * Rate * A) / D);
  end;
end;

procedure TValueStack.DoAddress(AArgCount: integer);
var
  SheetText: AxUCString;
  A1: boolean;
  AbsNum: integer;
  Col,Row: integer;
  AbsCol,AbsRow: boolean;
begin
  SheetText := '';
  A1 := True;
  AbsNum := 1;
  if AArgCount = 5 then
    if not PopResStr(SheetText) then Exit;
  if AArgCount >= 4 then
    if not PopResBoolean(A1) then Exit;
  if AArgCount >= 3 then
    if not PopResInt(AbsNum) then Exit;
  if not PopResInt(Col) then Exit;
  if not PopResInt(Row) then Exit;

  if not InsideExtent(Col,Row,Col,Row) then begin
    Push(errValue);
    Exit;
  end;

  case AbsNum of
    2: begin
      AbsCol := False;
      AbsRow := True;
    end;
    3: begin
      AbsCol := True;
      AbsRow := False;
    end;
    4: begin
      AbsCol := False;
      AbsRow := False;
    end;
    else begin
      AbsCol := True;
      AbsRow := True;
    end;
  end;

  // TODO R1C1

  if SheetText <> '' then
    Push(SheetText + '!' + ColRowToRefStr(Col,Row,AbsCol,AbsRow))
  else
    Push(ColRowToRefStr(Col,Row,AbsCol,AbsRow));
end;

procedure TValueStack.DoAggregate(AArgCount: integer);
var
  FuncNum: integer;
  Options: integer;
  V1,V2: double;
//  IgnoreNested: boolean;
begin
  if not PeekFloat(AArgCount - 2,V1) then Exit;
  if not PeekFloat(AArgCount - 1,V2) then Exit;

  Options := Trunc(V1);
  if (Options < 0) or (Options > 7) then begin
    Push(errValue);
    Exit;
  end;

  FuncNum := Trunc(V2);
  if (FuncNum < 1) or (FuncNum > 19) then begin
    Push(errValue);
    Exit;
  end;

  if (FuncNum in [14,15,16,17,18,19]) and (AArgCount < 4) then begin
    Push(errValue);
    Exit;
  end;

//  IgnoreNested := False;

  try
    case Options of
      0: begin
//        IgnoreNested := True;
      end;
      1: begin
//        IgnoreNested := True;
        FIterator.IgnoreHidden := True;
      end;
      3: begin
//        IgnoreNested := True;
        FIterator.IgnoreHidden := True;
        FIterator.IgnoreError := True;
      end;
      4: ;
      5: FIterator.IgnoreHidden := True;
      6: FIterator.IgnoreError := True;
      7: begin
        FIterator.IgnoreHidden := True;
        FIterator.IgnoreError := True;
      end;
    end;

    case FuncNum of
       1: DoAverage(AArgCount - 2);
       2: DoCount(AArgCount - 2);
       3: DoCountA(AArgCount - 2);
       4: DoMax(AArgCount - 2);
       5: DoMin(AArgCount - 2);
       6: DoProduct(AArgCount - 2);
       7: DoStdev_S(AArgCount - 2);
       8: DoStdev_P(AArgCount - 2);
       9: DoSum(AArgCount - 2);
      10: DoVar_S(AArgCount - 2);
      11: DoVar_P(AArgCount - 2);
      12: DoMedian(AArgCount - 2);
      13: DoMode_Sngl(AArgCount - 2);
      14: DoLarge;
      15: DoSmall;
      16: DoPercentile_Inc;
      17: DoQuartile_Inc;
      18: DoPercentile_Exc;
      19: DoQuartile_Exc;
    end;
  finally
    FIterator.IgnoreHidden := False;
    FIterator.IgnoreError := False;
  end;
end;

procedure TValueStack.DoAnnuity(AFuncId, AArgCount: integer);
var
  Rate : double;
  NPer : integer;
  Pmt  : double;
  Fv   : double;
  PType: integer;
  PT   : TPaymentTime;
begin
  PType := 0;
  Fv := 0;
  if AArgCount >= 6 then
    PopResFloat(Fv); // Ignore.
  if AArgCount >= 5 then
    if not PopResInt(PType) then Exit;
  if AArgCount >= 4 then
    if not PopResFloat(Fv) then Exit;

  if not PopResFloat(Pmt) then Exit;
  if not PopResInt(NPer) then Exit;
  if not PopResFloat(Rate) then Exit;

  if PType = 0 then
    PT := ptEndOfPeriod
  else
    PT := ptStartOfPeriod;

  if NPer = 0 then
    Push(errDiv0)
  else begin
    case AFuncId of
      056: Push(PresentValue(Rate,NPer,Pmt,Fv,PT));       // PV
      057: Push(FutureValue(Rate,NPer,Pmt,Fv,PT));        // FV
      058: Push(NumberOfPeriods(Rate,NPer,Pmt,Fv,PT));    // NPER
      059: Push(Payment(Rate,NPer,Pmt,Fv,PT));            // PMT
      060: Push(InterestRate(Round(Rate),NPer,Pmt,Fv,PT));// RATE
    end;
  end;
end;

procedure TValueStack.DoAreas;
var
  n: integer;
  FV: TXLSFormulaValue;
begin
  FV := PopRes;

  case FV.RefType of
    xfrtRef,
    xfrtArea,
    xfrtXArea   : n := 1;
    xfrtAreaList: n := TCellAreas(FV.vSource).Count;
    else begin
      Push(errValue);
      Exit;
    end;
  end;
  Push(n);
end;

procedure TValueStack.DoAveDev(AArgCount: integer);
var
  i: integer;
  Average: double;
  Res: double;
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValues(AArgCount,Vals) then begin
    if Length(Vals) <= 0 then
      Push(errValue)
    else begin
      Average := 0;
      for i := 0 to High(Vals) do
        Average := Average + Vals[i];
      Average := Average / Length(Vals);
      Res := 0;
      for i := 0 to High(Vals) do
        Res := Res + Abs(Vals[i] - Average);
      Res := Res * (1 / Length(Vals));
      Push(Res);
    end;
  end;
end;

procedure TValueStack.DoAverage(AArgCount: integer);
var
  N: integer;
  Sum: double;
begin
  N := 0;
  Sum := 0;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloat do begin
      Sum := Sum + FIterator.Result.vFloat;
      Inc(N);
    end;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  if N = 0 then
    Push(errDiv0)
  else
    Push(Sum / N);
end;

procedure TValueStack.DoAveragea(AArgCount: integer);
var
  N: integer;
  Sum: double;
begin
  N := 0;
  Sum := 0;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloatA do begin
      Sum := Sum + FIterator.AsFloatA;
      Inc(N);
    end;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  if N = 0 then
    Push(errDiv0)
  else
    Push(Sum / N);
end;

procedure TValueStack.DoAverageIf(AArgCount: integer);
var
  Criteria: TXLSFormulaValue;
  CritVal: TXLSVarValue;
  AveRange: TXLSFormulaValue;
  Val: TXLSVarValue;
  FV: TXLSFormulaValue;
  Op: TXLSDbCondOperator;
  Sum: double;
  Sum2: double;
  n: integer;
begin
  if AArgCount = 3 then
    AveRange := PopRes;
  Criteria := PopRes;
  case Criteria.RefType of
    xfrtNone,
    xfrtRef     : FV := Criteria;
    xfrtArea    : GetCellValue(Criteria.Cells,Criteria.Col1,Criteria.Row1,@FV);
    xfrtXArea   : ;
    xfrtAreaList: begin
      Push(errValue);
      Exit;
    end;
    xfrtArray   : FV := TXLSArrayItem(Criteria.vSource).GetAsFormulaValue(0,0);
  end;
  FVToVarValue(@FV,@CritVal);

  Op := xdcoEQ;
  if CritVal.ValType = xfvtString then
    ConditionStrToVarValue(CritVal.vStr,Op,@CritVal);

  if AArgCount = 3 then begin
    FV := PopRes;
    if not (AveRange.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) then begin
      Push(errValue);
      Exit;
    end;
    if (FV.Col2 - FV.Col1) > (FV.Row2 - FV.Row1) then
      AveRange.Col2 := Max(FV.Col2,AveRange.Col2)
    else
      AveRange.Row2 := Max(FV.Row2,AveRange.Row2);
    FIterator.BeginIterate(FV,AveRange,False);
  end
  else
    FIterator.BeginIterate(PopRes,False);

  n := 0;
  Sum := 0;
  Sum2 := 0;
  while FIterator.Next do begin
    FVToVarValue(@FIterator.Result,@Val);
    if CompareVarValue(@Val,@CritVal,Op) then begin
      if FIterator.ResIsFloat then
        Sum := Sum + FIterator.Result.vFloat;
      if (AArgCount = 3) and FIterator.Linked.ResIsFloat then
        Sum2 := Sum2 + FIterator.Linked.Result.vFloat;
      Inc(n);
    end;
  end;
  if n <= 0 then
    Push(errDiv0)
  else if AArgCount = 3 then
    Push(Sum2 / n)
  else
    Push(Sum / n);
end;

procedure TValueStack.DoAverageIfs(AArgCount: integer);
var
  i: integer;
  Val: TXLSVarValue;
  Sum: double;
  n: integer;
  Iter: TValStackIterator;
  Ok: boolean;
  Data: TDynCriteriaDataArray;
begin
  if not Odd(AArgCount) then begin
    Push(errValue);
    Exit;
  end;

  SetupIfs(AArgCount,Data);

  n := 0;
  Sum := 0;
  while FIterator.Next do begin
    Iter := FIterator;
    Ok := True;
    for i := 0 to High(Data) do begin
      FVToVarValue(@Iter.Result,@Val);
      Ok := CompareVarValue(@Val,@Data[i].CritVal,Data[i].Op);
      if not Ok then
        Break;
      Iter := Iter.Linked;
    end;
    if Ok then begin
      if Iter.ResIsFloat then
        Sum := Sum + Iter.Result.vFloat;
      Inc(n);
    end;
  end;
  if n = 0 then
    Push(errDiv0)
  else
    Push(Sum / n);
end;

procedure TValueStack.DoBetaDist(AArgCount: integer; AVer2010: boolean);
var
  B: double;
  A: double;
  Cumulative: boolean;
  Beta: double;
  Alpha: double;
  X: double;
begin
  A := 0;
  B := 1;
  Cumulative := True;
  if AVer2010 then begin
    if not PopResBoolean(Cumulative) then Exit;
    if AArgCount = 6 then
      if not PopResFloat(B) then Exit;
    if AArgCount >= 5 then
      if not PopResFloat(A) then Exit;
  end
  else begin
    if AArgCount = 5 then
      if not PopResFloat(B) then Exit;
    if AArgCount >= 4 then
      if not PopResFloat(A) then Exit;
  end;
  if not PopResFloat(Beta) then Exit;
  if not PopResFloat(Alpha) then Exit;
  if not PopResFloat(X) then Exit;

  if (X < A) or (X > B) or (A = B) or (Alpha <= 0) or (Beta <= 0) then
    Push(errNum)
  else begin
    if Cumulative then
      Push(BetaCDF(Alpha,Beta,X))
    else
      Push(BetaPDF(Alpha,Beta,X));
  end;
end;

procedure TValueStack.DoBetaInv(AArgCount: integer);
begin
  DoBeta_Inv(AArgCount);
end;

procedure TValueStack.DoBeta_Inv(AArgCount: integer);
var
  B: double;
  A: double;
  Beta: double;
  Alpha: double;
  Prob: double;
begin
  A := 0;
  B := 1;
  if AArgCount = 5 then
    if not PopResFloat(B) then Exit;
  if AArgCount >= 4 then
    if not PopResFloat(A) then Exit;
  if not PopResFloat(Beta) then Exit;
  if not PopResFloat(Alpha) then Exit;
  if not PopResFloat(Prob) then Exit;

  if (Prob < A) or (Prob > B) or (Prob = 0) or (Prob = 1) or (A = B) or (Alpha <= 0) or (Beta <= 0) then
    Push(errNum)
  else
    Push(BetaInvCDF(Alpha,Beta,Prob));
end;

procedure TValueStack.DoBinomdist;
begin
  DoBinom_Dist;
end;

procedure TValueStack.DoBinom_Dist;
var
  Cumulative: boolean;
  Probability_s: double;
  Trials: integer;
  Number_s: integer;
begin
  if not PopResBoolean(Cumulative) then Exit;
  if not PopResFloat(Probability_s) then Exit;
  if not PopResInt(Trials) then Exit;
  if not PopResInt(Number_s) then Exit;

  if (Number_s < 0) or (Number_s > Trials) or (Probability_s < 0) or (Probability_s > 1) then
    Push(errNum)
  else begin
    if Cumulative then
      Push(BinomialCDF(Number_s,Trials,Probability_s))
    else
      Push(BinomialPDF(Number_s,Trials,Probability_s));
  end;
end;

procedure TValueStack.DoBinom_Inv;
var
  alpha: double;
  probability_s: double;
  trials: integer;
begin
  if not PopResFloat(alpha) then Exit;
  if not PopResFloat(probability_s) then Exit;
  if not PopResInt(trials) then Exit;

  if (trials < 0) or (probability_s < 0) or (probability_s > 1) or (alpha < 0) or (alpha > 1) then
    Push(errNum)
  else
    Push(BinomialQuantile(trials,alpha,probability_s));
end;

procedure TValueStack.DoAND(AArgCount: integer);
var
  i: integer;
  B: boolean;
  Res: boolean;
  Err: TXc12CellError;
begin
  Res := True;

  Err := errUnknown;
  for i := 0 to AArgCount - 1 do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.Next do begin
      if Err = errUnknown then begin
        Err := GetAsBoolean(FIterator.Result,B);
        if Err = errUnknown then
          Res := Res and B;
      end;
    end;
  end;
  if Err <> errUnknown then
    Push(Err)
  else
    Push(Res);
end;

procedure TValueStack.DoOR(AArgCount: integer);
var
  i: integer;
  B: boolean;
  Err: TXc12CellError;
begin
  Err := errUnknown;
  for i := 0 to AArgCount - 1 do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.Next do begin
      if Err = errUnknown then begin
        Err := GetAsBoolean(FIterator.Result,B);
        if (Err = errUnknown) and B then begin
          Push(True);
          Exit;
        end;
      end;
    end;
  end;
  if Err <> errUnknown then
    Push(Err)
  else
    Push(False);
end;

procedure TValueStack.DoCeiling;
var
  number: double;
  significance: double;
begin
  if not PopResFloat(significance) then Exit;
  if not PopResFloat(number) then Exit;

  Push(Round(number / significance + 0.4) * significance);
end;

procedure TValueStack.DoCeiling_Precise(AArgCount: integer);
var
  number: double;
  significance: double;
begin
  significance := 1;
  if AArgCount = 2 then
    if not PopResFloat(significance) then Exit;
  if not PopResFloat(number) then Exit;

  significance := Abs(significance);
  if significance = 0 then
    Push(0)
  else
    Push(Round(number / significance + 0.4) * significance);
end;

procedure TValueStack.DoCell(AArgCount: integer);
var
  Cell: TXLSCellItem;
  FV: TXLSFormulaValue;
  S: AxUCString;
  Sheet: TXc12DataWorksheet;
begin
  if AArgCount = 2 then begin
    FV := PopRes;
    if not (FV.RefType in [xfrtRef,xfrtXArea,xfrtArea]) then begin
      Push(errValue,FV);
      Exit;
    end;
  end
  else begin
    // TODO? this shall be the latest edited cell.
    FV.Col1 := 0;
    FV.Row1 := 0;
  end;
  if not PopResStr(S,False) then
    Exit;

  S := Lowercase(S);
  if S = 'address' then
    Push(ColRowToRefStr(FV.Col1,FV.Row1,True,True))
  else if S = 'col' then
    Push(FV.Col1 + 1)
  else if S = 'color' then
    Push(0)
  else if S = 'filename' then begin
    S := FManager.FilenameAsXLS;
    if S <> '' then begin
      // If you read the docs it don't looks like the sheet name is included
      // but this is what excel does.
      case FV.RefType of
        xfrtRef,
        xfrtArea    : begin
          if FV.Cells <> Nil then begin
            Sheet := FManager.Worksheets.FindSheetFromCells(FV.Cells);
            if Sheet <> Nil then
              Push(S + Sheet.Name)
            else
              Push(S + '[???]');
          end
          else
            Push(S + FManager.Worksheets[FSheetIndex].Name);
        end;
        xfrtXArea   : Push(S); // TODO
        else          Push(S + FManager.Worksheets[FSheetIndex].Name);
      end;
    end
    else
      Push('');
  end
  else if S = 'format' then
    Push('')
  else if S = 'parentheses' then
    Push(0)
  else if S = 'prefix' then
    Push('')
  else if S = 'protect' then
    Push(0)
  else if S = 'row' then
    Push(FV.Row1 + 1)
  else if S = 'type' then begin
    if FV.Cells.FindCell(FV.Col1,FV.Row1,Cell) then
      Push('b')
    else begin
      if Fv.Cells.CellType(@Cell) = xctBlank then
        Push('b')
      else
        Push('v');
    end;
  end
  else if S = 'width' then
    Push(0)
  else
    Push(errValue);
end;

procedure TValueStack.DoChiDist;
begin
  DoChisq_Dist_rt;
end;

procedure TValueStack.DoChiInv;
begin
  DoChi_Inv_RT;
end;

procedure TValueStack.DoChisq_Dist;
var
  cumulative: boolean;
  degrees_freedom: double;
  x: double;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResFloat(degrees_freedom) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (degrees_freedom < 1) or (degrees_freedom > 10e10) then
    Push(errNum)
  else begin
    if cumulative then
      Push(ChiSquaredDistributionCDF(degrees_freedom,x))
    else
      Push(ChiSquaredDistributionPDF(degrees_freedom,x) * 0.5);
  end;
end;

procedure TValueStack.DoChisq_Dist_rt;
var
  degrees_freedom: double;
  x: double;
begin
  if not PopResFloat(degrees_freedom) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (degrees_freedom < 1) or (degrees_freedom > 10e10) then
    Push(errNum)
  else
    Push(1 - ChiSquaredDistributionCDF(degrees_freedom,x));
end;

procedure TValueStack.DoChiSq_Inv;
var
  degrees_freedom: double;
  probability: double;
begin
  if not PopResFloat(degrees_freedom) then Exit;
  if not PopResFloat(probability) then Exit;

  if (probability < 0) or (probability >= 1) or (degrees_freedom < 1) or (degrees_freedom > 10e10) then
    Push(errNum)
  else
    Push(ChiSquaredQuantile(degrees_freedom,probability));
end;

procedure TValueStack.DoChisq_Test;
var
  FV1: TXLSFormulaValue;
  FV2: TXLSFormulaValue;
  V1,V2: double;
  Res: double;
  nR,nC: integer;
  df: integer;
begin
  FV1 := PopRes;
  FV2 := PopRes;

  if not AreasEqualSize(@FV2,@FV1) or (AreaCellCount(@FV1) <= 1) then
    Push(errNA)
  else begin
    Res := 0;
    FIterator.BeginIterate(FV1,FV2,False);
    while FIterator.NextFloat do begin
      V2 := FIterator.Result.vFloat;
      V1 := FIterator.Linked.Result.vFloat;
      if V2 = 0 then begin
        Push(errDiv0);
        Exit;
      end;
      if V2 < 0 then begin
        Push(errNum);
        Exit;
      end;
      Res := Res + ((V1 - V2) * (V1 - V2) / V2);
    end;
    nR := FV1.Row2 - FV1.Row1 + 1;
    nC := FV1.Col2 - FV1.Col1 + 1;
    df := 0;
    if (nR > 1) and (nC > 1) then
      df := (nR - 1) * (nC - 1)
    else if (nR = 1) and (nC > 1) then
      df := nC - 1
    else if (nR > 1) and (nC = 1) then
      df := nR - 1
    else
      Push(errNA);
    if df > 0 then
      Push(1 - ChiSquaredDistributionCDF(df,Res));
  end;
end;

procedure TValueStack.DoChiTest;
begin
  DoChisq_Test
end;

procedure TValueStack.DoChi_Inv_RT;
var
  degrees_freedom: double;
  probability: double;
begin
  if not PopResFloat(degrees_freedom) then Exit;
  if not PopResFloat(probability) then Exit;

  if (probability < 0) or (probability >= 1) or (degrees_freedom < 1) or (degrees_freedom > 10e10) then
    Push(errNum)
  else
    Push(ChiSquaredQuantile(degrees_freedom,1 - probability));
end;

procedure TValueStack.DoChoose(AArgCount: integer);
var
  n: integer;
  V: double;
  Err: TXc12CellError;
  FV: TXLSFormulaValue;
begin
  FV := PeekRes(AArgCount - 1);

  if not GetAsFloat(FV,V,Err) then begin
     Push(Err);
     Exit;
  end;
  n := Trunc(V);

  if (n < 1) or (n >= AArgCount) then begin
    Push(errValue);
    Exit;
  end;

  FV := PeekRes(AArgCount - n - 1);
  Push(FV);
end;

procedure TValueStack.DoClean;
var
  i,j: integer;
  S1,S2: AxUCString;
begin
  if not PopResStr(S1) then
    Exit;
  SetLength(S2,Length(S1));

  j := 0;
  for i := 1 to Length(S2) do begin
    if Ord(S1[i]) > 32 then begin
      Inc(j);
      S2[j] := S1[i];
    end;
  end;
  SetLength(S2,j);
  Push(S2);
end;

function TValueStack.DoCollectAllValues(AArgCount: integer; var AResult: TDynDoubleArray): boolean;
begin
  Result := AArgCount > 0;
  while AArgCount > 0 do begin
    if Result then
      Result := DoCollectValues(AResult);
    Dec(AArgCount);
  end;
end;

function TValueStack.DoCollectAllValuesA(AArgCount: integer; var AResult: TDynDoubleArray): boolean;
begin
  Result := AArgCount > 0;
  while AArgCount > 0 do begin
    if Result then
      Result := DoCollectValuesA(AResult);
    Dec(AArgCount);
  end;
end;

function TValueStack.DoCollectValues(var AResult: TDynDoubleArray): boolean;
var
  V: double;
  N: integer;

procedure AddValue(AVal: double);
begin
  if N > High(AResult) then
    SetLength(AResult,Length(AResult) + 64);
  AResult[N] := AVal;
  Inc(N)
end;

begin
  Result := False;
  N := Length(AResult);
  SetLength(AResult,Length(AResult) + 64);
  FIterator.BeginIterate(PopRes,False);
  while FIterator.Next do begin
    if not FIterator.Result.Empty then begin
      case FIterator.Result.ValType of
        xfvtBoolean: begin
          if FIterator.Result.RefType = xfrtNone then
            if FIterator.Result.vBool then AddValue(1) else AddValue(0);
        end;
        xfvtFloat  : AddValue(FIterator.Result.vFloat);
        xfvtString : begin
          if (FIterator.Result.RefType = xfrtNone) and TryStrToFloat(FIterator.Result.vStr,V) then
            AddValue(V);
        end;
        xfvtError  : begin
          Push(FIterator.Result.vErr);
          Exit;
        end;
      end;
    end;
  end;
  SetLength(AResult,N);
  Result := True;
end;

function TValueStack.DoCollectValuesA(var AResult: TDynDoubleArray): boolean;
var
  V: double;
  N: integer;

procedure AddValue(AVal: double);
begin
  if N > High(AResult) then
    SetLength(AResult,Length(AResult) + 64);
  AResult[N] := AVal;
  Inc(N)
end;

begin
  Result := False;
  N := Length(AResult);
  SetLength(AResult,Length(AResult) + 64);
  FIterator.BeginIterate(PopRes,False);
  while FIterator.Next do begin
    if not FIterator.Result.Empty then begin
      case FIterator.Result.ValType of
        xfvtBoolean:  if FIterator.Result.vBool then AddValue(1) else AddValue(0);
        xfvtFloat  : AddValue(FIterator.Result.vFloat);
        xfvtString : begin
          if TryStrToFloat(FIterator.Result.vStr,V) then
            AddValue(V)
          else
            AddValue(0);
        end;
        xfvtError  : begin
          Push(FIterator.Result.vErr);
          Exit;
        end;
      end;
    end;
  end;
  SetLength(AResult,N);
  Result := True;
end;

procedure TValueStack.DoCombin;
var
  n: integer;
  k: integer;
begin
  if not PopResPosInt(k) then Exit;
  if not PopResPosInt(n) then Exit;

  if k > n then
    Push(errValue)
  else if n > MAX_QUICKFACTORIAL then
    Push(errNA)
  else
    Push(QuickFactorial(n) / (QuickFactorial(k) * QuickFactorial(n - k)));
end;

procedure TValueStack.DoConcatenate(AArgCount: integer);
var
  i: integer;
  S: AxUCString;
  Res: AxUCString;
  Err: TXc12CellError;
begin
  Res := '';
  for i := AArgCount - 1 downto 0 do begin
    FIterator.BeginIterate(PeekRes(i),False);
    while FIterator.Next do begin
      Err := GetAsString(FIterator.Result,S);
      if Err = errUnknown then
        Res := Res + S;
    end;
  end;
  Push(Res);
end;

procedure TValueStack.DoConfidence;
begin
  DoConfidence_Norm;
end;

procedure TValueStack.DoConfidence_Norm;
var
  size: double;
  standard_dev: double;
  alpha: double;
begin
  if not PopResFloat(size) then Exit;
  if not PopResFloat(standard_dev) then Exit;
  if not PopResFloat(alpha) then Exit;

  if (alpha <= 0) or (alpha >= 1) or (standard_dev <= 0) or (size < 1) then
    Push(errNum)
  else
    Push(NormalQuantile(1.0 - Alpha / 2,0,1) * standard_dev / Sqrt(Size));
end;

procedure TValueStack.DoConfidence_T;
var
  size: double;
  standard_dev: double;
  alpha: double;
begin
  if not PopResFloat(size) then Exit;
  if not PopResFloat(standard_dev) then Exit;
  if not PopResFloat(alpha) then Exit;

  if (alpha <= 0) or (alpha >= 1) or (standard_dev <= 0) or (size < 1) then
    Push(errNum)
  else
    Push(NormalQuantile(1.0 - Alpha / 2,0,1) * standard_dev / Sqrt(Size));
end;

procedure TValueStack.DoCorrel;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  X: double;
  Y: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  Sum2: double;
  Sum3: double;
  N: integer;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    Sum2 := 0;
    Sum3 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      X := FIterator.Result.vFloat;
      Y := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((X - MeanX) * (Y - MeanY));
      Sum2 := Sum2 + Sqr(X - MeanX);
      Sum3 := Sum3 + Sqr(Y - MeanY);
    end;
    if (Sum2 = 0) or (Sum3 = 0) then begin
      Push(errDiv0);
      Exit;
    end;
    Push(Sum1 / Sqrt(Sum2 * Sum3));
  end;
end;

function TValueStack.DoCount(AArgCount: integer): integer;
var
  FV: TXLSFormulaValue;
  i,j: integer;
  V: double;
  CA: TCellAreas;

function DoCountArea(Cells: TXLSCellMMU; AC1,AR1,AC2,AR2: integer): integer;
var
  C,R: integer;
begin
  Result := 0;
  for R := AR1 to AR2 do begin
    for C := AC1 to AC2 do begin
      if GetCellType(Cells,C,R) = xfvtFloat then
        Inc(Result);
    end;
  end;
end;

begin
  Result := 0;
  for i := 0 to AArgCount - 1 do begin
    FV := PopRes;
    case FV.RefType of
      xfrtNone: begin
        case FV.ValType of
          xfvtBoolean,
          xfvtFloat   : Inc(Result);
          xfvtString  : begin
            if TryStrToFloat(FV.vStr,V) then
              Inc(Result);
          end;
        end;
      end;
      xfrtRef: begin
        case FV.ValType of
          xfvtFloat  : Inc(Result);
          xfvtString : ;
        end;
      end;
      xfrtArea    : Inc(Result,DoCountArea(FV.Cells,FV.Col1,FV.Row1,FV.Col2,FV.Row2));
      xfrtAreaList: begin
        CA := TCellAreas(FV.vSource);
        for j := 0 to CA.Count - 1 do
          Inc(Result,DoCountArea(TXLSCellMMU(CA[j].Obj),CA[j].Col1,CA[j].Row1,CA[j].Col2,CA[j].Row2));
      end;
    end;
  end;
end;

procedure TValueStack.DoCounta(AArgCount: integer);
var
  FVT: TXLSFormulaValueType;
  Res: integer;
begin
  Res := 0;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.Next do begin
      GetAsValueType(FIterator.Result,FVT);
      if FVT in [xfvtFloat,xfvtBoolean,xfvtError,xfvtString] then
        Inc(Res);
    end;
    Dec(AArgCount)
  end;
  Push(Res);
end;

procedure TValueStack.DoCountblank;
var
  Res: integer;
begin
  Res := 0;
  FIterator.BeginIterate(PopRes,False);
  while FIterator.Next do begin
    if FIterator.Result.Empty or (FIterator.Result.ValType = xfvtUnknown) or ((FIterator.Result.ValType = xfvtString) and (FIterator.Result.vStr = '')) then
      Inc(Res);
  end;
  Push(Res);
end;

procedure TValueStack.DoCountif;
var
  Condition: TXLSFormulaValue;
  CondVal: TXLSVarValue;
  Val: TXLSVarValue;
  FV: TXLSFormulaValue;
  Op: TXLSDbCondOperator;
  Res: integer;
begin
  Condition := PopRes;
  case Condition.RefType of
    xfrtNone,
    xfrtRef     : FV := Condition; // FV := Condition;
    xfrtArea    : GetCellValue(Condition.Cells,Condition.Col1,Condition.Row1,@FV);
    xfrtXArea   : ;
    xfrtAreaList: begin
      Push(errValue);
      Exit;
    end;
    xfrtArray   : FV := TXLSArrayItem(Condition.vSource).GetAsFormulaValue(0,0);
  end;
  FVToVarValue(@FV,@CondVal);

  Op := xdcoEQ;
  if CondVal.ValType = xfvtString then
    ConditionStrToVarValue(CondVal.vStr,Op,@CondVal);

  FIterator.BeginIterate(PopRes,False);

  Res := 0;
  while FIterator.Next do begin
    FVToVarValue(@FIterator.Result,@Val);
    if CompareVarValue(@Val,@CondVal,Op) then
      Inc(Res);
  end;
  Push(Res);
end;

procedure TValueStack.DoCountIfs(AArgCount: integer);
var
  i: integer;
  Val: TXLSVarValue;
  n: integer;
  Iter: TValStackIterator;
  Ok: boolean;
  Data: TDynCriteriaDataArray;
begin
  if Odd(AArgCount) then begin
    Push(errValue);
    Exit;
  end;

  SetupIfs(AArgCount,Data);

  n := 0;
  while FIterator.Next do begin
    Iter := FIterator;
    Ok := True;
    for i := 0 to High(Data) do begin
      FVToVarValue(@Iter.Result,@Val);
      Ok := CompareVarValue(@Val,@Data[i].CritVal,Data[i].Op);
      if not Ok then
        Break;
      Iter := Iter.Linked;
    end;
    if Ok then
      Inc(n);
  end;
  Push(n);
end;

procedure TValueStack.DoCovar;
begin
  DoCovariance_P;
end;

procedure TValueStack.DoCovariance_P;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  X: double;
  Y: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  N: integer;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      X := FIterator.Result.vFloat;
      Y := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((X - MeanX) * (Y - MeanY));
    end;
    Push(Sum1 / N);
  end;
end;

procedure TValueStack.DoCovariance_S;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  X: double;
  Y: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  N: integer;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      X := FIterator.Result.vFloat;
      Y := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((X - MeanX) * (Y - MeanY));
    end;
    Push((1 / (N - 1)) * Sum1);
  end;
end;

procedure TValueStack.DoCritbinom;
begin
  DoBinom_Inv;
end;

procedure TValueStack.DoDatabase(AFuncId: integer);
var
  fvDb      : TXLSFormulaValue;
  sField    : AxUCString;
  fvCriteria: TXLSFormulaValue;
  Db        : TXLSFmlaDatabase;
  FV        : TXLSFormulaValue;
  Err       : TXc12CellError;
begin
  fvCriteria := PopRes;

  // TODO Relative col number is also ok, e.g. "4"
  if not PopResStr(sField) then
    Exit;

  fvDb := PopRes;
  if not (fvDb.RefType in [xfrtArea,xfrtXArea]) or not (fvCriteria.RefType in [xfrtArea,xfrtXArea]) then begin
    Push(errValue);
    Exit;
  end;

  Db := TXLSFmlaDatabase.Create(fvDb,fvCriteria);
  try
    Err := Db.BuildDb;
    if Err <> errUnknown then begin
      Push(Err);
      Exit;
    end;

    Db.SearchDb(AFuncId,sField,FV);
    Push(FV);
  finally
    Db.Free;
  end;
end;

procedure TValueStack.DoDate;
var
  D : TDateTime;
  YY,MM,DD: integer;
begin
  if not PopResInt(DD) or not PopResInt(MM) or not PopResInt(YY) then
    Exit;

  if YY < 0 then begin
    Push(errNum);
    Exit;
  end;

{$ifndef DELPHI_6_OR_LATER}
  raise XLSRWException.Create('DATE not supported by Delphi 5');
{$endif}

{$ifdef DELPHI_6_OR_LATER}
  D := EncodeDate(YY,1,1);
  D := IncMonth(D,MM - 1);
  D := IncDay(D,DD - 1);
{$else}
  D := EncodeDate(YY,1,1);
{$endif}

  Push(D);
end;

procedure TValueStack.DoDays360(AArgCount: integer);
var
  Method: boolean;
  StartDate: TDateTime;
  EndDate: TDateTime;
  YY1,MM1,DD1: word;
  YY2,MM2,DD2: word;
  N: integer;
begin
  Method := False;
  if AArgCount = 3 then
    if not PopResBoolean(Method) then  Exit;
  if not PopResFloat(Double(EndDate)) then  Exit;
  if not PopResFloat(Double(StartDate)) then  Exit;
  if (EndDate < 0) or (StartDate < 0) then
    Push(errValue)
  else begin
    DecodeDate(StartDate,YY1,MM1,DD1);
    DecodeDate(EndDate,YY2,MM2,DD2);
{$ifdef DELPHI_6_OR_LATER}
    if not Method then begin
      if DaysInMonth(EndDate) = DD2 then begin
        if DD1 < 30 then begin
          DD2 := 1;
          Inc(MM2);
          if MM2 > 12 then begin
            MM2 := 1;
            Inc(YY2);
          end;
        end
        else
          DD2 := 30;
      end;
      if DaysInMonth(EndDate) = DD1 then
        DD1 := 30;
    end
    else begin
      if DaysInMonth(StartDate) = DD1 then
        DD1 := 30;
      if DaysInMonth(EndDate) = DD2 then
        DD2 := 30;
    end;
{$else}
    raise XLSRWException.Create('DATE360 not supported by Delphi 5');
{$endif}
    N := ((YY2 * 360) + (MM2 * 30) + DD2) - ((YY1 * 360) + (MM1 * 30) + DD1);
    Push(N);
  end;
end;

procedure TValueStack.DoDB(AArgCount: integer);
var
  Cost   : double;
  Salvage: double;
  Life   : integer;
  Period : integer;
  Month  : integer;
begin
  Month := 12;
  if AArgCount = 5 then
    if not PopResPosInt(Month,1) then Exit;
  if not PopResPosInt(Period,1) then Exit;
  if not PopResPosInt(Life,1) then Exit;
  if not PopResFloat(Salvage) then Exit;
  if not PopResFloat(Cost) then Exit;
  if (Cost <= 0.0) or (Cost < Salvage) or (Period < 1) or (Life < 2) or (Period > (Life + 1)) then begin
    Push(errValue);
    Exit;
  end;
  Push(DB(Cost,Salvage,Life,Period,Month));
end;

procedure TValueStack.DoDDB(AArgCount: integer);
var
  Cost: double;
  Salvage: double;
  Life: double;
  Per: double;
  Factor: double;
begin
  Factor := 2;
  if AArgCount = 5 then
    if not PopResFloat(Factor) then Exit;

  if not PopResFloat(Per) then Exit;
  if not PopResFloat(Life) then Exit;
  if not PopResFloat(Salvage) then Exit;
  if not PopResFloat(Cost) then Exit;

  if (Salvage < 0) or (Cost < 0) or (Per <= 0) or (Life <= 0) or (Factor <= 0) or (Per > Life) then
    Push(errNum)
  else
    Push(DDB(Cost,Salvage,Life,Per,Factor));
end;

procedure TValueStack.DoDegrees;
var
  Angle: double;
begin
  if not PopResFloat(Angle) then Exit;

  Push(RadToDeg(Angle));
end;

procedure TValueStack.DoDevSq(AArgCount: integer);
var
  i: integer;
  Arr: TDynDoubleArray;
  Sum: double;
  Mean: double;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;

  if Length(Arr) <= 0 then
    Push(errDiv0)
  else begin
    Sum := 0;
    for i := 0 to High(Arr) do
      Sum := Sum + Arr[i];

    Mean := Sum / Length(Arr);
    Sum := 0;
    for i := 0 to High(Arr) do
      Sum := Sum + ((Arr[i] - Mean) * (Arr[i] - Mean));

    Push(Sum);
  end;
end;

procedure TValueStack.DoEDate;
var
  D: double;
  n: integer;
begin
  if not PopResInt(n) then
    Exit;
  if not PopResDateTime(D) then
    Exit;

  D := IncMonth(D,N);
  Push(Trunc(D));
end;

procedure TValueStack.DoEOMonth;
var
  D: double;
  n: integer;
begin
  if not PopResInt(n) then
    Exit;
  if not PopResDateTime(D) then
    Exit;

  D := IncMonth(D,N);
{$ifdef DELPHI_6_OR_LATER}
  D := EndOfTheMonth(D);
{$else}
  raise XLSRWException.Create('EOMONTH not supported by Delphi 5');
{$endif}
  Push(Trunc(D));
end;

procedure TValueStack.DoErfc;
var
  x: double;
begin
  if not PopResFloat(x) then Exit;

  Push(erfc(x));
end;

procedure TValueStack.DoErfc_Precise;
var
  x: double;
begin
  if not PopResFloat(x) then Exit;

  Push(erfc(x));
end;

procedure TValueStack.DoErf_Precise;
begin

end;

procedure TValueStack.DoError_Type;
var
  FV: TXLSFormulaValue;
begin
  FV := PopRes;
  if not (FV.RefType in [xfrtNone,xfrtRef]) or (FV.ValType <> xfvtError) then
    Push(errNA)
  else begin
    case FV.vErr of
      errUnknown    : Push(errNA);
      errNull       : Push(1);
      errDiv0       : Push(2);
      errValue      : Push(3);
      errRef        : Push(4);
      errName       : Push(5);
      errNum        : Push(6);
      errNA         : Push(7);
      errGettingData: Push(8);
    end;
  end;
end;

procedure TValueStack.DoExpondist;
begin
  DoExpon_Dist;
end;

procedure TValueStack.DoExpon_Dist;
var
  Cumulative: boolean;
  lambda: double;
  x: double;
begin
  if not PopResBoolean(Cumulative) then Exit;
  if not PopResFloat(lambda) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (lambda <= 0) then
    Push(errNum)
  else begin
    if Cumulative then
      Push(ExponentialDistCDF(x,lambda))
    else
      Push(ExponentialDistPDF(x,lambda));
  end;
end;

procedure TValueStack.DoExtractDateTime(AFuncId, AArgCount: integer);
var
  n: integer;
  DT: TDateTime;
  S: AxUCString;
  ReturnType: integer;
  YY,MM,DD,HH,NN,SS,MS: word;
begin
  ReturnType := 1;
  if (AArgCount = 2) and not PopResInt(ReturnType) then
    Exit;

  if PeekRes(0).ValType = xfvtString then begin
    PopResStr(S);
    if not TryStrToDateTime(S,DT) then begin
      Push(errValue);
      Exit;
    end;
  end
  else if not PopResFloat(Double(DT)) then
    Exit;

{$ifdef DELPHI_5}
  DecodeDate(DT,YY,MM,DD);
  DecodeTime(DT,HH,NN,SS,MS);
{$else}
  DecodeDateTime(DT,YY,MM,DD,HH,NN,SS,MS);
{$endif}

  case AFuncId of
    067: Push(DD);   // DAY
    068: Push(MM);   // MONTH
    069: Push(YY);   // YEAR
    070: begin       // WEEKDAY
{$ifdef DELPHI_6_OR_LATER}
      n := DayOfTheWeek(DT);
{$else}
      n := 1;
{$endif}
      case ReturnType of
        1: begin
          Inc(n);
          if n > 7 then
            n := 1;
        end;
        3: Dec(n);
      end;
      Push(n);
    end;
    071: Push(HH);   // HOUR
    072: Push(NN);   // MINUTE
    073: Push(SS);   // SECOND
  end;
end;

procedure TValueStack.DoFact;
var
  V: double;
  Err: TXc12CellError;
begin
  if not GetAsFloatNum(PopRes,V,Err) then
    Push(errValue)
  else if Err <> errUnknown then
    Push(errValue)
  else if (V < 0) or (V > MAX_QUICKFACTORIAL) then
    Push(errNum)
  else
    Push(QuickFactorial(Trunc(V)));
end;

procedure TValueStack.DoFDist;
begin
  DoF_dist_RT;
end;

procedure TValueStack.DoFInv;
begin
  DoF_Inv_RT;
end;

procedure TValueStack.DoFisher;
var
  x: double;
begin
  if not PopResFloat(x) then Exit;

  if (x <= -1) or (x >= 1) then
    Push(errNum)
  else begin
    x := 0.5 * Ln((1 + x) / (1 - x));
    Push(x);
  end;
end;

procedure TValueStack.DoFisherInv;
var
  y: double;
begin
  if not PopResFloat(y) then Exit;

  y := (Exp(2 * y) - 1) / (Exp(2 * y) + 1);
  Push(y);
end;

procedure TValueStack.DoFloor;
var
  significance: double;
  number: double;
begin
  if not PopResFloat(significance) then Exit;
  if not PopResFloat(number) then Exit;

  if (number > 0) and (significance < 0) then
    Push(errNum)
  else if significance = 0 then
    Push(errDiv0)
  else
    Push(Round(number / significance - 0.4) * significance);
end;

procedure TValueStack.DoFloor_Precise(AArgCount: integer);
var
  number: double;
  significance: double;
begin
  significance := 1;
  if AArgCount = 2 then
    if not PopResFloat(significance) then Exit;
  if not PopResFloat(number) then Exit;

  significance := Abs(significance);
  if significance = 0 then
    Push(0)
  else
    Push(Round(number / significance - 0.4) * significance);
end;

procedure TValueStack.DoForecast;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  x: double;
  vX: double;
  vY: double;
  a: double;
  b: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  Sum2: double;
  N: integer;
begin
  FVX := PopRes;
  FVY := PopRes;
  if not PopResFloat(x) then Exit;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    Sum2 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      vX := FIterator.Result.vFloat;
      vY := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((vX - MeanX) * (vY - MeanY));
      Sum2 := Sum2 + Sqr(vX - MeanX);
    end;
    if Sum2 <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    b := Sum1 / Sum2;
    a := MeanY - b * MeanX;
    Push(a + b * x);
  end;
end;

procedure TValueStack.DoFrequency;
var
  i      : integer;
  iInsert: integer;
  ResLen : integer;
  V1,V2  : double;
  csArray: TXLSCellsSource;
  csBin  : TXLSCellsSource;
  csRes  : TXLSArrayItem;
begin
  if not PopResCellsSource(csBin) then
    Exit;
  if not PopResCellsSource(csArray) then
    Exit;
  ResLen := csBin.Width * csBin.Height + 1;
  csRes := TXLSArrayItem.Create(1,ResLen);

  csArray.BeginIterate;
  while csArray.IterateNext do begin
    V1 := csArray.IterAsFloat;
    if not csArray.LastWasEmpty then begin
      if csArray.LastError <> errUnknown then begin
        Push(csArray.LastError);
        Exit;
      end;
      i := 0;
      iInsert := -1;
      csBin.BeginIterate;
      while csBin.IterateNext do begin
        V2 := csBin.IterAsFloat;
        if not csBin.LastWasEmpty then begin
          if csBin.LastError <> errUnknown then begin
            Push(errNA);
            Exit;
          end;
          if V2 >= V1 then begin
            iInsert := i;
            Break;
          end;
        end;
        Inc(i);
      end;
      if iInsert < 0 then
        csRes.AsFloat[0,ResLen - 1] := csRes.AsFloat[0,ResLen - 1] + 1
      else
        csRes.AsFloat[0,iInsert] := csRes.AsFloat[0,iInsert] + 1;
    end;
  end;
  Push(csRes);
end;

procedure TValueStack.DoFTest;
begin
  DoF_Test;
end;

procedure TValueStack.DoF_Dist;
var
  cumulative: boolean;
  deg_freedom2: integer;
  deg_freedom1: integer;
  x: double;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResPosInt(deg_freedom2,1) then Exit;
  if not PopResPosInt(deg_freedom1,1) then Exit;
  if not PopResFloat(x) then Exit;

  if x < 0 then
    Push(errNum)
  else begin
    if cumulative then
      Push(FisherFDistCDF(x,deg_freedom1,deg_freedom2))
    else
      Push(FisherFDistPDF(x,deg_freedom1,deg_freedom2));
  end;
end;

procedure TValueStack.DoF_dist_RT;
var
  deg_freedom2: integer;
  deg_freedom1: integer;
  x: double;
begin
  if not PopResPosInt(deg_freedom2,1) then Exit;
  if not PopResPosInt(deg_freedom1,1) then Exit;
  if not PopResFloat(x) then Exit;

  if x < 0 then
    Push(errNum)
  else
    Push(FDist(x,deg_freedom1,deg_freedom2));
end;

procedure TValueStack.DoF_Inv;
var
  deg_freedom2: integer;
  deg_freedom1: integer;
  probability: double;
begin
  if not PopResPosInt(deg_freedom2,1) then Exit;
  if not PopResPosInt(deg_freedom1,1) then Exit;
  if not PopResFloat(probability) then Exit;

  if (probability < 0) or (probability >= 1) or (deg_freedom1 < 1) or (deg_freedom2 < 1) then
    Push(errNum)
  else
    Push(FisherFDistQuantile(probability,deg_freedom1,deg_freedom2));
end;

procedure TValueStack.DoF_Inv_RT;
var
  deg_freedom2: integer;
  deg_freedom1: integer;
  probability: double;
begin
  if not PopResPosInt(deg_freedom2,1) then Exit;
  if not PopResPosInt(deg_freedom1,1) then Exit;
  if not PopResFloat(probability) then Exit;

  if (probability < 0) or (probability >= 1) or (deg_freedom1 < 1) or (deg_freedom2 < 1) then
    Push(errNum)
  else
    Push(FisherFDistQuantile(1 - probability,deg_freedom1,deg_freedom2));
end;

procedure TValueStack.DoF_Test;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  V: double;
  Count: integer;
  Sum1: double;
  SumSq1: double;
  Sum2: double;
  SumSq2: double;
  df: double;
  df1,df2: integer;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) <= 1 then
    Push(errDiv0)
  else begin
    Count := 0;
    Sum1 := 0;
    SumSq1:= 0;
    Sum2 := 0;
    SumSq2:= 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      V := FIterator.Result.vFloat;
      Sum1 := Sum1 + V;
      SumSq1 := SumSq1 + V * V;

      V := FIterator.Linked.Result.vFloat;
      Sum2 := Sum2 + V;
      SumSq2 := SumSq2 + V * V;

      Inc(Count);
    end;
    if Count <  2then begin
      Push(errNum);
      Exit;
    end;
    Sum1 := (SumSq1 - Sum1 * Sum1 / Count) / (Count - 1);
    Sum2 := (SumSq2 - Sum2 * Sum2 / Count) / (Count - 1);
    if (Sum1 = 0) or (Sum2 = 0) then begin
      Push(errNA);
      Exit;
    end;
    if (Sum1 > Sum2) then begin
      df := Sum1 / Sum2;
      df1 := Count - 1;
      df2 := Count - 1;
    end
    else begin
      df := Sum2 / Sum1;
      df1 := Count - 1;
      df2 := Count - 1;
    end;
    Push(2 * FDist(df,df1,df2));
  end;
end;

procedure TValueStack.DoGammadist;
begin
  DoGamma_Dist;
end;

procedure TValueStack.DoGammaInv;
begin
  DoGamma_Inv;
end;

procedure TValueStack.DoGammaLn;
var
  x: double;
begin
  if not PopResFloat(x) then Exit;

  if x <= 0 then
    Push(errNum)
  else
    Push(Ln(TGamma(x)));
end;

procedure TValueStack.DoGammaln_Precise;
var
  x: double;
begin
  if not PopResFloat(x) then Exit;

  if x <= 0 then
    Push(errNum)
  else
    Push(Ln(TGamma(x)));
end;

procedure TValueStack.DoGamma_Dist;
var
  Cumulative: boolean;
  beta: double;
  alpha: double;
  x: double;
begin
  if not PopResBoolean(Cumulative) then Exit;
  if not PopResFloat(beta) then Exit;
  if not PopResFloat(alpha) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (alpha < 0) or (beta < 0) then
    Push(errNum)
  else begin
    if x = 0 then
      Push(0)
    else if Cumulative then
      Push(GammaCDF(x,alpha,beta))
    else
      Push(GammaPDF(x,alpha,beta));
  end;
end;

procedure TValueStack.DoGamma_Inv;
var
  beta: double;
  alpha: double;
  probability: double;
begin
  if not PopResFloat(beta) then Exit;
  if not PopResFloat(alpha) then Exit;
  if not PopResFloat(probability) then Exit;

  if (probability < 0) or (probability >= 1) or (alpha <= 0) or (beta <= 0) then
    Push(errNum)
  else
    Push(GammaDistributionQuantile(alpha,beta,probability));
end;

procedure TValueStack.DoGeomean(AArgCount: integer);
var
  i: integer;
  Arr: TDynDoubleArray;
  Sum: double;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  if Length(Arr) <= 0 then
    Push(errNum)
  else begin
    Sum := 0;
    for i := 0 to High(Arr) do begin
      if Arr[i] <= 0 then begin
        Push(errNum);
        Exit;
      end;
      Sum := Sum + Ln(Arr[i]);
    end;
    Push(Exp(Sum / Length(Arr)));
  end;
end;

procedure TValueStack.DoGrowth(AArgCount: integer);
var
  i: integer;
  const_: boolean;
  new_x: TDynDoubleArray;
  known_x: TDynDoubleArray;
  known_y: TDynDoubleArray;
  Res: TXLSArrayItem;
begin
  if AArgCount = 4 then
    if not PopResBoolean(const_) then Exit;
  if AArgCount >= 3 then
    DoCollectValues(new_x);
  if AArgCount >= 2 then
    DoCollectValues(known_x);

  DoCollectValues(known_y);

  if AArgCount < 2 then begin
    SetLength(known_x,Length(known_y));
    for i := 0 to High(known_y) do
      known_x[i] := i + 1;
  end;
  if AArgCount < 3 then begin
    SetLength(new_x,Length(known_x));
    for i := 0 to High(known_x) do
      new_x[i] := known_x[i];
  end;

  if (Length(known_y) <> Length(known_x)) then
    Push(errRef)
  else if (Length(known_y) < 2) then
    Push(errNum)
  else begin
    Res := TXLSArrayItem.Create(1,Length(new_x));;
    for i := 0 to High(new_x) do
      Res.AsFloat[0,i] := ForecastExponential(new_x[i],known_y,known_x);
    Push(Res);
  end;
end;

procedure TValueStack.DoHarmean(AArgCount: integer);
var
  i: integer;
  Sum: double;
  Arr: TDynDoubleArray;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  if Length(Arr) <= 0 then
    Push(errNA)
  else begin
    Sum := 0;
    for i := 0 to High(Arr) do begin
      if Arr[i] <= 0 then begin
        Push(errNum);
        Exit;
      end;
      Sum := Sum + (1 / Arr[i]);
    end;
    Push(Length(Arr) / Sum);
  end;
end;

procedure TValueStack.DoHLookup(AArgCount: integer);
var
  C       : integer;
  FoundC  : integer;
  Range   : boolean;
  RowIndex: integer;
  Table   : TXLSFormulaValue;
  Lookup  : TXLSFormulaValue;
  FV      : TXLSFormulaValue;
begin
  Range := True;
  if AArgCount = 4 then
    if not PopResBoolean(Range) then Exit;
  if not PopResInt(RowIndex) then Exit;
  Table := PopRes;
  Lookup := PopRes;

  if (Table.RefType = xfrtNone) and (Table.ValType = xfvtError) then begin
    Push(Table.vErr);
    Exit;
  end;


  if not (Table.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) or (RowIndex <= 0) then begin
    Push(errValue);
    Exit;
  end;

  MakeIntersect(@Lookup);

  if Lookup.RefType in [xfrtArea,xfrtXArea,xfrtAreaList] then begin
    Push(errNA);
    Exit;
  end;

  Dec(RowIndex);

  if RowIndex > (Table.Row2 - Table.Row1) then begin
    Push(errREF);
    Exit;
  end;


  if Lookup.RefType = xfrtArray then
    Lookup := TXLSArrayItem(Lookup.vStr).GetAsFormulaValue(0,0);

  FoundC := -1;

  for C := Table.Col1 to Table.Col2 do begin
    GetCellValue(Table.Cells,C,Table.Row1,@FV);
    if FV.ValType = Lookup.ValType then begin
      if Range then begin
        case FV.ValType of
          xfvtFloat  : if Lookup.vFloat >= FV.vFloat then FoundC := C;
          xfvtBoolean: if Lookup.vBool = FV.vBool then FoundC := C;
          xfvtError  : if Lookup.vErr = FV.vErr then FoundC := C;
          xfvtString : if CompareText(FV.vStr,Lookup.vStr) <= 0 then FoundC := C;
        end;
      end
      else begin
        case FV.ValType of
          xfvtFloat  : if Lookup.vFloat = FV.vFloat then FoundC := C;
          xfvtBoolean: if Lookup.vBool = FV.vBool then FoundC := C;
          xfvtError  : if Lookup.vErr = FV.vErr then FoundC := C;
{$ifdef DELPHI_7_OR_LATER}
          xfvtString : if MatchesMask(FV.vStr,Lookup.vStr) then FoundC := C;
{$endif}          
        end;
      end;
    end;
  end;
  if FoundC >= 0 then begin
    GetCellValue(Table.Cells,FoundC,Table.Row1 + RowIndex,@FV);
    Push(FV);
  end
  else
    Push(errNA);
end;

procedure TValueStack.DoHyperlink(AArgCount: integer);
var
  sLink: AxUCString;
  sText: AxUCString;
begin
  if AArgCount = 2 then
    if not PopResStr(sText) then Exit;
  if not PopResStr(sLink) then Exit;
  if AArgCount = 1 then
    sText := sLink;

  Push(sText);
end;

procedure TValueStack.DoHypGeomDist;
begin
  DoHypgeom_Dist;
end;

procedure TValueStack.DoHypgeom_Dist;
var
  cumulative: boolean;
  sample_s: integer;
  number_sample: integer;
  population_s: integer;
  number_pop: integer;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResInt(number_pop) then Exit;
  if not PopResInt(population_s) then Exit;
  if not PopResInt(number_sample) then Exit;
  if not PopResInt(sample_s) then Exit;

  if (sample_s < 0) or (sample_s > number_sample) or (number_sample <= 0) or (number_sample > number_pop) or (population_s <= 0) or (population_s > number_pop) or (number_pop <= 0) then
    Push(errNum)
  else begin
    if cumulative then
      Push(HypergeometricDistCDF(sample_s,population_s,number_sample,number_pop))
    else
      Push(HypergeometricDistPDF(sample_s,population_s,number_sample,number_pop));
  end;
end;

procedure TValueStack.DoIferror;
var
  Val: TXLSFormulaValue;
  ValIfError: TXLSFormulaValue;
begin
  ValIfError := PopRes;
  Val := PopRes;

  if Val.ValType = xfvtError then
    Push(ValIfError)
  else
    Push(Val);
end;

procedure TValueStack.DoIndex(AArgCount: integer);
var
  CA: TCellAreas;
  Arr: TXLSArrayItem;
  fvSource: TXLSFormulaValue;
  Col,Row: integer;
  Area: integer;
begin
  Col := 0;
  Area := 0;
  case AArgCount of
    2: begin
      if not PopResInt(Row) then
        Exit;
      Dec(Row);
    end;
    3: begin
      if not PopResInt(Col) then
        Exit;
      if not PopResInt(Row) then
        Exit;
      Dec(Col);
      Dec(Row);
    end;
    4: begin
      if not PopResInt(Area) then
        Exit;
      if not PopResInt(Col) then
        Exit;
      if not PopResInt(Row) then
        Exit;
      Dec(Col);
      Dec(Row);
      Dec(Area);
    end;
  end;
  fvSource := PopRes;

  if (AArgCount = 2) and (fvSource.Row1 = fvSource.Row2) then begin
    Col := Row;
    Row := 0;
  end;

  if (Area < 0) or (Col > (fvSource.Col2 - fvSource.Col1)) or (Row > fvSource.Row2 - fvSource.Row1) then begin
    Push(errRef);
    Exit;
  end;

  case fvSource.RefType of
    xfrtNone,
    xfrtRef     : begin
      if (Col <= 0) and (Row <= 0) and (Area = 0) then
        Push(fvSource)
      else
        Push(errRef);
    end;
    xfrtArea    : begin
      if (Col < 0) and (Row < 0) then begin
        Push(fvSource);
        Exit;
      end
      else begin
        if (Col < 0) then begin
          if fvSource.Col1 = fvSource.Col2 then
            Push(fvSource.Cells,fvSource.Col1,fvSource.Row1)
          else
            Push(fvSource.Cells,fvSource.Col1,fvSource.Row1 + Row,fvSource.Col2,fvSource.Row1 + Row);
          Exit;
        end
        else if (Row < 0) then begin
          if fvSource.Row1 = fvSource.Row2 then
            Push(fvSource.Cells,fvSource.Col1,fvSource.Row1)
          else
            Push(fvSource.Cells,fvSource.Col1 + Col,fvSource.Row1,fvSource.Col1 + Col,fvSource.Row2);
          Exit;
        end;
        Push(fvSource.Cells,fvSource.Col1 + Col,fvSource.Row1 + Row)
      end;
    end;
    xfrtXArea   : ;
    xfrtAreaList: begin
      CA := TCellAreas(fvSource.vSource);
      if (Area < CA.Count) and (Col <= (CA[Area].Col2 - CA[Area].Col1)) and (Row <= (CA[Area].Row2 - CA[Area].Row1))then
        Push(TXLSCellMMU(CA[Area].Obj),CA[Area].Col1 + Col,CA[Area].Row1 + Row)
      else
        Push(errRef);
    end;
    xfrtArray   : begin
      Arr := TXLSArrayItem(fvSource.vSource);
      if (Col < Arr.Width) and (Row < Arr.Height) and (Area = 0) then
        Push(Arr.GetAsFormulaValue(Col,Row))
      else
        Push(errRef);
    end;
  end;
end;

procedure TValueStack.DoInfo;
var
  S: AxUCString;
begin
  if PopResStr(S,False) then begin
    S := Lowercase(S);
    if S = 'pig' then
      Push('Oink!')
    else
      Push(errNA);
//    if S := 'directory' then
//    else if S := 'numfile' then
//    else if S := 'origin' then
//    else if S := 'osversion' then
//    else if S := 'recalc' then
//    else if S := 'release' then
//    else if S := 'system' then


  end;
end;

procedure TValueStack.DoIntercept;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  vX: double;
  vY: double;
  b: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  Sum2: double;
  N: integer;
begin
  FVX := PopRes;
  FVY := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    Sum2 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      vX := FIterator.Result.vFloat;
      vY := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((vX - MeanX) * (vY - MeanY));
      Sum2 := Sum2 + Sqr(vX - MeanX);
    end;
    if Sum2 = 0 then begin
      Push(errDiv0);
      Exit;
    end;
    b := Sum1 / Sum2;
    Push(MeanY - b * MeanX);
  end;
end;

procedure TValueStack.DoIPMT(AArgCount: integer);
var
  Typ : integer;
  FV  : double;
  PV  : double;
  NPer: integer;
  Per : integer;
  Rate: double;
begin
  Typ := 0;
  FV := 0;
  if AArgCount = 6 then
    if not PopResPosInt(Typ) then Exit;
  if AArgCount >= 5 then
    if not PopResFloat(FV) then Exit;
  if not PopResFloat(PV) then Exit;
  if not PopResInt(NPer) then Exit;
  if not PopResInt(Per) then Exit;
  if not PopResFloat(Rate) then Exit;

  if (Per < 0) or (Per > NPer) then
    Push(errNum)
  else begin
    if Typ = 0 then
      Push(InterestPayment(Rate,Per,NPer,PV,FV,ptEndOfPeriod))
    else
      Push(InterestPayment(Rate,Per,NPer,PV,FV,ptStartOfPeriod));
  end;
end;

procedure TValueStack.DoIRR(AArgCount: integer);
var
  Vals: TDynDoubleArray;
  Guess: double;
begin
  Guess := 0.1;
  if AArgCount = 2 then
    if not PopResFloat(Guess) then Exit;

  if not DoCollectValues(Vals) then
    Exit
  else if Length(Vals) < 2 then
    Push(errNum)
  else
    Push(InternalRateOfReturn(Guess,Vals));
end;

procedure TValueStack.DoIso_Ceiling(AArgCount: integer);
var
  number: double;
  significance: double;
begin
  significance := 1;
  if AArgCount = 2 then
    if not PopResFloat(significance) then Exit;
  if not PopResFloat(number) then Exit;

  significance := Abs(significance);
  if significance = 0 then
    Push(0)
  else
    Push(Round(number / significance + 0.4) * significance);
end;

procedure TValueStack.DoISPMT;
var
  pv: double;
  nper: integer;
  per: integer;
  rate: double;
begin
  if not PopResFloat(pv) then Exit;
  if not PopResPosInt(nper,1) then Exit;
  if not PopResPosInt(per) then Exit;
  if not PopResFloat(rate) then Exit;

  Push(-((pv - (pv / nper) * per) * rate));
end;

procedure TValueStack.DoKurt(AArgCount: integer);
var
  i: integer;
  Len: integer;
  Sum1: double;
  Sum2: double;
  dx: double;
  SumP4: double;
  stdev: double;
  Mean: double;
  d: double;
  l: double;
  t: double;
  Arr: TDynDoubleArray;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  Len := Length(Arr);
  if Len <= 0 then
    Push(errNA)
  else begin
    Sum1 := 0;
    for i := 0 to High(Arr) do
      Sum1 := Sum1 + Arr[i];
    Mean := Sum1 / Len;

    Sum2 := 0;
    for i := 0 to High(Arr) do
      Sum2 := Sum2 + (Arr[i] - Mean) * (Arr[i] - Mean);
    stdev := Sqrt(Sum2 / (Len - 1));
    if stdev = 0 then begin
      Push(errDiv0);
      Exit;
    end;

    SumP4 := 0;
    for i := 0 to High(Arr) do begin
      dx := (Arr[i] - Mean) / stdev;
      SumP4 := SumP4 + (dx * dx * dx * dx);
    end;

    d := (Len - 2) * (Len - 3);
    l := Len * (Len + 1) / ((Len - 1) * d);
    t := 3 * (Len - 1) * (Len - 1) / d;

    Push(SumP4 * l - t);
  end;
end;

procedure TValueStack.DoLarge;
var
  Arr: TDynDoubleArray;
  k: integer;
begin
  if not PopResPosInt(k,1) then Exit;
  if not DoCollectValues(Arr) then
    Exit;
  SortDoubleArray(Arr);
  if k > Length(Arr) then
    Push(errNum)
  else
    Push(Arr[Length(Arr) - k]);
end;

procedure TValueStack.DoLinest(AArgCount: integer);
var
  i: integer;
  stats: boolean;
  const_: boolean;
  known_x: TDynDoubleArray;
  known_y: TDynDoubleArray;
  LED: TLinEstData;
  Res: TXLSArrayItem;
begin
  if AArgCount = 4 then
    if not PopResBoolean(stats) then Exit;
  if AArgCount >= 3 then
    if not PopResBoolean(const_) then Exit;
  if AArgCount >= 2 then
    DoCollectValues(known_x);

  DoCollectValues(known_y);

  if AArgCount = 1 then begin
    SetLength(known_x,Length(known_y));
    for i := 0 to High(known_y) do
      known_x[i] := i + 1;
  end;

  if Length(known_x) <> Length(known_y) then
    Push(errRef)
  else if (Length(known_y) < 2) then
    Push(errNum)
  else begin
    if not LinEst(known_y,known_x,LED,True) then
      Push(errNum)
    else begin
      Res := TXLSArrayItem.Create(2,5);
      Res.AsFloat[0,0] := LED.B1;
      Res.AsFloat[1,0] := LED.B0;
      Res.AsFloat[0,1] := LED.seB1;
      Res.AsFloat[1,1] := LED.seB0;
      Res.AsFloat[0,2] := LED.R2;
      Res.AsFloat[1,2] := LED.sigma;
      Res.AsFloat[0,3] := LED.F0;
      Res.AsFloat[1,3] := LED.df;
      Res.AsFloat[0,4] := LED.SSr;
      Res.AsFloat[1,4] := LED.SSe;
      Push(Res);
    end;
  end;
end;

procedure TValueStack.DoLogEst(AArgCount: integer);
var
  i: integer;
  stats: boolean;
  const_: boolean;
  known_x: TDynDoubleArray;
  known_y: TDynDoubleArray;
  LED: TLinEstData;
  Res: TXLSArrayItem;
begin
  if AArgCount = 4 then
    if not PopResBoolean(stats) then Exit;
  if AArgCount >= 3 then
    if not PopResBoolean(const_) then Exit;
  if AArgCount >= 2 then
    DoCollectValues(known_x);

  DoCollectValues(known_y);

  if AArgCount = 1 then begin
    SetLength(known_x,Length(known_y));
    for i := 0 to High(known_y) do
      known_x[i] := i + 1;
  end;

  if Length(known_x) <> Length(known_y) then
    Push(errRef)
  else begin
    LogEst(known_y,known_x,LED,True);

    Res := TXLSArrayItem.Create(2,5);
    Res.AsFloat[0,0] := LED.B1;
    Res.AsFloat[1,0] := LED.B0;
    Res.AsFloat[0,1] := LED.seB1;
    Res.AsFloat[1,1] := LED.seB0;
    Res.AsFloat[0,2] := LED.R2;
    Res.AsFloat[1,2] := LED.sigma;
    Res.AsFloat[0,3] := LED.F0;
    Res.AsFloat[1,3] := LED.df;
    Res.AsFloat[0,4] := LED.SSr;
    Res.AsFloat[1,4] := LED.SSe;
    Push(Res);
  end;
end;

procedure TValueStack.DoLoginv;
begin
  DoLognorm_Inv;
end;

procedure TValueStack.DoLognormdist;
begin
  // cumulative
  Push(True);
  DoLognorm_dist;
end;

procedure TValueStack.DoLognorm_dist;
var
  cumulative: boolean;
  stdev: double;
  mean: double;
  x,x2: double;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResFloat(stdev) then Exit;
  if not PopResFloat(mean) then Exit;
  if not PopResFloat(x) then Exit;

  if (stdev <= 0) or (x <= 0) then
    Push(errNum)
  else begin
    if cumulative then
      Push(0.5 * erfc(-((ln(x) - mean) / stdev) * 0.7071067811865475))
    else begin
      x2 := (ln(x) - mean) / stdev;
      Push((0.39894228040143268 * exp(-(x2 * x2) / 2.0)) / stdev / x);
    end;
  end;
end;

procedure TValueStack.DoLognorm_Inv;
var
  standard_dev: double;
  mean: double;
  Probability: double;
begin
  if not PopResFloat(standard_dev) then Exit;
  if not PopResFloat(mean) then Exit;
  if not PopResFloat(Probability) then Exit;

  if (Probability <= 0) or (Probability >= 1) or (standard_dev <= 0) then
    Push(errNum)
  else
    Push(LogNormalQuantile(Probability,mean,standard_dev));
end;

procedure TValueStack.DoLookup(AArgCount: integer);
var
  i,j: integer;
  fvValue: TXLSFormulaValue;
  fvLookup: TXLSFormulaValue;
  fvResult: TXLSFormulaValue;
begin
  if AArgCount = 3 then begin
    fvResult := PopRes;
    if not IsVector(@fvResult) then begin
      if fvResult.ValType = xfvtError then
        Push(fvResult.vErr)
      else
        Push(errNA);
      Exit;
    end;
  end;
  fvLookup := PopRes;
  fvValue := PopRes;
  if fvValue.ValType = xfvtError then begin
    Push(fvValue.vErr);
    Exit;
  end;

  MakeIntersect(@fvValue);

  if fvValue.RefType in [xfrtArea,xfrtXArea,xfrtAreaList] then begin
    Push(errNA);
    Exit;
  end;

  if fvValue.RefType = xfrtArray then
    fvValue := TXLSArrayItem(fvValue.vSource).GetAsFormulaValue(0,0);

  if AArgCount = 2 then begin
    fvResult := fvLookup;
    if fvResult.RefType in [xfrtArea,xfrtXArea] then
      MakeVector(@fvResult,False);
  end;
  if fvLookup.RefType in [xfrtArea,xfrtXArea] then
    MakeVector(@fvLookup,True);

  i := 0;
  j := -1;
  FIterator.BeginIterate(fvLookup,False);
  while FIterator.Next do begin
    if FIterator.Result.ValType = fvValue.ValType then begin
      case fvValue.ValType of
        xfvtFloat   : begin
          if FIterator.Result.vFloat <= fvValue.vFloat then
            j := i;
        end;
        xfvtBoolean : begin
          if FIterator.Result.vBool = fvValue.vBool then
            j := i;
        end;
        xfvtString  : begin
          if CompareText(FIterator.Result.vStr,fvValue.vStr) <= 0 then
            j := i;
        end;
      end;
    end;
    Inc(i);
  end;

  if j >= 0 then begin
    case fvResult.RefType of
      xfrtNone    : ;
      xfrtRef     : ;
      xfrtArea    : begin
        if fvResult.Col1 = fvResult.Col2 then
          GetCellValue(fvResult.Cells,fvResult.Col1,fvResult.Row1 + j,@fvResult)
        else
          GetCellValue(fvResult.Cells,fvResult.Col1 + j,fvResult.Row1,@fvResult);
        Push(fvResult);
      end;
      xfrtXArea   : ;
      xfrtAreaList: ;
      xfrtArray   : begin
        // TODO. Check that two dimensional works.
        if j < TXLSArrayItem(fvResult.vSource).Width then
          Push(TXLSArrayItem(fvResult.vSource).GetAsFormulaValue(j,0))
        else
          Push(errNA);
      end;
    end;
  end
  else
    Push(errNA);
end;

procedure TValueStack.DoMatch(AArgCount: integer);
var
  iLast: integer;
  MatchType: integer;
  FV: TXLSFormulaValue;
  LookupVal: TXLSFormulaValue;
  IsFirst: boolean;

function ResultIndex: integer;
begin
  if FV.RefType = xfrtArea then begin
    if FV.Col1 = FV.Col2 then
      Result := FIterator.Row - FV.Row1 + 1
    else
      Result := FIterator.Col - FV.Col1 + 1;
  end
  else if FV.RefType = xfrtArray then begin
    if TXLSArrayItem(FV.vSource).Width = 1 then
      Result := FIterator.Row
    else
      Result := FIterator.Col;
  end
  else // Single cell ref
    Result := 0;
end;

begin
  MatchType := 1;
  if AArgCount = 3 then
    if not PopResInt(MatchType) then Exit;

  FV := PopRes;
  if FV.ValType = xfvtError then begin
    Push(FV);
    Exit;
  end;

  LookupVal := PopRes;
  if LookupVal.ValType = xfvtError then begin
    Push(LookupVal);
    Exit;
  end;
  MakeIntersect(@LookupVal);

  if not IsVector(@FV) then begin
    Push(errNA);
    Exit;
  end;

  IsFirst := True;
  iLast := -1;
  FIterator.BeginIterate(FV,False);
  while FIterator.NextNotEmpty do begin
    if LookupVal.ValType = FIterator.Result.ValType then begin
      case MatchType of
       -1: begin
          iLast := ResultIndex;
          if IsFirst then begin
            if CompareFVValue(@LookupVal,@FIterator.Result,xdcoGT) then begin
              Push(errNA);
              Exit;
            end;
            IsFirst := False;
          end;
          if CompareFVValue(@FIterator.Result,@LookupVal,xdcoLT) then begin
            if not CompareFVValue(@FIterator.Result,@LookupVal,xdcoEQ) then
              Push(ResultIndex - 1)
            else
              Push(ResultIndex);
            Exit;
          end;
        end;
        0: begin
          if CompareFVValue(@FIterator.Result,@LookupVal,xdcoEQ) then begin
            Push(ResultIndex);
            Exit;
          end;
        end;
        1: begin
          iLast := ResultIndex;
          if IsFirst then begin
            if CompareFVValue(@LookupVal,@FIterator.Result,xdcoLT) then begin
              Push(errNA);
              Exit;
            end;
            IsFirst := False;
          end;
          if CompareFVValue(@FIterator.Result,@LookupVal,xdcoGT) then begin
            if not CompareFVValue(@FIterator.Result,@LookupVal,xdcoEQ) then
              Push(ResultIndex - 1)
            else
              Push(ResultIndex);
            Exit;
          end;
        end;
      end;
    end;
  end;

  if ((MatchType = -1) or (MatchType = 1)) and (iLast >= 0) then
    Push(iLast)
  else
    Push(errNA);
end;

procedure TValueStack.DoMax(AArgCount: integer);
var
  n: integer;
  vMax: double;
begin
  n := 0;
  vMax := MinDouble;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloat do begin
      Inc(n);
      if FIterator.Result.vFloat > vMax then
        vMax := FIterator.Result.vFloat;
    end;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  if n = 0 then
    Push(0)
  else
    Push(vMax);
end;

procedure TValueStack.DoMaxa(AArgCount: integer);
var
  n: integer;
  vMax: double;
begin
  n := 0;
  vMax := MinDouble;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloatA do begin
      Inc(n);
      if FIterator.Result.vFloat > vMax then
        vMax := FIterator.AsFloatA;
    end;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  if n = 0 then
    Push(0)
  else
    Push(vMax);
end;

procedure TValueStack.DoMDeterm;
var
  FV: TXLSFormulaValue;
  V: double;
  M: TXLSMatrix;
begin
  FV := PopRes;
  if not (FV.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) then begin
    Push(errValue);
    Exit;
  end;
  // TODO arrays
  if FV.RefType = xfrtArray then begin
    Push(errNA);
    Exit;
  end
  else begin
    M := TXLSMatrix.Create(FV.Col2 - FV.Col1 + 1,FV.Row2 - FV.Row1 + 1);
    try
      if not PopulateMatrix(M,@FV) then
        Exit;
      if M.Determ(V) then
        Push(V)
      else
        Push(errValue);
    finally
      M.Free;
    end;
  end;
end;

procedure TValueStack.DoMedian(AArgCount: integer);
var
  L: integer;
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValues(AArgCount,Vals) then begin
    L := Length(Vals);
    if L <= 0 then
      Push(errNum)
    else begin
      SortDoubleArray(Vals);
      if Odd(L) then
        Push(Vals[L div 2])
      else if L > 1 then
        Push((Vals[L div 2 - 1] + Vals[L div 2]) / 2)
      else
      Push(Vals[0]);
    end;
  end;
end;

procedure TValueStack.DoMin(AArgCount: integer);
var
  n: integer;
  vMin: double;
begin
  n := 0;
  vMin := MaxDouble;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloat do begin
      Inc(n);
      if FIterator.Result.vFloat < vMin then
        vMin := FIterator.Result.vFloat;
    end;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  if n = 0 then
    Push(0)
  else
    Push(vMin);
end;

procedure TValueStack.DoMina(AArgCount: integer);
var
  n: integer;
  vMin: double;
begin
  n := 0;
  vMin := MaxDouble;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloatA do begin
      Inc(n);
      if FIterator.Result.vFloat < vMin then
        vMin := FIterator.AsFloatA;
    end;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  if n = 0 then
    Push(0)
  else
    Push(vMin);
end;

procedure TValueStack.DoMInverse;
var
  FV: TXLSFormulaValue;
  M,MI: TXLSMatrix;
begin
  FV := PopRes;
  if not (FV.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) then
    Push(errValue)
  // TODO arrays
  else if FV.RefType = xfrtArray then
    Push(errNA)
  else begin
    M := TXLSMatrix.Create(FV.Col2 - FV.Col1 + 1,FV.Row2 - FV.Row1 + 1);
    try
      if not PopulateMatrix(M,@FV) then
        Exit;
      MI := M.Inverse;
      try
        Push(MI);
      finally
        MI.Free;
      end;
    finally
      M.Free;
    end;
  end;
end;

procedure TValueStack.DoMIRR;
var
  i: integer;
  L: double;
  V: double;
  Values: TDynDoubleArray;
  FinanceRate: double;
  ReinvestRate: double;
  npvPos,npvNeg: double;
  Res: double;
begin
  if not PopResFloat(ReinvestRate) then
    Exit;
  if not PopResFloat(FinanceRate) then
    Exit;

  if not DoCollectValues(Values) then
    Exit
  else if Length(Values) < 2 then
    Push(errValue)
  else begin
    L := Length(Values);

    npvPos := 0.0;
    npvNeg := 0.0;
    for i := 0 to High(Values) do begin
      V := Values[i];
      if (V > 0.0) then
        npvPos := npvPos + V / Power(1.0 + ReinvestRate, i + 1.0)
      else
        npvNeg := npvNeg + V / Power(1.0 + FinanceRate, i + 1.0);
    end;
    ReinvestRate := Power(1.0 + ReinvestRate,L);
    FinanceRate := 1.0 + FinanceRate;
    Res := Power(-npvPos * ReinvestRate / (npvNeg * FinanceRate), 1.0 / (L - 1.0)) - 1.0;

    Push(Res);
  end;
end;

procedure TValueStack.DoMMult;
var
  FV1,FV2: TXLSFormulaValue;
  M1,M2,MRes: TXLSMatrix;
begin
  FV1 := PopRes;
  FV2 := PopRes;
  if not (FV1.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) or not (FV2.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) then
    Push(errValue)
  // TODO arrays
  else if (FV1.RefType = xfrtArray) or (FV2.RefType = xfrtArray) then
    Push(errNA)
  else begin
    M1 := TXLSMatrix.Create(FV1.Col2 - FV1.Col1 + 1,FV1.Row2 - FV1.Row1 + 1);
    try
      M2 := TXLSMatrix.Create(FV2.Col2 - FV2.Col1 + 1,FV2.Row2 - FV2.Row1 + 1);
      try
        if not PopulateMatrix(M1,@FV1) then
          Exit;
        if not PopulateMatrix(M2,@FV2) then
          Exit;
        MRes := M1.Mult(M2);
        if MRes = Nil then begin
          Push(errValue);
          Exit;
        end;
        try
          Push(MRes);
        finally
          MRes.Free;
        end;
      finally
        M2.Free;
      end;
    finally
      M1.Free;
    end;
  end;
end;

procedure TValueStack.DoMode(AArgCount: integer);
var
  i: integer;
  L: integer;
  Arr: TDynDoubleArray;
  last: double;
  max: double;
  f: integer;
  maxf: double;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  L := Length(Arr);
  if L < 0 then begin
    Push(errNum);
    Exit;
  end;
  SortDoubleArray(Arr);

  last := Arr[0];
  max := last;
  f := 1;
  maxf := f;

  for i := 1 to L - 1 do begin
    if Arr[i] = last then
      Inc(f)
    else begin
      if f > maxf then begin
        max := last;
        maxf := f;
      end;
      last := Arr[i];
      f := 1;
    end;
  end;

  if maxf <= 1 then
    Push(errNA)
  else begin
    if f > maxf then
      max := last;

    Push(max);
  end;
end;

procedure TValueStack.DoMode_Mult(AArgCount: integer);
var
  i,j,k: integer;
  L: integer;
  V: double;
  Arr: TDynDoubleArray;
  Res: TXLSArrayItem;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  L := Length(Arr);
  if L < 0 then begin
    Push(errNum);
    Exit;
  end;
  L := Length(Arr);
  SortDoubleArray(Arr);
  Res := TXLSArrayItem.Create(1,L div 2);

  i := 0;
  k := 0;
  while i < L do begin
    V := Arr[i];
    j := i + 1;
    while (j < L) and (V = Arr[j]) do
      Inc(j);
    if j > (i + 1) then begin
      Res.AsFloat[0,k] := V;
      Inc(k);
    end;
    i := j;
  end;
  if k <= 0 then begin
    Res.Free;
    Push(errNA);
  end
  else
    Push(Res);
end;

procedure TValueStack.DoMode_Sngl(AArgCount: integer);
begin

end;

procedure TValueStack.DoNegbinomdist;
begin
  // cumulative
  Push(True);
  DoNegbinom_Dist;
end;

procedure TValueStack.DoNegbinom_Dist;
var
  cumulative: boolean;
  Probability_s: double;
  Number_s: double;
  Number_f: double;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResFloat(Probability_s) then Exit;
  if not PopResFloat(Number_s) then Exit;
  if not PopResFloat(Number_f) then Exit;

  if (Number_s < 1) or (Number_f <0) or (Probability_s < 0) or (Probability_s > 1) then
    Push(errNum)
  else begin
    if cumulative then
      Push(NegBinomialCDF(Number_f,Number_s,Probability_s))
    else
      Push(NegBinomialPDF(Number_f,Number_s,Probability_s));
  end;
end;

//Public Function WorkingDays(StartDate As Variant, EndDate As Variant, Optional Holidays As Variant) As Variant
//
//On Error GoTo exitclean            ' make sure we exit the routine cleanly
//Application.EnableEvents = False   ' Turn events off for speed
// '=========================================================================================
//Dim WholeDays As Double            ' Total days inclusive of both containing days
//Dim DaysFirstWeek As Double        ' Number of days till the first saturday
//Dim DaysSecondWeek As Double       ' number of days in last week till Saturday
//Dim Weeks As Double                ' The number of whole weeks beginning with saturday
//Dim DaysToAdd As Double            ' How many days in the 2 partial weeks
//Dim MyDay As Variant               ' Used in calculating the holidays
//Dim Day As Date                    ' Used in check for holiday dates
//Dim Swap As Boolean                ' Flag incase Dates are Other order
//Dim LocalStartDate As Date         ' Incase dates are reversed
//Dim LocalEndDate As Date           ' Incase dates are reversed
//Dim TempStartDate As Date          'used to take care of the #num! for overflows
//Dim TempEndDate As Date
// Dim HolidayCount As Long
// '=========================================================================================
//' If the parameters are EndDate is less than start date
//' then swap the calculations and flag to negate at end
//
//
//' Tests for both parameters being blank cells for Network Days compatibility
//
//If IsEmpty(StartDate) And IsEmpty(EndDate) Then
//   WorkingDays = 0
//  Application.EnableEvents = True
//  Exit Function
//End If
//
//'if either date is 0 or blank then set to day 0 (#12/31/1899#)
//If IsEmpty(StartDate) Or StartDate = 0 Then StartDate = #12/31/1899#
//If IsEmpty(EndDate) Or EndDate = 0 Then EndDate = #12/31/1899#
//
//' this will create an error if the values arn't valid dates!
//TempStartDate = StartDate
//TempEndDate = EndDate
//
//' dont allow dates before 1/1/1900
//If TempStartDate < #12/31/1899# Or TempEndDate < #12/31/1899# Then
//    WorkingDays = CVErr(xlErrNum)
//    Application.EnableEvents = True
//   Exit Function
//End If
//
//' allow for the case where the start date is less than the end date
//If TempStartDate > TempEndDate Then
//    Swap = True
//    LocalStartDate = TempEndDate
//    LocalEndDate = TempStartDate
//Else
//    Swap = False
//    LocalStartDate = TempStartDate
//    LocalEndDate = EndDate
//End If
// '=========================================================================================
//'Calculate total days between dates inclusive
// WholeDays = LocalEndDate - LocalStartDate + 1
//
//' This will calculate the number of days leading up to the first saturday
//DaysFirstWeek = (7 - (LocalStartDate Mod 7))               ' day of week 0-sat,6-fri
//DaysFirstWeek = DaysFirstWeek * -(DaysFirstWeek < 7)  ' if its 7 then its a full week, so ignore
//
//' days in second week
//DaysSecondWeek = (LocalEndDate Mod 7)
//' if we subtract from total the two values above we should be left with whole weeks starting sat-Fri
//Weeks = Int((WholeDays - DaysFirstWeek - DaysSecondWeek) / 7)
//
//' Calculates the number of days in the end partial week
//DaysToAdd = ((DaysSecondWeek - 1) * -(DaysSecondWeek > 1))
//DaysToAdd = DaysToAdd + (DaysFirstWeek * -(DaysFirstWeek < 6)) + (-5 * (DaysFirstWeek = 6))
//' if the days in the firs week = 6,7 then ignore as they are weekends
//' so full weeks*5 = number of working days+Days in second week + days in first week
//' this gives us the total weekdays between the two dates
//WorkingDays = DaysToAdd + (Weeks * 5)
//'=======================================================================================
//' Check for holidays
//' if a holiday is sat or sun ignore it
//
//If Not IsMissing(Holidays) Then        ' Ignore if we have no parameter save time
//
//          ' Number of holidays we have encountered
//    HolidayCount = 0                   ' initialise holiday count
//
//    For Each MyDay In Holidays
//
//        Day = MyDay
//        If (Day >= LocalStartDate And Day <= LocalEndDate) And (Day Mod 7) > 1 Then HolidayCount = HolidayCount + 1
//
//    Next
//
//
//   WorkingDays = WorkingDays - HolidayCount ' Adjust for holidays
//End If
//'==========================================================================================
//   ' If LocalStartDate is bigget than end date figure should be negative
//   If Swap = True Then
//         WorkingDays = -WorkingDays
//   End If
// '=========================================================================================
//   Application.EnableEvents = True          ' Turn Events back on
//   On Error GoTo 0
//   Exit Function                            ' Exit Function
//
//'===========================================================================================
//' Error handler to make sure we turn events back on if there are any problems
//exitclean:
//   ' Take care of special cases in the function for compatability with networkdays
//    Select Case Err.Number
//        Case 6
//                WorkingDays = CVErr(xlErrNum)
//        Case 13
//                If VarType(StartDate) = vbString Then
//                    WorkingDays = CVErr(xlErrValue)
//                Else
//                    WorkingDays = CVErr(xlErrNum)
//                End If
//
//   End Select
//
//    Application.EnableEvents = True
//
//    On Error GoTo 0
//
//End Function

procedure TValueStack.DoNetWorkdays(AArgCount: integer);
var
  holidays: TDynDoubleArray;
  days: integer;
  start_date: double;
  end_date: double;
  Res: integer;
begin
  if AArgCount = 3 then
    if not DoCollectValues(holidays) then Exit;
  if not PopResInt(days) then Exit;
  if not PopResDateTime(start_date) then Exit;

  end_date := start_date + days;
  if (end_date < 1) or (end_date > XLSFMLA_MAXDATE) then
    Push(errValue)
  else begin
{$ifdef DELPHI_5}
    raise XLSRWException.Create('NETWORKDAYS not supported by Delphi 5');
{$else}
    Res := WeeksBetween(start_date,end_date) * 2;
{$endif}

    Push(Res);
  end;
end;

procedure TValueStack.DoNetworkdays_Intl(AArgCount: integer);
begin
  Push(errNA);
end;

procedure TValueStack.DoNormdist;
begin
  DoNorm_Dist;
end;

procedure TValueStack.DoNormInv;
begin
  DoNorm_Inv;
end;

procedure TValueStack.DoNormsdist;
begin
  DoNorm_S_Dist;
end;

procedure TValueStack.DoNormsInv;
begin
  DoNorm_S_Inv;
end;

procedure TValueStack.DoNorm_S_Dist;
var
  z: double;
begin
  if not PopResFloat(z) then Exit;

  Push(NormalCDF(z,0,1));
end;

procedure TValueStack.DoNorm_S_Inv;
var
  Probability: double;
begin
  if not PopResFloat(Probability) then Exit;
  if (Probability <= 0) or (Probability >= 1) then
    Push(errNum)
  else
    Push(NormalQuantile(Probability,0,1));
end;

procedure TValueStack.DoNorm_Dist;
var
  cumulative: boolean;
  stdev: double;
  mean: double;
  x: double;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResFloat(stdev) then Exit;
  if not PopResFloat(mean) then Exit;
  if not PopResFloat(x) then Exit;

  if stdev < 0 then
    Push(errNum)
  else begin
    if cumulative then
      Push(NormalCDF(x,mean,stdev))
    else
      Push(NormalPDF(x,mean,stdev));
  end;
end;

procedure TValueStack.DoNorm_Inv;
var
  standard_dev: double;
  mean: double;
  Probability: double;
begin
  if not PopResFloat(standard_dev) then Exit;
  if not PopResFloat(mean) then Exit;
  if not PopResFloat(Probability) then Exit;

  if (Probability < 0) or (Probability > 1) or (standard_dev <= 0) then
    Push(errNum)
  else
    Push(NormalQuantile(Probability,mean,standard_dev));
end;

function TValueStack.DoNPV(AArgCount: integer; out AResult: double): boolean;
var
  Vals: TDynDoubleArray;
  Rate: double;
  Res: double;
begin
  Result := False;
  if DoCollectAllValues(AArgCount - 1,Vals) then begin
    if not PopResFloat(Rate) then
      Exit;

    Res := NPV(Rate,Vals);

    Push(Res);
    Result := True;
  end;
end;

procedure TValueStack.DoOdd;
var
  number: double;
begin
  if not PopResFloat(number) then Exit;

  if (number = Trunc(number)) and Odd(Trunc(number)) then
    Push(number)
  else if number < 0 then
    Push(Floor(number / 2) * 2 - 1)
  else
    Push(Ceil(number / 2) * 2 + 1);
end;

procedure TValueStack.DoOffset(AArgCount: integer);
var
  FV: TXLSFormulaValue;
  fvRes: TXLSFormulaValue;
  vWidth,vHeight: integer;
  vCols,vRows: integer;
begin
  vWidth := 1;
  vHeight := 1;
  if AArgCount = 5 then
    if not PopResInt(vWidth) then Exit;
  if AArgCount >= 4 then
    if not PopResInt(vHeight) then Exit;
  if not PopResInt(vCols) then Exit;
  if not PopResInt(vRows) then Exit;

  FV := PopRes;

  if not (FV.RefType in [xfrtRef,xfrtArea,xfrtXArea]) then begin
    Push(errValue);
    Exit;
  end;

  if (vWidth = 0) or (vHeight = 0) then begin
    Push(errRef);
    Exit;
  end;

  fvRes.Col1 := FV.Col1 + vCols;
  fvRes.Row1 := FV.Row1 + vRows;
  fvRes.Col2 := FV.Col2 + vCols;
  fvRes.Row2 := FV.Row2 + vRows;

  if vWidth <> 1 then begin
    if vWidth > 1 then
      fvRes.Col2 := fvRes.Col2 + vWidth - 1
    else if vWidth < 0 then
      fvRes.Col1 := fvRes.Col1 + vWidth + 1
  end;

  if vHeight <> 1 then begin
    if vHeight > 1 then
      fvRes.Row2 := fvRes.Row2 + vHeight - 1
    else if vHeight < 0 then
      fvRes.Row1 := fvRes.Row1 + vHeight + 1
  end;

  if not InsideExtent(fvRes.Col1,fvRes.Row1,fvRes.Col2,fvRes.Row2) then begin
    Push(errRef);
    Exit;
  end;

  fvRes.Cells := FV.Cells;
  if (fvRes.Col1 = fvRes.Col2) and (fvRes.Row1 = fvRes.Row2) then begin
    GetCellValue(fvRes.Cells,fvRes.Col1,fvRes.Row1,@fvRes);
    fvRes.RefType := xfrtRef;
  end
  else
    fvRes.RefType := xfrtArea;

  Push(fvRes);
end;

procedure TValueStack.DoPercentile;
begin
  DoPercentile_Inc;
end;

procedure TValueStack.DoPercentile_Exc;
begin
  Push(errNA);
end;

procedure TValueStack.DoPercentile_Inc;
var
  k: double;
  Res: double;
begin
  if not PopResFloat(k) then Exit;
  if (k < 0) or (k > 1) then begin
    PopRes;
    Push(errNum);
  end
  else begin
    if not Percentile(k,Res) then
      Push(errNum)
    else
      Push(Res);
  end;
end;

procedure TValueStack.DoPercentRank(AArgCount: integer);
begin
  DoPercentrank_Inc(AArgCount);
end;

procedure TValueStack.DoPercentrank_Exc(AArgCount: integer);
begin
  Push(errNA);
end;

procedure TValueStack.DoPercentrank_Inc(AArgCount: integer);
var
  L: integer;
  significance: double;
  x: double;
  b: integer;
  t: integer;
  m: integer;
  Arr: TDynDoubleArray;
begin
  if AArgCount = 3 then
    if not PopResFloat(significance) then Exit;
  if not PopResFloat(x) then Exit;

  if not DoCollectValues(Arr) then
    Exit;
  L := Length(Arr);
  if L < 0 then begin
    Push(errNum);
    Exit;
  end;
  SortDoubleArray(Arr);

  if (x < Arr[0]) or (x > Arr[L - 1]) then begin
    Push(errNA);
    Exit;
  end;

  b := 0;
  t := L - 1;
  while (t - b) > 1 do begin
    m := (b + t) shr 1;
    if (x >= Arr[m]) then
      b := m
    else
      t := m;
  end;

  while (b > 0) and (Arr[b - 1] = x) do
    Dec(b);

  if L = 1 then
    Push(x)
  else if (Arr[b] = x) then
    Push(b / (L - 1))
  else
    Push((b + (X - Arr[b]) / (Arr[b + 1] - Arr[b])) / (L - 1));
end;

procedure TValueStack.DoPermut;
var
  number: integer;
  number_choosen: integer;
begin
  if not PopResPosInt(number_choosen) then Exit;
  if not PopResPosInt(number) then Exit;

  if number_choosen = 0 then
    Push(1)
  else if (number <= 0) or (number_choosen > Number) then
    Push(errNum)
  else if number > MAX_QUICKFACTORIAL then
    Push(errNum)
  else begin
    Push(QuickFactorial(number) / QuickFactorial(number - number_choosen))
  end;
end;

procedure TValueStack.DoPhonetic;
begin
//  Do nothing. Best value is on stack.
end;

procedure TValueStack.DoPoisson;
begin
  DoPoisson_Dist;
end;

procedure TValueStack.DoPoisson_Dist;
var
  x: double;
  mean: double;
  cumulative: boolean;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResFloat(mean) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (mean < 0) then
    Push(errNum)
  else begin
    if cumulative then
      Push(PoissonCDF(x,mean))
    else
      Push(PoissonPDF(x,mean));
  end;
end;

procedure TValueStack.DoPower;
var
  P: double;
  N: double;
begin
  if not PopResFloat(P) then Exit;
  if not PopResFloat(N) then Exit;

  Push(Power(N,P));
end;

procedure TValueStack.DoPPMT(AArgCount: integer);
var
  type_: integer;
  fv: double;
  pv: double;
  nper: integer;
  per: integer;
  rate: double;
  PT: TPaymentTime;
begin
  type_ := 0;
  fv := 0;
  if AArgCount = 6 then
    if not PopResPosInt(type_) then Exit;
  if AArgCount >= 5 then
    if not PopResFloat(fv) then Exit;
  if not PopResFloat(pv) then Exit;
  if not PopResInt(nper) then Exit;
  if not PopResInt(per) then Exit;
  if not PopResFloat(rate) then Exit;

  if type_ = 0 then
    PT := ptEndOfPeriod
  else
    PT := ptStartOfPeriod;

  if (nper < 1) or (per < 0) or (per > nper) then
    Push(errNum)
  else
    Push(Payment(rate,nper,pv,fv,PT) - InterestPayment(rate,per,nper,pv,fv,PT))
end;

procedure TValueStack.DoProb(AArgCount: integer);
var
  UpperLim: double;
  LowerLim: double;
  vProb: double;
  vVal: double;
  Res: double;
begin
  if AArgCount = 4 then
    if not PopResFloat(UpperLim) then Exit;
  if not PopResFloat(LowerLim) then Exit;

  if AArgCount < 4 then
    UpperLim := LowerLim;

  Res := 0;
  FIterator.BeginIterate(PopRes,PopRes,False);
  while FIterator.NextFloat do begin
    vVal := FIterator.Result.vFloat;
    vProb := FIterator.Linked.Result.vFloat;
    if (vProb < 0) or (vProb > 1) then begin
      Push(errNum);
      Exit;
    end;
    if (vVal >= LowerLim) and (vVal <= UpperLim) then
      Res := Res + vProb;
  end;
  Push(Res);
end;

procedure TValueStack.DoProduct(AArgCount: integer);
var
  Res: double;
  V: double;
  Err: TXc12CellError;
  Found: boolean;
begin
  Found := False;
  Res := 1;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.Next do begin
      if GetAsFloatNum(FIterator.Result,V,Err) then begin
        if Err <> errUnknown then begin
          Push(Err);
          Exit;
        end;
        Found := True;
        Res := Res * V;
      end;
    end;
    Dec(AArgCount);
  end;
  if Found then
    Push(Res)
  else
    Push(0);
end;

procedure TValueStack.DoQuartile;
var
  q: integer;
  R: boolean;
  Res: double;
begin
  if not PopResInt(q) then Exit;
  if (q < 0) or (q > 4) then
    Push(errNum)
  else begin
    case q of
      0: R := Percentile(0,Res);
      1: R := Percentile(0.25,Res);
      2: R := Percentile(0.5,Res);
      3: R := Percentile(0.75,Res);
      else R := Percentile(1,Res);
    end;
    if not R then
      Push(errNum)
    else
      Push(Res);
  end;
end;

procedure TValueStack.DoQuartile_Exc;
begin
  Push(errNA);
end;

procedure TValueStack.DoQuartile_Inc;
begin

end;

procedure TValueStack.DoQuotient;
var
  V1,V2: double;
begin
  if not PopResFloat(V2) then Exit;
  if not PopResFloat(V1) then Exit;

  if V2 = 0 then
    Push(errNum)
  else
    Push(Trunc(V1 / V2));
end;

procedure TValueStack.DoRadians;
var
  Angle: double;
begin
  if not PopResFloat(Angle) then Exit;

  Push(DegToRad(Angle));
end;

procedure TValueStack.DoRank(AArgCount: integer);
begin
  DoRank_Eq(AArgCount);
end;

procedure TValueStack.DoRank_Eq(AArgCount: integer);
var
  Num: double;
  V: double;
  FV: TXLSFormulaValue;
  Order: integer;
  Err: TXc12CellError;
  nAsc,nDesc: integer;
  Found: boolean;
begin
  Order := 0;
  if AArgCount = 3 then
    if not PopResInt(Order) then Exit;
  FV := PopRes;
  if not PopResFloat(Num) then Exit;

  nAsc := 0;
  nDesc := 0;
  Found := False;
  FIterator.BeginIterate(FV,False);
  while FIterator.Next do begin
    GetAsFloat(FIterator.Result,V,Err);
    if Err <> errUnknown then begin
      Push(Err);
      Exit;
    end;
    if V <= Num then begin
      Inc(nAsc);
      if V = Num then
        Found := True;
    end;
    if V >= Num then begin
      Inc(nDesc);
      if V = Num then
        Found := True;
    end;
  end;
  if Found then begin
    if Order = 0 then
      Push(nDesc)
    else
      Push(nAsc);
  end
  else
    Push(errNA);
end;

procedure TValueStack.DoReplace;
var
  OldText: AxUCString;
  NewText: AxUCString;
  S: AxUCString;
  i,n: integer;
begin
  if not PopResStr(NewText) then Exit;
  if not PopResPosInt(n) then Exit;
  if not PopResPosInt(i,1) then Exit;
  if not PopResStr(OldText) then Exit;

  S := Copy(OldText,1,i - 1) + NewText + Copy(OldText,i + n,MAXINT);
  Push(S);
end;

procedure TValueStack.DoRoman(AArgCount: integer);
const
  Arabics: Array[0..12] of Integer = (1,4,5,9,10,40,50,90,100,400,500,900,1000) ;
  Romans: Array[0..12] of String = ('I','IV','V','IX','X','XL','L','XC','C','CD','D','CM','M') ;
var
  Mode: integer;
  Num: integer;
  i: integer;
  Res: AxUCString;
begin
  Mode := 0;
  if AArgCount = 2 then
    if not PopResInt(Mode) then Exit;
  if not PopResPosInt(Num,0) then Exit;

  if (Mode < 0) or (Mode > 4) or (Num > 3999) then
    Push(errValue)
  else begin
    Res := '';
    for i := High(Arabics) downto 0 do begin
      while (Num >= Arabics[i]) do begin
        Num := Num - Arabics[i];
        Res := Res + Romans[i];
      end;
    end;
    Push(Res);
  end;
end;

procedure TValueStack.DoRoundUp;
var
  NumDigits: integer;
  V: double;
  Sign: double;
begin
  if not PopResInt(NumDigits) then Exit;
  if not PopResFloat(V) then Exit;

  if V < 0 then
    Sign := -1
  else
    Sign := 1;

  V := Abs(V);
  if NumDigits >= 0 then
    V := Ceil(V * Array10[NumDigits]) / Array10[NumDigits]
  else if NumDigits < 0 then
    V := Ceil(V / Array10[Abs(NumDigits)]) * Array10[Abs(NumDigits)];
  Push(V * Sign);
end;

procedure TValueStack.DoRowsColumns(AFuncId: integer);
var
  nC,nR: integer;
  FV: TXLSFormulaValue;
begin
  FV := PopRes;

  case FV.RefType of
    xfrtRef  : begin
      nC := 1;
      nr := 1;
    end;
    xfrtArea : begin
      nC := FV.Col2 - FV.Col1 + 1;
      nR := FV.Row2 - FV.Row1 + 1
    end;
    xfrtXArea: begin
      nC := FV.Col2 - FV.Col1 + 1;
      nR := FV.Row2 - FV.Row1 + 1
    end;
    xfrtArray: begin
      nC := TXLSArrayItem(FV.vSource).Width;
      // TODO arrays are only one dimensional.
      nR := nC;
    end;
    else begin
      Push(errRef);
      Exit;
    end;
  end;

  if AFuncId = 076 { ROWS } then
    Push(nR)
  else
    Push(nC);
end;

procedure TValueStack.DoRsq;
var
  V: double;
begin
  DoCorrel;
  if (PeekRes(0).RefType = xfrtNone) and (PeekRes(0).ValType = xfvtFloat) then begin
    PopResFloat(V);
    Push(V * V);
  end;
end;

procedure TValueStack.DoSearch(AArgCount: integer; ACaseSensitive: boolean);
var
  p: integer;
  FindText: AxUCString;
  WithinText: AxUCString;
  StartNum: integer;
begin
  StartNum := 1;
  if AArgCount = 3 then
    if not PopResInt(StartNum) then Exit;
  if not PopResStr(WithinText) then Exit;
  if not PopResStr(FindText) then Exit;

  if StartNum < 1 then begin
    Push(errValue);
    Exit;
  end;

  if not ACaseSensitive then begin
    FindText := Lowercase(FindText);
    WithinText := Lowercase(WithinText);
  end;

  p := PosEx(FindText,WithinText,StartNum);
  if p < 1 then begin
    Push(errValue);
    Exit;
  end
  else
    Push(p);
end;

procedure TValueStack.DoSkew(AArgCount: integer);
var
  i: integer;
  Len: integer;
  Sum1: double;
  Sum2: double;
  dx: double;
  SumP3: double;
  stdev: double;
  Mean: double;
  Arr: TDynDoubleArray;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  Len := Length(Arr);
  if Len <= 0 then
    Push(errNA)
  else begin
    Sum1 := 0;
    for i := 0 to High(Arr) do
      Sum1 := Sum1 + Arr[i];
    Mean := Sum1 / Len;

    Sum2 := 0;
    for i := 0 to High(Arr) do
      Sum2 := Sum2 + (Arr[i] - Mean) * (Arr[i] - Mean);
    stdev := Sqrt(Sum2 / (Len - 1));
    if stdev = 0 then begin
      Push(errDiv0);
      Exit;
    end;

    SumP3 := 0;
    for i := 0 to High(Arr) do begin
      dx := (Arr[i] - Mean) / stdev;
      SumP3 := SumP3 + (dx * dx * dx);
    end;

    Push(((SumP3 * Len) / (Len - 1)) / (Len - 2));
  end;
end;

procedure TValueStack.DoSLN;
var
  Cost: double;
  Salvage: double;
  Life: double;
begin
  if not PopResFloat(Life) then Exit;
  if not PopResFloat(Salvage) then Exit;
  if not PopResFloat(Cost) then Exit;

  if Life = 0 then
    Push(errDiv0)
  else
    Push((Cost - Salvage) / Life);
end;

procedure TValueStack.DoSlope;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  vX: double;
  vY: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  Sum2: double;
  N: integer;
begin
  FVX := PopRes;
  FVY := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 0 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    Sum2 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      vX := FIterator.Result.vFloat;
      vY := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((vX - MeanX) * (vY - MeanY));
      Sum2 := Sum2 + Sqr(vX - MeanX);
    end;
    if Sum2 = 0 then begin
      Push(errDiv0);
      Exit;
    end;
    Push(Sum1 / Sum2);
  end;
end;

procedure TValueStack.DoSmall;
var
  Arr: TDynDoubleArray;
  k: integer;
begin
  if not PopResPosInt(k,1) then Exit;
  if not DoCollectValues(Arr) then
    Exit;
  SortDoubleArray(Arr);
  if k > Length(Arr) then
    Push(errNum)
  else
    Push(Arr[k - 1]);
end;

procedure TValueStack.DoStandardize;
var
  standard_dev: double;
  mean: double;
  x: double;
begin
  if not PopResFloat(standard_dev) then Exit;
  if not PopResFloat(mean) then Exit;
  if not PopResFloat(x) then Exit;

  if standard_dev < 0 then
    Push(errNum)
  else
    Push((x - mean) / standard_dev)
end;

procedure TValueStack.DoStDev(AArgCount: integer);
var
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValues(AArgCount,Vals) then begin
    if Length(Vals) <= 1 then
      Push(errDiv0)
    else
      Push(StdDev(Vals));
  end;
end;

procedure TValueStack.DoStdeva(AArgCount: integer);
var
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValuesA(AArgCount,Vals) then begin
    if Length(Vals) <= 1 then
      Push(errDiv0)
    else
      Push(StdDev(Vals));
  end;
end;

procedure TValueStack.DoStdevp(AArgCount: integer);
var
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValues(AArgCount,Vals) then begin
    if Length(Vals) <= 1 then
      Push(errDiv0)
    else
      Push(PopnStdDev(Vals));
  end;
end;

procedure TValueStack.DoStdevpa(AArgCount: integer);
var
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValuesA(AArgCount,Vals) then begin
    if Length(Vals) <= 1 then
      Push(errDiv0)
    else
      Push(PopnStdDev(Vals));
  end;
end;

procedure TValueStack.DoStdev_P(AArgCount: integer);
begin

end;

procedure TValueStack.DoStdev_S(AArgCount: integer);
begin

end;

procedure TValueStack.DoSteyx;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  X: double;
  Y: double;
  MeanX: double;
  Meany: double;
  Sum1: double;
  Sum2: double;
  Sum3: double;
  N: integer;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) < 1 then
    Push(errDiv0)
  else begin
    N := Sum2Arrays(FVX,FVY,MeanX,MeanY);
    if N <= 3 then begin
      Push(errDiv0);
      Exit;
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    Sum2 := 0;
    Sum3 := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do begin
      X := FIterator.Result.vFloat;
      Y := FIterator.Linked.Result.vFloat;
      Sum1 := Sum1 + ((X - MeanX) * (Y - MeanY));
      Sum2 := Sum2 + Sqr(X - MeanX);
      Sum3 := Sum3 + Sqr(Y - MeanY);
    end;
    if (Sum2 = 0) or (Sum3 = 0) then begin
      Push(errDiv0);
      Exit;
    end;
    Push(Sqrt((Sum2 - Sum1 * Sum1 / Sum3) / (N - 2)));
  end;
end;

procedure TValueStack.DoSubstitute(AArgCount: integer);
var
  i: integer;
  InstanceNum: integer;
  Text: AxUCString;
  OldText: AxUCString;
  NewText: AxUCString;
begin
  InstanceNum := -1;
  if AArgCount = 4 then
    if not PopResPosInt(InstanceNum,1) then Exit;
  if not PopResStr(NewText) then Exit;
  if not PopResStr(OldText) then Exit;
  if not PopResStr(Text) then Exit;

  if InstanceNum < 1 then
    Push(StringReplace(Text,OldText,NewText,[rfReplaceAll]))
  else begin
    for i := 1 to Length(Text) - Length(OldText) + 1 do begin
      if Copy(Text,i,Length(OldText)) = OldText then begin
        Dec(InstanceNum);
        if InstanceNum <= 0 then begin
          Push(Copy(Text,1,i - 1) + NewText + Copy(Text,i + Length(OldText),MAXINT));
          Exit;
        end;
      end;
    end;
    Push(Text);
  end;
end;

procedure TValueStack.DoSubtotal(AArgCount: integer);
var
  V: double;
  Err: TXc12CellError;
  FuncNum: integer;
  FV: TXLSFormulaValue;
begin
  FV := Peek(AArgCount - 1);
  Dec(AArgCount);
  if not GetAsFloat(FV,V,Err) then begin
    ClearStack;
    if Err <> errUnknown then
      Push(Err)
    else
      Push(errValue);
    Exit;
  end;
  FuncNum := Trunc(V);
  if not ((FuncNum in [1..11]) or (FuncNum in [101..111])) then begin
    ClearStack;
    Push(errValue);
    Exit;
  end;

  if FuncNum > 100 then
    FIterator.IgnoreHidden := True;
  try
    case FuncNum of
       1: DoAverage(AArgCount);
       2: DoCount(AArgCount);
       3: DoCounta(AArgCount);
       4: DoMax(AArgCount);
       5: DoMin(AArgCount);
       6: DoProduct(AArgCount);
       7: DoStdev(AArgCount);
       8: DoStdevp(AArgCount);
       9: DoSum(AArgCount);
      10: DoVar(AArgCount);
      11: DoVarp(AArgCount);

     101: DoAverage(AArgCount);
     102: DoCount(AArgCount);
     103: DoCounta(AArgCount);
     104: DoMax(AArgCount);
     105: DoMin(AArgCount);
     106: DoProduct(AArgCount);
     107: DoStdev(AArgCount);
     108: DoStdevp(AArgCount);
     109: DoSum(AArgCount);
     110: DoVar(AArgCount);
     111: DoVarp(AArgCount);
    end;
  finally
    FIterator.IgnoreHidden := False;
  end;
end;

procedure TValueStack.DoSum(AArgCount: integer);
var
  Sum: double;
begin
  Sum := 0;
  while AArgCount > 0 do begin
    FIterator.BeginIterate(PopRes,True);
    while FIterator.NextFloat do
      Sum := Sum + FIterator.Result.vFloat;
    if FIterator.Result.ValType = xfvtError then begin
      Push(FIterator.Result.vErr);
      Exit;
    end;
    Dec(AArgCount);
  end;
  Push(Sum);
end;

procedure TValueStack.DoSumif(AArgCount: integer);
var
  Criteria: TXLSFormulaValue;
  CritVal: TXLSVarValue;
  SumRange: TXLSFormulaValue;
  Val: TXLSVarValue;
  FV: TXLSFormulaValue;
  Op: TXLSDbCondOperator;
  Sum: double;
  Sum2: double;
  Err: TXc12CellError;
begin
  if AArgCount = 3 then
    SumRange := PopRes;
  Criteria := PopRes;
  case Criteria.RefType of
    xfrtNone,
    xfrtRef     : FV := Criteria;
    xfrtArea    : begin
      MakeIntersect(@Criteria);
      GetCellValue(Criteria.Cells,Criteria.Col1,Criteria.Row1,@FV);
    end;
    xfrtXArea   : ;
    xfrtAreaList: begin
      PopRes;
      Push(errValue);
      Exit;
    end;
    xfrtArray   : FV := TXLSArrayItem(Criteria.vSource).GetAsFormulaValue(0,0);
  end;
  FVToVarValue(@FV,@CritVal);

  Op := xdcoEQ;
  if CritVal.ValType = xfvtString then
    ConditionStrToVarValue(CritVal.vStr,Op,@CritVal);

  if AArgCount = 3 then begin
    FV := PopRes;
    if not (SumRange.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) then begin
      Push(errValue);
      Exit;
    end;
    if (FV.Col2 - FV.Col1) > (FV.Row2 - FV.Row1) then
      SumRange.Col2 := Max(FV.Col2,SumRange.Col2)
    else
      SumRange.Row2 := Max(FV.Row2,SumRange.Row2);
    FIterator.BeginIterate(FV,SumRange,False);
  end
  else
    FIterator.BeginIterate(PopRes,False);

  Sum := 0;
  Sum2 := 0;
  while FIterator.Next do begin
    FVToVarValue(@FIterator.Result,@Val);
    if CompareVarValue(@Val,@CritVal,Op) then begin
      if FIterator.ResIsError(Err) then begin
        Push(Err);
        Exit;
      end;
      if FIterator.ResIsFloat then
        Sum := Sum + FIterator.Result.vFloat;
      if (AArgCount = 3) and FIterator.Linked.ResIsFloat then
        Sum2 := Sum2 + FIterator.Linked.Result.vFloat;
    end;
  end;
  if AArgCount = 3 then
    Push(Sum2)
  else
    Push(Sum);
end;

procedure TValueStack.DoSumifs(AArgCount: integer);
var
  i: integer;
  Val: TXLSVarValue;
  Sum: double;
  Iter: TValStackIterator;
  Ok: boolean;
  Data: TDynCriteriaDataArray;
begin
  if not Odd(AArgCount) then begin
    Push(errValue);
    Exit;
  end;

  SetupIfs(AArgCount,Data);

  Sum := 0;
  while FIterator.Next do begin
    Iter := FIterator;
    Ok := True;
    for i := 0 to High(Data) do begin
      FVToVarValue(@Iter.Result,@Val);
      Ok := CompareVarValue(@Val,@Data[i].CritVal,Data[i].Op);
      if not Ok then
        Break;
      Iter := Iter.Linked;
    end;
    if Ok then begin
      if Iter.ResIsFloat then
        Sum := Sum + Iter.Result.vFloat;
    end;
  end;
  Push(Sum);
end;

procedure TValueStack.DoSumProduct(AArgCount: integer);
var
  Res: double;
  C1,R1: integer;
  C2,R2: integer;
  V1,V2: double;
  Empty1,Empty2: boolean;
  FV1,FV2: TXLSFormulaValue;
  Err1,Err2: TXc12CellError;
  Mat: array of array of double;
begin
  Res := 0;

  FV1 := PeekRes(0);
  if not (FV1.RefType in [xfrtArea,xfrtXArea]) then begin
    Push(errValue);
    Exit;
  end;
  C1 := FV1.Col2 - FV1.Col1;
  R1 := FV1.Row2 - FV1.Row1;
  SetLength(Mat,C1 + 1,R1 + 1);
  for R2 := 0 to R1 do begin
    for C2 := 0 to C1 do
      Mat[C2,R2] := 1;
  end;
  while AArgCount > 0 do begin
    FV1 := PopRes;
    if not (FV1.RefType in [xfrtArea,xfrtXArea]) then begin
      Push(errValue);
      Exit;
    end;

    Dec(AArgCount);
    if AArgCount > 0 then begin
      FV2 := PopRes;
      C1 := FV1.Col2 - FV1.Col1;
      R1 := FV1.Row2 - FV1.Row1;
      C2 := FV2.Col2 - FV2.Col1;
      R2 := FV2.Row2 - FV2.Row1;
      if not (FV2.RefType in [xfrtArea,xfrtXArea]) or not ((C1 = C2) and (R1 = R2)) then begin
        Push(errValue);
        Exit;
      end;
      for R2 := 0 to R1 do begin
        for C2 := 0 to C1 do
          Mat[C2,R2] := 0;
      end;
      Dec(AArgCount);

      for R2 := 0 to R1 do begin
        for C2 := 0 to C1 do begin
          Err1 := GetCellValueFloat(FV1.Cells,FV1.Col1 + C2,FV1.Row1 + R2,False,V1,Empty1);
          Err2 := GetCellValueFloat(FV2.Cells,FV2.Col1 + C2,FV2.Row1 + R2,False,V2,Empty2);
          if Err1 <> errUnknown then begin
            Push(Err1);
            Exit;
          end
          else if Err2 <> errUnknown then begin
            Push(Err2);
            Exit;
          end
          else if not Empty1 and not Empty2 then begin
            Mat[C2,R2] := V1 * V2;
            Res := Res + V1 * V2;
          end;
        end;
      end;
    end
    else begin
      Res := 0;
      for R2 := FV1.Row1 to FV1.Row2 do begin
        for C2 := FV1.Col1 to FV1.Col2 do begin
          Err1 := GetCellValueFloat(FV1.Cells,C2,R2,False,V1,Empty1);
          if Err1 <> errUnknown then begin
            Push(Err1);
            Exit;
          end
          else if not Empty1 then
            Res := Res + V1 * Mat[C2 - FV1.Col1,R2 - FV1.Row1];
        end;
      end;
    end;
  end;
  Push(Res);
end;

procedure TValueStack.DoSumsq(AArgCount: integer);
var
  i: integer;
  Sum: double;
  Arr: TDynDoubleArray;
begin
  if not DoCollectAllValues(AArgCount,Arr) then
    Exit;
  if Length(Arr) <= 0 then
    Push(errNA)
  else begin
    Sum := 0;
    for i := 0 to High(Arr) do
      Sum := Sum + Arr[i] * Arr[i];
    Push(Sum);
  end;
end;

procedure TValueStack.DoSumX2mY2;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  Res: double;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) <= 1 then
    Push(errDiv0)
  else begin
    Res := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do
      Res := Res + Sqr(FIterator.Result.vFloat) - Sqr(FIterator.Linked.Result.vFloat);
    Push(Res);
  end;
end;

procedure TValueStack.DoSumX2pY2;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  Res: double;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) <= 1 then
    Push(errDiv0)
  else begin
    Res := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do
      Res := Res + Sqr(FIterator.Result.vFloat) + Sqr(FIterator.Linked.Result.vFloat);
    Push(Res);
  end;
end;

procedure TValueStack.DoSumXmY2;
var
  FVX: TXLSFormulaValue;
  FVY: TXLSFormulaValue;
  Res: double;
begin
  FVY := PopRes;
  FVX := PopRes;

  if not AreasEqualSize(@FVX,@FVY) then
    Push(errNA)
  else if AreaCellCount(@FVX) <= 1 then
    Push(errDiv0)
  else begin
    Res := 0;
    FIterator.BeginIterate(FVX,FVY,False);
    while FIterator.NextFloat do
      Res := Res + Sqr(FIterator.Result.vFloat - FIterator.Linked.Result.vFloat);
  end;
end;

procedure TValueStack.DoSYD;
var
  Cost: double;
  Salvage: double;
  Life: double;
  Per: double;
begin
  if not PopResFloat(Per) then Exit;
  if not PopResFloat(Life) then Exit;
  if not PopResFloat(Salvage) then Exit;
  if not PopResFloat(Cost) then Exit;

  if (Salvage < 0) or (Life <= 0) or (Per <= 0) or (Per > Life) then
    Push(errNum)
  else
    Push(((Cost - Salvage) * (Life - Per + 1) * 2) / (Life * (Life + 1)));
end;

procedure TValueStack.DoTDist;
var
  tails: integer;
begin
  if not PopResInt(tails) then Exit;

  if (tails < 1) or (tails > 2) then
    Push(errNum)
  else begin
    if tails = 1 then
      DoT_Dist_RT
    else
      DoT_Dist_2T;
  end;
end;

procedure TValueStack.DoText;
var
  V: double;
  S: AxUCString;
  Mask: TExcelMask;
begin
  if not PopResStr(S) then
    Exit;
  if not PopResFloat(V) then
    Exit;

  Mask := TExcelMask.Create;
  try
    Mask.Mask := S;
    Push(Mask.FormatNumber(V));
  finally
    Mask.Free;
  end;
end;

procedure TValueStack.DoTime;
var
  T: TDateTime;
  HH,MM,SS: integer;
begin
  if not PopResInt(SS) or not PopResInt(MM) or not PopResInt(HH) then
    Exit;

  if HH < 0 then begin
    Push(errNum);
    Exit;
  end;

  if HH > 24 then
    HH := HH div 24;

  T := EncodeTime(HH,0,0,0);
{$ifdef DELPHI_6_OR_LATER}
  T := IncMinute(T,MM);
  T := IncSecond(T,SS);
{$else}
  raise XLSRWException.Create('TIME not supported by Delphi 5');
{$endif}

  if T < 0 then
    Push(errNum)
  else
    Push(TimeToStr(T));
end;

procedure TValueStack.DoTInv;
begin
  DoT_Inv_2T;
end;

procedure TValueStack.DoTranspose;
var
  C,R,C2,R2: integer;
  FV: TXLSFormulaValue;
  fvCell: TXLSFormulaValue;
  Res: TXLSArrayItem;
begin
  FV := PopRes;
  if not (FV.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) then
    Push(errValue)
  // TODO arrays
  else if FV.RefType = xfrtArray then
    Push(errNA)
  else begin
    Res := TXLSArrayItem.Create(FV.Row2 - FV.Row1 + 1,FV.Col2 - FV.Col1 + 1);
    for R := FV.Row1 to FV.Row2 do begin
      for C := FV.Col1 to FV.Col2 do begin
        GetCellValue(FV.Cells,C,R,@fvCell);
        if not fvCell.Empty then begin
          C2 := R - FV.Row1;
          R2 := C - FV.Col1;
          case fvCell.ValType of
            xfvtFloat  : Res.AsFloat[C2,R2]   := fvCell.vFloat;
            xfvtBoolean: Res.AsBoolean[C2,R2] := fvCell.vBool;
            xfvtError  : Res.AsError[C2,R2]   := fvCell.vErr;
            xfvtString : Res.AsString[C2,R2]  := fvCell.vStr;
          end;
        end;
      end;
    end;
    Push(Res);
  end;
end;

procedure TValueStack.DoTrend(AArgCount: integer);
var
  i: integer;
  V: double;
  a: double;
  b: double;
  MeanX: double;
  MeanY: double;
  Sum1: double;
  Sum2: double;
  N: integer;
  const_: boolean;
  new_x: TDynDoubleArray;
  known_x: TDynDoubleArray;
  known_y: TDynDoubleArray;
  Res: TXLSArrayItem;
begin
  if AArgCount = 4 then
    if not PopResBoolean(const_) then Exit;
  if AArgCount >= 3 then
    DoCollectValues(new_x);
  if AArgCount >= 2 then
    DoCollectValues(known_x);

  DoCollectValues(known_y);

  if AArgCount < 2 then begin
    SetLength(known_x,Length(known_y));
    for i := 0 to High(known_y) do
      known_x[i] := i + 1;
  end;
  if AArgCount < 3 then begin
    SetLength(new_x,Length(known_x));
    for i := 0 to High(known_x) do
      new_x[i] := known_x[i];
  end;

  N := Length(known_y);

  if (Length(known_y) <> Length(known_x)) then
    Push(errRef)
  else if (N < 2) then
    Push(errNum)
  else begin
    MeanX := 0;
    MeanY := 0;
    for i := 0 to High(Known_y) do begin
      MeanX := MeanX + known_x[i];
      MeanY := MeanY + known_y[i];
    end;
    MeanX := MeanX / N;
    MeanY := MeanY / N;
    Sum1 := 0;
    Sum2 := 0;
    for i := 0 to High(known_y) do begin
      Sum1 := Sum1 + ((known_x[i] - MeanX) * (known_y[i] - MeanY));
      Sum2 := Sum2 + Sqr(known_x[i] - MeanX);
    end;
    if Sum2 <= 0 then begin
      Push(errDiv0);
      Exit;
    end;

    Res := TXLSArrayItem.Create(1,Length(new_x));
    for i := 0 to High(new_x) do begin
      b := Sum1 / Sum2;
      a := MeanY - b * MeanX;
      V := a + b * new_x[i];
      Res.AsFloat[0,i] := V;
    end;
    Push(Res);
  end;
end;

procedure TValueStack.DoTrimMean;
var
  i: integer;
  n: integer;
  Per: double;
  Arr: TDynDoubleArray;
  Sum: double;
begin
  if not PopResFloat(Per) then
    Push(errValue)
  else if (Per < 0) or (Per >= 1) then
    Push(errNum)
  else begin
    if not DoCollectValues(Arr) then
      Exit;
    SortDoubleArray(Arr);
    if Length(Arr) <= 0 then
      Push(errNum)
    else begin
      n := Floor(Length(Arr) * Per);
      if (n mod 2) <> 0 then
        Dec(n);
      n := n div 2;
      Sum := 0;
      for i := n to Length(Arr) - n - 1 do
        Sum := Sum + Arr[i];
      Push(Sum / (Length(Arr) - 2 * n));
    end;
  end;
end;

procedure TValueStack.DoTrunc(AArgCount: integer);
const
  Array10: array[0..15] of double = (1,10,100,1000,10000,100000,1000000,10000000,1000000000,10000000000,100000000000,1000000000000,10000000000000,100000000000000,1000000000000000,10000000000000000);
var
  V: double;
  NumDigits: integer;
begin
  NumDigits := 0;
  if AArgCount = 2 then
    if not PopResInt(NumDigits) then Exit;
  if not PopResFloat(V) then Exit;

  if NumDigits > High(Array10) then
    Push(V)
  else if NumDigits = 0 then
    Push(Trunc(V))
  else if NumDigits > 0 then
    Push(Trunc(V * Array10[NumDigits]) / Array10[NumDigits])
  else
    Push(Trunc(V / Array10[Abs(NumDigits)]) * Array10[Abs(NumDigits)]);
end;

procedure TValueStack.DoTTest;
begin
  DoT_Test;
end;

procedure TValueStack.DoType;
var
  FV: TXLSFormulaValue;
  Res: integer;
begin
  FV := PopRes;

  Res := 16;
  case FV.RefType of
    xfrtNone,
    xfrtRef     : begin
      case FV.ValType of
        xfvtFloat  : Res := 1;
        xfvtBoolean: Res := 4;
        xfvtError  : Res := 16;
        xfvtString : Res := 2;
        else         Res := 16;
      end;
    end;
    xfrtArea:      begin
      Res := 16;
      if MakeIntersect(@FV) then begin
        case FV.ValType of
          xfvtFloat  : Res := 1;
          xfvtBoolean: Res := 4;
          xfvtError  : Res := 16;
          xfvtString : Res := 2;
          else         Res := 16;
        end;
      end;
    end;
    xfrtXArea,
    xfrtAreaList: Res := 16;
    xfrtArray   : Res := 64;
  end;

  Push(Res);
end;

procedure TValueStack.DoT_Dist;
var
  cumulative: boolean;
  degrees_freedom: integer;
  x: double;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResInt(degrees_freedom) then Exit;
  if not PopResFloat(x) then Exit;

  if cumulative and ((x < 0) or (degrees_freedom < 1)) then
    Push(errNum)
  else if not cumulative and (degrees_freedom < 1) then
    Push(errDiv0)
  else begin
    if cumulative then
      Push(StudentTDistCDF(x,degrees_freedom))
    else
      Push(StudentTDistPDF(x,degrees_freedom));
  end;
end;

procedure TValueStack.DoT_Dist_2T;
var
  degrees_freedom: integer;
  x: double;
begin
  if not PopResInt(degrees_freedom) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (degrees_freedom < 1) then
    Push(errNum)
  else
    Push((1 - StudentTDistCDF(x,degrees_freedom)) * 2);
end;

procedure TValueStack.DoT_Dist_RT;
var
  degrees_freedom: integer;
  x: double;
begin
  if not PopResInt(degrees_freedom) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (degrees_freedom < 1) then
    Push(errNum)
  else
    Push(1 - StudentTDistCDF(x,degrees_freedom));
end;

procedure TValueStack.DoValue;
var
  S: AxUCString;
  V: double;
  VDT: TDateTime;
begin
  if PopResStr(S) then begin
    if TryStrToFloat(S,V) then
      Push(V)
    else if TryStrToDateTime(S,VDT) then
      Push(VDT)
    else if TryCurrencyStrToFloat(S,V) then
      Push(V)
    else
      Push(errValue);
  end
  else
    Push(errValue);
end;

procedure TValueStack.DoVar(AArgCount: integer);
var
  i: integer;
  TempStackPtr: integer;
  Sum: double;
  n: integer;
  Average: double;
  Variance: double;
begin
  TempStackPtr := FStackPtr;

  Sum := 0;
  n := 0;
  for i := 1 to AArgCount do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.NextFloat do begin
      Sum := Sum + FIterator.Result.vFloat;
      Inc(n);
    end;
  end;

  if n <= 0 then begin
    Push(errDiv0);
    Exit;
  end;

  Average := Sum / n;

  FStackPtr := TempStackPtr;
  Variance := 0;
  for i := 0 to AArgCount - 1 do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.Next do begin
      case FIterator.Result.ValType of
        xfvtFloat   : Variance := Variance + Sqr(FIterator.Result.vFloat - Average);
        xfvtError   : begin
          Push(FIterator.Result.vErr);
          Exit;
        end;
      end;
    end;
  end;
  Push(Variance / (n - 1));
end;

procedure TValueStack.DoVara(AArgCount: integer);
var
  i: integer;
  TempStackPtr: integer;
  Sum: double;
  n: integer;
  Average: double;
  Variance: double;
begin
  TempStackPtr := FStackPtr;

  Sum := 0;
  n := 0;
  for i := 1 to AArgCount do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.NextFloatA do begin
      Sum := Sum + FIterator.Result.vFloat;
      Inc(n);
    end;
  end;

  if n <= 0 then begin
    Push(errDiv0);
    Exit;
  end;

  Average := Sum / n;

  FStackPtr := TempStackPtr;
  Variance := 0;
  for i := 0 to AArgCount - 1 do begin
    FIterator.BeginIterate(PopRes,False);
    while FIterator.Next do begin
      case FIterator.Result.ValType of
        xfvtFloat,
        xfvtBoolean,
        xfvtString  :  Variance := Variance + Sqr(FIterator.AsFloatA - Average);
        xfvtError   : begin
          Push(FIterator.Result.vErr);
          Exit;
        end;
      end;
    end;
  end;
  Push(Variance / (n - 1));
end;

procedure TValueStack.DoVarp(AArgCount: integer);
var
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValues(AArgCount,Vals) then begin
    if Length(Vals) <= 1 then
      Push(0)
    else
      Push(PopnVariance(Vals));
  end;
end;

procedure TValueStack.DoVarpa(AArgCount: integer);
var
  Vals: TDynDoubleArray;
begin
  if DoCollectAllValuesA(AArgCount,Vals) then begin
    if Length(Vals) <= 1 then
      Push(0)
    else
      Push(PopnVariance(Vals));
  end;
end;

procedure TValueStack.DoVar_P(AArgCount: integer);
begin

end;

procedure TValueStack.DoVar_S(AArgCount: integer);
begin

end;

procedure TValueStack.DoVDB(AArgCount: integer);
var
  Cost: double;
  Salvage: double;
  Life: double;
  StartPer: double;
  EndPer: double;
  Factor: double;
  NoSwitch: boolean;
begin
  NoSwitch := False;
  Factor := 2;
  if AArgCount = 7 then
    if not PopResBoolean(NoSwitch) then Exit;
  if AArgCount >= 6 then
    if not PopResFloat(Factor) then Exit;

  if not PopResFloat(EndPer) then Exit;
  if not PopResFloat(StartPer) then Exit;
  if not PopResFloat(Life) then Exit;
  if not PopResFloat(Salvage) then Exit;
  if not PopResFloat(Cost) then Exit;

  if (Salvage < 0) or (Cost < 0) or (StartPer < 0) or (EndPer < 0) or (Life <= 0) or (Factor <= 0) or (StartPer > Life) or (EndPer > Life) or (StartPer > EndPer) then
    Push(errNum)
  else
    Push(VDB(Cost,Salvage,Life,StartPer,EndPer,Factor,NoSwitch));
end;

procedure TValueStack.DoVLookup(AArgCount: integer);
var
  R       : integer;
  FoundR  : integer;
  Range   : boolean;
  ColIndex: integer;
  Table   : TXLSFormulaValue;
  Lookup  : TXLSFormulaValue;
  FV      : TXLSFormulaValue;
begin
  Range := True;
  if AArgCount = 4 then
    if not PopResBoolean(Range) then Exit;
  if not PopResInt(ColIndex) then Exit;
  Table := PopRes;
  Lookup := PopRes;

  if not (Table.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray]) or (ColIndex <= 0) then begin
    Push(errValue);
    Exit;
  end;

  MakeIntersect(@Lookup);

  if Lookup.RefType in [xfrtArea,xfrtXArea,xfrtAreaList] then begin
    Push(errNA);
    Exit;
  end;

  Dec(ColIndex);

  if ColIndex > (Table.Col2 - Table.Col1) then begin
    Push(errREF);
    Exit;
  end;


  if Lookup.RefType = xfrtArray then
    Lookup := TXLSArrayItem(Lookup.vStr).GetAsFormulaValue(0,0);

  FoundR := -1;

  for R := Table.Row1 to Table.Row2 do begin
    GetCellValue(Table.Cells,Table.Col1,R,@FV);
    if FV.ValType = Lookup.ValType then begin
      if Range then begin
        case FV.ValType of
          xfvtFloat  : if Lookup.vFloat >= FV.vFloat then FoundR := R;
          xfvtBoolean: if Lookup.vBool = FV.vBool then FoundR := R;
          xfvtError  : if Lookup.vErr = FV.vErr then FoundR := R;
          xfvtString : if CompareText(FV.vStr,Lookup.vStr) <= 0 then FoundR := R;
        end;
      end
      else begin
        case FV.ValType of
          xfvtFloat  : if Lookup.vFloat = FV.vFloat then FoundR := R;
          xfvtBoolean: if Lookup.vBool = FV.vBool then FoundR := R;
          xfvtError  : if Lookup.vErr = FV.vErr then FoundR := R;
{$ifdef DELPHI_7_OR_LATER}
          xfvtString : if MatchesMask(FV.vStr,Lookup.vStr) then FoundR := R;
{$endif}          
        end;
      end;
    end;
  end;
  if FoundR >= 0 then begin
    GetCellValue(Table.Cells,Table.Col1 + ColIndex,FoundR,@FV);
    Push(FV);
  end
  else
    Push(errNA);
end;

procedure TValueStack.DoWeibull;
begin
  DoWeibull_Dist;
end;

procedure TValueStack.DoWeibull_Dist;
var
  x: double;
  alpha: double;
  beta: double;
  cumulative: boolean;
begin
  if not PopResBoolean(cumulative) then Exit;
  if not PopResFloat(beta) then Exit;
  if not PopResFloat(alpha) then Exit;
  if not PopResFloat(x) then Exit;

  if (x < 0) or (alpha <= 0) or (beta <= 0) then
    Push(errNum)
  else begin
    if cumulative then
      Push(WeibullCDF(x,alpha,beta))
    else
      Push(WeibullPDF(x,alpha,beta));
  end;
end;

procedure TValueStack.DoWorkday(AArgCount: integer);
var
  holidays: TDynDoubleArray;
  days: integer;
  start_date: double;
  n: integer;
begin
  if AArgCount = 3 then
    if not DoCollectValues(holidays) then Exit;
  if not PopResInt(days) then Exit;
  if not PopResDateTime(start_date) then Exit;

  n := (Trunc(days) div 7) * 2;
  Inc(days,n);

  n := DayOfWeek(start_date);
  if n = 6 then
    Inc(days,2)
  else if n = 7 then
    Inc(days);

  Push(start_date + days);
end;

procedure TValueStack.DoWorkday_Int(AArgCount: integer);
begin
  Push(errNA);
end;

procedure TValueStack.DoXIRR(AArgCount: integer);
const
  MAX_STEPS = 100;
var
  Dates: TDynDoubleArray;
  Values: TDynDoubleArray;
  Rate: double;
  Rate1, Rate2, RateN: double;
  F1, F2, FN, dF, Scale: double;
  Quit: boolean;
  N: integer;
  FoundRoot: boolean;

function CalcValue(Rate: double): double;
var
  i: integer;

function disc(d: TDateTime; v: double): double;
var
  Exp, Coef: double;
begin
  Exp := (d - Dates[0]) / 365;
  Coef := Power(1 + Rate,Exp);
  Result := v / Coef;
end;

begin
  Result := 0;
  for i := 0 to High(Dates) do
    Result := Result + disc(Dates[i], Values[i]);
end;

begin
  Rate := 0.1;
  if AArgCount = 3 then
    if not PopResFloat(Rate) then
      Exit;

  if not DoCollectValues(Dates) then
    Exit;

  if not DoCollectValues(Values) then
    Exit;

  if (Length(Values) < 2) or (Length(Values) <> Length(Dates)) then begin
    Push(errNum);
    Exit;
  end;

  FoundRoot := True;
  Rate1 := Rate;
  Rate2 := Rate + 1;
  Quit := False;
  N := 0;
  FN := 0;
  Scale := 1;
  RateN := 0;

  F1 := CalcValue(Rate1);
  F2 := CalcValue(Rate2);
  while not Quit do begin
    if (F2 = F1) or (Rate2 = Rate1) then begin
      Quit := True;
      FoundRoot := False;
    end
    else begin
      dF := (F2 - F1) / (Rate2 - Rate1);

      RateN := Rate1 + (0 - F1)/dF/Scale;
      N := N + 1;

      if RateN > -100 then
        FN := CalcValue(RateN);

      if Abs(RateN - Rate1) / Abs(Rate1) < 0.0000005 then
        Quit := True
      else if N >= MAX_STEPS then begin
        Push(errValue);
        Exit;
      end
      else if not (RateN > -100) then begin
        //RateN := Rate1 - 3*(Rate1 + 100)/4;
        Scale := Scale * 2;
      end
      else begin
        Scale := 1;

        Rate2 := Rate1;
        F2 := F1;

        Rate1 := RateN;
        F1 := FN;
      end;
    end;
  end;
  if FoundRoot then
    Push(RateN)
  else
    Push(0);
end;

procedure TValueStack.DoXNVP;
var
  i: Integer;
  Dates: TDynDoubleArray;
  Values: TDynDoubleArray;
  Rate: double;
  Res: double;
  Day1: double;
  Diff: Double;
begin
  if not DoCollectValues(Dates) then
    Exit;
  if not DoCollectValues(Values) then
    Exit;
  if not PopResFloat(Rate) then
    Exit;

  if (Length(Values) < 2) or (Length(Values) <> Length(Dates)) then begin
    Push(errNum);
    Exit;
  end;

  Res := 0.0;
  Day1 := Dates[0];
  for i := 0 to High(Dates) do begin
    Diff := Dates[i] - Day1;
    if (Diff < 0) then begin
      Push(errNum);
      Exit;
    end;
    Res := Res + Values[i] / Power(1.0 + Rate, Diff / 365);
  end;
  Push(Res);
end;

procedure TValueStack.DoYearfrac(AArgCount: integer);
var
  Basis: integer;
  D,D1,D2: double;
  Res: double;
  YearLen: integer;
  NumYears: integer;
  DaysInYears: double;
  AveYearLen: double;
  DayDiff360: integer;
  YY1,MM1,DD1: word;
  YY2,MM2,DD2: word;

function AppearsOneYearApart(const AD1,AD2: double): boolean;
begin
  if YY1 = YY2 then
    Result := True
  else
    Result := (((YY1 + 1) = YY2) and ((MM1 > MM2) or ((MM1 = MM2) and (DD1 >= DD2))));
end;

function Feb29Between(const AD1,AD2: double): boolean;
var
  Y: word;
  D: double;
begin
  for Y := YY1 to YY2 do begin
    if IsLeapYear(Y) then begin
      D := EncodeDate(Y,2,29);
      Result := (D >= AD1) and (D <= AD2);
      if Result then
        Exit;
    end;
  end;
  Result := False;
end;

begin
  if AArgCount = 3 then begin
    if not PopResPosInt(Basis) then
      Exit;
  end
  else
    Basis := 0;

  if not PopResDateTime(D2) then
    Exit;
  if not PopResDateTime(D1) then
    Exit;

  if not Basis in [0..4] then begin
    Push(errValue);
    Exit;
  end;

  D1 := Trunc(D1);
  D2 := Trunc(D2);
  if D1 > D2 then begin
    D := D1;
    D1:= D2;
    D2 := D;
  end;

  Res := 0;
  case Basis of
    0: begin
{$ifdef DELPHI_6_OR_LATER}
      if D1 = D2 then
        Res := 00
      else begin
        DecodeDate(D1,YY1,MM1,DD1);
        DecodeDate(D2,YY2,MM2,DD2);
        if (DD1 = 31) and (DD2 = 31) then begin
          DD1 := 30;
          DD2 := 30;
        end
        else if (DD1 = 30) and (DD2 = 31) then
          DD2 := 30
        else if (MM1 = 2) and (MM2 = 2) and (D1 = EndOfAMonth(YY1,MM1)) and (D2 = EndOfAMonth(YY2,MM2)) then begin
          DD1 := 30;
          DD2 := 30;
        end
        else if (MM1 = 2) and (D1 = EndOfAMonth(YY1,MM1)) then
          DD1 := 30;
        Res := ((DD2 + MM2 * 30 + YY2 * 360) - (DD1 + MM1 * 30 + YY1 * 360)) / 360;
      end;
{$else}
     raise XLSRWException.Create('YEARFRAC not supported by Delphi 5');
{$endif}
    end;
    1: begin
      if D1 = D2 then
        Res := 00
      else begin
        DecodeDate(D1,YY1,MM1,DD1);
        DecodeDate(D2,YY2,MM2,DD2);
        if AppearsOneYearApart(D1,D2) then begin
          if (YY1 = YY2) and IsLeapYear(YY1) then
            YearLen := 366
          else if Feb29Between(D1,D2) or ((MM2 = 2) and (DD2 = 29)) then
            YearLen := 366
          else
            YearLen := 365;
          Res := (D2 - D1) / YearLen;
        end
        else begin
          NumYears := (YY2 - YY1) + 1;
          DaysInYears := EncodeDate(YY2 + 1,1,1) - EncodeDate(YY1,1,1);
          AveYearLen := DaysInYears / NumYears;
          Res := (D2 - D1) / AveYearLen;
        end;
      end;
    end;
    2: begin
      Res := (D2 - D1) / 360;
    end;
    3: begin
      Res := (D2 - D1) / 365;
    end;
    4: begin
      if D1 = D2 then
        Res := 00
      else begin
        DecodeDate(D1,YY1,MM1,DD1);
        DecodeDate(D2,YY2,MM2,DD2);
        if DD1 = 31 then
          DD1 := 30;
        if DD2 = 31 then
          DD2 := 30;
        Daydiff360 := ((DD2 + MM2 * 30 + YY2 * 360) - (DD1 + MM1 * 30 + YY1 * 360));
        Res := Daydiff360 / 360;
      end;
    end;
  end;
  Push(Res);
end;

procedure TValueStack.DoZtest(AArgCount: integer);
begin
  DoZ_Test(AArgCount);
end;

procedure TValueStack.DoZ_Test(AArgCount: integer);
var
  Sigma: double;
  x: double;
  Arr: TDynDoubleArray;
begin
  Sigma := 1000;
  if AArgCount = 3 then
    if not PopResFloat(Sigma) then Exit;
  if not PopResFloat(x) then Exit;

  if not DoCollectValues(Arr) then
    Exit;

  if Length(Arr) < 1 then
    Push(errNA)
  else if (Sigma = 0) or (Length(Arr) = 1) then
    Push(errDiv0)
  else
    Push(ZTest(Arr,x,Sigma));
end;

function TValueStack.FDist(x: double; df1, df2: integer): double;
begin
  x := df2 / (df2 + df1 * x);
  Result := BetaCDF(df2 / 2,df1 / 2,x);
end;

procedure TValueStack.FillArrayItem(AItem: TXLSArrayItem; AValue: TXLSFormulaValue);
var
  C,R: integer;
  FV: TXLSFormulaValue;
begin
  case AValue.RefType of
    xfrtNone,
    xfrtRef     :  AItem.Add(0,0,AValue);
    xfrtArea    : begin
      for R := AValue.Row1 to AValue.Row2 do begin
        for C := AValue.Col1 to AValue.Col2 do begin
          GetCellValue(AValue.Cells,C,R,@FV);
          AItem.Add(C - AValue.Col1,R - AValue.Row1,FV);
        end;
      end;
    end;
    xfrtAreaList: raise XLSRWException.Create('Invalid value');
    xfrtArray   : AItem.Assign(TXLSArrayItem(AValue.vSource));
  end;
end;

procedure TValueStack.FillArrayItemFloat(AItem: TXLSArrayItem; AValue: TXLSFormulaValue);
var
  C,R: integer;
  V: double;
  Err: TXc12CellError;
  FV: TXLSFormulaValue;
begin
  case AValue.RefType of
    xfrtNone,
    xfrtRef     : begin
      if GetAsFloat(AValue,V,Err) then
        AItem.AsFloat[0,0] := V
      else
        AItem.AsError[0,0] := Err;
    end;
    xfrtArea    : begin
      for R := AValue.Row1 to AValue.Row2 do begin
        for C := AValue.Col1 to AValue.Col2 do begin
          GetCellValue(AValue.Cells,C,R,@FV);
          if GetAsFloat(FV,V,Err) then
            AItem.AsFloat[C - AValue.Col1,R - AValue.Row1] := V
          else
            AItem.AsError[C - AValue.Col1,R - AValue.Row1] := Err;
        end;
      end;
    end;
    xfrtAreaList: raise XLSRWException.Create('Invalid value');
    xfrtArray   : AItem.Assign(TXLSArrayItem(AValue.vSource));
  end;
end;

function TValueStack.ForecastExponential(X: Double; const KnownY, KnownX: TDynDoubleArray): Double;
var
  LF : TLinEstData;
begin
  LogEst(KnownY, KnownX, LF, false);
  Result := LF.B0 * Power(LF.B1, X);
end;

function TValueStack.GetAsBoolean(AValue: TXLSFormulaValue; out AResult: boolean): TXc12CellError;
var
  C,R: integer;
  V: double;
begin
  Result := errUnknown;

  case AValue.RefType of
    xfrtArea    : begin
      C := FCol;
      R := FRow;
      if Intersect(@AValue,C,R) then
        GetCellValue(AValue.Cells,C,R,@AValue)
      else begin
        Result := errValue;
        Exit;
      end;
    end;
    xfrtAreaList: begin
      Result := errValue;
      Exit;
    end;
  end;
  case AValue.ValType of
    xfvtUnknown: AResult := False;
    xfvtFloat  : AResult := AValue.vFloat <> 0;
    xfvtBoolean: AResult := AValue.vBool;
    xfvtError  : Result := AValue.vErr;
    xfvtString : begin
      if TryStrToFloat(AValue.vStr,V) then
        AResult := V <> 0
      else if Uppercase(AValue.vStr) = G_StrTRUE then
        AResult := True
      else if Uppercase(AValue.vStr) = G_StrFALSE then
        AResult := False
      else
        Result := errValue;
    end;
  end;
end;

function TValueStack.GetAsError(AValue: TXLSFormulaValue): TXc12CellError;
var
  C,R: integer;
begin
  Result := errUnknown;

  case AValue.RefType of
    xfrtArea    : begin
      C := FCol;
      R := FRow;
      if Intersect(@AValue,C,R) then
        GetCellValue(AValue.Cells,C,R,@AValue)
      else begin
        Result := errValue;
        Exit;
      end;
    end;
    xfrtAreaList: begin
      Result := errValue;
      Exit;
    end;
  end;
  if AValue.ValType = xfvtError then
    Result := AValue.vErr;
end;

function TValueStack.GetAsFloat(AValue: TXLSFormulaValue; out AResult: double; out AError: TXc12CellError): boolean;
var
  C,R: integer;
begin
  Result := True;
  AError := errUnknown;

  case AValue.RefType of
    xfrtRef    : begin
      Result := not AValue.Empty;
      if not Result then begin
        AResult := 0;
        Exit;
      end;
    end;
    xfrtArea    : begin
      C := FCol;
      R := FRow;
      if Intersect(@AValue,C,R) then begin
        GetCellValue(AValue.Cells,C,R,@AValue);
        Result := not AValue.Empty;
      end
      else begin
        AError := errValue;
        Exit;
      end;
    end;
    xfrtXArea   : ;
    xfrtAreaList: begin
      AError := errValue;
      Exit;
    end;
  end;
  case AValue.ValType of
    xfvtUnknown: begin
      AResult := 0;
      Result := False;
    end;
    xfvtFloat  : AResult := AValue.vFloat;
    xfvtBoolean: if AValue.vBool then AResult := 1 else AResult := 0;
    xfvtError  : AError := AValue.vErr;
    xfvtString : begin
      if TryStrToFloat(AValue.vStr,AResult) then
        Exit;
      if TryStrToDateTime(AValue.vStr,TDateTime(AResult)) then
        Exit;
      AError := errValue;
    end;
  end;
end;

function TValueStack.GetAsFloatNum(AValue: TXLSFormulaValue; out AResult: double; out AError: TXc12CellError): boolean;
begin
  Result := True;
  AError := errUnknown;

  case AValue.RefType of
    xfrtNone    : begin
      case AValue.ValType of
        xfvtFloat   : AResult := AValue.vFloat;
        xfvtBoolean : if AValue.vBool then AResult := 1 else AResult := 0;
        xfvtString  : if not TryStrToFloat(AValue.vStr,AResult) then AError := errValue;
        xfvtError   : AError := AValue.vErr;
        else          Result := False;
      end;
    end;
    xfrtRef     : begin
      if AValue.ValType = xfvtError then
        AError := AValue.vErr
      else if AValue.ValType = xfvtFloat then
        AResult := AValue.vFloat
      else begin
        AResult := 0;
        Result := False;
      end;
    end;
    xfrtArea,
    xfrtXArea,
    xfrtAreaList: raise XLSRWException.Create('Invalid ref type in calc formula');
  end;
end;

function TValueStack.GetAsFloatNumOnly(AValue: TXLSFormulaValue; out AResult: double; out  AError: TXc12CellError): boolean;
begin
  AError := AValue.vErr;
  Result := AError = errUnknown;
  if Result then begin
    Result := (AValue.RefType in [xfrtNone,xfrtRef]) and (AValue.ValType = xfvtFloat);
    if Result then
      AResult := AValue.vFloat;
  end;
end;

function TValueStack.GetAsString(AValue: TXLSFormulaValue; out AResult: AxUCString): TXc12CellError;
var
  C,R: integer;
begin
  Result := errUnknown;
  if AValue.RefType = xfrtArea then begin
    C := FCol;
    R := FRow;
    if Intersect(@AValue,C,R) then
      GetCellValue(AValue.Cells,C,R,@AValue)
    else begin
      Result := errValue;
      Exit;
    end;
  end;
  case AValue.ValType of
    xfvtUnknown: AResult := '';
    xfvtFloat  : AResult := FloatToStr(AValue.vFloat);
    xfvtBoolean: if AValue.vBool then AResult := G_StrTRUE else AResult := G_StrFALSE;
    xfvtError  : begin
      Result := AValue.vErr;
      AResult := Xc12CellErrorNames[AValue.vErr];
    end;
    xfvtString : AResult := AValue.vStr;
  end;
end;

function TValueStack.GetAsString2(AValue: TXLSFormulaValue; out AResult: AxUCString): boolean;
var
  Err: TXc12CellError;
begin
  Err := GetAsString(AValue,AResult);
  Result := Err = errUnknown;
  if not Result then
    Push(Err);
end;

function TValueStack.GetAsValueType(AValue: TXLSFormulaValue; out AResult: TXLSFormulaValueType): TXc12CellError;
var
  C,R: integer;

function GetRefVal(ACol,ARow: integer; out AResult: TXLSFormulaValueType): TXc12CellError;
var
  Cell: TXLSCellItem;
begin
  Result := errUnknown;
  if not AValue.Cells.FindCell(ACol,ARow,Cell) then
    AResult := xfvtUnknown
  else begin
    case AValue.Cells.CellType(@Cell) of
      xctNone,
      xctBlank         : AResult := xfvtUnknown;
//      xctCurrency      : AResult := xfvtFloat;
      xctBoolean       : AResult := xfvtBoolean;
      xctError         : Result := AValue.Cells.GetError(@Cell);
      xctString        : AResult := xfvtString;
      xctFloat,
      xctFloatFormula  : AResult := xfvtFloat;
      xctStringFormula : AResult := xfvtString;
      xctBooleanFormula: AResult := xfvtBoolean;
      xctErrorFormula  : Result := AValue.Cells.GetFormulaValError(@Cell);
    end;
  end;
end;

begin
  Result := errUnknown;
  if AValue.RefType = xfrtArea then begin
    C := FCol;
    R := FRow;
    if Intersect(@AValue,C,R) then
      Result := GetRefVal(C,R,AResult)
    else
      Result := errValue;
  end
  else
    AResult := AValue.ValType;
end;

procedure TValueStack.GetCellValue(Cells: TXLSCellMMU; ACol, ARow: integer; AFormulaVal: PXLSFormulaValue);
var
  Cell: TXLSCellItem;
begin
  AFormulaVal.Empty := not Cells.FindCell(ACol,ARow,Cell);
  if AFormulaVal.Empty then begin
    // Is this correct? Is default always float/zero?
    // Default value depends on if it is read as a float or string. Function/operator decides that.
    AFormulaVal.ValType := xfvtUnknown
  end
  else begin
    AFormulaVal.RefType := xfrtNone;
    case Cells.CellType(@Cell) of
      xctNone,
      xctBlank         : begin
        AFormulaVal.ValType := xfvtUnknown;
      end;
      xctError         : begin
        AFormulaVal.ValType := xfvtError;
        AFormulaVal.vErr := Cells.GetError(@Cell);
      end;
      xctString        : begin
        AFormulaVal.ValType := xfvtString;
        AFormulaVal.vStr := Cells.GetString(@Cell);
      end;
      xctFloat         : begin
        AFormulaVal.ValType := xfvtFloat;
        AFormulaVal.vFloat := Cells.GetFloat(@Cell);
      end;
//      xctCurrency      : begin
//        AFormulaVal.ValType := xfvtFloat;
//        AFormulaVal.vFloat := Cells.GetFloat(@Cell);
//      end;
      xctBoolean       : begin
        AFormulaVal.ValType := xfvtBoolean;
        AFormulaVal.vBool := Cells.GetBoolean(@Cell);
      end;
      xctFloatFormula  : begin
        AFormulaVal.ValType := xfvtFloat;
        AFormulaVal.vFloat := Cells.GetFormulaValFloat(@Cell);
      end;
      xctStringFormula : begin
        AFormulaVal.ValType := xfvtString;
        AFormulaVal.vStr := Cells.GetFormulaValString(@Cell);
      end;
      xctBooleanFormula: begin
        AFormulaVal.ValType := xfvtBoolean;
        AFormulaVal.vBool := Cells.GetFormulaValBoolean(@Cell);
      end;
      xctErrorFormula  : begin
        AFormulaVal.ValType := xfvtError;
        AFormulaVal.vErr := Cells.GetFormulaValError(@Cell);
      end;
    end;
  end;
end;

function TValueStack.GetCellValueFloat(Cells: TXLSCellMMU; ACol, ARow: integer; var AResult: double; out AEmptyCell: boolean): TXc12CellError;
var
  Cell: TXLSCellItem;
begin
  Result := errUnknown;
  AResult := 0;
  AEmptyCell := not Cells.FindCell(ACol,ARow,Cell);
  if not AEmptyCell then begin
    case Cells.CellType(@Cell) of
      xctNone,
      xctBlank         : AEmptyCell := True;
      xctError         : Result := Cells.GetError(@Cell);
      xctString        : begin
        if not TryStrToFloat(Cells.GetString(@Cell),AResult) then begin
          AEmptyCell := True;
          AResult := 0;
        end;
      end;
      xctFloat         : AResult := Cells.GetFloat(@Cell);
//      xctCurrency      : AResult := Cells.GetFloat(@Cell);
      xctBoolean       : AResult := Integer(Cells.GetBoolean(@Cell));
      xctFloatFormula  : AResult := Cells.GetFormulaValFloat(@Cell);
      xctStringFormula : begin
        if not TryStrToFloat(Cells.GetFormulaValString(@Cell),AResult) then begin
          AEmptyCell := True;
          AResult := 0;
        end;
      end;
      xctBooleanFormula: AResult := Integer(Cells.GetFormulaValBoolean(@Cell));
      xctErrorFormula  : Result := Cells.GetFormulaValError(@Cell);
    end;
  end;
end;

function TValueStack.GetCellValueFloat(Cells: TXLSCellMMU; ACol, ARow: integer; AAcceptText: boolean; var Aresult: double; out AEmptyCell: boolean): TXc12CellError;
begin
  if AAcceptText then
    Result := GetCellValueFloat(Cells,ACol,ARow,AResult,AEmptyCell)
  else
    Result := Cells.AsFloat(ACol,ARow,AResult,AEmptyCell);
end;

function TValueStack.GetCellBlank(Cells: TXLSCellMMU; ACol, ARow: integer): boolean;
var
  Cell: TXLSCellItem;
begin
  Result := not Cells.FindCell(ACol,ARow,Cell) or (Cells.CellType(@Cell) = xctBlank);
end;

function TValueStack.GetCellType(Cells: TXLSCellMMU; ACol, ARow: integer): TXLSFormulaValueType;
var
  Cell: TXLSCellItem;
begin
  Result := xfvtUnknown;
  if Cells.FindCell(ACol,ARow,Cell) then begin
    case Cells.CellType(@Cell) of
      xctNone,
      xctBlank         : ;
//      xctCurrency      : Result := xfvtFloat;
      xctBoolean       : Result := xfvtBoolean;
      xctError         : Result := xfvtError;
      xctString        : Result := xfvtString;
      xctFloat,
      xctFloatFormula  : Result := xfvtFloat;
      xctStringFormula : Result := xfvtString;
      xctBooleanFormula: Result := xfvtBoolean;
      xctErrorFormula  : Result := xfvtError;
    end;
  end;
end;

procedure TValueStack.IncStackPtr;
begin
  Inc(FStackPtr);
  if FStackPtr > High(FStack) then begin
    SetLength(FStack,Length(FStack) + $7F);
{$ifdef _AXOLOT_DEBUG }
    if Length(FStack) > 1024 then
      raise XLSRWException.Create('Evaluate stack have grown to horrendeous size!');
{$endif}
  end;
end;

function TValueStack.Intersect(AValue: PXLSFormulaValue; var ACol, ARow: integer): boolean;
begin
  if AValue.Col1 = AValue.Col2 then begin
    Result := (ARow >= AValue.Row1) and (ARow <= AValue.Row2);
    if Result then
      ACol := AValue.Col1;
  end
  else if AValue.Row1 = AValue.Row2 then begin
    Result := (ACol >= AValue.Col1) and (ACol <= AValue.Col2);
    if Result then
      ARow := AValue.Row1;
  end
  else
    Result := False;
end;

function TValueStack.IsVector(AFormulaVal: PXLSFormulaValue): boolean;
begin
  Result := AFormulaVal.RefType in [{xfrtNone,xfrtRef,}xfrtArea,xfrtXArea,xfrtArray];
  if Result then begin
    if (AFormulaVal.RefType in [xfrtArea,xfrtXArea]) then
      Result := ((AFormulaVal.Col2 - AFormulaVal.Col1) = 0) xor ((AFormulaVal.Row2 - AFormulaVal.Row1) = 0)
    else
      Result := (TXLSArrayItem(AFormulaVal.vSource).Width > 1) xor (TXLSArrayItem(AFormulaVal.vSource).Height > 1);
  end;
end;

function TValueStack.LinEst(const KnownY,KnownX: TDynDoubleArray; var LF : TLinEstData; ErrorStats : boolean): boolean;
var
  i : Integer;
  sx, sy, xmean, ymean, sxx, sxy, syy, x, y : Extended;
  NData: integer;
begin
  FillChar(LF,SizeOf(TLinEstData),#0);

  NData := Length(KnownY);
  {compute basic sums}
  sx := 0.0;
  sy := 0.0;
  sxx := 0.0;
  sxy := 0.0;
  syy := 0.0;
  for i := 0 to NData-1 do begin
    x := KnownX[i];
    y := KnownY[i];
    sx := sx+x;
    sy := sy+y;
    sxx := sxx+x*x;
    syy := syy+y*y;
    sxy := sxy+x*y;
  end;
  xmean := sx/NData;
  ymean := sy/NData;
  sxx := sxx-NData*xmean*xmean;
  syy := syy-NData*ymean*ymean;
  sxy := sxy-NData*xmean*ymean;

  {check for zero variance}
  Result := (sxx > 0.0) and (syy > 0.0);
  if not Result then
    Exit;

  {initialize returned parameters}
  fillchar(LF, sizeof(LF), 0);

  {regression coefficients}
  LF.B1 := sxy/sxx;
  LF.B0 := ymean-LF.B1*xmean;

  {error statistics}
  if (ErrorStats) then begin
    LF.ssr := LF.B1*sxy;
    LF.sse := syy-LF.ssr;
    LF.R2 := LF.ssr/syy;
    LF.df := NData-2;
    LF.sigma := sqrt(LF.sse/LF.df);
    if LF.sse = 0.0 then
      {pick an arbitrarily large number for perfect fit}
      LF.F0 := 1.7e+308
    else
      LF.F0 := (LF.ssr*LF.df)/LF.sse;
    LF.seB1 := LF.sigma/sqrt(sxx);
    LF.seB0 := LF.sigma*sqrt((1.0/NData)+(xmean*xmean/sxx));
  end;
end;

procedure TValueStack.LogEst(const KnownY, KnownX: TDynDoubleArray; var LF: TLinEstData; ErrorStats: Boolean);
var
  i : Integer;
  lny : TDynDoubleArray;
  NData: integer;
begin
  NData := Length(KnownY);

  {allocate array for the log-transformed data}
  {f (Size > MaxBlockSize) then}
  { RaiseStatError(stscStatBadCount);}
  SetLength(lny, NData);

  {initialize transformed data}
  for i := 0 to NData-1 do
    lny[i] := ln(KnownY[i]);

  {fit transformed data}
  LinEst(lny, KnownX, LF, ErrorStats);

  {return values for B0 and B1 in exponential model y=B0*B1^x}
  LF.B0 := exp(LF.B0);
  LF.B1 := exp(LF.B1);
  {leave other values in LF in log form}
end;

function TValueStack.MakeIntersect(AValue: PXLSFormulaValue): boolean;
var
  C,R: integer;
begin
  Result := AValue.RefType = xfrtArea;
  if Result then begin
    C := FCol;
    R := FRow;
    Result := Intersect(AValue,C,R);
    if Result then begin
      AValue.RefType := xfrtRef;
      AValue.Col1 := C;
      AValue.Row1 := R;
      AValue.Col2 := C;
      AValue.Row2 := R;
      GetCellValue(AValue.Cells,C,R,AValue);
    end;
  end;
end;

procedure TValueStack.MakeVector(AFormulaVal: PXLSFormulaValue; ALeftTop: boolean);
begin
  if ALeftTop then begin
    if (AFormulaVal.Row2 - AFormulaVal.Row1) >= (AFormulaVal.Col2 - AFormulaVal.Col1) then
      AFormulaVal.Col2 := AFormulaVal.Col1
    else
      AFormulaVal.Row2 := AFormulaVal.Row1;
  end
  else begin
    if (AFormulaVal.Row2 - AFormulaVal.Row1) >= (AFormulaVal.Col2 - AFormulaVal.Col1) then
      AFormulaVal.Col1 := AFormulaVal.Col2
    else
      AFormulaVal.Row1 := AFormulaVal.Row2;
  end;
end;

function TValueStack.NPV(ARate: double; AValues: TDynDoubleArray): double;
var
  i: integer;
begin
  ARate := 1 / (1 + ARate);

  Result := 0;
  for i := 0 to High(AValues) do
    Result := Result * ARate + AValues[i];
  Result := Result * ARate;
end;

procedure TValueStack.OpArea(AOperator: integer);
var
  V1,V2: PXLSFormulaValue;
  A,A1,A2: TXLSCellArea;
  CA: TCellAreas;

function GetAreas(AVal: PXLSFormulaValue; out AArea: TXLSCellArea): boolean;
begin
  Result := True;
  case AVal.RefType of
    xfrtNone : Result := False;
    xfrtRef  : SetCellArea(AArea,AVal.Col1,AVal.Row1);
    xfrtArea : SetCellArea(AArea,AVal.Col1,AVal.Row1,AVal.Col2,AVal.Row2);
    else      Result := False;
  end;
end;

begin
  DecStackPtr;
  V1 := @FStack[FStackPtr];
  V2 := @FStack[FStackPtr + 1];
  case AOperator of
    xptgOpIsect: begin
      if not GetAreas(V1,A1) or not GetAreas(V2,A2) then
        SetError(errValue)
      else begin
        if not IntersectCellArea(A1,A2,A) then
          SetError(errNull)
        else begin
          V1.Col1 := A.Col1;
          V1.Row1 := A.Row1;
          V1.Col2 := A.Col2;
          V1.Row2 := A.Row2;
          if AreaIsRef(A) then begin
            FStack[FStackPtr].RefType := xfrtRef;
            GetCellValue(FStack[FStackPtr].Cells,A.Col1,A.Row1,@FStack[FStackPtr]);
          end
          else
            FStack[FStackPtr].RefType := xfrtArea;
        end;
      end;
    end;
    xptgOpUnion: begin
      if (V1.RefType = xfrtAreaList) and (V2.RefType = xfrtAreaList) then begin
        CA := TCellAreas(V1.vSource);
        CA.Assign(TCellAreas(V2.vSource));
      end
      else if V1.RefType = xfrtAreaList then begin
        CA := TCellAreas(V1.vSource);
        if not GetAreas(V2,A2) then
          SetError(errNull)
        else
          CA.Add(A2.Col1,A2.Row1,A2.Col2,A2.Row2).Obj := V2.Cells;
      end
      else if V2.RefType = xfrtAreaList then begin
        CA := TCellAreas(V2.vSource);
        if not GetAreas(V1,A1) then
          SetError(errNull)
        else begin
          CA.Add(A1.Col1,A1.Row1,A1.Col2,A1.Row2).Obj := V1.Cells;
          V1.RefType := xfrtAreaList;
          V1.vSource := V2.vSource;
        end;
      end
      else begin
        if not GetAreas(V1,A1) or not GetAreas(V2,A2) then
          SetError(errValue)
        else begin
          CA := TCellAreas.Create;
          CA.Add(A1.Col1,A1.Row1,A1.Col2,A1.Row2).Obj := V1.Cells;
          CA.Add(A2.Col1,A2.Row1,A2.Col2,A2.Row2).Obj := V2.Cells;
          V1.RefType := xfrtAreaList;
          V1.vSource := CA;
          FGarbage.Add(CA);
        end;
      end;
    end;
    xptgOpRange: begin
      if not GetAreas(V1,A1) or not GetAreas(V2,A2) then
        SetError(errValue)
      else begin
        ExtendCellArea(A1,A2,A);
        V1.Col1 := A.Col1;
        V1.Row1 := A.Row1;
        V1.Col2 := A.Col2;
        V1.Row2 := A.Row2;
        if AreaIsRef(A) then begin
          FStack[FStackPtr].RefType := xfrtRef;
          GetCellValue(FStack[FStackPtr].Cells,A.Col1,A.Row1,@FStack[FStackPtr]);
        end
        else
          FStack[FStackPtr].RefType := xfrtArea;
      end;
    end;
  end;
end;

procedure TValueStack.OpArrayBoolean(AOperator: integer);
var
  C,R: integer;
  W,H: integer;
  Arr1,Arr2: TXLSArrayItem;
begin
  PopArrays(Arr1,Arr2,False);

  W := Max(Arr1.Width,Arr2.Width);
  H := Max(Arr1.Height,Arr2.Height);

  for R := 0 to H - 1 do begin
    for C := 0 to W - 1 do begin
      if Arr1.Hit[C,R] and Arr2.Hit[C,R] then begin
        if Arr2[C,R].ValType = xfvtError then
          Arr1.AsError[C,R] := Arr2.AsError[C,R]
        else if Arr1[C,R].ValType <> xfvtError then begin
          Push(Arr1.GetAsFormulaValue(C,R));
          Push(Arr2.GetAsFormulaValue(C,R));
          OpBoolean(AOperator);
          Arr1.Add(C,R,Pop);
        end;
      end
      else
        Arr1.AsError[C,R] := errNA;
    end;
  end;
  FGarbage.Add(Arr2);
  Push(Arr1);
end;

procedure TValueStack.OpArrayConcat;
var
  C,R: integer;
  W,H: integer;
  Arr1,Arr2: TXLSArrayItem;
begin
  PopArrays(Arr1,Arr2,False);

  W := Max(Arr1.Width,Arr2.Width);
  H := Max(Arr1.Height,Arr2.Height);

  for R := 0 to H - 1 do begin
    for C := 0 to W - 1 do begin
      if Arr1.Hit[C,R] and Arr2.Hit[C,R] then begin
        if Arr2[C,R].ValType = xfvtError then
          Arr1.AsError[C,R] := Arr2.AsError[C,R]
        else if Arr1[C,R].ValType <> xfvtError then begin
          Push(Arr1.GetAsFormulaValue(C,R));
          Push(Arr2.GetAsFormulaValue(C,R));
          OpConcat;
          Arr1.Add(C,R,Pop);
        end;
      end
      else
        Arr1.AsError[C,R] := errNA;
    end;
  end;
  FGarbage.Add(Arr2);
  Push(Arr1);
end;

procedure TValueStack.OpArrayFloat(AOperator: integer);
var
  C,R: integer;
  W,H: integer;
  Arr1,Arr2,Arr3: TXLSArrayItem;
begin
  PopArrays(Arr1,Arr2,True);
  FGarbage.Add(Arr1);
  FGarbage.Add(Arr2);

  if Arr1.IsVector and Arr2.IsVector then begin
    if Arr1.Width = Arr2.Width then begin
      W := Arr1.Width;
      H := Min(Arr1.Height,Arr2.Height);
    end
    else if Arr1.Height = Arr2.Height then begin
      W := Min(Arr1.Width,Arr2.Width);
      H := Arr1.Height;
    end
    else begin
      W := Max(Arr1.Width,Arr2.Width);
      H := Max(Arr1.Height,Arr2.Height);
    end;
  end
  else if Arr1.IsVector then begin
    W := Arr2.Width;
    H := Arr2.Height
  end
  else if Arr2.IsVector then begin
    W := Arr1.Width;
    H := Arr1.Height
  end
  else begin
    W := Min(Arr1.Width,Arr2.Width);
    H := Min(Arr1.Height,Arr2.Height);
  end;

  Arr3 := TXLSArrayItem.Create(W,H);

  for R := 0 to H - 1 do begin
    for C := 0 to W - 1 do begin
      if Arr1.Hit[C,R] and Arr2.Hit[C,R] and Arr3.Hit[C,R] then begin
        if (Arr1[C,R].ValType <> xfvtFloat) or (Arr2[C,R].ValType <> xfvtFloat) then begin
          if (Arr1[C,R].ValType <> xfvtError) and (Arr2[C,R].ValType = xfvtError) then
            Arr1.AsError[C,R] := Arr2.AsError[C,R]
          else
            Arr1.AsError[C,R] := errValue;
        end
        else begin
          case AOperator of
            xptgOpAdd   : Arr3.AsFloat[C,R] := Arr1.AsFloat[C,R] + Arr2.AsFloat[C,R];
            xptgOpSub   : Arr3.AsFloat[C,R] := Arr1.AsFloat[C,R] - Arr2.AsFloat[C,R];
            xptgOpMult  : Arr3.AsFloat[C,R] := Arr1.AsFloat[C,R] * Arr2.AsFloat[C,R];
            xptgOpDiv   : begin
              if Arr2.AsFloat[C,R] = 0 then
                Arr3.AsError[C,R] := errDiv0
              else
                Arr3.AsFloat[C,R] := Arr1.AsFloat[C,R] / Arr2.AsFloat[C,R];
            end;
            xptgOpPower : Arr3.AsFloat[C,R] := Power(Arr1.AsFloat[C,R],Arr2.AsFloat[C,R]);
          end;
        end;
      end
      else
        Arr3.AsError[C,R] := errNA;
    end;
  end;
  Push(Arr3);
end;

procedure TValueStack.OpArrayUnary(AOperator: integer);
var
  C,R: integer;
  Arr: TXLSArrayItem;
begin
  PopArray(Arr,True);

  for R := 0 to Arr.Height - 1 do begin
    for C := 0 to Arr.Width - 1 do begin
      if Arr.Hit[C,R] then begin
        if Arr[C,R].ValType <> xfvtError then begin
          Push(Arr.GetAsFormulaValue(C,R));
          OpUnary(AOperator);
          Arr.Add(C,R,Pop);
        end;
      end
      else
        Arr.AsError[C,R] := errNA;
    end;
  end;
  Push(Arr);
end;

procedure TValueStack.OpBoolean(AOperator: integer);
var
  VT1,VT2: TXLSFormulaValueType;
  V1,V2: double;
  E1,E2: TXc12CellError;
  Res: boolean;
begin
  DecStackPtr;

  Res := False;
  E1 := GetAsValueType(FStack[FStackPtr],VT1);
  E2 := GetAsValueType(FStack[FStackPtr + 1],VT2);
  if E1 <> errUnknown then
    SetError(E1)
  else if E2 <> errUnknown then
    SetError(E2)
  else if (VT1 = xfvtError) or (VT2 = xfvtError) then begin
    if VT1 = xfvtError then
      SetError(FStack[FStackPtr].vErr)
    else
      SetError(FStack[FStackPtr + 1].vErr);
  end
  else begin
    if VT1 = VT2 then begin
      case VT1 of
        xfvtUnknown: FStack[FStackPtr].vBool := False;
        xfvtFloat,
        xfvtBoolean: begin
          GetAsFloat(FStack[FStackPtr],V1,E1);
          GetAsFloat(FStack[FStackPtr + 1],V2,E2);
          case AOperator of
            xptgOpLT    : Res := V1 < V2;
            xptgOpLE    : Res := V1 <= V2;
            xptgOpEQ    : Res := V1 = V2;
            xptgOpGE    : Res := V1 >= V2;
            xptgOpGT    : Res := V1 > V2;
            xptgOpNE    : Res := V1 <> V2;
          end;
          FStack[FStackPtr].vBool := Res;
        end;
        xfvtString: begin
          case AOperator of
            xptgOpLT    : Res := CompareText(FStack[FStackPtr].vStr,FStack[FStackPtr + 1].vStr) < 0;
            xptgOpLE    : Res := CompareText(FStack[FStackPtr].vStr,FStack[FStackPtr + 1].vStr) <= 0;
            xptgOpEQ    : Res := CompareText(FStack[FStackPtr].vStr,FStack[FStackPtr + 1].vStr) = 0;
            xptgOpGE    : Res := CompareText(FStack[FStackPtr].vStr,FStack[FStackPtr + 1].vStr) >= 0;
            xptgOpGT    : Res := CompareText(FStack[FStackPtr].vStr,FStack[FStackPtr + 1].vStr) > 0;
            xptgOpNE    : Res := CompareText(FStack[FStackPtr].vStr,FStack[FStackPtr + 1].vStr) <> 0;
          end;
          FStack[FStackPtr].vBool := Res;
        end;
        xfvtError : raise XLSRWException.Create('Invalid value type');
      end;
      FStack[FStackPtr].RefType := xfrtNone;
      FStack[FStackPtr].ValType := xfvtBoolean;
    end
    else begin
      case VT1 of
        xfvtUnknown: V1 := 0;
        xfvtFloat  : GetAsFloat(FStack[FStackPtr],V1,E1);
        xfvtString : V1 := MAXDOUBLE - 1;
        xfvtBoolean: V1 := MAXDOUBLE;
        xfvtError  : V1 := 4; // Errors are caught above
      end;
      case VT2 of
        xfvtUnknown: V2 := 0;
        xfvtFloat  : GetAsFloat(FStack[FStackPtr + 1],V2,E2);
        xfvtString : V2 := MAXDOUBLE - 1;
        xfvtBoolean: V2 := MAXDOUBLE;
        xfvtError  : V2 := 4; // Errors are caught above
      end;
      case AOperator of
        xptgOpLT    : Res := V1 < V2;
        xptgOpLE    : Res := V1 <= V2;
        xptgOpEQ    : begin
          if ((VT1 = xfvtUnknown) or (VT2 = xfvtUnknown)) and ((VT1 = xfvtString) or (VT2 = xfvtString)) then begin
            if (VT1 = xfvtString) and (FStack[FStackPtr].vStr = '') or (VT2 = xfvtString) and (FStack[FStackPtr + 1].vStr = '') then
              Res := True
            else
              Res := V1 = V2;
          end
          else
            Res := V1 = V2;
        end;
        xptgOpGE    : Res := V1 >= V2;
        xptgOpGT    : Res := V1 > V2;
        xptgOpNE    : Res := V1 <> V2;
      end;
      FStack[FStackPtr].vBool := Res;
      FStack[FStackPtr].RefType := xfrtNone;
      FStack[FStackPtr].ValType := xfvtBoolean;
    end;
  end;
end;

procedure TValueStack.OpFloat(AOperator: integer);
var
  V1,V2: double;
  E1,E2: TXc12CellError;
begin
  DecStackPtr;
  if FStackPtr < 0 then
    raise XLSRWException.Create('Empty stack');
  GetAsFloat(FStack[FStackPtr],V1,E1);
  GetAsFloat(FStack[FStackPtr + 1],V2,E2);
  if E1 <> errUnknown then
    SetError(E1)
  else if E2 <> errUnknown then
    SetError(E2)
  else begin
    if FIsArrayRes and ((FStack[FStackPtr].RefType in [xfrtArea,xfrtXArea,xfrtAreaList,xfrtArray]) or (FStack[FStackPtr + 1].RefType in [xfrtArea,xfrtXArea,xfrtAreaList,xfrtArray])) then begin
      IncStackPtr;
      OpArrayFloat(AOperator);
    end
    else begin
      try
        case AOperator of
          xptgOpAdd   : V1 := V1 + V2;
          xptgOpSub   : V1 := V1 - V2;
          xptgOpMult  : V1 := V1 * V2;
          xptgOpDiv   : begin
            if V2 = 0 then begin
              SetError(errDiv0);
              Exit;
            end
            else
              V1 := V1 / V2;
          end;
          xptgOpPower : begin
            if V1 >= 0 then
              V1 := Power(V1,V2)
            else
              SetError(errNum);
          end;
        end;
      except
        SetError(errNum);
      end;
      FStack[FStackPtr].RefType := xfrtNone;
      FStack[FStackPtr].ValType := xfvtFloat;
      FStack[FStackPtr].vFloat := V1;
    end;
  end;
end;

function TValueStack.Peek(ALevel: integer): TXLSFormulaValue;
begin
  if (FStackPtr - ALevel) < 0  then
    raise XLSRWException.Create('Empty stack');
  Result := FStack[FStackPtr - ALevel];
end;

function TValueStack.PeekFloat(ALevel: integer; out AValue: double): boolean;
var
  FV: TXLSFormulaValue;
  Err: TXc12CellError;
begin
  Result := False;
  FV := Peek(ALevel);
  if not GetAsFloat(FV,AValue,Err) then begin
    Push(errValue);
    Exit;
  end;
  if Err <> errUnknown then begin
    Push(Err);
    Exit;
  end;
  Result := True;
end;

function TValueStack.PeekRes(ALevel: integer): TXLSFormulaValue;
begin
  if (FResultPtr - ALevel) < 0  then
    raise XLSRWException.Create('Empty stack');
  Result := FStack[FResultPtr - ALevel];
end;

function TValueStack.Percentile(APercent: double; out AResult: double): boolean;
var
  Arr: TDynDoubleArray;
  L: integer;
  D: double;
  Index: integer;
begin
  Result := DoCollectValues(Arr);
  if not Result then
    Exit;

  SortDoubleArray(Arr);

  L := Length(Arr);
  Result := L > 0;
  if not Result then
    Exit;
  Index := Floor(APercent * (L - 1));
  D := APercent * (L - 1) - Index;

  if D = 0 then
    AResult := Arr[Index]
  else
    AResult := Arr[Index] + D * (Arr[Index + 1] - Arr[Index]);
end;

function TValueStack.Pop: TXLSFormulaValue;
begin
  if FStackPtr < 0 then
    raise XLSRWException.Create('Empty stack');
  Result := FStack[FStackPtr];
  Dec(FStackPtr);
end;

procedure TValueStack.PopArray(out AArr: TXLSArrayItem; AFloat: boolean);
var
  FV: TXLSFormulaValue;
begin
  FV := Pop;

  AArr := TXLSArrayItem.CreateFV(FV);

  AArr.VectorMode := True;

  if AFloat then
    FillArrayItemFloat(AArr,FV)
  else
    FillArrayItem(AArr,FV);
end;

procedure TValueStack.PopArrays(out AArr1, AArr2: TXLSArrayItem; AFloat: boolean);
begin
  PopArray(AArr2,AFloat);
  PopArray(AArr1,AFloat);
end;

function TValueStack.PopResult(out AValue: TXLSFormulaValue): boolean;
begin
  if FStackPtr < 0 then
    raise XLSRWException.Create('Empty stack');
  AValue := Pop;

  if not FIsArrayRes and (AValue.RefType = xfrtArea) then begin
    if (FCol >= AValue.Col1) and (FCol <= AValue.Col2) and (FRow >= AValue.Row1) and (FRow <= AValue.Row2) then
      Push(errRef)
    else begin
      if (FCol >= AValue.Col1) and (FCol <= AValue.Col2) then begin
        if FRow < AValue.Row1 then
          Push(AValue.Cells,FCol,AValue.Row1)
        else
          Push(AValue.Cells,FCol,AValue.Row2);
      end
      else if (FRow >= AValue.Row1) and (FRow <= AValue.Row2) then begin
        if FCol < AValue.Col1 then
          Push(AValue.Cells,AValue.Col1,FRow)
        else
          Push(AValue.Cells,AValue.Col2,FRow);
      end
      else
        Push(errRef);
    end;
//    PopResult(AValue);
  end;

  Result := AValue.RefType = xfrtNone;
  if Result and (AValue.ValType = xfvtUnknown) then begin
    AValue.ValType := xfvtFloat;
    AValue.vFloat := 0;
  end;
end;

function TValueStack.PopulateMatrix(AMatrix: TXLSMatrix; AValue: PXLSFormulaValue): boolean;
var
  R,C: integer;
  V: double;
  Empty: boolean;
begin
  Result := True;
  for R := AValue.Row1 to AValue.Row2 do begin
    for C := AValue.Col1 to AValue.Col2 do begin
      Result := GetCellValueFloat(AValue.Cells,C,R,V,Empty) = errUnknown;
      if not Result then begin
        Push(errValue);
        Exit;
      end;
      AMatrix[C - AValue.Col1,R - AValue.Row1] := V;
    end;
  end;
end;

procedure TValueStack.Push;
begin
  IncStackPtr;
  FStack[FStackPtr].Empty := False;
  FStack[FStackPtr].RefType := xfrtNone;
  FStack[FStackPtr].ValType := xfvtUnknown;
end;

procedure TValueStack.Push(const ACol, ARow: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.RefType := xfrtRef;
//  P.Cells := FManager.Worksheets.IdOrder[FSheetIndex].Cells;
  P.Cells := FManager.Worksheets[FSheetIndex].Cells;
  P.Col1 := ACol;
  P.Row1 := ARow;
  P.Col2 := ACol;
  P.Row2 := ARow;

  GetCellValue(P.Cells,ACol,ARow,P);
  P.RefType := xfrtRef;
end;

procedure TValueStack.Push(const AValue: boolean);
begin
  IncStackPtr;
  FStack[FStackPtr].Empty := False;
  FStack[FStackPtr].RefType := xfrtNone;
  FStack[FStackPtr].ValType := xfvtBoolean;
  FStack[FStackPtr].vBool := AValue;
end;

procedure TValueStack.Push(const AValue: TXc12CellError);
begin
  IncStackPtr;
  FStack[FStackPtr].Empty := False;
  FStack[FStackPtr].RefType := xfrtNone;
  FStack[FStackPtr].ValType := xfvtError;
  FStack[FStackPtr].vErr := AValue;
end;

procedure TValueStack.Push(const AValue: AxUCString);
begin
  IncStackPtr;
  FStack[FStackPtr].Empty := False;
  FStack[FStackPtr].RefType := xfrtNone;
  FStack[FStackPtr].ValType := xfvtString;
  FStack[FStackPtr].vStr := AValue;
end;

procedure TValueStack.SetError(AError: TXc12CellError);
begin
  FStack[FStackPtr].RefType := xfrtNone;
  FStack[FStackPtr].ValType := xfvtError;
  FStack[FStackPtr].vErr := AError;
end;

procedure TValueStack.SetResult(const ACount: integer);
begin
  FResultPtr := FStackPtr;
  Dec(FStackPtr,ACount);
  if FStackPtr < -1 then
    raise XLSRWException.Create('Empty stack');
end;

procedure TValueStack.SetupIfs(AArgCount: integer; var AData: TDynCriteriaDataArray);
var
  i: integer;
  Criteria: TXLSFormulaValue;
  SourceRange: TXLSFormulaValue;
  FV,PrevFV: TXLSFormulaValue;
  Iter: TValStackIterator;
begin
  SetLength(AData,AArgCount div 2);
  Iter := FIterator;

  for i := 0 to (AArgCount div 2) - 1 do begin
    Criteria := PopRes;
    case Criteria.RefType of
      xfrtNone,
      xfrtRef     : FV := Criteria;
      xfrtArea    : GetCellValue(Criteria.Cells,Criteria.Col1,Criteria.Row1,@FV);
      xfrtXArea   : ;
      xfrtAreaList: begin
        Push(errValue);
        Exit;
      end;
      xfrtArray   : FV := TXLSArrayItem(Criteria.vSource).GetAsFormulaValue(0,0);
    end;
    FVToVarValue(@FV,@AData[i].CritVal);

    AData[i].Op := xdcoEQ;
    if AData[i].CritVal.ValType = xfvtString then
      ConditionStrToVarValue(AData[i].CritVal.vStr,AData[i].Op,@AData[i].CritVal);

    FV := PopRes;

    if i > 0 then begin
      Iter.AddLinked;
      Iter := Iter.Linked;
      if not AreasEqualSize(@FV,@PrevFV) then begin
        Push(errValue);
        Exit;
      end;
    end;
    Iter.BeginIterate(FV,False);

    PrevFV := FV;
  end;

  if Odd(AArgCount) then begin
    SourceRange := PopRes;
    Iter.AddLinked;
    Iter := Iter.Linked;
    Iter.BeginIterate(SourceRange,False);
  end;
end;

function TValueStack.SolveChiInvEvent(AData: PDoubleArray; AValue: double): double;
begin
  Result := ChiSquaredDistributionPDF(AData[0],AValue)
end;

function TValueStack.SolveFInvEvent(AData: PDoubleArray; AValue: double): double;
begin
  Result := FDist(AValue,Trunc(AData[0]),Trunc(AData[1]));
end;

function TValueStack.StackSize: integer;
begin
  Result := FStackPtr;
end;

function TValueStack.Sum2Arrays(AFV1, AFV2: TXLSFormulaValue; out ASum1,ASum2: double): integer;
begin
  Result := 0;
  ASum1 := 0;
  ASum2 := 0;
  FIterator.BeginIterate(AFV1,AFV2,False);
  while FIterator.NextFloat do begin
    ASum1 := ASum1 + FIterator.Result.vFloat;
    ASum2 := ASum2 + FIterator.Linked.Result.vFloat;
    Inc(Result);
  end;
end;

function TValueStack.TDist(X: double; DegreesFreedom: Integer; TwoTails: Boolean): double;
begin
  if TwoTails then
    Result := (StudentTDistCDF(x,DegreesFreedom)) * 2
  else
    Result := 1 - StudentTDistCDF(x,DegreesFreedom);
end;

// 1;2
function TValueStack.TTest(Arr1, Arr2: TDynDoubleArray; Tails, Type_: integer): double;
var
  Sm1,Sd1,Sm2,Sd2: double;
  v,sp,t_stat: double;
  n: integer;
begin
  n := Length(Arr1);
  Sm1 := Mean(Arr1);
  Sm2 := Mean(Arr2);
  Sd1 := StdDev(Arr1);
  Sd2 := StdDev(Arr2);

  v := n + n - 2;

  sp := sqrt(((n - 1) * Sd1 * Sd1 + (n - 1) * Sd2 * Sd2) / v);
  t_stat := (Sm1 - Sm2) / (sp * sqrt(1.0 / n + 1 / n));
  Result := 1 - StudentTDistCDF(t_stat,v);
  if Tails = 2 then
    Result := Result / 2;
end;

function TValueStack.TTest2(Arr1, Arr2: TDynDoubleArray; Tails, Type_: integer): double;
var
  i: integer;
  Diff: TDynDoubleArray;
  MD,SD: double;

function StDev_(Arr: TDynDoubleArray; M : double) : double;
var
  D, SD, SD2, V : double;
  I             : Integer;
begin
  SD  := 0.0;  { Sum of deviations (used to reduce roundoff error) }
  SD2 := 0.0;  { Sum of squared deviations }

  for i := 0 to High(Arr) do
  begin
    D := Arr[I] - M;
    SD := SD + D;
    SD2 := SD2 + Sqr(D)
  end;

  V := (SD2 - Sqr(SD) / Length(Arr)) / (Length(Arr) - 1);  { Variance }
  Result := Sqrt(V);
end;

begin
  SetLength(Diff,Length(Arr1));

  for i := 0 to High(Arr1) do
    Diff[i] := Arr1[i] - Arr2[i];

  MD := Mean(Diff);
  SD := StDev_(Diff,MD);

  Result := MD * Sqrt(Length(Diff)) / SD;
end;

procedure TValueStack.DoT_Inv;
var
  DegreesFreedom: integer;
  Probability: double;
begin
  if not PopResPosInt(DegreesFreedom,1) then Exit;
  if not PopResFloat(Probability) then Exit;

  if (Probability < 0.0) or (Probability > 1.0) then
    Push(errNum)
  else
    Push(InverseStudentsT(DegreesFreedom,Probability,1 - Probability));
end;

procedure TValueStack.DoT_Inv_2T;
var
  DegreesFreedom: integer;
  Probability: double;
begin
  if not PopResPosInt(DegreesFreedom,1) then Exit;
  if not PopResFloat(Probability) then Exit;

  if (Probability < 0.0) or (Probability > 1.0) then
    Push(errNum)
  else begin
    Probability := Probability / 2;
    Push(-InverseStudentsT(DegreesFreedom,Probability,1 - Probability));
  end;
end;

procedure TValueStack.DoT_Test;
var
  tails: integer;
  type_: integer;
  Arr1: TDynDoubleArray;
  Arr2: TDynDoubleArray;
begin
  if not PopResInt(type_) then Exit;
  if not PopResInt(tails) then Exit;

  if (type_ < 1) or (type_ > 3) or (tails < 1) or (tails > 2) then
    Push(errNum)
  else begin
    if not DoCollectValues(Arr2) then
      Exit;
    if not DoCollectValues(Arr1) then
      Exit;
    if (Length(Arr1) <> Length(Arr2)) or (Length(Arr1) < 2) then
      Push(errNA)
    else
      Push(TTest(Arr1,Arr2,type_,tails));
  end;
end;

function TValueStack.VDB(const Cost, Salvage, Life, StartPeriod, EndPeriod, Factor: double; NoSwitch: boolean): double;
var
  VDB : Extended;
  SLD : Extended;
  Rate: Extended;
begin
  if (Factor = 0.0) then
    Rate := 2.0 / Life
  else
    Rate := Factor / Life;
  SLD := (Cost - Salvage) * (EndPeriod - StartPeriod) / Life;
  VDB := Cost * (Power(1.0 - Rate, StartPeriod) - Power(1.0 - Rate, EndPeriod));
  if (not NoSwitch) and (SLD > VDB) then
//  if NoSwitch then
    Result := SLD
  else
    Result := VDB;
end;

procedure TValueStack.WriteMatrix(AMatrix: TXLSMatrix; ACol, ARow: integer);
var
  C,R: integer;
  Arr: TXLSArrayItem;
begin
  Arr := TXLSArrayItem.Create(AMatrix.Cols,AMatrix.Rows);
  for R := 0 to AMatrix.Rows - 1 do begin
    for C := 0 to AMatrix.Cols - 1 do
      Arr.AsFloat[C,R] := AMatrix[C,R];
  end;
end;

function TValueStack.ZTest(Arr: TDynDoubleArray; x, sigma: double): double;
var
  i: integer;
  n: integer;
  u: double;
  Sum: double;
  SumSq: double;
begin
  n := Length(Arr);
  Sum := 0;
  SumSq := 0;
  for i := 0 to n - 1 do begin
    Sum := Sum + Arr[i];
    SumSq := Sumsq + Arr[i] * Arr[i];
  end;
  u := Sum / n;
  if Sigma = 1000 then begin
    Sigma := (SumSq - Sum * Sum / n) / (n - 1);
    Result := 1 - NormalCDF(((u - x) / Sqrt(sigma / n)),0,1)
  end
  else
    Result := 1 - NormalCDF(((u - x) * sqrt(n) / sigma),0,1);
end;

procedure TValueStack.Push(const AValue: double);
begin
  IncStackPtr;
  FStack[FStackPtr].Empty := False;
  FStack[FStackPtr].RefType := xfrtNone;
  FStack[FStackPtr].ValType := xfvtFloat;
  FStack[FStackPtr].vFloat := AValue;
end;

procedure TValueStack.OpUnary(AOperator: integer);
var
  V: double;
  Err: TXc12CellError;
begin
  GetAsFloat(FStack[FStackPtr],V,Err);
  if Err <> errUnknown then
    SetError(Err)
  else begin
    // xptgOpUPlus don't change value type. =+TRUE == TRUE (=-TRUE == -1)
    case AOperator of
      xptgOpUMinus : begin
        FStack[FStackPtr].RefType := xfrtNone;
        FStack[FStackPtr].ValType := xfvtFloat;
        FStack[FStackPtr].vFloat := -V;
      end;
      xptgOpPercent: begin
        FStack[FStackPtr].RefType := xfrtNone;
        FStack[FStackPtr].ValType := xfvtFloat;
        FStack[FStackPtr].vFloat := V * 0.01;
      end;
    end;
  end;
end;

{ TXLSFormulaDecoder }

constructor TXLSFormulaEvaluator.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
  FStack := TValueStack.Create(FManager);
end;

procedure TXLSFormulaEvaluator.DebugEvaluate(ADebugData : TFmlaDebugItems; ASheetIndex, ACol, ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea);
begin
  FVolatile := False;
  FDebugData := ADebugData;
  FDebugSteps := FDebugData.Steps;
  FSheetIndex := ASheetIndex;
  FCol := ACol;
  FRow := ARow;
  FTargetArea := ATargetArea;
  FStack.Clear(ATargetArea <> Nil,ASheetIndex,ACol,ARow);
  DoEvaluate(APtgs,APtgsSize,FTargetArea);
end;

destructor TXLSFormulaEvaluator.Destroy;
begin
  FStack.Free;
  inherited;
end;

procedure TXLSFormulaEvaluator.DoFunction(AId, AArgCount: integer);
var
  S: AxUCString;
  Arg1,Arg2: TXLSFormulaValue;
  E: TXc12CellError;
  N1,N2: integer;
  ValFloat1,ValFloat2: double;
  ValBool: boolean;
  FloatFormat: TFloatFormat;
begin
  G_XLSExcelFuncNames.ArgCount(AId,N1,N2);
  if AArgCount < N1 then begin
    FManager.Errors.Error(Format('%s(%d)',[G_XLSExcelFuncNames.FindName(AId),N1]),XLSERR_FMLA_MISSINGARG);
    Exit;
  end
  else if AArgCount > N2 then begin
    FManager.Errors.Error(Format('%s(%d)',[G_XLSExcelFuncNames.FindName(AId),N2]),XLSERR_FMLA_TOMANYGARGS);
    Exit;
  end;

  FStack.SetResult(AArgCount);
  case AId of
    000: begin  // COUNT
      FStack.Push(FStack.DoCount(AArgCount));
    end;
    001: begin  // IF
      if AArgCount = 3 then
        Arg2 := FStack.PopRes
      else begin
        Arg2.RefType := xfrtNone;
        Arg2.ValType := xfvtBoolean;
        Arg2.vBool := False;
      end;
      Arg1 := FStack.PopRes;
      E := FStack.GetAsBoolean(FStack.PopRes,ValBool);
      if E <> errUnknown then
        FStack.Push(E)
      else begin
        if ValBool then
          FStack.Push(Arg1)
        else
          FStack.Push(Arg2);
      end;
    end;
    004: begin  // SUM
      FStack.DoSum(AArgCount);
    end;
    005: begin  // AVERAGE
      FStack.DoAverage(AArgCount);
    end;
    006: begin  // MIN
      FStack.DoMin(AArgCount);
    end;
    007: begin  // MAX
      FStack.DoMax(AArgCount);
    end;

    008: begin  // ROW
      if AArgCount = 0 then
        // Shall NOT be +1. If that is required, the error is elsewhere.
        // Now it is. Can't remember why it was wrong.
        FStack.Push(FStack.FRow + 1)
      else begin
        Arg1 := FStack.PopRes;
        if Arg1.RefType <> xfrtNone then
          // Shall NOT be +1. If that is required, the error is elsewhere.
          // Now it is. Can't remember why it was wrong.
          FStack.Push(Arg1.Row1 + 1)
        else
          FStack.Push(errREF);
      end;
    end;
    009: begin  // COLUMN
      if AArgCount = 0 then
        // Shall NOT be +1. If that is required, the error is elsewhere.
        // Now it is. Can't remember why it was wrong.
        FStack.Push(FStack.FCol + 1)
      else begin
        Arg1 := FStack.PopRes;
        if Arg1.RefType <> xfrtNone then
          // Shall NOT be +1. If that is required, the error is elsewhere.
          // Now it is. Can't remember why it was wrong.
          FStack.Push(Arg1.Col1 + 1)
        else
          FStack.Push(errREF);
      end;
    end;
    011: begin  // NPV
      FStack.DoNPV(AArgCount,ValFloat1);
    end;
    012: begin  // STDEV
      FStack.DoStDev(AArgCount);
    end;
    013: begin  // DOLLAR
      if (AArgCount = 2) and not FStack.PopResFloat(ValFloat2) then
        Exit
      else
        ValFloat2 :=  FormatSettings.CurrencyDecimals;
      if FStack.GetAsFloat(FStack.PopRes,ValFloat1,E) and (E = errUnknown)then
        // Bug in FloatToStrF? if decimal digits is negative (-2 = rounding integer part,
        // 1285 -> 1300), the result has about 20 zeros after the decimal point.
        FStack.Push(FloatToStrF(ValFloat1,ffCurrency,15,Round(ValFloat2)))
      else if E <> errUnknown then
        FStack.Push(E);
    end;
    014: begin  // FIXED
      FloatFormat := ffNumber;
      ValBool := True;
      if (AArgCount = 3) and not FStack.PopResBoolean(ValBool) then
        Exit;
      if not ValBool then
        FloatFormat := ffFixed;
      if (AArgCount in [2,3]) and  not FStack.PopResFloat(ValFloat2) then
        Exit
      else
        ValFloat2 := FormatSettings.CurrencyDecimals;
      if FStack.GetAsFloat(FStack.PopRes,ValFloat1,E) and (E = errUnknown)then
        FStack.Push(FloatToStrF(ValFloat1,FloatFormat,15,Round(ValFloat2)))
      else if E <> errUnknown then
        FStack.Push(E);
    end;
    028: begin  // LOOKUP
      FStack.DoLookup(AArgCount);
    end;
    029: begin  // INDEX
      FStack.DoIndex(AArgCount);
    end;
    036: begin  // AND
      FStack.DoAND(AArgCount);
    end;
    037: begin  // OR
      FStack.DoOR(AArgCount);
    end;
    046: begin  // VAR
      FStack.DoVar(AArgCount);
    end;
    049: begin
      FStack.DoLinest(AArgCount);
    end;
    050: begin
      FStack.DoTrend(AArgCount);
    end;
    051: begin
      FStack.DoLogest(AArgCount);
    end;
    052: begin
      FStack.DoGrowth(AArgCount);
    end;

    056,        // PV
    057,        // FV
    058,        // NPER
    059,        // PMT
    060: begin  // RATE
      FStack.DoAnnuity(AId,AArgCount);
    end;
    062: begin  // IRR
      FStack.DoIRR(AArgCount);
    end;
    064: begin  // MATCH
      FStack.DoMatch(AArgCount);
    end;
    070: begin // WEEKDAY
      FStack.DoExtractDateTime(AId,AArgCount);
    end;
    078: begin  // OFFSET
      FStack.DoOffset(AArgCount);
    end;
    082: begin  // SEARCH
      FStack.DoSearch(AArgCount,False);
    end;
    100: begin  // CHOOSE
      FStack.DoChoose(AArgCount);
    end;
    101: begin  // HLOOKUP
      FStack.DoHLookup(AArgCount);
    end;
    102: begin  // VLOOKUP
      FStack.DoVLookup(AArgCount);
    end;
    109: begin  // LOG
      ValFloat1 := 10;
      if AArgCount = 2 then
        if not FStack.PopResFloat(ValFloat1) then Exit;
      if not FStack.PopResFloat(ValFloat2) then Exit;
      if ValFloat1 = 10 then
        FStack.Push(Log10(ValFloat2))
      else
        FStack.Push(LogN(ValFloat1,ValFloat2));
    end;
    115: begin  // LEFT
      N1 := 1;
      if AArgCount = 2 then
        if not FStack.PopResInt(N1) then Exit;
      if not FStack.PopResStr(S) then Exit;
      FStack.Push(Copy(S,1,N1));
    end;
    116: begin  // RIGHT
      N1 := 1;
      if AArgCount = 2 then
        if not FStack.PopResInt(N1) then Exit;
      if not FStack.PopResStr(S) then Exit;
      FStack.Push(RightStr(S,N1));
    end;
    120: begin  // SUBSTITUTE
      FStack.DoSubstitute(AArgCount);
    end;
    124: begin  // FIND
      FStack.DoSearch(AArgCount,True);
    end;
    125: begin  // CELL
      FStack.DoCell(AArgCount);
    end;
    144: begin  // DDB
      FStack.DoDDB(AArgCount);
    end;
    148: begin  // INDIRECT
      DoIndirect(AArgCount);
    end;
    167: begin  // IPMT
      FStack.DoIPMT(AArgCount);
    end;
    168: begin  // PPMT
      FStack.DoPPMT(AArgCount);
    end;
    169: begin  // COUNTA
      FStack.DoCounta(AArgCount);
    end;
    183: begin  // PRODUCT
      FStack.DoProduct(AArgCount);
    end;
    193: begin  // PRODUCT
      FStack.DoStdevp(AArgCount);
    end;
    194: begin  // VARP
      FStack.DoVarp(AArgCount);
    end;
    197: begin  // TRUNC
      FStack.DoTrunc(AArgCount);
    end;
    205: begin  // FINDB
      FStack.DoSearch(AArgCount,True);
    end;
    206: begin  // SEARCHB
      FStack.DoSearch(AArgCount,False);
    end;
    208: begin  // LEFTB
      N1 := 1;
      if AArgCount = 2 then
        if not FStack.PopResInt(N1) then Exit;
      if not FStack.PopResStr(S) then Exit;
      FStack.Push(Copy(S,1,N1));
    end;
    209: begin  // RIGHTB
      N1 := 1;
      if AArgCount = 2 then
        if not FStack.PopResInt(N1) then Exit;
      if not FStack.PopResStr(S) then Exit;
      FStack.Push(RightStr(S,N1));
    end;
    216: begin  // RANK
      FStack.DoRank(AArgCount);
    end;
    219: begin  // ADDRESS
      FStack.DoAddress(AArgCount);
    end;
    220: begin  // DAYS360
      FStack.DoDays360(AArgCount);
    end;
    222: begin  // VDB
      FStack.DoVDB(AArgCount);
    end;
    227: begin  // MEDIAN
      FStack.DoMedian(AArgCount);
    end;
    228: begin  // SUMPRODUCT
      FStack.DoSumProduct(AArgCount);
    end;
    247: begin  // DB
      FStack.DoDb(AArgCount);
    end;
    269: begin  // AVEDEV
      FStack.DoAveDev(AArgCount);
    end;
    270: begin  // BETADIST
      FStack.DoBetaDist(AArgCount,False);
    end;
    272: begin  // BETAINV
      FStack.DoBetaInv(AArgCount);
    end;
    317: FStack.DoProb(AArgCount);
    318: FStack.DoDevSq(AArgCount);
    319: FStack.DoGeomean(AArgCount);
    320: FStack.DoHarmean(AArgCount);
    321: FStack.DoSumsq(AArgCount);
    322: FStack.DoKurt(AArgCount);
    323: FStack.DoSkew(AArgCount);
    324: FStack.DoZtest(AArgCount);
    329: FStack.DoPercentRank(AArgCount);
    330: FStack.DoMode(AArgCount);
    336: FStack.DoConcatenate(AArgCount);
    344: FStack.DoSubtotal(AArgCount);
    345: FStack.DoSumif(AArgCount);
    354: FStack.DoRoman(AArgCount);
    359: FStack.DoHyperlink(AArgCount);
    361: FStack.DoAveragea(AArgCount);
    362: FStack.DoMaxa(AArgCount);
    363: FStack.DoMina(AArgCount);
    364: FStack.DoStdevpa(AArgCount);
    365: FStack.DoVarpa(AArgCount);
    366: FStack.DoStdeva(AArgCount);
    367: FStack.DoVara(AArgCount);
    818: FStack.DoAggregate(AArgCount);
    819: FStack.DoStdev_S(AArgCount);
    820: FStack.DoStdev_P(AArgCount);
    821: FStack.DoVar_S(AArgCount);
    822: FStack.DoVar_P(AArgCount);
    823: FStack.DoMode_Sngl(AArgCount);
    828: FStack.DoAverageIf(AArgCount);
    829: FStack.DoAverageIfs(AArgCount);
    830: FStack.DoCeiling_Precise(AArgCount);
    833: FStack.DoCountIfs(AArgCount);
    837: FStack.DoFloor_Precise(AArgCount);
    840: FStack.DoIso_Ceiling(AArgCount);
    841: FStack.DoMode_Mult(AArgCount);
    842: FStack.DoNetWorkdays_Intl(AArgCount);
    843: FStack.DoPercentrank_Exc(AArgCount);
    844: FStack.DoSumifs(AArgCount);
    845: FStack.DoWorkday_Int(AArgCount);
    846: FStack.DoBetaDist(AArgCount,True);
    847: FStack.DoBeta_Inv(AArgCount);
    861: FStack.DoPercentrank_Inc(AArgCount);
    863: FStack.DoRank_Eq(AArgCount);
    865: FStack.DoZ_Test(AArgCount);
    866: FStack.DoWorkday(AArgCount);
    867: FStack.DoNetWorkdays(AArgCount);
    868: FStack.DoAccrInt(AArgCount);
    870: FStack.DoYearfrac(AArgCount);
    871: FStack.DoXIRR(AArgCount);
    else FStack.Push(errNA);
  end;
end;

procedure TXLSFormulaEvaluator.DoArray(APtgs: PXLSPtgs);
var
  C,R: integer;
  FV: TXLSFormulaValue;
  Arr: TXLSArrayItem;
begin
  Arr := TXLSArrayItem.Create(PXLSPtgsArray(APtgs).Cols,PXLSPtgsArray(APtgs).Rows);

  FStack.SetResult(PXLSPtgsArray(APtgs).Cols * PXLSPtgsArray(APtgs).Rows);
  for R := PXLSPtgsArray(APtgs).Rows - 1 downto 0 do begin
    for C := PXLSPtgsArray(APtgs).Cols - 1 downto 0 do begin
      FV := FStack.PopRes;
      if FV.RefType <> xfrtNone then begin
        FStack.Push(errValue);
        Exit;
      end;

      if not Arr.Add(C,R,FV) then begin
        FStack.Push(errValue);
        Exit;
      end;
    end;
  end;
  FStack.Push(Arr);
end;

{$ifdef XLS_BIFF}
procedure TXLSFormulaEvaluator.DoArray97(var APtgs: PXLSPtgs);
var
  C,R: integer;
  Col,Row: integer;
  Arr: TXLSArrayItem;
begin
  Col := PPTGArray97(APtgs).Cols;
  Row := PPTGArray97(APtgs).Rows;
  APtgs := Pointer(NativeInt(Aptgs) + SizeOf(TPTGArray97));
  Arr := TXLSArrayItem.Create(Col + 1,Row + 1);
  for C := 0 to Col do begin
    for R := 0 to Row do begin
      case Byte(APtgs^) of
        $00: APtgs := Pointer(NativeInt(APtgs) + 9);
        $01: begin
          Arr.Add(C,R,PArrayFloat97(APtgs).Value);
          APtgs := Pointer(NativeInt(APtgs) + SizeOf(TArrayFloat97));
        end;
        $02: begin
          Arr.Add(C,R,ByteStrToWideString(@PArrayString97(APtgs).Data,PArrayString97(APtgs).Len));
          APtgs := Pointer(NativeInt(APtgs) + PArrayString97(APtgs).Len + SizeOf(TArrayString97) + 1);
        end;
        $04: begin
          APtgs := Pointer(NativeInt(APtgs) + 1);
          Arr.Add(C,R,Boolean(Byte(APtgs^)));
          APtgs := Pointer(NativeInt(APtgs) + 8);
        end;
        $10: begin
          APtgs := Pointer(NativeInt(APtgs) + 1);
          Arr.Add(C,R,TXc12CellError(Byte(APtgs^)));
          APtgs := Pointer(NativeInt(APtgs) + 8);
        end;
      end;
    end;
  end;
  FStack.Push(Arr);
end;
{$endif}

procedure TXLSFormulaEvaluator.DoArrayFunction(APtgs: PXLSPtgs; ATargetArea: PXLSFormulaArea);
var
  i: integer;
  S: AxUCString;
  FuncId: integer;
  ArgCount: integer;
  ArgType: TXLSExcelFuncArgType;
  Arr: TXLSArrayItem;
  FV: TXLSFormulaValue;
  OldStack: array of TXLSFormulaValue;
  C,R: integer;
  W,H: integer;
  Res: TXLSArrayItem;
begin
  case APtgs.Id of
    xptgFunc    : begin
      FuncId := PXLSPtgsFunc(APtgs).FuncId;
      ArgCount := G_XLSExcelFuncNames[FuncId].Min;
    end;
    xptgFuncVar : begin
      FuncId := PXLSPtgsFuncVar(APtgs).FuncId;
      ArgCount := PXLSPtgsFuncVar(APtgs).ArgCount;
    end;
    xptgUserFunc: begin
      FuncId := -1;
      ArgCount := PXLSPtgsUserFunc(APtgs).ArgCount;
    end;
    else begin
      FuncId := $FFFF;
      ArgCount := 0;
    end;
  end;

  if (FuncId = $FFFF) or (G_XLSExcelFuncNames[FuncId].Args = Nil) then begin
    FStack.Push(errNA);
    Exit;
  end;

  if FuncId in [83] then begin
    case APtgs.Id of
      xptgFunc    : DoFunction(FuncId);
      xptgFuncVar : DoFunction(FuncId,ArgCount);
    end;
    Exit;
  end;

  W := ATargetArea.Col2 - ATargetArea.Col1 + 1;
  H := ATargetArea.Row2 - ATargetArea.Row1 + 1;
  SetLength(OldStack,ArgCount);
  for i := ArgCount - 1 downto 0 do begin
    if FuncId < 0 then
      ArgType := xefatRef  // TODO What is the arg type for a user function?
    else begin
      if G_XLSExcelFuncNames[FuncId].Args.ArgCount = 255 then
        ArgType := xefatArea
      else if G_XLSExcelFuncNames[FuncId].Args.ArgCount <= ArgCount  then
        ArgType := G_XLSExcelFuncNames[FuncId].Args.Args[ArgCount - i - 1]
      else
        raise XLSRWException.Create('To many arguments in function');
    end;
    FV := FStack.Pop;
    // TODO Check how to treat xefatRange
    if (ArgType in [xefatRef,xefatValue]) and (FV.RefType in [xfrtArea,xfrtXArea,xfrtArray]) then begin
      Arr := TXLSArrayItem.CreateFV(FV);
      Arr.VectorMode := True;
      FStack.FillArrayItem(Arr,FV);
      OldStack[i].RefType := xfrtArrayArg;
      OldStack[i].vSource := Arr;
      FStack.FGarbage.Add(Arr);
    end
    else
      OldStack[i] := FV;
  end;

//  FStack.SetResult(ArgCount);
  Res :=  TXLSArrayItem.Create(W,H);
  for R := 0 to H - 1 do begin
    for C := 0 to W - 1 do begin
      for i := 0 to ArgCount - 1 do begin
        if OldStack[i].RefType = xfrtArrayArg then begin
          Arr := TXLSArrayItem(OldStack[i].vSource);
          if Arr.Hit[C,R] then
            FStack.Push(Arr.GetAsFormulaValue(C,R))
          else
            FStack.Push(errNA);
        end
        else
          FStack.Push(OldStack[i]);
      end;
      case APtgs.Id of
        xptgFunc    : DoFunction(FuncId);
        xptgFuncVar : DoFunction(FuncId,ArgCount);
        xptgUSerFunc: begin
          SetLength(S,PXLSPtgsUserFunc(APtgs).Len);
          Move(PXLSPtgsUserFunc(APtgs).Name[0],Pointer(S)^,PXLSPtgsUserFunc(APtgs).Len * 2);
          DoFunction(S,ArgCount);
        end;
      end;
      Res.Add(C,R,FStack.Pop);
    end;
  end;
  FStack.Push(Res);
end;

procedure TXLSFormulaEvaluator.DoDataTable(APtgs: PXLSPtgsDataTableFmla);
var
  C,R: longword;
  Cell: TXLSCellItem;
  Cells: TXLSCellMMU;
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
  SrcRef1: TXLSPtgsRef;
  SrcRef2: TXLSPtgsRef;
  IdRef1: TXLSPtgsRef;
  IdRef2: TXLSPtgsRef;
  FV: TXLSFormulaValue;
  Arr: TXLSArrayItem;
  Ok: boolean;
begin
  Cells := FManager.Worksheets[FSheetIndex].Cells;

  Arr := TXLSArrayItem.Create(FTargetArea.Col2 - FTargetArea.Col1 + 1,FTargetArea.Row2 - FTargetArea.Row1 + 1);
  IdRef1.Col := APtgs.R1Col;
  IdRef1.Row := APtgs.R1Row;
  IdRef2.Col := APtgs.R2Col;
  IdRef2.Row := APtgs.R2Row;

  // Row and Column
  if (APtgs.Options and Xc12FormulaTableOpt_DT2D) <> 0 then begin
    if Cells.FindCell(FTargetArea.Col1 - 1,FTargetArea.Row1 - 1,Cell) and (Cells.CellType(@Cell) in XLSCellTypeFormulas) then begin
      PtgsSz := Cells.GetFormulaPtgs(@Cell,Ptgs);
      if PtgsSz > 0 then begin
        SrcRef2.Col := FTargetArea.Col1 - 1;
        SrcRef1.Row := FTargetArea.Row1 - 1;
        for R := FTargetArea.Row1 to FTargetArea.Row2 do begin
          SrcRef2.Row := R;
          for C := FTargetArea.Col1 to FTargetArea.Col2 do begin
            SrcRef1.Col := C;
            FV := DoDataTableFormula(Ptgs,PtgsSz,@IdRef1,@IdRef2,@SrcRef1,@SrcRef2);
            Arr.Add(C - FTargetArea.Col1,R - FTargetArea.Row1,FV);
          end;
        end;
      end;
    end
    else begin
      for R := FTargetArea.Row1 to FTargetArea.Row2 do begin
        for C := FTargetArea.Col1 to FTargetArea.Col2 do
          Arr.AsFloat[C - FTargetArea.Col1,R - FTargetArea.Row1] := 0;
      end;
    end;
  end
  // Row
  else if (APtgs.Options and Xc12FormulaTableOpt_DTR) <> 0 then begin
    for R := FTargetArea.Row1 to FTargetArea.Row2 do begin
      Ok := False;
      if Cells.FindCell(FTargetArea.Col1 - 1,R,Cell) and (Cells.CellType(@Cell) in XLSCellTypeFormulas) then begin
        PtgsSz := Cells.GetFormulaPtgs(@Cell,Ptgs);
        if PtgsSz > 0 then begin
          SrcRef1.Row := FTargetArea.Row1 - 1;
          for C := FTargetArea.Col1 to FTargetArea.Col2 do begin
            SrcRef1.Col := C;
            FV := DoDataTableFormula(Ptgs,PtgsSz,@IdRef1,Nil,@SrcRef1,Nil);
            Arr.Add(C - FTargetArea.Col1,R - FTargetArea.Row1,FV);
          end;
          Ok := True;
        end;
      end;
      if not Ok then begin
        for C := FTargetArea.Col1 to FTargetArea.Col2 do
          Arr.AsFloat[C - FTargetArea.Col1,R - FTargetArea.Row1] := 0;
      end;
    end;
  end
  // Column
  else begin
    for C := FTargetArea.Col1 to FTargetArea.Col2 do begin
      Ok := False;
      if Cells.FindCell(C,FTargetArea.Row1 - 1,Cell) and (Cells.CellType(@Cell) in XLSCellTypeFormulas) then begin
        PtgsSz := Cells.GetFormulaPtgs(@Cell,Ptgs);
        if PtgsSz > 0 then begin
          SrcRef1.Col := FTargetArea.Col1 - 1;
          for R := FTargetArea.Row1 to FTargetArea.Row2 do begin
            SrcRef1.Row := R;
            FV := DoDataTableFormula(Ptgs,PtgsSz,@IdRef1,Nil,@SrcRef1,Nil);
            Arr.Add(C - FTargetArea.Col1,R - FTargetArea.Row1,FV);
          end;
          Ok := True;
        end;
      end;
      if not Ok then begin
        for R := FTargetArea.Row1 to FTargetArea.Row2 do
          Arr.AsFloat[C - FTargetArea.Col1,R - FTargetArea.Row1] := 0;
      end;
    end;
  end;
  FStack.Push(Arr);
end;

function TXLSFormulaEvaluator.DoDataTableFormula(APtgs: PXLSPtgs; APtgsSz: integer; AIdRef1,AIdRef2,ASrcRef1,ASrcRef2: PXLSPtgsRef): TXLSFormulaValue;
var
  L: integer;
  Ptgs: PXLSPtgs;
  NewPtgs: PXLSPtgs;
  Sz: integer;
begin
  GetMem(NewPtgs,APtgsSz);
  try
    Move(APtgs^,NewPtgs^,APtgsSz);
    Ptgs := NewPtgs;
    Sz := APtgsSz;
    while Sz > 0 do begin
      case Ptgs.Id of
        xptgRef: begin
          if (PXLSPtgsRef(Ptgs).Col = AIdRef1.Col) and (PXLSPtgsRef(Ptgs).Row = AIdRef1.Row) then begin
            PXLSPtgsRef(Ptgs).Col := ASrcRef1.Col;
            PXLSPtgsRef(Ptgs).Row := ASrcRef1.Row;
          end
          else if (AIdRef2 <> Nil) and (PXLSPtgsRef(Ptgs).Col = AIdRef2.Col) and (PXLSPtgsRef(Ptgs).Row = AIdRef2.Row) then begin
            PXLSPtgsRef(Ptgs).Col := ASrcRef2.Col;
            PXLSPtgsRef(Ptgs).Row := ASrcRef2.Row;
          end;
          L := SizeOf(TXLSPtgsRef);
        end;

        xptgStr       : L := FixedSzPtgsStr + PXLSPtgsStr(Ptgs).Len * 2;
        xptgUserFunc  : L := FixedSzPtgsUserFunc + PXLSPtgsUserFunc(Ptgs).Len * 2;

        else            L := G_XLSPtgsSize[Ptgs.Id];
      end;
      Ptgs := PXLSPtgs(NativeInt(Ptgs) + L);
      Dec(Sz,L);
    end;
    Result := DoEvaluate(NewPtgs,APtgsSz,Nil);
  finally
    FreeMem(NewPtgs);
  end;
end;

procedure TXLSFormulaEvaluator.DoFunction(const AName: AxUCString; AArgCount: integer);
var
  i: integer;
  Data: TXLSUserFuncData;
begin
  if Assigned(FUserFuncEvent) then begin
    Data := TXLSUserFuncData.Create;
    try
      SetLength(Data.FArgs,AArgCount);
      for i := 0 to AArgCount - 1 do
        Data.FArgs[i] := FStack.Pop;

      Data.FResult.ValType := xfvtError;
      Data.FResult.vErr := errName;

      FUserFuncEvent(AName,Data);

      FStack.Push(Data.FResult);
    finally
      Data.Free;
    end;
  end
  else
    FStack.Push(errNA);
end;

procedure TXLSFormulaEvaluator.DoIndirect(AArgCount: integer);
var
  p: integer;
  SheetId: integer;
  S: AxUCString;
  C1,R1,C2,R2: integer;
  RelC1,RelR1,RelC2,RelR2: boolean;
  Ref: AxUCString;
  RefOk: boolean;
  A1Mode: boolean;
begin
  A1Mode := True;
  if AArgCount = 2 then
    if not FStack.PopResBoolean(A1Mode) then Exit;
  if not FStack.PopResStr(Ref) then Exit;

  p := CPos('!',Ref);

  if p > 1 then begin
    S := Copy(Ref,1,p - 1);
    Ref := Copy(Ref,p + 1,MAXINT);
    if (S[1] = '"') or (S[1] = '''') then begin
      S := Copy(S,2,Length(S) - 2);
      p := CPos(']',S);
      if p > 1 then
        S := Copy(S,p + 1,MAXINT);
    end;
    SheetId := FManager.Worksheets.Find(S);
    if SheetId < 0 then begin
      FStack.Push(errRef);
      Exit;
    end;
  end
  else
    SheetId := -1;

  if A1Mode then begin
    if FIsExcel97 then
      RefOk := AreaStrToColRow97(Uppercase(Ref),C1,R1,C2,R2)
    else
      RefOk := AreaStrToColRow(Uppercase(Ref),C1,R1,C2,R2);
  end
  else begin
    RefOk := R1C1ToArea(Uppercase(Ref),C1,R1,C2,R2,RelC1,RelR1,RelC2,RelR2);
    if RefOk then begin
      if RelC1 then C1 := FCol + C1;
      if RelR1 then R1 := FRow + R1;
      if RelC2 then C2 := FCol + C2;
      if RelR2 then R2 := FRow + R2;
    end;
  end;

  if RefOk then begin
    if SheetId >= 0 then begin
      if (C1 = C2) and (R1 = R2) then
        FStack.Push(SheetId,C1,R1)
      else
        FStack.Push(SheetId,C1,R1,C2,R2);
    end
    else begin
      if (C1 = C2) and (R1 = R2) then
        FStack.Push(C1,R1)
      else
        FStack.Push(C1,R1,C2,R2);
    end;
  end
  else begin
    if SheetId < 0 then
      DoName(Ref,FSheetIndex)
    else
      DoName(Ref,SheetId);
  end;
end;

procedure TXLSFormulaEvaluator.DoName(const AName: AxUCString; const ASheetIndex: integer);
var
  Id: integer;
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
  FV: TXLSFormulaValue;
begin
  if ASheetIndex = FSheetIndex then begin
    Id := FManager.Workbook.DefinedNames.FindId(AName,ASheetIndex);
    if Id < 0 then begin
      Id := FManager.Workbook.DefinedNames.FindId(AName,-1);
//      if Id < 0 then begin
//        Id := FManager.Workbook.DefinedNames.FindId('_' + AName,ASheetIndex);
//        if Id < 0 then
//          Id := FManager.Workbook.DefinedNames.FindId('_' + AName,-1);
//      end;
    end;
  end
  else begin
    Id := FManager.Workbook.DefinedNames.FindId(AName,-1);
    if Id < 0 then begin
      Id := FManager.Workbook.DefinedNames.FindId(AName,ASheetIndex);
//      if Id < 0 then begin
//        Id := FManager.Workbook.DefinedNames.FindId('_' + AName,-1);
//        if Id < 0 then
//          Id := FManager.Workbook.DefinedNames.FindId('_' + AName,ASheetIndex);
  //      end;
    end;
  end;
  if Id >= 0 then begin
    Ptgs := FManager.Workbook.DefinedNames[Id].Ptgs;
    PtgsSz := FManager.Workbook.DefinedNames[Id].PtgsSz;
    FV := DoEvaluate(Ptgs,PtgsSz,Nil,True);
    FStack.Push(FV);
  end
  else
    FStack.Push(errRef);
end;

procedure TXLSFormulaEvaluator.DoTable(ATable: PXLSPtgsTable);
var
  FV: TXLSFormulaValue;
  Cnt: integer;
  AreaSpec: TXLSCellArea;
  AreaSrc: TXLSCellArea;
begin
  Cnt := ATable.ArgCount;

  ClearCellArea(AreaSpec);
  ClearCellArea(AreaSrc);
  while Cnt > 0 do begin
    FV := FStack.Pop;

    if FV.ValType = xfvtError then begin
      FStack.Push(FV.vErr);
      Exit;
    end
    else if FV.ValType = xfvtTableSpecial then begin
      // xtssNone means invalid value.
      if TXLSTableSpecialSpecifier(FV.vSpec) > xtssNone then begin
        if CellAreaAssigned(AreaSpec) then
          ExtendCellArea(AreaSpec,FV.Col1,FV.Row1,FV.Col2,FV.Row2)
        else
          SetCellArea(AreaSpec,FV.Col1,FV.Row1,FV.Col2,FV.Row2);
      end;
    end
    else begin
      if CellAreaAssigned(AreaSpec) then begin
        AreaSrc.Col1 := FV.Col1;
        AreaSrc.Col2 := FV.Col2;
        AreaSrc.Row1 := AreaSpec.Row1;
        AreaSrc.Row2 := AreaSpec.Row2;
      end
      else
        SetCellArea(AreaSrc,FV.Col1,FV.Row1,FV.Col2,FV.Row2);
    end;

    Dec(Cnt);
  end;
  if CellAreaAssigned(AreaSpec) then begin
    if not CellAreaAssigned(AreaSrc) then begin
      AreaSrc.Col1 := AreaSpec.Col1;
      AreaSrc.Col2 := AreaSpec.Col2;
    end;
    AreaSrc.Row1 := AreaSpec.Row1;
    AreaSrc.Row2 := AreaSpec.Row2;
  end;

  if CellAreaAssigned(AreaSrc) or CellAreaAssigned(AreaSpec) then begin
    if not CellAreaAssigned(AreaSrc) then
      AreaSrc := AreaSpec;
    if ATable.Sheet > XLS_MAXSHEETS then
      FStack.Push(AreaSrc.Col1,AreaSrc.Row1,AreaSrc.Col2,AreaSrc.Row2)
    else
      FStack.Push(ATable.Sheet,AreaSrc.Col1,AreaSrc.Row1,AreaSrc.Col2,AreaSrc.Row2);
  end
  else
    FStack.Push(errRef);

//  if CellAreaAssigned(AreaSrc) then
//    FStack.Push(AreaSrc.Col1,AreaSrc.Row1,AreaSrc.Col2,AreaSrc.Row2)
//  else if CellAreaAssigned(AreaSpec) then
//    FStack.Push(AreaSpec.Col1,AreaSpec.Row1,AreaSpec.Col2,AreaSpec.Row2)
//  else
//    FStack.Push(errRef);
end;

procedure TXLSFormulaEvaluator.DoFunction(AId: integer);
var
  ArgCnt: integer;
  ValFloat1,ValFloat2: double;
  I1,I2: integer;
  S1,S2: AxUCString;
  ValBool: boolean;
  FV: TXLSFormulaValue;
  FVT: TXLSFormulaValueType;
  DT: TDateTime;
begin
  G_XLSExcelFuncNames.ArgCount(AId,ArgCnt,ArgCnt);
  FStack.SetResult(ArgCnt);
  case AId of
    002: begin  // ISNA
      FStack.Push(FStack.GetAsError(FStack.PopRes) = errNA);
    end;
    003: begin  // ISERROR
      FStack.Push(FStack.GetAsError(FStack.PopRes) <> errUnknown);
    end;
    010: begin  // NA
      FStack.Push(errNA);
    end;
    015: begin  // SIN
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Sin(ValFloat1));
    end;
    016: begin  // COS
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Cos(ValFloat1));
    end;
    017: begin  // TAN
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Tan(ValFloat1));
    end;
    018: begin  // ATAN
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(ArcTan(ValFloat1));
    end;
    019: begin  // PI
      FStack.Push(PI);
    end;
    020: begin // SQRT
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Sqrt(ValFloat1));
    end;
    021: begin // EXP
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Exp(ValFloat1));
    end;
    022: begin // LN
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Ln(ValFloat1));
    end;
    023: begin // LOG10
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Log10(ValFloat1));
    end;
    024: begin // ABS
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Abs(ValFloat1));
    end;
    025: begin // INT
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Int(ValFloat1));
    end;
    026: begin // SIGN
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Sign(ValFloat1));
    end;
    027: begin // ROUND
      if not FStack.PopResFloat(ValFloat2) then
        Exit;
      if (ValFloat2 < -20) or (ValFloat2 > 20) then begin
        FStack.Push(errValue);
        Exit;
      end;
{$ifdef DELPHI_5}
      raise XLSRWException.Create('ROUND not supported by Delphi 5');
{$else}
      if FStack.PopResFloat(ValFloat1) then begin
        if ValFloat2 = 0 then
          FStack.Push(SimpleRoundTo(ValFloat1,0))
        else
          FStack.Push(RoundTo(ValFloat1,Round(-ValFloat2)));
      end;
{$endif}
    end;
    030: begin  // REPT
      if FStack.PopResInt(I1) and FStack.PopResStr(S1) then begin
        S2 := '';
        for I2 := 0 to I1 - 1 do
          S2 := S2 + S1;
        FStack.Push(S2);
      end;
    end;
    031: begin  // MID
      if FStack.PopResInt(I1) and FStack.PopResInt(I2) and FStack.PopResStr(S1) then begin
        S2 := Copy(S1,I2,I1);
        FStack.Push(S2);
      end;
    end;
    032: begin  // LEN
      if FStack.PopResStr(S1) then
        FStack.Push(Length(S1));
    end;
    033: begin  // VALUE
      FStack.DoValue;
    end;
    034: begin  // TRUE
      FStack.Push(True);
    end;
    035: begin  // FALSE
      FStack.Push(False);
    end;
    038: begin  // NOT
      // TODO Set correct error value on non-boolean cells.
      if FStack.PopResBoolean(ValBool) then
        FStack.Push(not ValBool)
      else
        FStack.Push(True);
    end;
    039: begin  // MOD
      if FStack.PopResFloat(ValFloat1) and FStack.PopResFloat(ValFloat2) then begin
        if ValFloat1 = 0 then
          FStack.Push(errDiv0)
        else
          FStack.Push(ValFloat2 - Trunc(ValFloat2 / ValFloat1) * ValFloat1);
      end;
    end;
    040,        // DCOUNT
    041,        // DSUM
    042,        // DAVERAGE
    043,        // DMIN
    044,        // DMAX
    045: begin  // DSTDEV
      FStack.DoDatabase(AId);
    end;

    047: begin // DVAR
      FStack.DoDatabase(AId);
    end;
    048: begin // DVAR
      FStack.DoText;
    end;

    061: begin // MIRR
      FStack.DoMIRR;
    end;
    063: begin // RAND
      FVolatile := True;
      FStack.Push(Random(MAXINT) / (MAXINT - 1));
    end;
    065: begin // DATE
      FVolatile := True;
      FStack.DoDate;
    end;
    066: begin // TIME
      FVolatile := True;
      FStack.DoTime;
    end;
    067,       // DAY
    068,       // MONTH
    069,       // YEAR
    071,       // HOUR
    072,       // MINUTE
    073: begin // SECOND
      FVolatile := True;
      FStack.DoExtractDateTime(AId,1);
    end;
    074: begin // NOW
      FVolatile := True;
      FStack.Push(Now);
    end;
    075: begin // AREAS
      FStack.DoAreas;
    end;
    076,       // ROWS
    077: begin // COLUMNS
      FStack.DoRowsColumns(AId);
    end;
    083: begin // TRANSPOSE
      FStack.DoTranspose;
    end;
    086: begin // TYPE
      FStack.DoType;
    end;
    097: begin // ATAN2
      if not FStack.PopResFloat(ValFloat1) then
        Exit;
      if not FStack.PopResFloat(ValFloat2) then
        Exit;
      FStack.Push(ArcTan2(ValFloat1,ValFloat2));
    end;
    098: begin // ASIN
      if not FStack.PopResFloat(ValFloat1) then
        Exit;
      FStack.Push(ArcSin(ValFloat1));
    end;
    099: begin // ACOS
      if not FStack.PopResFloat(ValFloat1) then
        Exit;
      FStack.Push(ArcCos(ValFloat1));
    end;
    105: begin // ISREF
      FV := FStack.PopRes;
      FStack.Push(FV.RefType in [xfrtRef,xfrtArea,xfrtXArea]);
    end;
    111: begin // CHAR
      if not FStack.PopResInt(I1) then
        Exit;
      FStack.Push(Char(I1));
    end;
    112: begin // LOWER
      if not FStack.PopResStr(S1) then
        Exit;
      FStack.Push(Lowercase(S1));
    end;
    113: begin // UPPER
      if not FStack.PopResStr(S1) then
        Exit;
      FStack.Push(Uppercase(S1));
    end;
    114: begin // PROPER
      if not FStack.PopResStr(S1) then
        Exit;
      FStack.Push(S1);
    end;
    117: begin // EXACT
      if not FStack.PopResStr(S1) then
        Exit;
      if not FStack.PopResStr(S2) then
        Exit;
      FStack.Push(S1 = S2);
    end;
    118: begin // TRIM
      if not FStack.PopResStr(S1) then
        Exit;
      FStack.Push(Trim(S1));
    end;
    119: begin // REPLACE
      FStack.DoReplace;
    end;
    121: begin // CODE
      if not FStack.PopResStr(S1,False) then
        Exit;
      FStack.Push(Ord(S1[1]));
    end;
    126: begin // ISERR
      FV := FStack.PopRes;
      FStack.Push(FV.ValType = xfvtError);
    end;
    127: begin // ISTEXT
      FV := FStack.PopRes;
      FStack.Push(FV.ValType = xfvtstring);
    end;
    128: begin // ISNUMBER
      FV := FStack.PopRes;
      FStack.Push(FV.ValType = xfvtFloat);
    end;
    129: begin // ISBLANK
      FV := FStack.PopRes;
      if FV.RefType in [xfrtRef,xfrtArea] then begin
        FStack.MakeIntersect(@FV);
        FStack.Push(FStack.GetCellBlank(FV.Cells,FV.Col1,FV.Row1));
      end
      else
        FStack.Push(False);
    end;
    130: begin // T
      FV := FStack.PopRes;
      if FV.ValType = xfvtString then
        FStack.Push(FV.vStr)
      else
        FStack.Push('');
    end;
    131: begin // N
      FV := FStack.PopRes;
      if FV.ValType = xfvtFloat then
        FStack.Push(FV.vFloat)
      else
        FStack.Push(0);
    end;
    140: begin // DATEVALUE
      if not FStack.PopResStr(S1) then
        Exit;
      if TryStrToDate(S1,DT) then
        FStack.Push(Trunc(DT))
      else
        FStack.Push(errValue);
    end;
    141: begin // TIMEVALUE
      if not FStack.PopResStr(S1) then
        Exit;
      if TryStrToTime(S1,DT) then
        FStack.Push(Frac(DT))
      else
        FStack.Push(errValue);
    end;
    142: begin // SLN
      FStack.DoSLN;
    end;
    143: begin // SYD
      FStack.DoSYD;
    end;
    162: begin // CLEAN
      FStack.DoClean;
    end;
    163: begin // MDETERM
      FStack.DoMDeterm;
    end;
    164: begin // MINVERSE
      FStack.DoMInverse;
    end;
    165: begin // MMULT
      FStack.DoMMult;
    end;
    184: begin // FACT
      FStack.DoFact;
    end;

    189: begin // DPRODUCT
      FStack.DoDatabase(AId);
    end;
    190: begin // ISNONTEXT
      FStack.GetAsValueType(FStack.PopRes,FVT);
      FStack.Push(not (FVT = xfvtString));
    end;
    195: begin // DSTDEVP
      FStack.DoDatabase(AId);
    end;
    196: begin // DVARP
      FStack.DoDatabase(AId);
    end;
    198: begin // ISLOGICAL
      FV := FStack.PopRes;
      FStack.Push(FV.ValType = xfvtBoolean);
    end;
    199: begin // DCOUNTA
      FStack.DoDatabase(AId);
    end;
    207: begin  // REPLACEB
      FStack.DoReplace;
    end;
    210: begin  // MIDB
      if FStack.PopResInt(I1) and FStack.PopResInt(I2) and FStack.PopResStr(S1) then begin
        S2 := Copy(S1,I2,I1);
        FStack.Push(S2);
      end;
    end;
    211: begin  // LENB
      if FStack.PopResStr(S1) then
        FStack.Push(Length(S1));
    end;
    212: begin  // ROUNDUP
      FStack.DoRoundUp;
    end;
    213: begin  // ROUNDDOWN
      FStack.DoTrunc(2);
    end;
    221: begin  // TODAY
      FVolatile := True;
      FStack.Push(Date);
    end;
    229: begin  // SINH
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Sinh(ValFloat1));
    end;
    230: begin  // COSH
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Cosh(ValFloat1));
    end;
    231: begin  // TANH
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(Tanh(ValFloat1));
    end;
    232: begin  // ASINH
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(ArcSinh(ValFloat1));
    end;
    233: begin  // ACOSH
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(ArcCosh(ValFloat1));
    end;
    234: begin  // ATANH
      if FStack.PopResFloat(ValFloat1) then
        FStack.Push(ArcTanh(ValFloat1));
    end;
    235: begin // DGET
      FStack.DoDatabase(AId);
    end;
    244: begin // INFO
      FStack.DoInfo;
    end;
    252: begin // FREQUENZY
      FStack.DoFrequency;
    end;
    261: begin // ERROR.TYPE
      FStack.DoError_Type;
    end;
    271: begin // GAMMALN
      if FStack.PopResFloat(ValFloat1) then begin
        if ValFloat1 <= 0 then
          FStack.Push(errNum)
        else
          FStack.Push(Ln(TGamma(ValFloat1)));
      end;
    end;
    273: FStack.DoBinomdist;
    274: FStack.DoChidist;
    275: FStack.DoChiInv;
    276: FStack.DoCombin;
    277: FStack.DoConfidence;
    278: FStack.DoCritbinom;
    279: begin // EVEN
      if FStack.PopResFloat(ValFloat1) then begin
        if ValFloat1 < 0 then
          FStack.Push(Floor(ValFloat1 / 2) * 2)
        else
          FStack.Push(Ceil(ValFloat1 / 2) * 2);
      end;
    end;
    280: FStack.DoExpondist;
    281: FStack.DoFDist;
    282: FStack.DoFInv;
    283: FStack.DoFisher;
    284: FStack.DoFisherInv;
    285: FStack.DoFloor;
    286: FStack.DoGammadist;
    287: FStack.DoGammaInv;
    289: FStack.DoHypGeomDist;
    288: FStack.DoCeiling;
    290: FStack.DoLognormdist;
    291: FStack.DoLoginv;
    292: FStack.DoNegbinomdist;
    293: FStack.DoNormdist;
    294: FStack.DoNormsdist;
    295: FStack.DoNormInv;
    296: FStack.DoNormsInv;
    297: FStack.DoStandardize;
    298: FStack.DoOdd;
    299: FStack.DoPermut;
    300: FStack.DoPoisson;
    301: FStack.DoTDist;
    302: FStack.DoWeibull;
    303: FStack.DoSumxmy2;
    304: FStack.DoSumx2my2;
    305: FStack.DoSumx2py2;
    306: FStack.DoChitest;
    307: FStack.DoCorrel;
    308: FStack.DoCovar;
    309: FStack.DoForecast;
    310: FStack.DoFtest;
    311: FStack.DoIntercept;
    312: FStack.DoCorrel;  // PEARSON Same as CORREL
    313: FStack.DoRsq;
    314: FStack.DoSteyx;
    315: FStack.DoSlope;
    316: FStack.DoTtest;
    325: FStack.DoLarge;
    326: FStack.DoSmall;
    327: FStack.DoQuartile;
    328: FStack.DoPercentile;
    331: FStack.DoTrimMean;
    332: FStack.DoTinv;
    337: FStack.DoPower;
    342: FStack.DoRadians;
    343: FStack.DoDegrees;
    346: FStack.DoCountif;
    347: FStack.DoCountblank;
    350: FStack.DoIspmt;
    360: FStack.DoPhonetic;
    381: FStack.DoQuotient;

    800: FStack.DoT_Dist_2T;
    801: FStack.DoT_Dist_RT;
    802: FStack.DoT_Dist;
    803: FStack.DoT_Inv;
    804: FStack.DoT_Inv_2T;
    805: FStack.DoChiSq_Inv;
    806: FStack.DoChi_Inv_RT;
    807: FStack.DoGamma_Inv;
    808: FStack.DoLognorm_Inv;
    809: FStack.DoNorm_Inv;
    810: FStack.DoNorm_S_Inv;
    811: FStack.DoF_Inv;
    812: FStack.DoF_Inv_RT;
    813: FStack.DoBinom_Inv;
    814: FStack.DoLognorm_dist;
    815: FStack.DoF_Test;
    816: FStack.DoErfc;
    817: FStack.DoErfc_Precise;
    824: FStack.DoPercentile_Inc;
    825: FStack.DoQuartile_Inc;
    826: FStack.DoPercentile_Exc;
    827: FStack.DoQuartile_Exc;
    831: FStack.DoChisq_Dist;
    832: FStack.DoConfidence_T;
    834: FStack.DoCovariance_S;
    835: FStack.DoErf_Precise;
    836: FStack.DoF_Dist;
    838: FStack.DoGammaln_Precise;
    839: FStack.DoIferror;
    848: FStack.DoBinom_Dist;
    849: FStack.DoChisq_Dist_rt;
    850: FStack.DoChisq_Test;
    851: FStack.DoConfidence_Norm;
    853: FStack.DoCovariance_P;
    854: FStack.DoExpon_Dist;
    855: FStack.DoF_dist_RT;
    856: FStack.DoGamma_Dist;
    857: FStack.DoHypgeom_Dist;
    858: FStack.DoNegbinom_Dist;
    859: FStack.DoNorm_Dist;
    860: FStack.DoNorm_S_Dist;
    862: FStack.DoPoisson_Dist;
    864: FStack.DoT_Test;
    865: FStack.DoWeibull_Dist;
    869: FStack.DoEOMonth;
    872: FStack.DoXNVP;
    873: FStack.DoEDATE;
    else FStack.Push(errNA);
  end;
end;

// =IF(ISBLANK(U40);10;OK)
function TXLSFormulaEvaluator.DoEvaluate(APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea; AIsName: boolean = False): TXLSFormulaValue;
var
  i1,i2: integer;
  S: AxUCString;
  P: Pointer;
  B: Byte;
  C1,R1,C2,R2: integer;
  LastPtgs: NativeInt;
  XValue: TXc12XCellData;
  Ptgs: PXLSPtgs;
  Sz: integer;
  FirstPtgs: PXLSPtgs;
  PrevPtgs: PXLSPtgs;
  ArrayData97: PXLSPtgs;
  Name: TXc12DefinedName;
begin
  FirstPtgs := APtgs;
  PrevPtgs := APtgs;
  LastPtgs := NativeInt(APtgs) + APtgsSize;
  while NativeInt(APtgs) < LastPtgs do begin
    if FDebugData <> Nil then
      FDebugData.CurrPtgsOffs := NativeInt(PrevPtgs) - NativeInt(FirstPtgs);
    if FIsExcel97 then
      B := GetBasePtgs(APtgs.Id)
    else
      B := APtgs.Id;
    case B of
      xptgNone    : raise XLSRWException.Create('Illegal ptgs');
      xptgOpAdd,
      xptgOpSub,
      xptgOpMult,
      xptgOpDiv,
      xptgOpPower : begin
        if ATargetArea <> Nil then
          FStack.OpArrayFloat(APtgs.Id)
        else
          FStack.OpFloat(APtgs.Id);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptgOpLT,
      xptgOpLE,
      xptgOpEQ,
      xptgOpGE,
      xptgOpGT,
      xptgOpNE: begin
        if ATargetArea <> Nil then
          FStack.OpArrayBoolean(APtgs.Id)
        else
          FStack.OpBoolean(APtgs.Id);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      xptgOpConcat: begin
        if ATargetArea <> Nil then
          FStack.OpArrayConcat
        else
          FStack.OpConcat;
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      // Unary plus shall be ignored. Can occure in formulas such as +IF(A1>2,"Yes","No")
      xptgOpUPlus  : APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      xptgOpUMinus,
      xptgOpPercent: begin
        if ATargetArea <> Nil then
          FStack.OpArrayUnary(APtgs.Id)
        else
          FStack.OpUnary(APtgs.Id);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      xptgOpIsect: begin
        FStack.OpArea(APtgs.Id);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsISect));
      end;
      xptgOpUnion,
      xptgOpRange: begin
        FStack.OpArea(APtgs.Id);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      xptgLPar:  APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));

      xptgStr: begin
        SetLength(S,PXLSPtgsStr(APtgs).Len);
        Move(PXLSPtgsStr(APtgs).Str[0],Pointer(S)^,PXLSPtgsStr(APtgs).Len * 2);
        FStack.Push(S);
        APtgs := PXLSPtgs(NativeInt(APtgs) + FixedSzPtgsStr + PXLSPtgsStr(APtgs).Len * 2);
      end;
      xptgErr: begin
        FStack.Push(TXc12CellError(PXLSPtgsErr(APtgs).Value));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsErr));
      end;
      xptgBool: begin
        FStack.Push(PXLSPtgsBool(APtgs).Value = 1);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsBool));
      end;

      xptgInt: begin
        FStack.Push(PXLSPtgsInt(APtgs).Value);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsInt));
      end;
      xptgNum: begin
        FStack.Push(PXLSPtgsNum(APtgs).Value);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsNum));
      end;

      xptgFunc: begin
        if (ATargetArea <> Nil) and (G_XLSExcelFuncNames[PXLSPtgsFunc(APtgs).FuncId].Type_ <> xeftArray) then
          DoArrayFunction(APtgs,ATargetArea)
        else
          DoFunction(PXLSPtgsFunc(APtgs).FuncId);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsFunc));
      end;
      xptgFuncVar: begin
        if (ATargetArea <> Nil) and (G_XLSExcelFuncNames[PXLSPtgsFuncVar(APtgs).FuncId].Type_ <> xeftArray) then
          DoArrayFunction(APtgs,ATargetArea)
        else
          DoFunction(PXLSPtgsFuncVar(APtgs).FuncId,PXLSPtgsFuncVar(APtgs).ArgCount);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsFuncVar));
      end;
      xptgUserFunc: begin
        if ATargetArea <> Nil then
          DoArrayFunction(APtgs,ATargetArea)
        else begin
          SetLength(S,PXLSPtgsUserFunc(APtgs).Len);
          Move(PXLSPtgsUserFunc(APtgs).Name[0],Pointer(S)^,PXLSPtgsUserFunc(APtgs).Len * 2);
          DoFunction(S,PXLSPtgsUserFunc(APtgs).ArgCount);
        end;
        APtgs := PXLSPtgs(NativeInt(APtgs) + FixedSzPtgsUserFunc + PXLSPtgsUserFunc(APtgs).Len * 2)
      end;
      xptgName: begin
        if PXLSPtgsName(APtgs).NameId = XLS_NAME_UNKNOWN then
          FStack.Push(errName)
        else begin
          Name := FManager.Workbook.DefinedNames[PXLSPtgsName(APtgs).NameId];

          case Name.SimpleName of
            xsntNone : begin
              Ptgs := Name.Ptgs;
              Sz := Name.PtgsSz;
              FStack.Push(DoEvaluate(Ptgs,Sz,ATargetArea,True));
            end;
            xsntRef  : begin
//              FStack.Push(Name.SimpleArea.SheetIndex,Name.SimpleArea.Col1,Name.SimpleArea.Row1);

              if (Name.SimpleArea.Col1 and COL_ABSFLAG) = COL_ABSFLAG then
                C1 := Name.SimpleArea.Col1 and not COL_ABSFLAG
              else
                C1 := FCol;
              if (Name.SimpleArea.Row1 and ROW_ABSFLAG) = ROW_ABSFLAG then
                R1 := Name.SimpleArea.Row1 and not ROW_ABSFLAG
              else
                R1 := FRow;
              FStack.Push(Name.SimpleArea.SheetIndex,C1,R1);
            end;
            xsntArea : FStack.Push(Name.SimpleArea.SheetIndex,Name.SimpleArea.Col1 and not COL_ABSFLAG,Name.SimpleArea.Row1 and not ROW_ABSFLAG,Name.SimpleArea.Col2 and not COL_ABSFLAG,Name.SimpleArea.Row2 and not ROW_ABSFLAG);
            xsntError: FStack.Push(errRef);
          end;
        end;

        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsName));
      end;
      xptgRef: begin
        if AIsName then begin
          if (PXLSPtgsRef(APtgs).Col and COL_ABSFLAG) = COL_ABSFLAG then
            C1 := PXLSPtgsRef(APtgs).Col and not COL_ABSFLAG
          else
            C1 := FCol;
          if (PXLSPtgsRef(APtgs).Row and ROW_ABSFLAG) = ROW_ABSFLAG then
            R1 := PXLSPtgsRef(APtgs).Row and not ROW_ABSFLAG
          else
            R1 := FRow;
          FStack.Push(C1,R1);
        end
        else
          FStack.Push(PXLSPtgsRef(APtgs).Col and not COL_ABSFLAG,PXLSPtgsRef(APtgs).Row and not ROW_ABSFLAG);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef));
      end;
      xptgArea: begin
        if AIsName then begin
          if (PXLSPtgsArea(APtgs).Col1 and COL_ABSFLAG) = COL_ABSFLAG then
            C1 := PXLSPtgsArea(APtgs).Col1 and not COL_ABSFLAG
          else
            C1 := FCol;
          if (PXLSPtgsArea(APtgs).Col2 and COL_ABSFLAG) = COL_ABSFLAG then
            C2 := PXLSPtgsArea(APtgs).Col2 and not COL_ABSFLAG
          else
            C2 := PXLSPtgsArea(APtgs).Col2 + FCol;
          if (PXLSPtgsArea(APtgs).Row1 and ROW_ABSFLAG) = ROW_ABSFLAG then
            R1 := PXLSPtgsArea(APtgs).Row1 and not ROW_ABSFLAG
          else
            R1 := FRow;
          if (PXLSPtgsArea(APtgs).Row2 and ROW_ABSFLAG) = ROW_ABSFLAG then
            R2 := PXLSPtgsArea(APtgs).Row2 and not ROW_ABSFLAG
          else
            R2 := NativeInt(PXLSPtgsArea(APtgs).Row2) + FRow;
          if (C2 > XLS_MAXCOL) then
            C2 := XLS_MAXCOL;
          if (R2 > XLS_MAXROW) then
            R2 := XLS_MAXROW;
          FStack.Push(C1,R1,C2,R2);
        end
        else
          FStack.Push(PXLSPtgsArea(APtgs).Col1 and not COL_ABSFLAG,PXLSPtgsArea(APtgs).Row1 and not ROW_ABSFLAG,PXLSPtgsArea(APtgs).Col2 and not COL_ABSFLAG,PXLSPtgsArea(APtgs).Row2 and not ROW_ABSFLAG);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea));
      end;

      xptgRefErr: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef));
      end;
      xptgAreaErr: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea));
      end;

      xptgWS: APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsWS));
//      xptgRPar: ;
      xptgRef1d: begin
        if AIsName and (FSheetIndex = PXLSPtgsRef1d(APtgs).Sheet) then begin
          if (PXLSPtgsRef(APtgs).Col and COL_ABSFLAG) = COL_ABSFLAG then
            C1 := PXLSPtgsRef(APtgs).Col and not COL_ABSFLAG
          else
            C1 := FCol;
          if (PXLSPtgsRef(APtgs).Row and ROW_ABSFLAG) = ROW_ABSFLAG then
            R1 := PXLSPtgsRef(APtgs).Row and not ROW_ABSFLAG
          else
            R1 := FRow - 1;
          FStack.Push(C1,R1);
        end
        else
          FStack.Push(PXLSPtgsRef1d(APtgs).Sheet,PXLSPtgsRef1d(APtgs).Col and not COL_ABSFLAG,PXLSPtgsRef1d(APtgs).Row and not ROW_ABSFLAG);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef1d));
      end;
      xptgRef1dErr: begin
        FStack.Push(TXc12CellError(PXLSPtgsRef1dError(APtgs).Error));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef1dError));
      end;
      xptgArea1d: begin
        if AIsName and (FSheetIndex = PXLSPtgsArea1d(APtgs).Sheet) then begin
          if (PXLSPtgsArea(APtgs).Col1 and COL_ABSFLAG) = COL_ABSFLAG then
            C1 := PXLSPtgsArea(APtgs).Col1 and not COL_ABSFLAG
          else
            C1 := FCol;
          if (PXLSPtgsArea(APtgs).Col2 and COL_ABSFLAG) = COL_ABSFLAG then
            C2 := PXLSPtgsArea(APtgs).Col2 and not COL_ABSFLAG
          else
            C2 := PXLSPtgsArea(APtgs).Col2 + FCol;
          if (PXLSPtgsArea(APtgs).Row1 and ROW_ABSFLAG) = ROW_ABSFLAG then
            R1 := PXLSPtgsArea(APtgs).Row1 and not ROW_ABSFLAG
          else
            R1 := FRow;
          if (PXLSPtgsArea(APtgs).Row2 and ROW_ABSFLAG) = ROW_ABSFLAG then
            R2 := PXLSPtgsArea(APtgs).Row2 and not ROW_ABSFLAG
          else
            R2 := {PXLSPtgsArea(APtgs).Row2 +} FRow - 1;
          if (C2 > XLS_MAXCOL) then
            C2 := XLS_MAXCOL;
          if (R2 > XLS_MAXROW) then
            R2 := XLS_MAXROW;
          FStack.Push(C1,R1,C2,R2);
        end
        else
          FStack.Push(PXLSPtgsArea1d(APtgs).Sheet,PXLSPtgsArea1d(APtgs).Col1 and not COL_ABSFLAG,PXLSPtgsArea1d(APtgs).Row1 and not ROW_ABSFLAG,PXLSPtgsArea1d(APtgs).Col2 and not COL_ABSFLAG,PXLSPtgsArea1d(APtgs).Row2 and not ROW_ABSFLAG);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea1d));
      end;
      xptgArea1dErr: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea1dError));
      end;
      xptgRef3d: begin
        FStack.Push(PXLSPtgsRef3d(APtgs).Sheet1,PXLSPtgsRef3d(APtgs).Sheet2,PXLSPtgsRef3d(APtgs).Col and not COL_ABSFLAG,PXLSPtgsRef3d(APtgs).Row and not ROW_ABSFLAG,PXLSPtgsRef3d(APtgs).Col and not COL_ABSFLAG,PXLSPtgsRef3d(APtgs).Row);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef3d));
      end;
      xptgRef3dErr: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef3d));
      end;
      xptgArea3d: begin
        FStack.Push(PXLSPtgsArea3d(APtgs).Sheet1,PXLSPtgsArea3d(APtgs).Sheet2,PXLSPtgsArea3d(APtgs).Col1 and not COL_ABSFLAG,PXLSPtgsArea3d(APtgs).Row1 and not ROW_ABSFLAG,PXLSPtgsArea3d(APtgs).Col2 and not COL_ABSFLAG,PXLSPtgsArea3d(APtgs).Row2 and not ROW_ABSFLAG);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea3d));
      end;
      xptgArea3dErr: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea3dError));
      end;
      xptgXRef1d: begin
        if FManager.XLinks.GetValue(PXLSPtgsXRef1d(APtgs).XBook,PXLSPtgsXRef1d(APtgs).Sheet,PXLSPtgsXRef1d(APtgs).Col and not COL_ABSFLAG,PXLSPtgsXRef1d(APtgs).Row and not ROW_ABSFLAG,XValue) then begin
          case XValue.CT of
            x12ctNil    : FStack.Push;
            x12ctBoolean: FStack.Push(XValue.valBoolean);
            x12ctError  : FStack.Push(XValue.valError);
            x12ctFloat  : FStack.Push(XValue.valFloat);
            x12ctString : FStack.Push(XValue.valString);
          end;
        end
        else
          FStack.Push;
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXRef1d));
      end;

      xptgXRef3d: begin
        FStack.Push('[' + FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.FileName + ']' +
                  FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef3d(APtgs).Sheet1] + ':' +
                  FManager.XLinks[PXLSPtgsXRef3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXRef3d(APtgs).Sheet2] + '!' +
                  ColRowToRefStrEnc(PXLSPtgsXRef3d(APtgs).Col,PXLSPtgsXRef3d(APtgs).Row));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXRef3d));
      end;
      xptgXArea1d: begin
        FStack.Push('[' + FManager.XLinks[PXLSPtgsXArea1d(APtgs).XBook].ExternalBook.Filename + ']' +
                  FManager.XLinks[PXLSPtgsXArea1d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea1d(APtgs).Sheet] + '!' +
                  AreaToRefStrEnc(PXLSPtgsXArea1d(APtgs).Col1,PXLSPtgsXArea1d(APtgs).Row1,PXLSPtgsXArea1d(APtgs).Col2,PXLSPtgsXArea1d(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXArea1d));
      end;
      xptgXArea3d: begin
        FStack.Push('[' + FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.FileName + ']' +
                  FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea3d(APtgs).Sheet1] + ':' +
                  FManager.XLinks[PXLSPtgsXArea3d(APtgs).XBook].ExternalBook.SheetNames[PXLSPtgsXArea3d(APtgs).Sheet2] + '!' +
                  AreaToRefStrEnc(PXLSPtgsXArea3d(APtgs).Col1,PXLSPtgsXArea3d(APtgs).Row1,PXLSPtgsXArea3d(APtgs).Col2,PXLSPtgsXArea3d(APtgs).Row2));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsXArea3d));
      end;

      xptgArray: begin
        DoArray(APtgs);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArray));
      end;
      xptgDataTableFmla : begin
        DoDataTable(PXLSPtgsDataTableFmla(APtgs));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsDataTableFmla));
      end;
      // This shall never be calculated. Result is set when parent is calculated.
      xptgArrayFmlaChild,
      xptgDataTableFmlaChild : raise XLSRWException.Create('This ptgs shall not be calculated');

      xptgTable              : begin
        if PXLSPtgsTable(APtgs).ArgCount = 0 then
          APtgs := @PXLSPtgsTable(APtgs).AreaPtg
        else begin
          DoTable(PXLSPtgsTable(APtgs));
          APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsTable));
        end;
      end;
      xptgTableCol           : APtgs := @PXLSPtgsTableCol(APtgs).AreaPtg;
      xptgTableSpecial       : begin
        if PXLSPtgsTableSpecial(APtgs).SpecialId = Byte(xtssThisRow) then begin
          PXLSPtgsTableSpecial(APtgs).Row1 := FRow;
          PXLSPtgsTableSpecial(APtgs).Row2 := FRow;
        end;
        if TXc12CellError(PXLSPtgsTableSpecial(APtgs).Error) <> errUnknown then
          FStack.Push(xtssNone,0,0,0,0)
//          FStack.Push(TXc12CellError(PXLSPtgsTableSpecial(APtgs).Error))
        else
          FStack.Push(TXLSTableSpecialSpecifier(PXLSPtgsTableSpecial(APtgs).SpecialId),PXLSPtgsTableSpecial(APtgs).Col1,PXLSPtgsTableSpecial(APtgs).Row1,PXLSPtgsTableSpecial(APtgs).Col2,PXLSPtgsTableSpecial(APtgs).Row2);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsTableSpecial));
      end;

      xptgMissingArg: begin
        FStack.Push;
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
{$ifdef XLS_BIFF}
      // **************** Excel 97 *****************

      xptg_EXCEL_97: begin
        FIsExcel97 := True;
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;
      xptg_ARRAYCONSTS_97: begin
        FIsExcel97 := True;
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
        Dec(LastPtgs,PWord(APtgs)^);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(word));

        ArrayData97 := Pointer(LastPtgs);
      end;

      xptgArray97: begin
        DoArray97(ArrayData97);
        APtgs := PXLSPtgs(NativeInt(APtgs) + 1 + 7);
      end;

      xptgStr97: begin
        P := Pointer(NativeInt(APtgs) + 1);
        if PByteArray(P)[1] = 0 then begin
          B := Byte(P^);
          P := Pointer(NativeInt(P) + 1);
          S := ByteStrToWideString(P,B);
        end
        else begin
          B := Byte(P^) * 2;
          P := Pointer(NativeInt(P) + 1);
          S := ByteStrToWideString(P,B div 2);
        end;
        FStack.Push(S);
        APtgs := PXLSPtgs(NativeInt(APtgs) + B + 3);
      end;

      xptgInt97: begin
        FStack.Push(PXLSPtgsInt97(APtgs).Value);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsInt97));
      end;

      xptgAttr97:
      begin
        P := Pointer(NativeInt(APtgs) + 1);
        if (Byte(P^) and $04) = $04 then begin
          P := Pointer(NativeInt(P) + 1);
          P := Pointer(NativeInt(P) + (Word(P^) + 2) * SizeOf(word) - 3);
        end
        else if (Byte(P^) and $10) = $10 then
          DoFunction(004,1);
        APtgs := Pointer(NativeInt(P) + 3);
      end;

      xptgRef97: begin
        if AIsName then begin
          if (PXLSPtgsRef97(APtgs).Col and $4000) = 0 then
            C1 := PXLSPtgsRef(APtgs).Col and $3FFF
          else
            C1 := FCol;
          if (PXLSPtgsRef97(APtgs).Row and $8000) = 0 then
            R1 := PXLSPtgsRef(APtgs).Row
          else
            R1 := FRow;
          FStack.Push(C1,R1);
        end
        else
          FStack.Push(PXLSPtgsRef97(APtgs).Col and $3FFF,PXLSPtgsRef97(APtgs).Row);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef97));
      end;

      xptgArea97: begin
        if AIsName then begin
          if (PXLSPtgsArea97(APtgs).Col1 and $4000) = 0 then
            C1 := PXLSPtgsArea97(APtgs).Col1 and $3FFF
          else
            C1 := FCol;
          if (PXLSPtgsArea97(APtgs).Col2 and $4000) = 0 then
            C2 := PXLSPtgsArea97(APtgs).Col2 and $3FFF
          else
            C2 := PXLSPtgsArea97(APtgs).Col2 + FCol;
          if (PXLSPtgsArea97(APtgs).Row1 and $8000) = 0 then
            R1 := PXLSPtgsArea97(APtgs).Row1
          else
            R1 := FRow;
          if (PXLSPtgsArea97(APtgs).Row2 and $8000) = 0 then
            R2 := PXLSPtgsArea97(APtgs).Row2 and not ROW_ABSFLAG
          else
            R2 := NativeInt(PXLSPtgsArea97(APtgs).Row2) + FRow;
          if (C2 > XLS_MAXCOL_97) then
            C2 := XLS_MAXCOL_97;
          if (R2 > XLS_MAXROW_97) then
            R2 := XLS_MAXROW_97;
          FStack.Push(C1,R1,C2,R2);
        end
        else
          FStack.Push(PXLSPtgsArea97(APtgs).Col1 and $3FFF,PXLSPtgsArea97(APtgs).Row1,PXLSPtgsArea97(APtgs).Col2 and $3FFF,PXLSPtgsArea97(APtgs).Row2);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea97));
      end;

      xptgRef3d97: begin
        FManager._ExtNames97.Get3dSheets(PXLSPtgsRef3d97(APtgs).ExtSheetIndex,i1,i2);

        if AIsName and (i1 = i2) and (FSheetIndex = i1) then begin
          if (PXLSPtgsRef97(APtgs).Col and $4000) = 0 then
            C1 := PXLSPtgsRef97(APtgs).Col and $3FFF
          else
            C1 := FCol;
          if (PXLSPtgsRef97(APtgs).Row and $8000) = 0 then
            R1 := PXLSPtgsRef97(APtgs).Row
          else
            R1 := FRow - 1;
          FStack.Push(C1,R1);
        end
        else if i1 = i2 then
          FStack.Push(i1,PXLSPtgsRef3d97(APtgs).Col and $3FFF,PXLSPtgsRef3d97(APtgs).Row)
        else
          FStack.Push(i1,i2,PXLSPtgsRef3d97(APtgs).Col and $3FFF,PXLSPtgsRef3d97(APtgs).Row);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef3d97));
      end;

      xptgArea3d97: begin
        FManager._ExtNames97.Get3dSheets(PXLSPtgsArea3d97(APtgs).ExtSheetIndex,i1,i2);

        if AIsName and (i1 = i2) and (FSheetIndex = i1) then begin
          if (PXLSPtgsArea3d97(APtgs).Col1 and $4000) = 0 then
            C1 := PXLSPtgsArea3d97(APtgs).Col1 and $3FFF
          else
            C1 := FCol;
          if (PXLSPtgsArea3d97(APtgs).Col2 and $4000) = 0 then
            C2 := PXLSPtgsArea3d97(APtgs).Col2 and $3FFF
          else
            C2 := PXLSPtgsArea3d97(APtgs).Col2 + FCol;
          if (PXLSPtgsArea3d97(APtgs).Row1 and $8000) = 0 then
            R1 := PXLSPtgsArea3d97(APtgs).Row1
          else
            R1 := FRow;
          if (PXLSPtgsArea3d97(APtgs).Row2 and $8000) = 0 then
            R2 := PXLSPtgsArea3d97(APtgs).Row2
          else
            R2 := {PXLSPtgsArea3d97(APtgs).Row2 +} FRow - 1;
          if (C2 > XLS_MAXCOL_97) then
            C2 := XLS_MAXCOL_97;
          if (R2 > XLS_MAXROW_97) then
            R2 := XLS_MAXROW_97;
          FStack.Push(C1,R1,C2,R2);
        end
        else if i1 = i2 then
          FStack.Push(i1,PXLSPtgsArea3d97(APtgs).Col1 and $3FFF,PXLSPtgsArea3d97(APtgs).Row1,PXLSPtgsArea3d97(APtgs).Col2 and $3FFF,PXLSPtgsArea3d97(APtgs).Row2)
        else
          FStack.Push(i1,i2,PXLSPtgsArea3d97(APtgs).Col1 and $3FFF,PXLSPtgsArea3d97(APtgs).Row1,PXLSPtgsArea3d97(APtgs).Col2 and $3FFF,PXLSPtgsArea3d97(APtgs).Row2);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea3d97));
      end;

      xptgRefErr97: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsRef97));
      end;
      xptgAreaErr97: begin
        FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsArea97));
      end;

      xptgFunc97: begin
        if (ATargetArea <> Nil) and (G_XLSExcelFuncNames[PXLSPtgsFunc(APtgs).FuncId].Type_ <> xeftArray) then
          DoArrayFunction(APtgs,ATargetArea)
        else
          DoFunction(PXLSPtgsFunc(APtgs).FuncId);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsFunc));
      end;

      xptgFuncVar97: begin
        if (ATargetArea <> Nil) and (G_XLSExcelFuncNames[PXLSPtgsFuncVar(APtgs).FuncId].Type_ <> xeftArray) then
          DoArrayFunction(APtgs,ATargetArea)
        // Skip invalid functions
        else if PXLSPtgsFuncVar(APtgs).FuncId <> 255 then
          DoFunction(PXLSPtgsFuncVar(APtgs).FuncId,PXLSPtgsFuncVar(APtgs).ArgCount);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsFuncVar));
      end;

      xptgName97: begin
        Name := FManager.Workbook.DefinedNames[PXLSPtgsName97(APtgs).NameId - 1];

        if Name.VbProcedure then
          FStack.Push(errNA)
        else begin
          case Name.SimpleName of
            xsntNone: begin
              Ptgs := Name.Ptgs;
              Sz := Name.PtgsSz;
              if Sz <= 1 then
                FStack.Push(errName)
              else
                FStack.Push(DoEvaluate(Ptgs,Sz,ATargetArea,True));
            end;
            xsntRef  : FStack.Push(Name.SimpleArea.SheetIndex,Name.SimpleArea.Col1,Name.SimpleArea.Row1);
            xsntArea : FStack.Push(Name.SimpleArea.SheetIndex,Name.SimpleArea.Col1,Name.SimpleArea.Row1,Name.SimpleArea.Col2,Name.SimpleArea.Row2);
            xsntError: FStack.Push(errRef);
          end;
        end;

        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsName97));
      end;

      xptgNameX97: begin
        FManager._ExtNames97.Get3dSheets(PXLSPtgsNameX97(APtgs).ExtSheetIndex,i1,i2);
        if i1 = $FFFE then begin
          Name := FManager.Workbook.DefinedNames[PXLSPtgsNameX97(APtgs).NameIndex - 1];
          case Name.SimpleName of
            xsntNone: begin
              Ptgs := Name.Ptgs;
              Sz := Name.PtgsSz;
              if Sz <= 1 then
                FStack.Push(errName)
              else
                FStack.Push(DoEvaluate(Ptgs,Sz,ATargetArea,True));
            end;
            xsntRef  : FStack.Push(Name.SimpleArea.SheetIndex,Name.SimpleArea.Col1,Name.SimpleArea.Row1);
            xsntArea : FStack.Push(Name.SimpleArea.SheetIndex,Name.SimpleArea.Col1,Name.SimpleArea.Row1,Name.SimpleArea.Col2,Name.SimpleArea.Row2);
            xsntError: FStack.Push(errRef);
          end;
        end
        else
          FStack.Push(errRef);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgsNameX97));
      end;

      xptgMemFunc97: begin
        FStack.Push(errNA);
        P := Pointer(NativeInt(APtgs) + SizeOf(TXLSPtgs));
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs) + SizeOf(Word) + PWord(P)^);
      end;

      xptgMissArg97: begin
        FStack.Push(errNA);
        APtgs := PXLSPtgs(NativeInt(APtgs) + SizeOf(TXLSPtgs));
      end;

      // ************** End Excel 97 *****************
{$endif}
      else raise XLSRWException.CreateFmt('Invalid ptgs: %.2X',[APtgs.Id]);
    end;
    if FDebugData <> Nil then begin
      if (FDebugData.CollectFormulas) and (PrevPtgs.Id in [xptgFunc,xptgFuncVar,xptgUserFunc]) then
        FDebugData.FmlaItems.Add(FDebugData.Steps,FDebugData.CurrPtgsOffs,APtgs,FStack.Peek(0));
      Dec(FDebugSteps);
      if FDebugSteps <= 0 then begin
        FDebugData.CurrPtgs := APtgs;
        Exit;
      end;
      PrevPtgs := APtgs;
    end;
  end;
  if AIsName then
    Result := FStack.Pop
  else
    FStack.PopResult(Result);
end;

function TXLSFormulaEvaluator.Evaluate(ASheetIndex, ACol, ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea): TXLSFormulaValue;
begin
  FDebugData := Nil;
  FVolatile := False;
  FSheetIndex := ASheetIndex;
  FCol := ACol;
  FRow := ARow;
  FTargetArea := ATargetArea;
  FStack.Clear(ATargetArea <> Nil,ASheetIndex,ACol,ARow);
  Result := DoEvaluate(APtgs,APtgsSize,FTargetArea);
  if Result.RefType = xfrtArray then begin
    FillTargetArea(Result,Aptgs.Id <> xptgDataTableFmla);
    Result := FStack.Pop;
  end;
end;

function TXLSFormulaEvaluator.EvaluateStr(ASheetIndex, ACol, ARow: integer; APtgs: PXLSPtgs; APtgsSize: integer; ATargetArea: PXLSFormulaArea): AxUCString;
var
  Res: TXLSFormulaValue;
  Err: TXc12CellError;
begin
  FVolatile := False;
  Res := Evaluate(ASheetIndex,ACol,ARow,APtgs,APtgsSize,ATargetArea);
  Err := FStack.GetAsString(Res,Result);
  if Err <> errUnknown then
    Result := Xc12CellErrorNames[Err];
end;

procedure TXLSFormulaEvaluator.FillTargetArea(FV: TXLSFormulaValue; AHasParentFormula: boolean);
var
  Arr: TXLSArrayItem;
  C,R: integer;
  CT,RT: integer;
begin
  Arr := TXLSArrayItem(FV.vSource);
  FStack.Push(Arr.GetAsFormulaValue(0,0));
  if FTargetArea <> Nil then begin
    RT := FTargetArea.Row1;
    for R := 0 to Arr.Height - 1 do begin
      CT := FTargetArea.Col1;
      for C := 0 to Arr.Width - 1 do begin
        if not (AHasParentFormula and (R = 0) and (C = 0)) then begin
          FV := Arr.GetAsFormulaValue(C,R);
          case FV.ValType of
            xfvtFloat   : FStack.FCells.UpdateFloat(CT,RT,FV.vFloat);
            xfvtBoolean : FStack.FCells.UpdateBoolean(CT,RT,FV.vBool);
            xfvtError   : FStack.FCells.UpdateError(CT,RT,FV.vErr);
            xfvtString  : FStack.FCells.UpdateString(CT,RT,FV.vStr);
            else          raise XLSRWException.Create('Invalid array value');
          end;
        end;
        Inc(CT);
        if CT > FTargetArea.Col2 then
          CT := FTargetArea.Col1;
      end;
      Inc(RT);
      if RT > Integer(FTargetArea.Row2) then
        RT := FTargetArea.Row1;
    end;
  end;
end;

procedure TValueStack.Push(const ACol1, ARow1, ACol2, ARow2: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.Empty := False;
  P.RefType := xfrtArea;
  P.ValType := xfvtUnknown;
//  P.Cells := FManager.Worksheets.IdOrder[FSheetIndex].Cells;
  P.Cells := FManager.Worksheets[FSheetIndex].Cells;
  P.Col1 := ACol1;
  P.Row1 := ARow1;
  P.Col2 := ACol2;
  P.Row2 := ARow2;
end;

procedure TValueStack.Push(const AValue: TXLSFormulaValue);
begin
  IncStackPtr;
  FStack[FStackPtr] := AValue;
end;

procedure TValueStack.Push(const ASheetIndex, ACol, ARow: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.Empty := False;
  P.RefType := xfrtRef;
  P.Cells := FManager.Worksheets[ASheetIndex].Cells;
  P.Col1 := ACol;
  P.Row1 := ARow;
  P.Col2 := ACol;
  P.Row2 := ARow;

  GetCellValue(P.Cells,ACol,ARow,P);
  P.RefType := xfrtRef;
end;

procedure TValueStack.Push(const ASheetIndex, ACol1, ARow1, ACol2, ARow2: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.RefType := xfrtArea;
  P.ValType := xfvtUnknown;
  P.Empty := False;
  P.Cells := FManager.Worksheets[ASheetIndex].Cells;
  P.Col1 := ACol1;
  P.Row1 := ARow1;
  P.Col2 := ACol2;
  P.Row2 := ARow2;
end;

procedure TValueStack.Push(const ASheetIndex1, ASheetIndex2, ACol1, ARow1, ACol2, ARow2: integer);
var
  i: integer;
  P: PXLSFormulaValue;
  CA: TCellAreas;
begin
  IncStackPtr;
  CA := TCellAreas.Create;
  P := @FStack[FStackPtr];
  P.Empty := False;
  P.RefType := xfrtAreaList;
  P.ValType := xfvtUnknown;
  P.vSource := CA;
  FGarbage.Add(CA);
  for i := Min(ASheetIndex1,ASheetIndex2) to Max(ASheetIndex1,ASheetIndex2) do
    CA.Add(ACol1, ARow1, ACol2, ARow2).Obj := FManager.Worksheets[i].Cells;
end;

procedure TValueStack.Push(const AArray: TXLSArrayItem);
begin
  IncStackPtr;
  FStack[FStackPtr].RefType := xfrtArray;
  FStack[FStackPtr].ValType := xfvtUnknown;
  FStack[FStackPtr].vSource := AArray;
  FGarbage.Add(AArray);
end;

procedure TValueStack.Push(const AMatrix: TXLSMatrix);
var
  C,R: integer;
  Arr: TXLSArrayItem;
begin
  Arr := TXLSArrayItem.Create(AMatrix.Cols,AMatrix.Rows);
  for R := 0 to AMatrix.Rows - 1 do begin
    for C := 0 to AMatrix.Cols - 1 do
      Arr.AsFloat[C,R] := AMatrix[C,R];
  end;
  Push(Arr);
end;

procedure TValueStack.Push(const ACells: TXLSCellMMU; const ACol, ARow: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.RefType := xfrtRef;
//  P.Cells := FManager.Worksheets.IdOrder[FSheetIndex].Cells;
  P.Cells := ACells;
  P.Col1 := ACol;
  P.Row1 := ARow;
  P.Col2 := ACol;
  P.Row2 := ARow;

  GetCellValue(P.Cells,ACol,ARow,P);
  P.RefType := xfrtRef;
end;

procedure TValueStack.Push(const ATableSpecial: TXLSTableSpecialSpecifier; const ACol1,ARow1,ACol2,ARow2: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.Empty := False;
  P.RefType := xfrtNone;
  P.ValType := xfvtTableSpecial;
  P.vSpec := Integer(ATableSpecial);
  P.Col1 := ACol1;
  P.Row1 := ARow1;
  P.Col2 := ACol2;
  P.Row2 := ARow2;
end;

procedure TValueStack.Push(const ACells: TXLSCellMMU; const ACol1, ARow1, ACol2, ARow2: integer);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.RefType := xfrtRef;
//  P.Cells := FManager.Worksheets.IdOrder[FSheetIndex].Cells;
  P.Cells := ACells;
  P.Col1 := ACol1;
  P.Row1 := ARow1;
  P.Col2 := ACol2;
  P.Row2 := ARow2;

  P.RefType := xfrtArea;
end;

procedure TValueStack.PopResArray(out AArr: TXLSArrayItem; AFloat: boolean);
var
  FV: TXLSFormulaValue;
begin
  FV := PopRes;

  AArr := TXLSArrayItem.CreateFV(FV);

  AArr.VectorMode := True;

  if AFloat then
    FillArrayItemFloat(AArr,FV)
  else
    FillArrayItem(AArr,FV);
end;

procedure TValueStack.PopResArrays(out AArr1, AArr2: TXLSArrayItem; AFloat: boolean);
begin
  PopResArray(AArr2,AFloat);
  PopResArray(AArr1,AFloat);
end;

function TValueStack.PopResBoolean(out AValue: boolean): boolean;
var
  Err: TXc12CellError;
begin
  Err := GetAsBoolean(PopRes,AValue);
  Result := Err = errUnknown;
  if not Result then
    Push(Err);
end;

function TValueStack.PopResCellsSource(out AValue: TXLSCellsSource): boolean;
var
  FV: TXLSFormulaValue;
begin
  FV := PopRes;

  Result := FV.RefType in [xfrtRef,xfrtArea,xfrtXArea,xfrtArray];
  if Result then begin
    case FV.RefType of
      xfrtRef,
      xfrtArea : begin
        AValue := TXLSCellsSource.Create(FV.Cells,FV.Col1,FV.Row1,FV.Col2,FV.Row2);
      end;
      xfrtXArea: raise XLSRWException.Create('TODO XArea');
      xfrtArray: begin
        AValue := TXLSCellsSource.Create(TXLSArrayItem(FV.vSource));
      end;
    end;
    FGarbage.Add(AValue);
  end
  else
    Push(errValue);
end;

function TValueStack.PopResDateTime(out AValue: double): boolean;
begin
  Result := PopResFloat(AValue);
  if not Result then
    Exit;
  Result := (AValue >= 1) and (AValue <= XLSFMLA_MAXDATE);
  if not Result then
    Push(errValue);
end;

function TValueStack.PopResFloat(out AValue: double): boolean;
var
  Err: TXc12CellError;
begin
  if not GetAsFloat(PopRes,AValue,Err) then
    AValue := 0;
  Result := Err = errUnknown;
  if not Result then
    Push(Err);
end;

function TValueStack.PopResInt(out AValue: integer): boolean;
var
  V: double;
  Err: TXc12CellError;
begin
  if not GetAsFloatNum(PopRes,V,Err) then
    AValue := 0;

  Result := Err = errUnknown;
  if Result then
    AValue := Trunc(V)
  else
    Push(Err);
end;

function TValueStack.PopRes: TXLSFormulaValue;
begin
  if FResultPtr < 0 then
    raise XLSRWException.Create('Empty stack');
  Result := FStack[FResultPtr];
  Dec(FResultPtr);
end;

function TValueStack.PopResPosInt(out AValue: integer; AOkNum: integer): boolean;
var
  V: double;
  Err: TXc12CellError;
begin
  if not GetAsFloatNum(PopRes,V,Err) then
    AValue := 0;

  Result := Err = errUnknown;
  if Result then begin
    if V < AOkNum then
      Push(errValue)
    else
      AValue := Trunc(V);
  end
  else
    Push(Err);
end;

function TValueStack.PopResStr(out AValue: AxUCString; AEmptyOk: boolean): boolean;
var
  Err: TXc12CellError;
begin
  Err := GetAsString(PopRes,AValue);
  Result := Err = errUnknown;
  if not Result then begin
    Push(Err);
  end
  else if (AValue = '') and not AEmptyOk then begin
    Push(errValue);
    Result := False;
  end;
end;

procedure TValueStack.Push(const AValue: TXc12CellError; const AFVValue: TXLSFormulaValue);
begin
  if (AFVValue.RefType = xfrtNone) and (AFVValue.ValType = xfvtError) then
    Push(AFVValue.vErr)
  else
    Push(AValue);
end;

procedure TValueStack.Push(const ASheetIndex: integer; const AError: TXc12CellError);
var
  P: PXLSFormulaValue;
begin
  IncStackPtr;
  P := @FStack[FStackPtr];
  P.Empty := False;
  P.RefType := xfrtRef;
  P.Cells := Nil;
  P.ValType := xfvtError;
  P.vErr := AError;
end;

{ TValStackIterator }

procedure TValStackIterator.AddLinked;
begin
  if FLinked = Nil then
    FLinked := TValStackIterator.Create(FOwner);
end;

procedure TValStackIterator.BeginIterate(AValue: TXLSFormulaValue; AHaltOnError: boolean);
begin
  FDimension := Nil;
  FDone := False;
  FFV := AValue;
  FHaltOnError := AHaltOnError;
  if FLinked <> Nil then
    FLinked.Free;
  FLinked := Nil;
  case FFV.RefType of
    xfrtNone     : ;
    xfrtRef      : begin
      FCol := FFV.Col1;
      FRow := FFV.Row1;
    end;
    xfrtArea     : begin
      FSource := FFV.Cells;
      FDimension := @FFV.Cells.Dimension;
      FCol := Max(FFV.Col1 - 1,FDimension.Col1 - 1);
      FRow := Max(FFV.Row1,FDimension.Row1);
    end;
    xfrtXArea    : ;
    xfrtAreaList : begin
      FSource := FFV.vSource;
      FIndex := 0;
      FCol := TCellAreas(FSource)[0].Col1 - 1;
      FRow := TCellAreas(FSource)[0].Row1;
    end;
    xfrtArray    : begin
      FSource := FFV.vSource;
      FCol := -1;
      FRow := 0;
    end;
  end;
end;

function TValStackIterator.AsFloatA: double;
begin
  Result := 0;
  case FResult.ValType of
    xfvtFloat   : Result := FResult.vFloat;
    xfvtBoolean : if FResult.vBool then Result := 1 else Result := 0;
  end;
end;

procedure TValStackIterator.BeginIterate(AValue1, AValue2: TXLSFormulaValue; AHaltOnError: boolean);
begin
  BeginIterate(AValue1,AHaltOnError);
  AddLinked;
  FLinked.BeginIterate(AValue2,AHaltOnError);
end;

procedure TValStackIterator.Clear;
begin
  if FLinked <> Nil then begin
    FLinked.Free;
    FLinked := Nil;
  end;
end;

constructor TValStackIterator.Create(AOwner: TValueStack);
begin
  FOwner := AOwner;
end;

destructor TValStackIterator.Destroy;
begin
  Clear;
  inherited;
end;

function TValStackIterator.DoNext: boolean;
var
  CA: TCellArea;

function DoOptions: boolean;
var
  R: PXLSMMURowHeader;
begin
  R :=FOwner.FCells.FindRow(FRow);
  Result := (R <> Nil) and not (xroHidden in R.Options);
  if not Result then
    Result := DoNext;
end;

begin
  if not FIgnoreError and not FHaltOnError then
    FResult.ValType := xfvtUnknown;
  FResult.RefType := xfrtNone;
  Result := False;
  if FDone then
    Exit;

  case FFV.RefType of
    xfrtNone,
    xfrtRef      : begin
      FResult.Empty := FFV.Empty;
      if not FResult.Empty then begin
        FResult.ValType := FFV.ValType;
        case FFV.ValType of
          xfvtFloat   : FResult.vFloat := FFV.vFloat;
          xfvtBoolean : FResult.vBool := FFV.vBool;
          xfvtError   : FResult.vErr := FFV.vErr;
          xfvtString  : FResult.vStr := FFV.vStr;
        end;
      end;
      FDone := True;
      if FIgnoreHidden then
        Result := DoOptions
      else
        Result := True;
      if Result then
        FResult.RefType := FFV.RefType;
    end;
    xfrtArea     : begin
      Inc(FCol);
      if FCol > Min(FFV.Col2,FDimension.Col2) then begin
        FCol := FFV.Col1;
        Inc(FRow);
      end;
      FDone := (FRow > Min(FFV.Row2,FDimension.Row2));
      Result := not FDone;
      if Result and FIgnoreHidden then
        Result := DoOptions;
      if Result then
        FOwner.GetCellValue(TXLSCellMMU(FSource),FCol,FRow,@FResult);
    end;
    xfrtXArea    : ;
    xfrtAreaList : begin
      CA := TCellAreas(FSource)[FIndex];
      Inc(FCol);
      if FCol > CA.Col2 then begin
        FCol := CA.Col1;
        Inc(FRow);
        if FRow > CA.Row2 then begin
          Inc(FIndex);
          if FIndex < TCellAreas(FSource).Count then begin
            CA := TCellAreas(FSource)[FIndex];
            FCol := CA.Col1;
            FRow := CA.Row1;
          end;
        end;
      end;
      FDone := FIndex >= TCellAreas(FSource).Count;
      Result := not FDone;
      if Result and FIgnoreHidden then
        Result := DoOptions;
      if Result then
        FOwner.GetCellValue(TXLSCellMMU(CA.Obj),FCol,FRow,@FResult);
    end;
    xfrtArray    : begin
      Inc(FCol);
      if FCol >= TXLSArrayItem(FSource).Width then begin
        FCol := 0;
        Inc(FRow);
      end;
      FDone := FRow >= TXLSArrayItem(FSource).Height;
      Result := not FDone;
      if Result then
        FResult := TXLSArrayItem(FSource).GetAsFormulaValue(FCol,FRow);
    end;
  end;
  if not FIgnoreError and FHaltOnError and (FResult.ValType = xfvtError) then
    FDone := True;
  if FIgnoreError and not FDone and (FResult.ValType = xfvtError) then
    Result := DoNext;
end;

function TValStackIterator.Next: boolean;
begin
  if FLinked <> Nil then
    Result := DoNext and FLinked.Next
  else
    Result := DoNext;
end;

function TValStackIterator.NextFloat: boolean;
begin
  while Next do begin
    Result := ResIsFloat;
    if Result then
      Exit;
  end;
  Result := False;
end;

function TValStackIterator.NextFloatA: boolean;
begin
  while Next do begin
    Result := ResIsFloatA;
    if Result then
      Exit;
  end;
  Result := False;
end;

function TValStackIterator.NextNotEmpty: boolean;
begin
  while Next do begin
    Result := ResIsNotEmpty;
    if Result then
      Exit;
  end;
  Result := False;
end;

function TValStackIterator.ResIsError(out AError: TXc12CellError): boolean;
begin
  Result := FResult.ValType = xfvtError;
  if Result then
    AError := FResult.vErr
  else if FLinked <> Nil then
    Result := FLinked.ResIsError(AError);
end;

function TValStackIterator.ResIsFloat: boolean;
begin
  Result := FResult.ValType = xfvtFloat;
  if Result and (FLinked <> Nil) then
    Result := FLinked.ResIsFloat;
end;

function TValStackIterator.ResIsFloatA: boolean;
begin
  Result := FResult.ValType in [xfvtFloat,xfvtBoolean,xfvtString];
  if Result and (FLinked <> Nil) then
    Result := FLinked.ResIsFloatA;
end;

function TValStackIterator.ResIsNotEmpty: boolean;
begin
  Result := FResult.ValType <> xfvtUnknown;
  if Result and (FLinked <> Nil) then
    Result := FLinked.ResIsFloat;
end;

{ TXLSUserFuncData }

function TXLSUserFuncData.ArgCount: integer;
begin
  Result := Length(FArgs);
end;

procedure TXLSUserFuncData.Dimensions(const AIndex: integer; out ACol1, ARow1, ACol2, ARow2: integer);
begin
  if (AIndex < 0) or (AIndex > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[AIndex].RefType of
    xfrtNone,
    xfrtRef     : begin
      ACol1 := 0;
      ARow1 := 0;
      ACol2 := 0;
      ARow2 := 0;
    end;
    xfrtArea    : begin
      ACol1 := FArgs[AIndex].Col1;
      ARow1 := FArgs[AIndex].Row1;
      ACol2 := FArgs[AIndex].Col2;
      ARow2 := FArgs[AIndex].Row2;
    end;
    xfrtXArea   : begin
      ACol1 := FArgs[AIndex].Col1;
      ARow1 := FArgs[AIndex].Row1;
      ACol2 := FArgs[AIndex].Col2;
      ARow2 := FArgs[AIndex].Row2;
    end;
    xfrtAreaList: raise XLSRWException.Create('Argument is N/A');
    xfrtArray   : begin
      ACol1 := 0;
      ARow1 := 0;
      ACol2 := TXLSArrayItem(FArgs[AIndex].vSource).Width;
      ARow2 := TXLSArrayItem(FArgs[AIndex].vSource).Height;
    end;
    xfrtArrayArg: raise XLSRWException.Create('Argument is N/A');
  end;
end;

function TXLSUserFuncData.GetAreaAsBoolean(Index, ACol, ARow: integer): boolean;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtBoolean then
        Result := FArgs[Index].vBool
      else
        Result := False;
    end;
    xfrtRef     : FArgs[Index].Cells.GetBoolean(FArgs[Index].Col1,FArgs[Index].Row2,Result);
    xfrtArea    : FArgs[Index].Cells.GetBoolean(ACol,ARow,Result);
    xfrtXArea   : raise XLSRWException.Create('Argument is N/A');
    xfrtAreaList: raise XLSRWException.Create('Argument is N/A');
    xfrtArray   : Result := TXLSArrayItem(FArgs[Index].vSource).AsBoolean[ACol,ARow];
    xfrtArrayArg: raise XLSRWException.Create('Argument is N/A');
  end;
end;

function TXLSUserFuncData.GetAreaAsError(Index, ACol, ARow: integer): TXc12CellError;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtError then
        Result := FArgs[Index].vErr
      else
        Result := errNA;
    end;
    xfrtRef     : FArgs[Index].Cells.GetError(FArgs[Index].Col1,FArgs[Index].Row2,Result);
    xfrtArea    : FArgs[Index].Cells.GetError(ACol,ARow,Result);
    xfrtXArea   : raise XLSRWException.Create('Argument is N/A');
    xfrtAreaList: raise XLSRWException.Create('Argument is N/A');
    xfrtArray   : Result := TXLSArrayItem(FArgs[Index].vSource).AsError[ACol,ARow];
    xfrtArrayArg: raise XLSRWException.Create('Argument is N/A');
  end;
end;

function TXLSUserFuncData.GetAreaAsFloat(Index, ACol, ARow: integer): double;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtFloat then
        Result := FArgs[Index].vFloat
      else
        Result := 0;
    end;
    xfrtRef     : Result := FArgs[Index].Cells.GetFloat(FArgs[Index].Col1,FArgs[Index].Row2);
    xfrtArea    : Result := FArgs[Index].Cells.GetFloat(ACol,ARow);
    xfrtXArea   : raise XLSRWException.Create('Argument is N/A');
    xfrtAreaList: raise XLSRWException.Create('Argument is N/A');
    xfrtArray   : Result := TXLSArrayItem(FArgs[Index].vSource).AsFloat[ACol,ARow];
    xfrtArrayArg: raise XLSRWException.Create('Argument is N/A');
    else          Result := 0;
  end;
end;

function TXLSUserFuncData.GetAreaAsString(Index, ACol, ARow: integer): AxUCString;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtString then
        Result := FArgs[Index].vStr
      else
        Result := '';
    end;
    xfrtRef     : FArgs[Index].Cells.GetString(FArgs[Index].Col1,FArgs[Index].Row2,Result);
    xfrtArea    : FArgs[Index].Cells.GetString(ACol,ARow,Result);
    xfrtXArea   : raise XLSRWException.Create('Argument is N/A');
    xfrtAreaList: raise XLSRWException.Create('Argument is N/A');
    xfrtArray   : Result := TXLSArrayItem(FArgs[Index].vSource).AsString[ACol,ARow];
    xfrtArrayArg: raise XLSRWException.Create('Argument is N/A');
  end;
end;

function TXLSUserFuncData.GetArgType(Index: integer): TXLSFormulaValueType;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');
  Result := FArgs[Index].ValType;
end;

function TXLSUserFuncData.GetAsBoolean(Index: integer): boolean;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtBoolean then
        Result := FArgs[Index].vBool
      else
        Result := False;
    end;
    xfrtRef     : FArgs[Index].Cells.GetBoolean(FArgs[Index].Col1,FArgs[Index].Row2,Result);
    xfrtArea,
    xfrtXArea,
    xfrtAreaList,
    xfrtArray,
    xfrtArrayArg: raise XLSRWException.Create('Argument is not a single cell');
  end;
end;

function TXLSUserFuncData.GetAsError(Index: integer): TXc12CellError;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtError then
        Result := FArgs[Index].vErr
      else
        Result := errNA;
    end;
    xfrtRef     : FArgs[Index].Cells.GetError(FArgs[Index].Col1,FArgs[Index].Row2,Result);
    xfrtArea,
    xfrtXArea,
    xfrtAreaList,
    xfrtArray,
    xfrtArrayArg: raise XLSRWException.Create('Argument is not a single cell');
  end;
end;

function TXLSUserFuncData.GetAsFloat(Index: integer): double;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtFloat then
        Result := FArgs[Index].vFloat
      else
        Result := 0;
    end;
    xfrtRef     : Result := FArgs[Index].Cells.GetFloat(FArgs[Index].Col1,FArgs[Index].Row2);
    xfrtArea,
    xfrtXArea,
    xfrtAreaList,
    xfrtArray,
    xfrtArrayArg: raise XLSRWException.Create('Argument is not a single cell');
    else          Result := 0;
  end;
end;

function TXLSUserFuncData.GetAsString(Index: integer): AxUCString;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');

  case FArgs[Index].RefType of
    xfrtNone    : begin
      if FArgs[Index].ValType = xfvtString then
        Result := FArgs[Index].vStr
      else
        Result := '';
    end;
    xfrtRef     : FArgs[Index].Cells.GetString(FArgs[Index].Col1,FArgs[Index].Row2,Result);
    xfrtArea,
    xfrtXArea,
    xfrtAreaList,
    xfrtArray,
    xfrtArrayArg: raise XLSRWException.Create('Argument is not a single cell');
  end;
end;

function TXLSUserFuncData.GetCols(Index: integer): integer;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');
  case FArgs[Index].RefType of
    xfrtNone    : Result := 1;
    xfrtRef     : Result := 1;
    xfrtArea    : Result := FArgs[Index].Col2 - FArgs[Index].Col1 + 1;
    xfrtXArea   : Result := FArgs[Index].Col2 - FArgs[Index].Col1 + 1;
    xfrtAreaList: Result := 0;
    xfrtArray   : Result := TXLSArrayItem(FArgs[Index].vSource).Width;
    xfrtArrayArg: Result := 0;
    else          Result := 0;
  end;
end;

function TXLSUserFuncData.GetRefType(Index: integer): TXLSFormulaRefType;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');
  Result := FArgs[Index].RefType;
end;

function TXLSUserFuncData.GetRows(Index: integer): integer;
begin
  if (Index < 0) or (Index > High(FArgs)) then
    raise XLSRWException.Create('Index out of range');
  case FArgs[Index].RefType of
    xfrtNone    : Result := 1;
    xfrtRef     : Result := 1;
    xfrtArea    : Result := FArgs[Index].Row2 - FArgs[Index].Row1 + 1;
    xfrtXArea   : Result := FArgs[Index].Row2 - FArgs[Index].Row1 + 1;
    xfrtAreaList: Result := 0;
    xfrtArray   : Result := TXLSArrayItem(FArgs[Index].vSource).Height;
    xfrtArrayArg: Result := 0;
    else          Result := 0;
  end;
end;

procedure TXLSUserFuncData.SetResultAsBoolean(const Value: boolean);
begin
  FResult.ValType := xfvtBoolean;
  FResult.vBool := Value;
end;

procedure TXLSUserFuncData.SetResultAsError(const Value: TXc12CellError);
begin
  FResult.ValType := xfvtError;
  FResult.vErr := Value;
end;

procedure TXLSUserFuncData.SetResultAsFloat(const Value: double);
begin
  FResult.ValType := xfvtFloat;
  FResult.vFloat := Value;
end;

procedure TXLSUserFuncData.SetResultAsString(const Value: AxUCString);
begin
  FResult.ValType := xfvtString;
  FResult.vStr := Value;
end;

end.
