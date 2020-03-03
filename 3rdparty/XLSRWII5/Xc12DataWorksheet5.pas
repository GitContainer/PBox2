unit Xc12DataWorksheet5;

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

uses Classes, SysUtils, Contnrs,
     xpgParseDrawing, xpgParserPivot, xpgPUtils, xpgPSimpleDOM,
     Xc12Utils5, Xc12Common5, Xc12DataSST5, Xc12DataStyleSheet5, Xc12DataComments5,
     Xc12DataAutofilter5, Xc12DataTable5, Xc12FileData5,
     XLSUtils5, XLSCellMMU5, XLSCellAreas5, XLSMoveCopy5, XLSClassFactory5,
     XLSFormulaTypes5;


type TXc12CfType = (x12ctExpression,x12ctCellIs,x12ctColorScale,x12ctDataBar,x12ctIconSet,x12ctTop10,x12ctUniqueValues,x12ctDuplicateValues,x12ctContainsText,x12ctNotContainsText,x12ctBeginsWith,x12ctEndsWith,x12ctContainsBlanks,x12ctNotContainsBlanks,x12ctContainsErrors,x12ctNotContainsErrors,x12ctTimePeriod,x12ctAboveAverage);
type TXc12ConditionalFormattingOperator = (x12cfoLessThan,x12cfoLessThanOrEqual,x12cfoEqual,x12cfoNotEqual,x12cfoGreaterThanOrEqual,x12cfoGreaterThan,x12cfoBetween,x12cfoNotBetween,x12cfoContainsText,x12cfoNotContains,x12cfoBeginsWith,x12cfoEndsWith);
type TXc12TimePeriod = (x12tpToday,x12tpYesterday,x12tpTomorrow,x12tpLast7Days,x12tpThisMonth,x12tpLastMonth,x12tpNextMonth,x12tpThisWeek,x12tpLastWeek,x12tpNextWeek);
type TXc12CfvoType = (x12ctNum,x12ctPercent,x12ctMax,x12ctMin,x12ctFormula,x12ctPercentile);
type TXc12PageOrder = (x12poDownThenOver,x12poOverThenDown);
type TXc12Orientation = (x12oDefault,x12oPortrait,x12oLandscape);
type TXc12CellComments = (x12ccNone,x12ccAsDisplayed,x12ccAtEnd);
type TXc12PrintError = (x12peDisplayed,x12peBlank,x12peDash,x12peNA);
type TXc12SheetState = (x12ssVisible,x12ssHidden,x12ssVeryHidden);
type TXc12SheetViewType = (x12svtNormal,x12svtPageBreakPreview,x12svtPageLayout);
type TXc12PaneEnum = (x12pBottomRight,x12pTopRight,x12pBottomLeft,x12pTopLeft);
type TXc12PaneState = (x12psSplit,x12psFrozen,x12psFrozenSplit);
type TXc12Axis = (x12aAxisRow,x12aAxisCol,x12aAxisPage,x12aAxisValues,x12aUnknown);
type TXc12PivotAreaType = (x12patNone,x12patNormal,x12patData,x12patAll,x12patOrigin,x12patButton,x12patTopRight);
type TXc12DataValidationType = (x12dvtNone,x12dvtWhole,x12dvtDecimal,x12dvtList,x12dvtDate,x12dvtTime,x12dvtTextLength,x12dvtCustom);
type TXc12DataValidationErrorStyle = (x12dvesStop,x12dvesWarning,x12dvesInformation);
type TXc12DataValidationImeMode = (x12dvimNoControl,x12dvimOff,x12dvimOn,x12dvimDisabled,x12dvimHiragana,x12dvimFullKatakana,x12dvimHalfKatakana,x12dvimFullAlpha,x12dvimHalfAlpha,x12dvimFullHangul,x12dvimHalfHangul);
type TXc12DataValidationOperator = (x12dvoBetween,x12dvoNotBetween,x12dvoEqual,x12dvoNotEqual,x12dvoLessThan,x12dvoLessThanOrEqual,x12dvoGreaterThan,x12dvoGreaterThanOrEqual);
type TXc12WebSourceType = (x12wstSheet,x12wstPrintArea,x12wstAutoFilter,x12wstRange,x12wstChart,x12wstPivotTable,x12wstQuery,x12wstLabel);
type TXc12DvAspect = (x12daDVASPECT_CONTENT,x12daDVASPECT_ICON);
type TXc12OleUpdate = (x12ouOLEUPDATE_ALWAYS,x12ouOLEUPDATE_ONCALL);
type TXc12DataConsolidateFunction = (x12dcfAverage,x12dcfCount,x12dcfCountNums,x12dcfMax,x12dcfMin,x12dcfProduct,x12dcfStdDev,x12dcfStdDevp,x12dcfSum,x12dcfVar,x12dcfVarp);
type TXc12PhoneticType = (x12ptHalfwidthKatakana,x12ptFullwidthKatakana,x12ptHiragana,x12ptNoConversion);
type TXc12PhoneticAlignment = (x12paNoControl,x12paLeft,x12paCenter,x12paDistributed);

type TXc12ColumnOption = (xcoHidden,xcoBestFit,xcoCustomWidth,xcoPhonetic,xcoCollapsed);
     TXc12ColumnOptions = set of TXc12ColumnOption;

type TXc12ColumnHit = (xchLess,xchCol2,xchCol1Match,xchMatch,xchTargetInside,xchCol2Match,xchInside,xchCol1,xchGreater);
//================================
//   1----2                        xchLess
//          +------+
//================================
//       1----2....x               xchCol2 (can hit the entire target)
//          +------+
//================================
//          1----2                 xchCol1Match
//          +------+
//================================
//          1------2               xchMatch
//          +------+
//================================
//         1--------2              xchTargetInside
//          +------+
//================================
//            1----2               xchCol2Match
//          +------+
//================================
//           1----2                xchInside
//          +------+
//================================
//          x....1----2            xchCol1 (can hit the entire target)
//          +------+
//================================
//                   1----2        xchGreater
//          +------+
//================================

type TXc12Colors = class(TList)
private
     function GetItems(Index: integer): TXc12Color;
protected
public
     destructor Destroy; override;

     procedure Assign(AColors: TXc12Colors);

     procedure Add(AColor: TXc12Color); overload;
     procedure Add(ARGB: longword); overload;

     property Items[Index: integer]: TXc12Color read GetItems; default;
     end;

type TXc12Cfvo = class(TObject)
private
     function  GetAsFloat: double;
     procedure SetAsFloat(const Value: double);
     function  GetAsPercent: double;
protected
     FType: TXc12CfvoType;   // Required
     FVal: AxUCString;
     FGte: boolean;

     FPtgs: PXLSPtgs;
     FPtgsSz: integer;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(ACfvo: TXc12Cfvo);

     procedure Clear; overload;
     procedure Clear(AType: TXc12CfvoType; AVal: AxUCString; AGte: boolean = True); overload;

     property Type_: TXc12CfvoType read FType write FType;
     property Val: AxUCString read FVal write FVal;
     property AsFloat: double read GetAsFloat write SetAsFloat;
     property AsPercent: double read GetAsPercent;
     property Gte: boolean read FGte write FGte;

     property Ptgs: PXLSPtgs read FPtgs write FPtgs;
     property PtgsSz: integer read FPtgsSz write FPtgsSz;
     end;

type TXc12Cfvos = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Cfvo;
protected
public
     constructor Create;

     procedure Assign(ACfvos: TXc12Cfvos);

     function Add: TXc12Cfvo; overload;
     function Add(AType: TXc12CfvoType; const AVal: AxUCString; AGte: boolean = True): TXc12Cfvo; overload;

     property Items[Index: integer]: TXc12Cfvo read GetItems; default;
     end;

type TXc12IconSet = class(TObject)
protected
     FIconSet: TXc12IconSetType;
     FShowValue: boolean;
     FPercent: boolean;
     FReverse: boolean;
     FCfvos: TXc12Cfvos;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(AIconSet: TXc12IconSet);

     procedure Clear;

     property IconSet: TXc12IconSetType read FIconSet write FIconSet;
     property ShowValue: boolean read FShowValue write FShowValue;
     property Percent: boolean read FPercent write FPercent;
     property Reverse: boolean read FReverse write FReverse;
     property Cfvos: TXc12Cfvos read FCfvos;
     end;

type TXc12DataBar = class(TObject)
protected
     FMinLength: integer;
     FMaxLength: integer;
     FShowValue: boolean;
     FColor: TXc12Color;
     FCfvo1: TXc12Cfvo;
     FCfvo2: TXc12Cfvo;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(ADataBar: TXc12DataBar);

     procedure Clear;

     property MinLength: integer read FMinLength write FMinLength;
     property MaxLength: integer read FMaxLength write FMaxLength;
     property ShowValue: boolean read FShowValue write FShowValue;
     property Color: TXc12Color read FColor write FColor;
     property Cfvo1: TXc12Cfvo read FCfvo1;
     property Cfvo2: TXc12Cfvo read FCfvo2;
     end;

type TXc12ColorScale = class(TObject)
protected
     FCfvos: TXc12Cfvos;
     FColors: TXc12Colors;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(AColorScale: TXc12ColorScale);

     procedure Clear;

     property Cfvos: TXc12Cfvos read FCfvos;
     property Colors: TXc12Colors read FColors;
     end;

type TXc12CfRule = class(TObject)
private
     function  GetFormulas(Index: integer): AxUCString;
     procedure SerFormulas(Index: integer; const Value: AxUCString);
     function  GetPtgs(Index: integer): PXLSPtgs;
     function  GetPtgsSz(Index: integer): integer;
     procedure SetPtgs(Index: integer; const Value: PXLSPtgs);
     procedure SetPtgsSz(Index: integer; const Value: integer);
     function  GetValues(Index: integer): double;
     function  GetValuesCount: integer;
     procedure SetValues(Index: integer; const Value: double);
     procedure SetValuesCount(const Value: integer);
protected
     FFormula: array[0..2] of AxUCString;
     FPtgs: array[0..2] of PXLSPtgs;
     FPtgsSz: array[0..2] of integer;

     FType_: TXc12CfType;
     FDXF: TXc12DXF;
     FPriority: integer;
     FStopIfTrue: boolean;
     FAboveAverage: boolean;
     FPercent: boolean;
     FBottom: boolean;
     FOperator: TXc12ConditionalFormattingOperator;
     FText: AxUCString;
     FTimePeriod: TXc12TimePeriod;
     FRank: integer;
     FStdDev: integer;
     FEqualAverage: boolean;

     FColorScale: TXc12ColorScale;
     FDataBar: TXc12DataBar;
     FIconSet: TXc12IconSet;

     // Used by XLSSpreadSheet.
     FValAverage: double;
     FValMin    : double;
     FValMax    : double;
     FValues    : TDynDoubleArray;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(ARule: TXc12CfRule);

     procedure Clear;

     function  FormulaMaxCount: integer;

     property Formulas[Index: integer]: AxUCString read GetFormulas write SerFormulas;
     property Ptgs[Index: integer]: PXLSPtgs read GetPtgs write SetPtgs;
     property PtgsSz[Index: integer]: integer read GetPtgsSz write SetPtgsSz;

     property Type_: TXc12CfType read FType_ write FType_;
     property DXF: TXc12DXF read FDXF write FDXF;
     property Priority: integer read FPriority write FPriority;
     property StopIfTrue: boolean read FStopIfTrue write FStopIfTrue;
     property AboveAverage: boolean read FAboveAverage write FAboveAverage;
     property Percent: boolean read FPercent write FPercent;
     property Bottom: boolean read FBottom write FBottom;
     property Operator_: TXc12ConditionalFormattingOperator read FOperator write FOperator;
     property Text: AxUCString read FText write FText;
     property TimePeriod: TXc12TimePeriod read FTimePeriod write FTimePeriod;
     property Rank: integer read FRank write FRank;
     property StdDev: integer read FStdDev write FStdDev;
     property EqualAverage: boolean read FEqualAverage write FEqualAverage;

     property ColorScale: TXc12ColorScale read FColorScale;
     property DataBar: TXc12DataBar read FDataBar;
     property IconSet: TXc12IconSet read FIconSet;

     // Used by XLSSpreadSheet.
     procedure SortValues;
     property ValAverage: double read FValAverage write FValAverage;
     property ValMin: double read FValMin write FValMin;
     property ValMax: double read FValMax write FValMax;
     property ValuesCount: integer read GetValuesCount write SetValuesCount;
     property Values[Index: integer]: double read GetValues write SetValues;
     end;

type TXc12CfRules = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CfRule;
protected
public
     constructor Create;

     procedure Assign(ARules: TXc12CfRules);

     function Add: TXc12CfRule;

     property Items[Index: integer]: TXc12CfRule read GetItems; default;
     end;

type TXc12ConditionalFormatting = class(TXLSMoveCopyItem)
protected
     FPivot: boolean;
     FSQRef: TCellAreas;
     FCfRules: TXc12CfRules;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(ACondFmt: TXc12ConditionalFormatting); reintroduce;

     procedure Clear;

     property Pivot: boolean read FPivot write FPivot;
     property SQRef: TCellAreas read FSQRef;
     property CfRules: TXc12CfRules read FCfRules;
     end;

type TXc12ConditionalFormattings = class(TXLSMoveCopyList)
private
     function GetItems(Index: integer): TXc12ConditionalFormatting;
protected
     FClassFactory: TXLSClassFactory;
     FXc12Sheet: TObject;

     function  CreateMember: TXc12ConditionalFormatting; virtual; abstract;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     procedure Assign(ACondFmts: TXc12ConditionalFormattings);

     function Add: TXLSMoveCopyItem; override;

     function AddCF: TXc12ConditionalFormatting;

     property Xc12Sheet: TObject read FXc12Sheet write FXc12Sheet;
     property Items[Index: integer]: TXc12ConditionalFormatting read GetItems; default;
     end;

type TXc12HeaderFooter = class(TObject)
protected
     FDifferentOddEven: boolean;
     FDifferentFirst: boolean;
     FScaleWithDoc: boolean;
     FAlignWithMargins: boolean;
     FOddHeader: AxUCString;
     FOddFooter: AxUCString;
     FEvenHeader: AxUCString;
     FEvenFooter: AxUCString;
     FFirstHeader: AxUCString;
     FFirstFooter: AxUCString;
public
     constructor Create;

     procedure Clear;

     property DifferentOddEven: boolean read FDifferentOddEven write FDifferentOddEven;
     property DifferentFirst: boolean read FDifferentFirst write FDifferentFirst;
     property ScaleWithDoc: boolean read FScaleWithDoc write FScaleWithDoc;
     property AlignWithMargins: boolean read FAlignWithMargins write FAlignWithMargins;
     property OddHeader: AxUCString read FOddHeader write FOddHeader;
     property OddFooter: AxUCString read FOddFooter write FOddFooter;
     property EvenHeader: AxUCString read FEvenHeader write FEvenHeader;
     property EvenFooter: AxUCString read FEvenFooter write FEvenFooter;
     property FirstHeader: AxUCString read FFirstHeader write FFirstHeader;
     property FirstFooter: AxUCString read FFirstFooter write FFirstFooter;
     end;

type TXc12PageSetup = class(TObject)
private
     procedure SetPrinterSettings(const Value: TStream);
protected
     FPaperSize: integer;
     FScale: integer;
     FFirstPageNumber: integer;
     FFitToWidth: integer;
     FFitToHeight: integer;
     FPageOrder: TXc12PageOrder;
     FOrientation: TXc12Orientation;
     FUsePrinterDefaults: boolean;
     FBlackAndWhite: boolean;
     FDraft: boolean;
     FCellComments: TXc12CellComments;
     FUseFirstPageNumber: boolean;
     FErrors: TXc12PrintError;
     FHorizontalDpi: integer;
     FVerticalDpi: integer;
     FCopies: integer;

     FPrinterSettings: TStream;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property PaperSize: integer read FPaperSize write FPaperSize;
     property Scale: integer read FScale write FScale;
     property FirstPageNumber: integer read FFirstPageNumber write FFirstPageNumber;
     property FitToWidth: integer read FFitToWidth write FFitToWidth;
     property FitToHeight: integer read FFitToHeight write FFitToHeight;
     property PageOrder: TXc12PageOrder read FPageOrder write FPageOrder;
     property Orientation: TXc12Orientation read FOrientation write FOrientation;
     property UsePrinterDefaults: boolean read FUsePrinterDefaults write FUsePrinterDefaults;
     property BlackAndWhite: boolean read FBlackAndWhite write FBlackAndWhite;
     property Draft: boolean read FDraft write FDraft;
     property CellComments: TXc12CellComments read FCellComments write FCellComments;
     property UseFirstPageNumber: boolean read FUseFirstPageNumber write FUseFirstPageNumber;
     property Errors: TXc12PrintError read FErrors write FErrors;
     property HorizontalDpi: integer read FHorizontalDpi write FHorizontalDpi;
     property VerticalDpi: integer read FVerticalDpi write FVerticalDpi;
     property Copies: integer read FCopies write FCopies;
     property PrinterSettings: TStream read FPrinterSettings write SetPrinterSettings;
     end;

type TXc12PrintOptions = class(TObject)
protected
     FHorizontalCentered: boolean;
     FVerticalCentered: boolean;
     FHeadings: boolean;
     FGridLines: boolean;
     FGridLinesSet: boolean;
public
     constructor Create;

     procedure Clear;

     property HorizontalCentered: boolean read FHorizontalCentered write FHorizontalCentered;
     property VerticalCentered: boolean read FVerticalCentered write FVerticalCentered;
     property Headings: boolean read FHeadings write FHeadings;
     property GridLines: boolean read FGridLines write FGridLines;
     property GridLinesSet: boolean read FGridLinesSet write FGridLinesSet;
     end;

type TXc12PageMargins = class(TObject)
protected
     FLeft: double;
     FRight: double;
     FTop: double;
     FBottom: double;
     FHeader: double;
     FFooter: double;
public
     constructor Create;

     procedure Clear;

     property Left: double read FLeft write FLeft;
     property Right: double read FRight write FRight;
     property Top: double read FTop write FTop;
     property Bottom: double read FBottom write FBottom;
     property Header: double read FHeader write FHeader;
     property Footer: double read FFooter write FFooter;
     end;

type TXc12Break = class(TObject)
protected
     FId: integer;
     FMin: integer;
     FMax: integer;
     FMan: boolean;
     FPt: boolean;
public
     constructor Create;

     procedure Clear;

     property Id: integer read FId write FId;
     property Min: integer read FMin write FMin;
     property Max: integer read FMax write FMax;
     property Man: boolean read FMan write FMan;
     property Pt: boolean read FPt write FPt;
     end;

type TXc12PageBreaks = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Break;
protected
public
     constructor Create;

     procedure Clear; override;
     function Add: TXc12Break;
     function Find(const AId: integer): integer;
     function Hits(const AId: integer): integer;

     property Items[Index: integer]: TXc12Break read GetItems; default;
     end;

type TXc12Selection = class(TObject)
protected
     FPane: TXc12PaneEnum;
     FActiveCell: TXLSCellArea;
     FActiveCellId: integer;
     FSQRef: TCellAreas;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Pane: TXc12PaneEnum read FPane write FPane;
     property ActiveCell: TXLSCellArea read FActiveCell write FActiveCell;
     property ActiveCellId: integer read FActiveCellId write FActiveCellId;
     property SQRef: TCellAreas read FSQRef;
     end;

type TXc12Selections = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Selection;
protected
public
     constructor Create;

     function Add: TXc12Selection;

     property Items[Index: integer]: TXc12Selection read GetItems; default;
     end;

type TXc12Pane = class(TObject)
protected
     FXSplit: double;
     FYSplit: double;
     FTopLeftCell: TXLSCellArea;
     FActivePane: TXc12PaneEnum;
     FState: TXc12PaneState;

     FExcel97: boolean;
public
     constructor Create;

     procedure Clear;

     property XSplit: double read FXSplit write FXSplit;
     property YSplit: double read FYSplit write FYSplit;
     property TopLeftCell: TXLSCellArea read FTopLeftCell write FTopLeftCell;
     property ActivePane: TXc12PaneEnum read FActivePane write FActivePane;
     property State: TXc12PaneState read FState write FState;
     property Excel97: boolean read FExcel97 write FExcel97;
     end;

type TXc12PivotAreaReference = class(TObject)
protected
     FField: integer;
     FSelected: boolean;
     FByPosition: boolean;
     FRelative: boolean;
     FDefaultSubtotal: boolean;
     FSumSubtotal: boolean;
     FCountASubtotal: boolean;
     FAvgSubtotal: boolean;
     FMaxSubtotal: boolean;
     FMinSubtotal: boolean;
     FProductSubtotal: boolean;
     FCountSubtotal: boolean;
     FStdDevSubtotal: boolean;
     FStdDevPSubtotal: boolean;
     FVarSubtotal: boolean;
     FVarPSubtotal: boolean;
     FValues: TIntegerList;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Field: integer read FField write FField;
     property Selected: boolean read FSelected write FSelected;
     property ByPosition: boolean read FByPosition write FByPosition;
     property Relative: boolean read FRelative write FRelative;
     property DefaultSubtotal: boolean read FDefaultSubtotal write FDefaultSubtotal;
     property SumSubtotal: boolean read FSumSubtotal write FSumSubtotal;
     property CountASubtotal: boolean read FCountASubtotal write FCountASubtotal;
     property AvgSubtotal: boolean read FAvgSubtotal write FAvgSubtotal;
     property MaxSubtotal: boolean read FMaxSubtotal write FMaxSubtotal;
     property MinSubtotal: boolean read FMinSubtotal write FMinSubtotal;
     property ProductSubtotal: boolean read FProductSubtotal write FProductSubtotal;
     property CountSubtotal: boolean read FCountSubtotal write FCountSubtotal;
     property StdDevSubtotal: boolean read FStdDevSubtotal write FStdDevSubtotal;
     property StdDevPSubtotal: boolean read FStdDevPSubtotal write FStdDevPSubtotal;
     property VarSubtotal: boolean read FVarSubtotal write FVarSubtotal;
     property VarPSubtotal: boolean read FVarPSubtotal write FVarPSubtotal;
     property Values: TIntegerList read FValues;
     end;

type TXc12PivotAreaReferences = class(TObjectList)
private
     function GetItems(Index: integer): TXc12PivotAreaReference;
protected
public
     constructor Create;

     function Add: TXc12PivotAreaReference;

     property Items[Index: integer]: TXc12PivotAreaReference read GetItems; default;
     end;

type TXc12PivotArea = class(TObject)
protected
     FField: integer;
     FType_: TXc12PivotAreaType;
     FDataOnly: boolean;
     FLabelOnly: boolean;
     FGrandRow: boolean;
     FGrandCol: boolean;
     FCacheIndex: boolean;
     FOutline: boolean;
     FOffset: TXLSCellArea;
     FCollapsedLevelsAreSubtotals: boolean;
     FAxis: TXc12Axis;
     FFieldPosition: integer;
     FReferences: TXc12PivotAreaReferences;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Field: integer read FField write FField;
     property Type_: TXc12PivotAreaType read FType_ write FType_;
     property DataOnly: boolean read FDataOnly write FDataOnly;
     property LabelOnly: boolean read FLabelOnly write FLabelOnly;
     property GrandRow: boolean read FGrandRow write FGrandRow;
     property GrandCol: boolean read FGrandCol write FGrandCol;
     property CacheIndex: boolean read FCacheIndex write FCacheIndex;
     property Outline: boolean read FOutline write FOutline;
     property Offset: TXLSCellArea read FOffset write FOffset;
     property CollapsedLevelsAreSubtotals: boolean read FCollapsedLevelsAreSubtotals write FCollapsedLevelsAreSubtotals;
     property Axis: TXc12Axis read FAxis write FAxis;
     property FieldPosition: integer read FFieldPosition write FFieldPosition;
     property References: TXc12PivotAreaReferences read FReferences;
     end;

type TXc12PivotSelection = class(TObject)
protected
     FPane: TXc12PaneEnum;
     FShowHeader: boolean;
     FLabel_: boolean;
     FData: boolean;
     FExtendable: boolean;
     FCount: integer;
     FAxis: TXc12Axis;
     FDimension: integer;
     FStart: integer;
     FMin: integer;
     FMax: integer;
     FActiveRow: integer;
     FActiveCol: integer;
     FPreviousRow: integer;
     FPreviousCol: integer;
     FClick: integer;
     FRId: TXc12RId;

     FPivotArea: TXc12PivotArea;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Pane: TXc12PaneEnum read FPane write FPane;
     property ShowHeader: boolean read FShowHeader write FShowHeader;
     property Label_: boolean read FLabel_ write FLabel_;
     property Data: boolean read FData write FData;
     property Extendable: boolean read FExtendable write FExtendable;
     property Count: integer read FCount write FCount;
     property Axis: TXc12Axis read FAxis write FAxis;
     property Dimension: integer read FDimension write FDimension;
     property Start: integer read FStart write FStart;
     property Min: integer read FMin write FMin;
     property Max: integer read FMax write FMax;
     property ActiveRow: integer read FActiveRow write FActiveRow;
     property ActiveCol: integer read FActiveCol write FActiveCol;
     property PreviousRow: integer read FPreviousRow write FPreviousRow;
     property PreviousCol: integer read FPreviousCol write FPreviousCol;
     property Click: integer read FClick write FClick;
     property RId: TXc12RId read FRId write FRId;
     property PivotArea: TXc12PivotArea read FPivotArea;
     end;

type TXc12PivotSelections = class(TObjectList)
private
     function GetItems(Index: integer): TXc12PivotSelection;
protected
public
     constructor Create;
     destructor Destroy; override;

     function Add: TXc12PivotSelection;

     property Items[Index: integer]: TXc12PivotSelection read GetItems; default;
     end;

type TXc12CustomSheetView = class(TObject)
protected
     FClassFactory: TXLSClassFactory;

     FGuid: AxUCString;
     FScale: integer;
     FColorId: integer;
     FShowPageBreaks: boolean;
     FShowFormulas: boolean;
     FShowGridLines: boolean;
     FShowRowCol: boolean;
     FOutlineSymbols: boolean;
     FZeroValues: boolean;
     FFitToPage: boolean;
     FPrintArea: boolean;
     FFilter: boolean;
     FShowAutoFilter: boolean;
     FHiddenRows: boolean;
     FHiddenColumns: boolean;
     FState: TXc12SheetState;
     FFilterUnique: boolean;
     FView: TXc12SheetViewType;
     FShowRuler: boolean;
     FTopLeftCell: TXLSCellArea;
     FPane: TXc12Pane;
     FSelection: TXc12Selection;
     FRowBreaks: TXc12PageBreaks;
     FColBreaks: TXc12PageBreaks;
     FPageMargins: TXc12PageMargins;
     FPrintOptions: TXc12PrintOptions;
     FPageSetup: TXc12PageSetup;
     FHeaderFooter: TXc12HeaderFooter;
     FAutoFilter: TXc12AutoFilter;
public
     constructor Create(AClassFactory: TXLSClassFactory);
     destructor Destroy; override;

     procedure Clear;

     property GUID: AxUCString read FGUID write FGUID;
     property Scale: integer read FScale write FScale;
     property ColorId: integer read FColorId write FColorId;
     property ShowPageBreaks: boolean read FShowPageBreaks write FShowPageBreaks;
     property ShowFormulas: boolean read FShowFormulas write FShowFormulas;
     property ShowGridLines: boolean read FShowGridLines write FShowGridLines;
     property ShowRowCol: boolean read FShowRowCol write FShowRowCol;
     property OutlineSymbols: boolean read FOutlineSymbols write FOutlineSymbols;
     property ZeroValues: boolean read FZeroValues write FZeroValues;
     property FitToPage: boolean read FFitToPage write FFitToPage;
     property PrintArea: boolean read FPrintArea write FPrintArea;
     property Filter: boolean read FFilter write FFilter;
     property ShowAutoFilter: boolean read FShowAutoFilter write FShowAutoFilter;
     property HiddenRows: boolean read FHiddenRows write FHiddenRows;
     property HiddenColumns: boolean read FHiddenColumns write FHiddenColumns;
     property State: TXc12SheetState read FState write FState;
     property FilterUnique: boolean read FFilterUnique write FFilterUnique;
     property View: TXc12SheetViewType read FView write FView;
     property ShowRuler: boolean read FShowRuler write FShowRuler;
     property TopLeftCell: TXLSCellArea read FTopLeftCell write FTopLeftCell;
     property Pane: TXc12Pane read FPane;
     property Selection: TXc12Selection read FSelection;
     property RowBreaks: TXc12PageBreaks read FRowBreaks;
     property ColBreaks: TXc12PageBreaks read FColBreaks;
     property PageMargins: TXc12PageMargins read FPageMargins;
     property PrintOptions: TXc12PrintOptions read FPrintOptions;
     property PageSetup: TXc12PageSetup read FPageSetup;
     property HeaderFooter: TXc12HeaderFooter read FHeaderFooter;
     property AutoFilter: TXc12AutoFilter read FAutoFilter;
     end;

type TXc12CustomSheetViews = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CustomSheetView;
protected
     FClassFactory: TXLSClassFactory;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     function Add: TXc12CustomSheetView;

     property Items[Index: integer]: TXc12CustomSheetView read GetItems; default;
     end;

type TXc12SheetView = class(TObject)
protected
     FWindowProtection: boolean;
     FShowFormulas: boolean;
     FShowGridLines: boolean;
     FShowRowColHeaders: boolean;
     FShowZeros: boolean;
     FRightToLeft: boolean;
     FTabSelected: boolean;
     FShowRuler: boolean;
     FShowOutlineSymbols: boolean;
     FDefaultGridColor: boolean;
     FShowWhiteSpace: boolean;
     FView: TXc12SheetViewType;
     FTopLeftCell: TXLSCellArea;
     FColorId: integer;
     FZoomScale: integer;
     FZoomScaleNormal: integer;
     FZoomScaleSheetLayoutView: integer;
     FZoomScalePageLayoutView: integer;
     FWorkbookViewId: integer;
     FPane: TXc12Pane;
     FSelection: TXc12Selections;
     FPivotSelection: TXc12PivotSelections;

     // Used by chart sheets.
     FZoomToFit: boolean;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property WindowProtection: boolean read FWindowProtection write FWindowProtection;
     property ShowFormulas: boolean read FShowFormulas write FShowFormulas;
     property ShowGridLines: boolean read FShowGridLines write FShowGridLines;
     property ShowRowColHeaders: boolean read FShowRowColHeaders write FShowRowColHeaders;
     property ShowZeros: boolean read FShowZeros write FShowZeros;
     property RightToLeft: boolean read FRightToLeft write FRightToLeft;
     property TabSelected: boolean read FTabSelected write FTabSelected;
     property ShowRuler: boolean read FShowRuler write FShowRuler;
     property ShowOutlineSymbols: boolean read FShowOutlineSymbols write FShowOutlineSymbols;
     property DefaultGridColor: boolean read FDefaultGridColor write FDefaultGridColor;
     property ShowWhiteSpace: boolean read FShowWhiteSpace write FShowWhiteSpace;
     property View: TXc12SheetViewType read FView write FView;
     property TopLeftCell: TXLSCellArea read FTopLeftCell write FTopLeftCell;
     property ColorId: integer read FColorId write FColorId;
     property ZoomScale: integer read FZoomScale write FZoomScale;
     property ZoomScaleNormal: integer read FZoomScaleNormal write FZoomScaleNormal;
     property ZoomScaleSheetLayoutView: integer read FZoomScaleSheetLayoutView write FZoomScaleSheetLayoutView;
     property ZoomScalePageLayoutView: integer read FZoomScalePageLayoutView write FZoomScalePageLayoutView;
     property WorkbookViewId: integer read FWorkbookViewId write FWorkbookViewId;
     property Pane: TXc12Pane read FPane;
     property Selection: TXc12Selections read FSelection;
     property PivotSelection: TXc12PivotSelections read FPivotSelection;

     property ZoomToFit: boolean read FZoomToFit write FZoomToFit;
     end;

type TXc12SheetViews = class(TObjectList)
private
     function GetItems(Index: integer): TXc12SheetView;
protected
public
     constructor Create;

     function Add: TXc12SheetView;

     property Items[Index: integer]: TXc12SheetView read GetItems; default;
     end;

type TXc12CellSmartTagPr = class(TObject)
protected
     FKey: AxUCString;
     FVal: AxUCString;
public
     constructor Create;

     procedure Clear;

     property Key: AxUCString read FKey write FKey;
     property Val: AxUCString read FVal write FVal;
     end;

type TXc12CellSmartTag = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CellSmartTagPr;
protected
     FType_: integer;
     FDeleted: boolean;
     FXmlBased: boolean;
public
     constructor Create;

     function Add: TXc12CellSmartTagPr;
     procedure Clear; override;

     property Type_: integer read FType_ write FType_;
     property Deleted: boolean read FDeleted write FDeleted;
     property XmlBased: boolean read FXmlBased write FXmlBased;
     property Items[Index: integer]: TXc12CellSmartTagPr read GetItems; default;
     end;

type TXc12CellSmartTags = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CellSmartTag;
protected
     FRef: TCellRef;
public
     constructor Create;
     destructor Destroy; override;

     function Add: TXc12CellSmartTag;
     procedure Clear; override;

     property Ref: TCellRef read FRef write FRef;
     property Items[Index: integer]: TXc12CellSmartTag read GetItems;
     end;

type TXc12SmartTags = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CellSmartTags;
protected
public
     constructor Create;

     function Add: TXc12CellSmartTags;

     property Items[Index: integer]: TXc12CellSmartTags read GetItems; default;
     end;

type TXc12MergedCell = class(TCellArea)
protected
     FRef: TXLSCellArea;

     function  GetCol1: integer; override;
     function  GetRow1: integer; override;
     function  GetCol2: integer; override;
     function  GetRow2: integer; override;
     procedure SetCol1(const AValue: integer); override;
     procedure SetRow1(const AValue: integer); override;
     procedure SetCol2(const AValue: integer); override;
     procedure SetRow2(const AValue: integer); override;
public
     procedure Assign(const AItem: TXc12MergedCell);

     property Ref: TXLSCellArea read FRef write FRef;
     end;

type TXc12MergedCells = class(TSolidCellAreas)
private
     function GetItems(Index: integer): TXc12MergedCell;
protected
     FClassFactory: TXLSClassFactory;

     function  CreateObject: TCellArea; override;
     function  CreateMember: TXc12MergedCell;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     function Add(const ARef: TXLSCellArea): TXc12MergedCell; overload;

     property Items[Index: integer]: TXc12MergedCell read GetItems; default;
     end;

type TXc12Hyperlink = class(TCellArea)
private
     procedure SetRawAddress(const Value: AxUCString);
protected
     FHyperlinkType: TXLSHyperlinkType;

     FDisplay      : AxUCString;
     FAddress      : AxUCString;
     FLocation     : AxUCString;
     FRef          : TXLSCellArea;
     FTooltip      : AxUCString;
     // Used by Excel 97
     FTargetFrame  : AxUCString;
     FScreenTip    : AxUCString;
     FHyperlinkEnc : TXLSHyperlinkType;

     procedure Parse; virtual;
public
     constructor Create;

     procedure Assign(AHLink: TXc12Hyperlink);

     procedure Clear;

     property HyperlinkType: TXLSHyperlinkType read FHyperlinkType write FHyperlinkType;
     property Display      : AxUCString read FDisplay write FDisplay;
     property RawAddress   : AxUCString read FAddress write SetRawAddress;
     property Location     : AxUCString read FLocation write FLocation;
     property Ref          : TXLSCellArea read FRef write FRef;
     property Tooltip      : AxUCString read FTooltip write FTooltip;

     property TargetFrame  : AxUCString read FTargetFrame write FTargetFrame;
     property ScreenTip    : AxUCString read FScreenTip write FScreenTip;
     property HyperlinkEnc : TXLSHyperlinkType read FHyperlinkEnc write FHyperlinkEnc;
     end;

type TXc12Hyperlinks = class(TCellAreas)
private
     function GetItems(Index: integer): TXc12Hyperlink;
protected
     FClassFactory: TXLSClassFactory;

     function  CreateMember: TXc12Hyperlink;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     procedure Assign(AHLinks: TXc12Hyperlinks);

     function Add: TXc12Hyperlink;

     property Items[Index: integer]: TXc12Hyperlink read GetItems; default;
     end;

type TXc12Column = class(TXc12Data)
protected
     FMin: integer;
     FMax: integer;
     FWidth: double;
     FStyle: TXc12XF;
     FOutlineLevel: integer;
     FOptions: TXc12ColumnOptions;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;
     procedure Assign(ASource: TXc12Column);
     function  Equal(ACol: TXc12Column): boolean;
     // Equal, except for position.
     function  EqualProps(ACol: TXc12Column): boolean;
     function  Hit(ACol: integer): TXc12ColumnHit; overload;
     function  Hit(ACol1,ACol2: integer): TXc12ColumnHit; overload;

     property Min: integer read FMin write FMin;
     property Max: integer read FMax write FMax;
     property Width: double read FWidth write FWidth;
     property Style: TXc12XF read FStyle write FStyle;
     property OutlineLevel: integer read FOutlineLevel write FOutlineLevel;
     property Options: TXc12ColumnOptions read FOptions write FOptions;
     end;

type TXc12Columns = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Column;
protected
     FDefColWidth: double;
public
     constructor Create;

     function  Add: TXc12Column; overload;
     function  Add(AStyle: TXc12XF): TXc12Column; overload;
     procedure Add(ACol: TXc12Column; AStyle: TXc12XF); overload;
     procedure NilAndDelete(AIndex: integer);

     property DefColWidth: double read FDefColWidth write FDefColWidth;
     property Items[Index: integer]: TXc12Column read GetItems; default;
     end;

type TXc12WebPublishItem = class(TObject)
protected
     FId: integer;
     FDivId: AxUCString;
     FSourceType: TXc12WebSourceType;
     FSourceRef: TCellRef;
     FSourceObject: AxUCString;
     FDestinationFile: AxUCString;
     FTitle: AxUCString;
     FAutoRepublish: boolean;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Id: integer read FId write FId;
     property DivId: AxUCString read FDivId write FDivId;
     property SourceType: TXc12WebSourceType read FSourceType write FSourceType;
     property SourceRef: TCellRef read FSourceRef write FSourceRef;
     property SourceObject: AxUCString read FSourceObject write FSourceObject;
     property DestinationFile: AxUCString read FDestinationFile write FDestinationFile;
     property Title: AxUCString read FTitle write FTitle;
     property AutoRepublish: boolean read FAutoRepublish write FAutoRepublish;
     end;

type TXc12WebPublishItems = class(TObjectList)
private
     function GetItems(Index: integer): TXc12WebPublishItem;
protected
public
     constructor Create;

     function Add: TXc12WebPublishItem;

     property Items[Index: integer]: TXc12WebPublishItem read GetItems; default;
     end;

type TXc12Control = class(TObject)
protected
     FShapeId: integer;
     FRId: TXc12RId;
     FName: AxUCString;
public
     constructor Create;

     procedure Clear;

     property ShapeId: integer read FShapeId write FShapeId;
     property RId: TXc12RId read FRId write FRId;
     property Name: AxUCString read FName write FName;
     end;

type TXc12Controls = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Control;
protected
public
     constructor Create;

     function Add: TXc12Control;

     property Items[Index: integer]: TXc12Control read GetItems; default;
     end;

type TXc12OleObject = class(TObject)
protected
     FProgId: AxUCString;
     FDvAspect: TXc12DvAspect;
     FLink: AxUCString;
     FOleUpdate: TXc12OleUpdate;
     FAutoLoad: boolean;
     FShapeId: integer;
     FRId: TXc12RId;
public
     constructor Create;

     procedure Clear;

     property ProgId: AxUCString read FProgId write FProgId;
     property DvAspect: TXc12DvAspect read FDvAspect write FDvAspect;
     property Link: AxUCString read FLink write FLink;
     property OleUpdate: TXc12OleUpdate read FOleUpdate write FOleUpdate;
     property AutoLoad: boolean read FAutoLoad write FAutoLoad;
     property ShapeId: integer read FShapeId write FShapeId;
     property RId: TXc12RId read FRId write FRId;
     end;

type TXc12OleObjects = class(TObjectList)
private
     function GetItems(Index: integer): TXc12OleObject;
protected
public
     constructor Create;

     function Add: TXc12OleObject;

     property Items[Index: integer]: TXc12OleObject read GetItems; default;
     end;

type TXc12IgnoredError = class(TObject)
protected
     FSqref: TCellAreas;
     FEvalError: boolean;
     FTwoDigitTextYear: boolean;
     FNumberStoredAsText: boolean;
     FFormula: boolean;
     FFormulaRange: boolean;
     FUnlockedFormula: boolean;
     FEmptyCellReference: boolean;
     FListDataValidation: boolean;
     FCalculatedColumn: boolean;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Sqref: TCellAreas read FSqref;
     property EvalError: boolean read FEvalError write FEvalError;
     property TwoDigitTextYear: boolean read FTwoDigitTextYear write FTwoDigitTextYear;
     property NumberStoredAsText: boolean read FNumberStoredAsText write FNumberStoredAsText;
     property Formula: boolean read FFormula write FFormula;
     property FormulaRange: boolean read FFormulaRange write FFormulaRange;
     property UnlockedFormula: boolean read FUnlockedFormula write FUnlockedFormula;
     property EmptyCellReference: boolean read FEmptyCellReference write FEmptyCellReference;
     property ListDataValidation: boolean read FListDataValidation write FListDataValidation;
     property CalculatedColumn: boolean read FCalculatedColumn write FCalculatedColumn;
     end;

type TXc12IgnoredErrors = class(TObjectList)
private
     function GetItems(Index: integer): TXc12IgnoredError;
protected
public
     constructor Create;

     function Add: TXc12IgnoredError;

     property Items[Index: integer]: TXc12IgnoredError read GetItems; default;
     end;

type TXc12CustomProperty = class(TObject)
protected
     FName: AXUCString;
     FRId: AXUCString;
public
     constructor Create;

     procedure Clear;

     property Name: AXUCString read FName write FName;
     property RId: AXUCString read FRId write FRId;
     end;

type TXc12CustomProperties = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CustomProperty;
protected
public
     constructor Create;

     function Add: TXc12CustomProperty;

     property Items[Index: integer]: TXc12CustomProperty read GetItems; default;
     end;

type TXc12DataValidation = class(TXLSMoveCopyItem)
protected
     FType_: TXc12DataValidationType;
     FErrorStyle: TXc12DataValidationErrorStyle;
     FImeMode: TXc12DataValidationImeMode;
     FOperator_: TXc12DataValidationOperator;
     FAllowBlank: boolean;
     FShowDropDown: boolean;
     FShowInputMessage: boolean;
     FShowErrorMessage: boolean;
     FErrorTitle: AxUCString;
     FError: AxUCString;
     FPromptTitle: AxUCString;
     FPrompt: AxUCString;
     FSqref: TCellAreas;
     FFormula1: AxUCString;
     FFormula2: AxUCString;
public
     constructor Create;
     destructor Destroy; override;

     procedure Assign(AValidation: TXc12DataValidation); reintroduce;

     procedure Clear;

     property Type_: TXc12DataValidationType read FType_ write FType_;
     property ErrorStyle: TXc12DataValidationErrorStyle read FErrorStyle write FErrorStyle;
     property ImeMode: TXc12DataValidationImeMode read FImeMode write FImeMode;
     property Operator_: TXc12DataValidationOperator read FOperator_ write FOperator_;
     property AllowBlank: boolean read FAllowBlank write FAllowBlank;
     property ShowDropDown: boolean read FShowDropDown write FShowDropDown;
     property ShowInputMessage: boolean read FShowInputMessage write FShowInputMessage;
     property ShowErrorMessage: boolean read FShowErrorMessage write FShowErrorMessage;
     property ErrorTitle: AxUCString read FErrorTitle write FErrorTitle;
     property Error: AxUCString read FError write FError;
     property PromptTitle: AxUCString read FPromptTitle write FPromptTitle;
     property Prompt: AxUCString read FPrompt write FPrompt;
     property Sqref: TCellAreas read FSqref write FSqref;
     property Formula1: AxUCString read FFormula1 write FFormula1;
     property Formula2: AxUCString read FFormula2 write FFormula2;
     end;

type TXc12DataValidations = class(TXLSMoveCopyList)
private
     function GetItems(Index: integer): TXc12DataValidation;
protected
     FClassFactory: TXLSClassFactory;
     FDisablePrompts: boolean;
     FXWindow: integer;
     FYWindow: integer;

     function  CreateMember: TXc12DataValidation;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     procedure Assign(AValidations: TXc12DataValidations);

     procedure Clear; override;
     function Add: TXLSMoveCopyItem; override;
     function AddDV: TXc12DataValidation;

     property DisablePrompts: boolean read FDisablePrompts write FDisablePrompts;
     property XWindow: integer read FXWindow write FXWindow;
     property YWindow: integer read FYWindow write FYWindow;

     property Items[Index: integer]: TXc12DataValidation read GetItems; default;
     end;

type TXc12DataRef = class(TObject)
protected
     FRef: TXLSCellArea;
     FName: AxUCString;
     FSheet: AxUCString;
     FRId: TXc12RId;
public
     constructor Create;

     procedure Clear;

     property Ref: TXLSCellArea read FRef write FRef;
     property Name: AxUCString read FName write FName;
     property Sheet: AxUCString read FSheet write FSheet;
     property RId: TXc12RId read FRId write FRId;
     end;

type TXc12DataRefs = class(TObjectList)
private
     function GetItems(Index: integer): TXc12DataRef;
protected
public
     constructor Create;

     function Add: TXc12DataRef;

     property Items[Index: integer]: TXc12DataRef read GetItems; default;
     end;

type TXc12InputCell = class(TObject)
protected
     FRow: integer;
     FCol: integer;
     FDeleted: boolean;
     FUndone: boolean;
     FVal: AxUCString;
     FNumFmtId: integer;
public
     constructor Create;

     procedure Clear;

     property Row: integer read FRow write FRow;
     property Col: integer read FCol write FCol;
     property Deleted: boolean read FDeleted write FDeleted;
     property Undone: boolean read FUndone write FUndone;
     property Val: AxUCString read FVal write FVal;
     property NumFmtId: integer read FNumFmtId write FNumFmtId;
     end;

type Tx12InputCells = class(TObjectList)
private
     function GetItems(Index: integer): TXc12InputCell;
protected
public
     constructor Create;

     function Add: TXc12InputCell; overload;
     function Add(ACellRef: AxUCString): TXc12InputCell; overload;

     property Items[Index: integer]: TXc12InputCell read GetItems; default;
     end;

type TXc12Scenario = class(TObject)
protected
     FName: AxUCString;
     FLocked: boolean;
     FHidden: boolean;
     FCount: integer;
     FUser: AxUCString;
     FComment: AxUCString;
     FInputCells: Tx12InputCells;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Name: AxUCString read FName write FName;
     property Locked: boolean read FLocked write FLocked;
     property Hidden: boolean read FHidden write FHidden;
     property Count: integer read FCount write FCount;
     property User: AxUCString read FUser write FUser;
     property Comment: AxUCString read FComment write FComment;
     property InputCells: Tx12InputCells read FInputCells;
     end;

type TXc12Scenarios = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Scenario;
protected
     FCurrent: integer;
     FShow: integer;
     FSqref: TCellAreas;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear; override;
     function Add: TXc12Scenario;

     property Items[Index: integer]: TXc12Scenario read GetItems; default;

     property Current: integer read FCurrent write FCurrent;
     property Show: integer read FShow write FShow;
     property _Sqref: TCellAreas read FSqref;
     end;

type TXc12ProtectedRange = class(TObject)
protected
     FPassword: integer;
     FSqref: TCellAreas;
     FName: AxUCString;                    // Required
     FSecurityDescriptor: AxUCString;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Password: integer read FPassword write FPassword;
     property Sqref: TCellAreas read FSqref write FSqref;
     property Name: AxUCString read FName write FName;
     property SecurityDescriptor: AxUCString read FSecurityDescriptor write FSecurityDescriptor;
     end;

type TXc12ProtectedRanges = class(TObjectList)
private
     function GetItems(Index: integer): TXc12ProtectedRange;
protected
public
     constructor Create;

     function Add: TXc12ProtectedRange;

     property Items[Index: integer]: TXc12ProtectedRange read GetItems; default;
     end;

type TXc12OutlinePr = class(TObject)
protected
     FApplyStyles: boolean;
     FSummaryBelow: boolean;
     FSummaryRight: boolean;
     FShowOutlineSymbols: boolean;
public
     constructor Create;

     procedure Clear;

     property ApplyStyles: boolean read FApplyStyles write FApplyStyles;
     property SummaryBelow: boolean read FSummaryBelow write FSummaryBelow;
     property SummaryRight: boolean read FSummaryRight write FSummaryRight;
     property ShowOutlineSymbols: boolean read FShowOutlineSymbols write FShowOutlineSymbols;
     end;

type TXc12PageSetupPr = class(TObject)
protected
     FAutoPageBreaks: boolean;
     FFitToPage: boolean;
public
     constructor Create;

     procedure Clear;

     property AutoPageBreaks: boolean read FAutoPageBreaks write FAutoPageBreaks;
     property FitToPage: boolean read FFitToPage write FFitToPage;
     end;

type TXc12SheetPr = class(TObject)
protected
     FSyncHorizontal: boolean;
     FSyncVertical: boolean;
     FSyncRef: TXLSCellArea;
     FTransitionEvaluation: boolean;
     FTransitionEntry: boolean;
     FPublished_: boolean;
     FCodeName: AxUCString;
     FFilterMode: boolean;
     FEnableFormatConditionsCalculation: boolean;
     FTabColor: TXc12Color;
     FOutlinePr: TXc12OutlinePr;
     FPageSetupPr: TXc12PageSetupPr;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property SyncHorizontal: boolean read FSyncHorizontal write FSyncHorizontal;
     property SyncVertical: boolean read FSyncVertical write FSyncVertical;
     property SyncRef: TXLSCellArea read FSyncRef write FSyncRef;
     property TransitionEvaluation: boolean read FTransitionEvaluation write FTransitionEvaluation;
     property TransitionEntry: boolean read FTransitionEntry write FTransitionEntry;
     property Published_: boolean read FPublished_ write FPublished_;
     property CodeName: AxUCString read FCodeName write FCodeName;
     property FilterMode: boolean read FFilterMode write FFilterMode;
     property EnableFormatConditionsCalculation: boolean read FEnableFormatConditionsCalculation write FEnableFormatConditionsCalculation;
     property TabColor: TXc12Color read FTabColor write FTabColor;
     property OutlinePr: TXc12OutlinePr read FOutlinePr;
     property PageSetupPr: TXc12PageSetupPr read FPageSetupPr;
     end;

type TXc12SheetFormatPr = class(TObject)
protected
     FBaseColWidth    : integer;
     FDefaultColWidth : double;
     FDefaultRowHeight: double;
     FCustomHeight    : boolean;
     FZeroHeight      : boolean;
     FThickTop        : boolean;
     FThickBottom     : boolean;
     FOutlineLevelRow : integer;
     FOutlineLevelCol : integer;
public
     constructor Create;

     procedure Clear;

     procedure Assign(ASrc: TXc12SheetFormatPr);

     property BaseColWidth: integer read FBaseColWidth write FBaseColWidth;
     property DefaultColWidth: double read FDefaultColWidth write FDefaultColWidth;
     property DefaultRowHeight: double read FDefaultRowHeight write FDefaultRowHeight;
     property CustomHeight: boolean read FCustomHeight write FCustomHeight;
     property ZeroHeight: boolean read FZeroHeight write FZeroHeight;
     property ThickTop: boolean read FThickTop write FThickTop;
     property ThickBottom: boolean read FThickBottom write FThickBottom;
     property OutlineLevelRow: integer read FOutlineLevelRow write FOutlineLevelRow;
     property OutlineLevelCol: integer read FOutlineLevelCol write FOutlineLevelCol;
     end;

type TXc12SheetCalcPr = class(TObject)
protected
     FFullCalcOnLoad: boolean;
public
     constructor Create;

     procedure Clear;

     property FullCalcOnLoad: boolean read FFullCalcOnLoad write FFullCalcOnLoad;
     end;

type TXc12SheetProtection = class(TObject)
private
     function  GetPasswordAsString: AxUCString;
     procedure SetPasswordAsString(const Value: AxUCString);
protected
     FPassword: word;
     FSheet: boolean;
     FObjects: boolean;
     FScenarios: boolean;
     FFormatCells: boolean;
     FFormatColumns: boolean;
     FFormatRows: boolean;
     FInsertColumns: boolean;
     FInsertRows: boolean;
     FInsertHyperlinks: boolean;
     FDeleteColumns: boolean;
     FDeleteRows: boolean;
     FSelectLockedCells: boolean;
     FSort: boolean;
     FAutoFilter: boolean;
     FPivotTables: boolean;
     FSelectUnlockedCells: boolean;
public
     constructor Create;

     procedure Clear;

     property Password: word read FPassword write FPassword;
     property PasswordAsString: AxUCString read GetPasswordAsString write SetPasswordAsString;
     property Sheet: boolean read FSheet write FSheet;
     property Objects: boolean read FObjects write FObjects;
     property Scenarios: boolean read FScenarios write FScenarios;
     property FormatCells: boolean read FFormatCells write FFormatCells;
     property FormatColumns: boolean read FFormatColumns write FFormatColumns;
     property FormatRows: boolean read FFormatRows write FFormatRows;
     property InsertColumns: boolean read FInsertColumns write FInsertColumns;
     property InsertRows: boolean read FInsertRows write FInsertRows;
     property InsertHyperlinks: boolean read FInsertHyperlinks write FInsertHyperlinks;
     property DeleteColumns: boolean read FDeleteColumns write FDeleteColumns;
     property DeleteRows: boolean read FDeleteRows write FDeleteRows;
     property SelectLockedCells: boolean read FSelectLockedCells write FSelectLockedCells;
     property Sort: boolean read FSort write FSort;
     property AutoFilter: boolean read FAutoFilter write FAutoFilter;
     property PivotTables: boolean read FPivotTables write FPivotTables;
     property SelectUnlockedCells: boolean read FSelectUnlockedCells write FSelectUnlockedCells;
     end;

type TXc12DataConsolidate = class(TObject)
protected
     FFunction: TXc12DataConsolidateFunction;
     FLeftLabels: boolean;
     FTopLabels: boolean;
     FLink: boolean;
     FDataRefs: TXc12DataRefs;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Function_: TXc12DataConsolidateFunction read FFunction write FFunction;
     property LeftLabels: boolean read FLeftLabels write FLeftLabels;
     property TopLabels: boolean read FTopLabels write FTopLabels;
     property Link: boolean read FLink write FLink;
     property DataRefs: TXc12DataRefs read FDataRefs;
     end;

type TXc12PhoneticPr = class(TObject)
protected
     FFontId: integer;
     FType: TXc12PhoneticType;
     FAlignment: TXc12PhoneticAlignment;
public
     constructor Create;

     procedure Clear;

     property FontId: integer read FFontId write FFontId;
     property Type_: TXc12PhoneticType read FType write FType;
     property Alignment: TXc12PhoneticAlignment read FAlignment write FAlignment;
     end;

type TXc12DataWorksheet = class(TIndexObject)
private
     function  GetDimension: TXLSCellArea;
     procedure SetDimension(const Value: TXLSCellArea);
     function  GetQuotedName: AxUCString;
protected
     FClassFactory         : TXLSClassFactory;

     FSST                  : TXc12DataSST;
     FStyles               : TXc12DataStyleSheet;

     FCells                : TXLSCellMMU;

     FColumns              : TXc12Columns;
     FComments             : TXc12Comments;
     FHyperlinks           : TXc12Hyperlinks;
     FMergedCells          : TXc12MergedCells;
     FAutofilter           : TXc12AutoFilter;
     FCustomSheetViews     : TXc12CustomSheetViews;
     FSheetViews           : TXc12SheetViews;
     FPivotTables          : TCT_pivotTableDefinitions;

     FSheetPr              : TXc12SheetPr;
     FSheetFormatPr        : TXc12SheetFormatPr;
     FSheetCalcPr          : TXc12SheetCalcPr;
     FSheetProtection      : TXc12SheetProtection;
     FProtectedRanges      : TXc12ProtectedRanges;
     FScenarios            : TXc12Scenarios;
     FSortState            : TXc12SortState;
     FDataConsolidate      : TXc12DataConsolidate;
     FPhoneticPr           : TXc12PhoneticPr;
     FConditionalFormatting: TXc12ConditionalFormattings;
     FDataValidations      : TXc12DataValidations;
     FPrintOptions         : TXc12PrintOptions;
     FPageMargins          : TXc12PageMargins;
     FPageSetup            : TXc12PageSetup;
     FHeaderFooter         : TXc12HeaderFooter;
     FRowBreaks            : TXc12PageBreaks;
     FColBreaks            : TXc12PageBreaks;
     FCustomProperties     : TXc12CustomProperties;
     FCellWatches          : TCellRefs;
     FIgnoredErrors        : TXc12IgnoredErrors;
     FSmartTags            : TXc12SmartTags;
     FOleObjects           : TXc12OleObjects;
     FControls             : TXc12Controls;
     FWebPublishItems      : TXc12WebPublishItems;

     FTables               : TXc12Tables;
     FDrawing              : TXPGDocXLSXDrawing;
     FVmlDrawing           : TXpgSimpleDOM;
     FVmlDrawingRels       : TStringList;

     // Originally in TXc12DataWorkbook.Sheets[x], but moved here.
     FName          : AxUCString;
     FSheetId       : integer;
     FState         : TXc12Visibility;
     FRId           : TXc12RId;
     FIsChartSheet  : boolean;
public
     constructor Create(AClassFactory: TXLSClassFactory; AIndex: integer; ASST: TXc12DataSST; AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure SetDefaultValues;

     procedure Clear;

     property Cells                : TXLSCellMMU read FCells;

     property Columns              : TXc12Columns read FColumns;
     property Comments             : TXc12Comments read FComments;
     property Tables               : TXc12Tables read FTables;
     property Hyperlinks           : TXc12Hyperlinks read FHyperlinks;
     property MergedCells          : TXc12MergedCells read FMergedCells;
     property Autofilter           : TXc12AutoFilter read FAutofilter;
     property CustomSheetViews     : TXc12CustomSheetViews read FCustomSheetViews;
     property SheetViews           : TXc12SheetViews read FSheetViews;
     property PivotTables          : TCT_pivotTableDefinitions read FPivotTables;

     property SheetPr              : TXc12SheetPr read FSheetPr;
     property Dimension            : TXLSCellArea read GetDimension write SetDimension;
     property SheetFormatPr        : TXc12SheetFormatPr read FSheetFormatPr;
     property SheetCalcPr          : TXc12SheetCalcPr read FSheetCalcPr;
     property SheetProtection      : TXc12SheetProtection read FSheetProtection;
     property ProtectedRanges      : TXc12ProtectedRanges read FProtectedRanges;
     property Scenarios            : TXc12Scenarios read FScenarios;
     property SortState            : TXc12SortState read FSortState;
     property DataConsolidate      : TXc12DataConsolidate read FDataConsolidate;
     property PhoneticPr           : TXc12PhoneticPr read FPhoneticPr;
     property ConditionalFormatting: TXc12ConditionalFormattings read FConditionalFormatting;
     property DataValidations      : TXc12DataValidations read FDataValidations;
     property PrintOptions         : TXc12PrintOptions read FPrintOptions;
     property PageMargins          : TXc12PageMargins read FPageMargins;
     property PageSetup            : TXc12PageSetup read FPageSetup;
     property HeaderFooter         : TXc12HeaderFooter read FHeaderFooter;
     property RowBreaks            : TXc12PageBreaks read FRowBreaks;
     property ColBreaks            : TXc12PageBreaks read FColBreaks;
     property CustomProperties     : TXc12CustomProperties read FCustomProperties;
     property CellWatches          : TCellRefs read FCellWatches;
     property IgnoredErrors        : TXc12IgnoredErrors read FIgnoredErrors;
     property SmartTags            : TXc12SmartTags read FSmartTags;
     property OleObjects           : TXc12OleObjects read FOleObjects;
     property Controls             : TXc12Controls read FControls;
     property WebPublishItems      : TXc12WebPublishItems read FWebPublishItems;

     property Drawing              : TXPGDocXLSXDrawing read FDrawing;
     property VmlDrawing           : TXpgSimpleDOM read FVmlDrawing;
     property VmlDrawingRels       : TStringList read FVmlDrawingRels;

     property Name                 : AxUCString read FName write FName;
     // Returns the name quoted if the name contains other chars than letters and numbers, such as  spaces.
     property QuotedName           : AxUCString read GetQuotedName;
     // Don't use. It's only used for the display order of sheets.
     property _SheetId             : integer read FSheetId write FSheetId;
     property State                : TXc12Visibility read FState write FState;
     property RId                  : TXc12RId read FRId write FRId;
     property IsChartSheet         : boolean read FIsChartSheet write FIsChartSheet;
     end;

type TXc12DataWorksheets = class(TIndexObjectList)
private
     function GetItems(Index: integer): TXc12DataWorksheet;
protected
     FClassFactory: TXLSClassFactory;
     FSST         : TXc12DataSST;
     FStyles      : TXc12DataStyleSheet;
public
     constructor Create(AClassFactory: TXLSClassFactory; ASST: TXc12DataSST; AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure Clear; reintroduce;

     function  Add(const ASheetId: integer): TXc12DataWorksheet;
     function  Insert(const AIndex: integer): TXc12DataWorksheet;

     function  Find(const AName: AxUCString): integer;
     function  FindSheetFromCells(const ACells: TXLSCellMMU): TXc12DataWorksheet;

//     property IdOrder[Index: integer]: TXc12DataWorksheet read GetIdOrder;
     property Items[Index: integer]: TXc12DataWorksheet read GetItems; default;
     end;

implementation

{ TXc12DataWorksheet }

procedure TXc12DataWorksheet.Clear;
begin
  FCells.Clear;

  FColumns.Clear;
  FComments.Clear;
  FTables.Clear;
  FDrawing.WsDr.Clear;
  FVmlDrawing.Clear;
  FVmlDrawingRels.Clear;
  FHyperlinks.Clear;
  FMergedCells.Clear;
  FAutofilter.Clear;
  FCustomSheetViews.Clear;
  FSheetViews.Clear;
  FPivotTables.Clear;

  FSheetPr.Clear;
  FSheetFormatPr.Clear;
  FSheetCalcPr.Clear;
  FSheetProtection.Clear;
  FProtectedRanges.Clear;
  FScenarios.Clear;
  FSortState.Clear;
  FDataConsolidate.Clear;
  FPhoneticPr.Clear;
  FConditionalFormatting.Clear;
  FDataValidations.Clear;
  FPrintOptions.Clear;
  FPageMargins.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FRowBreaks.Clear;
  FColBreaks.Clear;
  FCustomProperties.Clear;
  FCellWatches.Clear;
  FIgnoredErrors.Clear;
  FSmartTags.Clear;
  FOleObjects.Clear;
  FControls.Clear;
  FWebPublishItems.Clear;

  SetDefaultValues;
end;

constructor TXc12DataWorksheet.Create(AClassFactory: TXLSClassFactory; AIndex: integer; ASST: TXc12DataSST; AStyles: TXc12DataStyleSheet);
begin
  FClassFactory := AClassFactory;

  FIndex := AIndex;
  FSST := ASST;
  FStyles := AStyles;

  FCells := TXLSCellMMU.Create(FIndex,Nil,FSST,FStyles);

  FColumns := TXc12Columns.Create;
  FComments := TXc12Comments.Create(AStyles);
  FTables := TXc12Tables.Create(FClassFactory);
  FDrawing := TXPGDocXLSXDrawing(FClassFactory.CreateAClass(xcftDrawing));
  FVmlDrawing := TXpgSimpleDOM.Create;
  FVmlDrawingRels := TStringList.Create;
  FHyperlinks := TXc12Hyperlinks(FClassFactory.CreateAClass(xcftHyperlinks));
  FMergedCells :=  TXc12MergedCells(FClassFactory.CreateAClass(xcftMergedCells));
  FAutofilter := TXc12AutoFilter(FClassFactory.CreateAClass(xcftAutofilter,Self));
  FCustomSheetViews := TXc12CustomSheetViews.Create(FClassFactory);
  FSheetViews := TXc12SheetViews.Create;
  FPivotTables := TCT_pivotTableDefinitions.Create;

  FSheetPr := TXc12SheetPr.Create;
  FSheetFormatPr := TXc12SheetFormatPr.Create;
  FSheetCalcPr := TXc12SheetCalcPr.Create;
  FSheetProtection := TXc12SheetProtection.Create;
  FProtectedRanges := TXc12ProtectedRanges.Create;
  FScenarios := TXc12Scenarios.Create;
  FSortState := TXc12SortState.Create;
  FDataConsolidate := TXc12DataConsolidate.Create;
  FPhoneticPr := TXc12PhoneticPr.Create;
  FConditionalFormatting := TXc12ConditionalFormattings(FClassFactory.CreateAClass(xcftConditionalFormats));
  FConditionalFormatting.Xc12Sheet := Self;
  FDataValidations := TXc12DataValidations(FClassFactory.CreateAClass(xcftDataValidations));
  FPrintOptions := TXc12PrintOptions.Create;
  FPageMargins := TXc12PageMargins.Create;
  FPageSetup := TXc12PageSetup.Create;
  FHeaderFooter := TXc12HeaderFooter.Create;
  FRowBreaks := TXc12PageBreaks.Create;
  FColBreaks := TXc12PageBreaks.Create;
  FCustomProperties := TXc12CustomProperties.Create;
  FCellWatches := TCellRefs.Create;
  FIgnoredErrors := TXc12IgnoredErrors.Create;
  FSmartTags := TXc12SmartTags.Create;
  FOleObjects := TXc12OleObjects.Create;
  FControls := TXc12Controls.Create;
  FWebPublishItems := TXc12WebPublishItems.Create;

  SetDefaultValues;
end;

destructor TXc12DataWorksheet.Destroy;
begin
  FCells.Free;

  FColumns.Free;
  FComments.Free;
  FTables.Free;
  FDrawing.Free;
  FVmlDrawing.Free;
  FVmlDrawingRels.Free;
  FHyperlinks.Free;
  FMergedCells.Free;
  FAutofilter.Free;
  FCustomSheetViews.Free;
  FSheetViews.Free;
  FPivotTables.Free;

  FSheetPr.Free;
  FSheetFormatPr.Free;
  FSheetCalcPr.Free;
  FSheetProtection.Free;
  FProtectedRanges.Free;
  FScenarios.Free;
  FSortState.Free;
  FDataConsolidate.Free;
  FPhoneticPr.Free;
  FConditionalFormatting.Free;
  FDataValidations.Free;
  FPrintOptions.Free;
  FPageMargins.Free;
  FPageSetup.Free;
  FHeaderFooter.Free;
  FRowBreaks.Free;
  FColBreaks.Free;
  FCustomProperties.Free;
  FCellWatches.Free;
  FIgnoredErrors.Free;
  FSmartTags.Free;
  FOleObjects.Free;
  FControls.Free;
  FWebPublishItems.Free;
  inherited;
end;

function TXc12DataWorksheet.GetDimension: TXLSCellArea;
begin
  Result := FCells.Dimension;
end;

function TXc12DataWorksheet.GetQuotedName: AxUCString;
var
  i: integer;
begin
  for i := 1 to Length(FName) do begin
    if not CharInSet(FName[i],['a'..'z','A'..'Z','0'..'9']) then begin
      Result := '''' + FName + '''';
      Exit;
    end;
  end;
  Result := FName;
end;

procedure TXc12DataWorksheet.SetDefaultValues;
begin
  FSheetViews.Add;
end;

procedure TXc12DataWorksheet.SetDimension(const Value: TXLSCellArea);
begin
  FCells.Dimension := Value;
end;

{ TXc12DataWorksheets }

function TXc12DataWorksheets.Add(const ASheetId: integer): TXc12DataWorksheet;
begin
  Result := TXc12DataWorksheet.Create(FClassFactory,Count,FSST,FStyles);
  if ASheetId < 0 then
    Result._SheetId := Count + 1
  else
    Result._SheetId := ASheetId;
  inherited Add(Result);
end;

procedure TXc12DataWorksheets.Clear;
begin
  inherited Clear;

  FSST.Clear;
  FStyles.Clear;
end;

constructor TXc12DataWorksheets.Create(AClassFactory: TXLSClassFactory; ASST: TXc12DataSST; AStyles: TXc12DataStyleSheet);
begin
  FClassFactory := AClassFactory;
  FSST := ASST;
  FStyles := AStyles;
  inherited Create;
end;

destructor TXc12DataWorksheets.Destroy;
begin
  inherited;
end;

function TXc12DataWorksheets.Find(const AName: AxUCString): integer;
var
  S: AxUCString;
begin
  S := AnsiUppercase(AName);
  for Result := 0 to Count - 1 do begin
    if S = AnsiUppercase(Items[Result].Name) then
      Exit;
  end;
  Result := -1;
end;

function TXc12DataWorksheets.FindSheetFromCells(const ACells: TXLSCellMMU): TXc12DataWorksheet;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FCells = ACells then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12DataWorksheets.GetItems(Index: integer): TXc12DataWorksheet;
begin
  Result := TXc12DataWorksheet(inherited Items[Index]);
end;

function TXc12DataWorksheets.Insert(const AIndex: integer): TXc12DataWorksheet;
begin
  Result := TXc12DataWorksheet.Create(FClassFactory,Count,FSST,FStyles);
  Result._SheetId := Count + 1;
  inherited Insert(AIndex,Result);
  ReIndex;
end;

{ TXc12Column }

procedure TXc12Column.Assign(ASource: TXc12Column);
begin
  FMin := ASource.FMin;
  FMax := ASource.FMax;
  FWidth := ASource.FWidth;
  FStyle := ASource.FStyle;
  FOutlineLevel := ASource.FOutlineLevel;
  FOptions := ASource.FOptions;
end;

procedure TXc12Column.Clear;
begin
  FMin := 0;
  FMax := 0;
  FWidth := 0;
  FStyle := Nil;
  FOutlineLevel := 0;
  FOptions := [];
end;

constructor TXc12Column.Create;
begin
  Clear;
end;

destructor TXc12Column.Destroy;
begin
  FStyle := Nil;
  inherited;
end;

function TXc12Column.Equal(ACol: TXc12Column): boolean;
begin
  Result := (FMin = ACol.FMin) and
            (FMax = ACol.FMax) and
            (FWidth = ACol.FWidth) and
            (FStyle = ACol.FStyle) and
            (FOutlineLevel = ACol.FOutlineLevel) and
            (FOptions = ACol.FOptions);
end;

function TXc12Column.EqualProps(ACol: TXc12Column): boolean;
begin
  Result := (FWidth = ACol.FWidth) and
            (FStyle = ACol.FStyle) and
            (FOutlineLevel = ACol.FOutlineLevel) and
            (FOptions = ACol.FOptions);
end;

function TXc12Column.Hit(ACol1, ACol2: integer): TXc12ColumnHit;
begin
  if (ACol1 < FMin) and (ACol2 < FMin) then
    Result := xchLess
  else if (ACol1 < FMin) and (ACol2 >= FMin) and (ACol2 <= FMax) then
    Result := xchCol2
  else if (ACol1 = FMin) and (ACol2 = FMax) then
    Result := xchMatch
  else if (ACol1 = FMin) and (ACol2 >= FMin) and (ACol2 <= FMax) then
    Result := xchCol1Match
  else if (ACol1 < FMin) and (ACol2 > FMax) then
    Result := xchTargetInside
  else if (ACol1 >= FMin) and (ACol1 <= FMax) and (ACol2 = FMax) then
    Result := xchCol2Match
  else if (ACol1 >= FMin) and (ACol1 <= FMax) and (ACol2 >= FMin) and (ACol2 <= FMax) then
    Result := xchInside
  else if (ACol1 >= FMin) and (ACol1 <= FMax) and (ACol2 > FMax) then
    Result := xchCol1
  else if (ACol1 > FMax) and (ACol2 > FMax)then
    Result := xchGreater
  else
    raise XLSRWException.Create('Column hit error');
end;

function TXc12Column.Hit(ACol: integer): TXc12ColumnHit;
begin
  if ACol < FMin then
    Result := xchLess
  else if (ACol >= FMin) and (ACol <= FMax) then
    Result := xchInside
  else
    Result := xchGreater;
end;

{ TXc12Columns }

function TXc12Columns.Add(AStyle: TXc12XF): TXc12Column;
begin
  Result := TXc12Column.Create;
  Result.Clear;
  Result.FStyle := AStyle;
  inherited Add(Result);
end;

procedure TXc12Columns.Add(ACol: TXc12Column; AStyle: TXc12XF);
begin
  ACol.FStyle := AStyle;
  inherited Add(ACol);
end;

function TXc12Columns.Add: TXc12Column;
begin
  Result := TXc12Column.Create;
  Result.Clear;
  inherited Add(Result);
end;

constructor TXc12Columns.Create;
begin
  inherited Create;
end;

function TXc12Columns.GetItems(Index: integer): TXc12Column;
begin
  Result := TXc12Column(inherited Items[Index]);
end;

procedure TXc12Columns.NilAndDelete(AIndex: integer);
var
  C: TXc12Column;
begin
  C := Items[AIndex];
  C.Free;
  // Prevent the list from deleting the object.
  List[AIndex] := Nil;
end;

{ TXc12Hyperlink }

procedure TXc12Hyperlink.Assign(AHLink: TXc12Hyperlink);
begin
  FHyperlinkType := AHLink.FHyperlinkType;

  FDisplay := AHLink.FDisplay;
  FAddress := AHLink.FAddress;
  FLocation := AHLink.FLocation;
  FRef := AHLink.FRef;
  FTooltip := AHLink.FTooltip;

  FTargetFrame := AHLink.FTargetFrame;
  FScreenTip := AHLink.FScreenTip;
  FHyperlinkEnc := AHLink.FHyperlinkEnc;
end;

procedure TXc12Hyperlink.Clear;
begin
  FHyperlinkType := xhltUnknown;
  FDisplay := '';
  FAddress := '';
  FLocation:= '';
  ClearCellArea(FRef);
  FTooltip := '';
end;

constructor TXc12Hyperlink.Create;
begin
  Clear;
end;

procedure TXc12Hyperlink.Parse;
begin

end;

procedure TXc12Hyperlink.SetRawAddress(const Value: AxUCString);
begin
  FAddress := Value;
  Parse;
end;

{ TXc12Hyperlinks }

function TXc12Hyperlinks.Add: TXc12Hyperlink;
begin
  Result := CreateMember;
  Result.Clear;
  inherited Add(Result);
end;

procedure TXc12Hyperlinks.Assign(AHLinks: TXc12Hyperlinks);
var
  i: integer;
begin
  Clear;

  for i := 0 to AHLinks.Count - 1 do
    Add.Assign(AHlinks[i]);
end;

constructor TXc12Hyperlinks.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;

  FClassFactory := AClassFactory;
end;

function TXc12Hyperlinks.CreateMember: TXc12Hyperlink;
begin
  Result := TXc12Hyperlink(FClassFactory.CreateAClass(xcftHyperlinksMember,Self));
end;

function TXc12Hyperlinks.GetItems(Index: integer): TXc12Hyperlink;
begin
  Result := TXc12Hyperlink(inherited Items[Index]);
end;

{ TXc12Colors }

procedure TXc12Colors.Add(AColor: TXc12Color);
var
  P: PXc12Color;
begin
  GetMem(P,SizeOf(TXc12Color));
  System.Move(AColor,P^,SizeOf(TXc12Color));
  inherited Add(P);
end;

procedure TXc12Colors.Add(ARGB: longword);
var
  C: TXc12Color;
begin
  C.ColorType := exctRgb;
  C.Tint := 0;
  C.OrigRGB := ARGB;
  Add(C);
end;

procedure TXc12Colors.Assign(AColors: TXc12Colors);
var
  i: integer;
begin
  Clear;

  for i := 0 to AColors.Count - 1 do
    Add(AColors[i]);
end;

destructor TXc12Colors.Destroy;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FreeMem(inherited Items[i]);
end;

function TXc12Colors.GetItems(Index: integer): TXc12Color;
var
  P: PXc12Color;
begin
  P := (inherited Items[Index]);
  System.Move(P^,Result,SizeOf(TXc12Color));
end;

{ TXc12Cfvo }

procedure TXc12Cfvo.Clear;
begin
  FType := x12ctNum;
  FVal := '';
  FGte := True;
end;

procedure TXc12Cfvo.Assign(ACfvo: TXc12Cfvo);
begin
  FType := ACfvo.FType;
  FVal := ACfvo.FVal;
  FGte := ACfvo.FGte;

  if ACfvo.FPtgsSz > 0 then begin
    FPtgsSz := ACfvo.FPtgsSz;
    GetMem(FPtgs,FPtgsSz);
    System.Move(ACfvo.FPtgs^,FPtgs^,FPtgsSz);
  end;
end;

procedure TXc12Cfvo.Clear(AType: TXc12CfvoType; AVal: AxUCString; AGte: boolean);
begin
  FType := AType;
  FVal := AVal;
  FGte := AGte;
end;

constructor TXc12Cfvo.Create;
begin
  Clear;
end;

destructor TXc12Cfvo.Destroy;
begin
  if FPtgs <> Nil then
    FreeMem(FPtgs);

  inherited;
end;

function TXc12Cfvo.GetAsFloat: double;
begin
  Result := XmlStrToFloatDef(FVal,0);
end;

function TXc12Cfvo.GetAsPercent: double;
begin
  Result := XmlStrToFloatDef(FVal,0) / 100;
end;

procedure TXc12Cfvo.SetAsFloat(const Value: double);
begin
  FVal := XmlFloatToStr(Value);
end;

{ TXc12Cfvos }

function TXc12Cfvos.Add(AType: TXc12CfvoType; const AVal: AxUCString; AGte: boolean): TXc12Cfvo;
begin
  Result := TXc12Cfvo.Create;
  Result.Clear(AType,AVal,AGte);
  inherited Add(Result);
end;

procedure TXc12Cfvos.Assign(ACfvos: TXc12Cfvos);
var
  i: integer;
begin
  Clear;

  for i := 0 to ACfvos.Count - 1 do
    Add.Assign(ACfvos[i]);
end;

function TXc12Cfvos.Add: TXc12Cfvo;
begin
  Result := TXc12Cfvo.Create;
  inherited Add(Result);
end;

constructor TXc12Cfvos.Create;
begin
  inherited Create;
end;

function TXc12Cfvos.GetItems(Index: integer): TXc12Cfvo;
begin
  Result := TXc12Cfvo(inherited Items[Index]);
end;

{ TXc12IconSet }

procedure TXc12IconSet.Assign(AIconSet: TXc12IconSet);
begin
  FIconSet := AIconSet.FIconSet;
  FShowValue := AIconSet.FShowValue;
  FPercent := AIconSet.FPercent;
  FReverse := AIconSet.FReverse;
  FCfvos.Assign(AIconSet.FCfvos);
end;

procedure TXc12IconSet.Clear;
begin
  FIconSet := x12ist3TrafficLights1;
  FShowValue := True;
  FPercent := True;
  FReverse := False;
end;

constructor TXc12IconSet.Create;
begin
  FCfvos := TXc12Cfvos.Create;
  Clear;
end;

destructor TXc12IconSet.Destroy;
begin
  FCfvos.Free;
  inherited;
end;

{ TXc12DataBar }

procedure TXc12DataBar.Assign(ADataBar: TXc12DataBar);
begin
  FMinLength := ADataBar.FMinLength;
  FMaxLength := ADataBar.FMaxLength;
  FShowValue := ADataBar.FShowValue;

  FColor := ADataBar.FColor;
  FCfvo1.Assign(ADataBar.FCfvo1);
  FCfvo2.Assign(ADataBar.FCfvo2);
end;

procedure TXc12DataBar.Clear;
begin
  FMinLength := 10;
  FMaxLength := 90;
  FShowValue := True;
  FColor.ColorType := exctUnassigned;
  FCfvo1.Clear;
  FCfvo2.Clear;
end;

constructor TXc12DataBar.Create;
begin
  FCfvo1 := TXc12Cfvo.Create;
  FCfvo2 := TXc12Cfvo.Create;
  Clear;
end;

destructor TXc12DataBar.Destroy;
begin
  FCfvo1.Free;
  FCfvo2.Free;
  inherited;
end;

{ TXc12ColorScale }

procedure TXc12ColorScale.Assign(AColorScale: TXc12ColorScale);
begin
  Clear;

  FCfvos.Assign(AColorScale.FCfvos);
  FColors.Assign(AColorScale.FColors);
end;

procedure TXc12ColorScale.Clear;
begin
  FCfvos.Clear;
  FColors.Clear;
end;

constructor TXc12ColorScale.Create;
begin
  FCfvos := TXc12Cfvos.Create;
  FColors := TXc12Colors.Create;
  Clear;
end;

destructor TXc12ColorScale.Destroy;
begin
  FCfvos.Free;
  FColors.Free;
  inherited;
end;

{ TXc12CfRule }

procedure TXc12CfRule.Assign(ARule: TXc12CfRule);
var
  i: integer;
begin
  FFormula[0] := ARule.FFormula[0];
  FFormula[1] := ARule.FFormula[1];
  FFormula[2] := ARule.FFormula[2];

  for i := 0 to High(FPtgs) do begin
    if ARule.FPtgs[i] <> Nil then begin
      FPtgsSz[i] := ARule.FPtgsSz[i];
      GetMem(FPtgs[i],FPtgsSz[i]);
      System.Move(ARule.FPtgs[i]^,FPtgs[i]^,FPtgsSz[i]);
    end;
  end;

  FType_ := ARule.FType_;

  if ARule.FDXF <> Nil then begin
    if FDXF = Nil then begin
      FDXF := TXc12DXFs(ARule.FDXF.Owner).Add;
      FDXF.Assign(ARule.FDXF);
    end
    else if FDXF.Owner = ARule.FDXF.Owner then
      FDXF := ARule.FDXF
    else begin
      FDXF := TXc12DXFs(ARule.FDXF.Owner).Find(ARule.FDXF);
      if FDXF = Nil then begin
        FDXF := TXc12DXFs(ARule.FDXF.Owner).Add;
        FDXF.Assign(ARule.FDXF);
      end;
    end;
  end;

  FPriority := ARule.FPriority;
  FStopIfTrue := ARule.FStopIfTrue;
  FAboveAverage := ARule.FAboveAverage;
  FPercent := ARule.FPercent;
  FBottom := ARule.FBottom;
  FOperator := ARule.FOperator;
  FText := ARule.FText;
  FTimePeriod := ARule.FTimePeriod;
  FRank := ARule.FRank;
  FStdDev := ARule.FStdDev;
  FEqualAverage := ARule.FEqualAverage;

  FColorScale.Assign(ARule.FColorScale);
  FDataBar.Assign(ARule.FDataBar);
  FIconSet.Assign(ARule.FIconSet);
end;

procedure TXc12CfRule.Clear;
begin
  FType_ := TXc12CfType(XPG_UNKNOWN_ENUM);
  FDXF := Nil;
  FPriority := 0;
  FStopIfTrue := False;
  FAboveAverage := True;
  FPercent := False;
  FBottom := False;
  FOperator := TXc12ConditionalFormattingOperator(XPG_UNKNOWN_ENUM);
  FText := '';
  FTimePeriod := TXc12TimePeriod(XPG_UNKNOWN_ENUM);
  FRank := 0;
  FStdDev := 0;
  FEqualAverage := False;
  FColorScale.Clear;
  FDataBar.Clear;
  FIconSet.Clear;
end;

constructor TXc12CfRule.Create;
begin
  FColorScale := TXc12ColorScale.Create;
  FDataBar := TXc12DataBar.Create;
  FIconSet := TXc12IconSet.Create;

  Clear;
end;

destructor TXc12CfRule.Destroy;
var
  i: integer;
begin
  FColorScale.Free;
  FDataBar.Free;
  FIconSet.Free;
  for i := 0 to High(FPtgs) do begin
    if FPtgs[i] <> Nil then
      FreeMem(FPtgs[i]);
  end;
  inherited;
end;

function TXc12CfRule.FormulaMaxCount: integer;
begin
  Result := Length(FFormula);
end;

function TXc12CfRule.GetFormulas(Index: integer): AxUCString;
begin
  if (Index < 0) or (Index > High(FFormula)) then
    raise XLSRWException.Create('Index out of range');
  Result := FFormula[Index];
end;

function TXc12CfRule.GetPtgs(Index: integer): PXLSPtgs;
begin
  if (Index < 0) or (Index > High(FFormula)) then
    raise XLSRWException.Create('Index out of range');
  Result := FPtgs[Index];
end;

function TXc12CfRule.GetPtgsSz(Index: integer): integer;
begin
  if (Index < 0) or (Index > High(FFormula)) then
    raise XLSRWException.Create('Index out of range');
  Result := FPtgsSz[Index];
end;

function TXc12CfRule.GetValues(Index: integer): double;
begin
  Result := FValues[Index];
end;

function TXc12CfRule.GetValuesCount: integer;
begin
  Result := Length(FValues);
end;

procedure TXc12CfRule.SerFormulas(Index: integer; const Value: AxUCString);
begin
  if (Index < 0) or (Index > High(FFormula)) then
    raise XLSRWException.Create('Index out of range');
  FFormula[Index] := Value;
  if FPtgs[Index] <> Nil then
    FreeMem(FPtgs[Index]);
end;

procedure TXc12CfRule.SetPtgs(Index: integer; const Value: PXLSPtgs);
begin
  if (Index < 0) or (Index > High(FFormula)) then
    raise XLSRWException.Create('Index out of range');
  FPtgs[Index] := Value;
end;

procedure TXc12CfRule.SetPtgsSz(Index: integer; const Value: integer);
begin
  if (Index < 0) or (Index > High(FFormula)) then
    raise XLSRWException.Create('Index out of range');
  FPtgsSz[Index] := Value;
end;

procedure TXc12CfRule.SetValues(Index: integer; const Value: double);
begin
  FValues[Index] := Value;
end;

procedure TXc12CfRule.SetValuesCount(const Value: integer);
begin
  SetLength(FValues,Value);
end;

procedure TXc12CfRule.SortValues;
begin
  SortDoubleArray(FValues);
end;

{ TXc12CfRules }

function TXc12CfRules.Add: TXc12CfRule;
begin
  Result := TXc12CfRule.Create;
  inherited Add(Result);
end;

procedure TXc12CfRules.Assign(ARules: TXc12CfRules);
var
  i: integer;
begin
  Clear;

  for i := 0 to ARules.Count - 1 do
    Add.Assign(ARules[i]);
end;

constructor TXc12CfRules.Create;
begin
  inherited Create;
end;

function TXc12CfRules.GetItems(Index: integer): TXc12CfRule;
begin
  Result := TXc12CfRule(inherited Items[Index]);
end;

{ TXc12ConditionalFormatting }

procedure TXc12ConditionalFormatting.Assign(ACondFmt: TXc12ConditionalFormatting);
begin
  FPivot := ACondFmt.FPivot;
  FSQRef.Assign(ACondFmt.FSQRef);;
  FCfRules.Assign(ACondFmt.FCfRules);
end;

procedure TXc12ConditionalFormatting.Clear;
begin
  FSQRef.Clear;
  FCfRules.Clear;
end;

constructor TXc12ConditionalFormatting.Create;
begin
  FSQRef := TCellAreas.Create;
  FCfRules := TXc12CfRules.Create;
  Clear;
end;

destructor TXc12ConditionalFormatting.Destroy;
begin
  FSQRef.Free;
  FCfRules.Free;
  inherited;
end;

{ TXc12ConditionalFormattings }

function TXc12ConditionalFormattings.Add: TXLSMoveCopyItem;
begin
  Result := CreateMember;
  inherited Add(Result);
end;

function TXc12ConditionalFormattings.AddCF: TXc12ConditionalFormatting;
begin
  Result := TXc12ConditionalFormatting(Add);
end;

procedure TXc12ConditionalFormattings.Assign(ACondFmts: TXc12ConditionalFormattings);
var
  i: integer;
begin
  Clear;

  for i := 0 to ACondFmts.Count - 1 do
    AddCF.Assign(ACondFmts[i]);
end;

constructor TXc12ConditionalFormattings.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;

  FClassFactory := AClassFactory;
end;

function TXc12ConditionalFormattings.GetItems(Index: integer): TXc12ConditionalFormatting;
begin
  Result := TXc12ConditionalFormatting(inherited Items[Index]);
end;

{ TXc12HeaderFooter }

procedure TXc12HeaderFooter.Clear;
begin
  FDifferentOddEven := False;
  FDifferentFirst := False;
  FScaleWithDoc := True;
  FAlignWithMargins := True;
  FOddHeader := '';
  FOddFooter := '';
  FEvenHeader := '';
  FEvenFooter := '';
  FFirstHeader := '';
  FFirstFooter := '';
end;

constructor TXc12HeaderFooter.Create;
begin
  Clear;
end;

{ TXc12PageSetup }

procedure TXc12PageSetup.Clear;
begin
  FPaperSize := 0;
  FScale := 100;
  FFirstPageNumber := 1;
  FFitToWidth := 1;
  FFitToHeight := 1;
  FPageOrder := x12poDownThenOver;
  FOrientation := x12oPortrait;
  FUsePrinterDefaults := True;
  FBlackAndWhite := False;
  FDraft := False;
  FCellComments := x12ccNone;
  FUseFirstPageNumber := False;
  FErrors := x12peDisplayed;
  FHorizontalDpi := 600;
  FVerticalDpi := 600;
  FCopies := 1;
  if FPrinterSettings <> Nil then
    FPrinterSettings.Free;
  FPrinterSettings := Nil;
end;

constructor TXc12PageSetup.Create;
begin
  Clear;
end;

destructor TXc12PageSetup.Destroy;
begin
  if FPrinterSettings <> Nil then
    FPrinterSettings.Free;
  inherited;
end;

procedure TXc12PageSetup.SetPrinterSettings(const Value: TStream);
begin
  if FPrinterSettings <> Nil then
    FPrinterSettings.Free;

  FPrinterSettings := TMemoryStream.Create;
  FPrinterSettings.CopyFrom(Value,Value.Size);
  Value.Free;
end;

{ TXc12PrintOptions }

procedure TXc12PrintOptions.Clear;
begin
  FHorizontalCentered := False;
  FVerticalCentered := False;
  FHeadings := False;
  FGridLines := False;
  FGridLinesSet := True;
end;

constructor TXc12PrintOptions.Create;
begin
  Clear;
end;

{ TXc12PageMargins }

procedure TXc12PageMargins.Clear;
begin
  FLeft   := 0.70;
  FRight  := 0.70;
  FTop    := 0.75;
  FBottom := 0.75;
  FHeader := 0.30;
  FFooter := 0.30;
end;

constructor TXc12PageMargins.Create;
begin
  Clear;
end;

{ TXc12Break }

procedure TXc12Break.Clear;
begin
  FId := 0;
  FMin := 0;
  FMax := 0;
  FMan := False;
  FPt := False;
end;

constructor TXc12Break.Create;
begin
  Clear;
end;

{ TXc12PageBreaks }

function TXc12PageBreaks.Add: TXc12Break;
begin
  Result := TXc12Break.Create;
  inherited Add(Result);
end;

procedure TXc12PageBreaks.Clear;
begin
  inherited Clear;
end;

constructor TXc12PageBreaks.Create;
begin
  inherited Create;
end;

function TXc12PageBreaks.Find(const AId: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Id = AId then
      Exit;
  end;
  Result := -1;
end;

function TXc12PageBreaks.GetItems(Index: integer): TXc12Break;
begin
  Result := TXc12Break(inherited Items[Index]);
end;

function TXc12PageBreaks.Hits(const AId: integer): integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do begin
    if (AId >= Items[i].Min) and (AId <= Items[i].Max) then
      Inc(Result);
  end;
end;

{ TXc12Selection }

procedure TXc12Selection.Clear;
begin
  FPane := x12pTopLeft;
  SetCellArea(FActiveCell,0,0);
  FActiveCellId := 0;
  FSQRef.Clear;
  FSQRef.Add(0,0);
end;

constructor TXc12Selection.Create;
begin
  FSQRef := TCellAreas.Create;
  Clear;
end;

destructor TXc12Selection.Destroy;
begin
  FSQRef.Free;
  inherited;
end;

{ TXc12Selections }

function TXc12Selections.Add: TXc12Selection;
begin
  Result := TXc12Selection.Create;
  inherited Add(Result);
end;

constructor TXc12Selections.Create;
begin
  inherited Create;
end;

function TXc12Selections.GetItems(Index: integer): TXc12Selection;
begin
  Result := TXc12Selection(inherited Items[Index]);
end;

{ TXc12Pane }

procedure TXc12Pane.Clear;
begin
  FXSplit := 0;
  FYSplit := 0;
  ClearCellArea(FTopLeftCell);
  FActivePane := x12pTopLeft;
  FState := x12psSplit;
  FExcel97 := False;
end;

constructor TXc12Pane.Create;
begin
  Clear;
end;

{ TXc12CustomSheetView }

procedure TXc12CustomSheetView.Clear;
begin
{$ifdef D2006PLUS}
  FGUID.Empty;
{$else}
  FillChar(FGUID, Sizeof(TGUID), 0);
{$endif}
  FScale := 100;
  FColorId := 64;
  FShowPageBreaks := False;
  FShowFormulas := False;
  FShowGridLines := True;
  FShowRowCol := True;
  FOutlineSymbols := True;
  FZeroValues := True;
  FFitToPage := False;
  FPrintArea := False;
  FFilter := False;
  FShowAutoFilter := False;
  FHiddenRows := False;
  FHiddenColumns := False;
  FState := x12ssVisible;
  FFilterUnique := False;
  FView := x12svtNormal;
  FShowRuler := True;
  ClearCellArea(FTopLeftCell);

  FPane.Clear;
  FSelection.Clear;
  FRowBreaks.Clear;
  FColBreaks.Clear;
  FPageMargins.Clear;
  FPrintOptions.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FAutoFilter.Clear;
end;

constructor TXc12CustomSheetView.Create(AClassFactory: TXLSClassFactory);
begin
  FClassFactory := AClassFactory;

  FPane := TXc12Pane.Create;
  FSelection := TXc12Selection.Create;
  FRowBreaks := TXc12PageBreaks.Create;
  FColBreaks := TXc12PageBreaks.Create;
  FPageMargins := TXc12PageMargins.Create;
  FPrintOptions := TXc12PrintOptions.Create;
  FPageSetup := TXc12PageSetup.Create;
  FHeaderFooter := TXc12HeaderFooter.Create;
  FAutoFilter := TXc12AutoFilter.Create(FClassFactory);
  Clear;
end;

destructor TXc12CustomSheetView.Destroy;
begin
  FPane.Free;
  FSelection.Free;
  FRowBreaks.Free;
  FColBreaks.Free;
  FPageMargins.Free;
  FPrintOptions.Free;
  FPageSetup.Free;
  FHeaderFooter.Free;
  FAutoFilter.Free;
  inherited;
end;

{ TXc12CustomSheetViews }

function TXc12CustomSheetViews.Add: TXc12CustomSheetView;
begin
  Result := TXc12CustomSheetView.Create(FClassFactory);
  inherited Add(Result);
end;

constructor TXc12CustomSheetViews.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;

  FClassFactory := AClassFactory;
end;

function TXc12CustomSheetViews.GetItems(Index: integer): TXc12CustomSheetView;
begin
  Result := TXc12CustomSheetView(inherited Items[Index]);
end;

{ TXc12PivotSelection }

procedure TXc12PivotSelection.Clear;
begin
  FPane := x12pTopLeft;
  FShowHeader := False;
  FLabel_ := False;
  FData := False;
  FExtendable := False;
  FAxis := x12aUnknown;
  FDimension := 0;
  FStart := 0;
  FMin := 0;
  FMax := 0;
  FActiveRow := 0;
  FActiveCol := 0;
  FPreviousRow := 0;
  FPreviousCol := 0;
  FClick := 0;
  FRId := '';
  FPivotArea.Clear;
end;

constructor TXc12PivotSelection.Create;
begin
  FPivotArea := TXc12PivotArea.Create;
  Clear;
end;

destructor TXc12PivotSelection.Destroy;
begin
  FPivotArea.Free;
  inherited;
end;

{ TXc12PivotArea }

constructor TXc12PivotArea.Create;
begin
  FReferences := TXc12PivotAreaReferences.Create;
  Clear;
end;

destructor TXc12PivotArea.Destroy;
begin
  FReferences.Free;
  inherited;
end;

procedure TXc12PivotArea.Clear;
begin
  FField := 0;
  FType_ := x12patNormal;
  FDataOnly := True;
  FLabelOnly := False;
  FGrandRow := False;
  FGrandCol := False;
  FCacheIndex := False;
  FOutline := True;
  ClearCellArea(FOffset);
  FCollapsedLevelsAreSubtotals := False;
  FAxis := x12aUnknown;
  FFieldPosition := 0;
  FReferences.Clear;
end;

{ TXc12PivotAreaReference }

procedure TXc12PivotAreaReference.Clear;
begin
  FField := 0;
  FSelected := True;
  FByPosition := False;
  FRelative := False;
  FDefaultSubtotal := False;
  FSumSubtotal := False;
  FCountASubtotal := False;
  FAvgSubtotal := False;
  FMaxSubtotal := False;
  FMinSubtotal := False;
  FProductSubtotal := False;
  FCountSubtotal := False;
  FStdDevSubtotal := False;
  FStdDevPSubtotal := False;
  FVarSubtotal := False;
  FVarPSubtotal := False;
  FValues.Clear;
end;

constructor TXc12PivotAreaReference.Create;
begin
  FValues := TIntegerList.Create;
  Clear;
end;

destructor TXc12PivotAreaReference.Destroy;
begin
  FValues.Free;
  inherited;
end;

{ TXc12PivotAreaReferences }

function TXc12PivotAreaReferences.Add: TXc12PivotAreaReference;
begin
  Result := TXc12PivotAreaReference.Create;
  inherited Add(Result);
end;

constructor TXc12PivotAreaReferences.Create;
begin
  inherited Create;
end;

function TXc12PivotAreaReferences.GetItems(Index: integer): TXc12PivotAreaReference;
begin
  Result := TXc12PivotAreaReference(inherited Items[Index]);
end;

{ TXc12PivotSelections }

function TXc12PivotSelections.Add: TXc12PivotSelection;
begin
  Result := TXc12PivotSelection.Create;
  inherited Add(Result);
end;

constructor TXc12PivotSelections.Create;
begin
  inherited Create;
end;

destructor TXc12PivotSelections.Destroy;
begin

  inherited;
end;

function TXc12PivotSelections.GetItems(Index: integer): TXc12PivotSelection;
begin
  Result := TXc12PivotSelection(inherited Items[Index]);
end;

{ TXc12SheetView }

procedure TXc12SheetView.Clear;
begin
  FWindowProtection := False;
  FShowFormulas := False;
  FShowGridLines := True;
  FShowRowColHeaders := True;
  FShowZeros := True;
  FRightToLeft := False;
  FTabSelected := False;
  FShowRuler := True;
  FShowOutlineSymbols := True;
  FDefaultGridColor := True;
  FShowWhiteSpace := True;
  FView := x12svtNormal;
  ClearCellArea(FTopLeftCell);
  FColorId := 64;
  FZoomScale := 100;
  FZoomScaleNormal := 0;
  FZoomScaleSheetLayoutView := 0;
  FZoomScalePageLayoutView := 0;
  FWorkbookViewId := 0;
  FPane.Clear;
  FSelection.Clear;
  FPivotSelection.Clear;
end;

constructor TXc12SheetView.Create;
begin
  FPane := TXc12Pane.Create;
  FSelection := TXc12Selections.Create;
  FPivotSelection := TXc12PivotSelections.Create;
  Clear;
end;

destructor TXc12SheetView.Destroy;
begin
  FPane.Free;
  FSelection.Free;
  FPivotSelection.Free;
  inherited;
end;

{ TXc12SheetViews }

function TXc12SheetViews.Add: TXc12SheetView;
begin
  Result := TXc12SheetView.Create;
  inherited Add(Result);
end;

constructor TXc12SheetViews.Create;
begin
  inherited Create;
end;

function TXc12SheetViews.GetItems(Index: integer): TXc12SheetView;
begin
  Result := TXc12SheetView(inherited Items[Index]);
end;

{ TXc12CellSmartTagPr }

procedure TXc12CellSmartTagPr.Clear;
begin
  FKey := '';
  FVal := '';
end;

constructor TXc12CellSmartTagPr.Create;
begin
  Clear;
end;

{ TXc12CellSmartTag }

function TXc12CellSmartTag.Add: TXc12CellSmartTagPr;
begin
  Result := TXc12CellSmartTagPr.Create;
  inherited Add(Result);
end;

procedure TXc12CellSmartTag.Clear;
begin
  inherited Clear;
  FType_ := 0;
  FDeleted := False;
  FXmlBased := False;
end;

constructor TXc12CellSmartTag.Create;
begin
  inherited Create;
end;

function TXc12CellSmartTag.GetItems(Index: integer): TXc12CellSmartTagPr;
begin
  Result := TXc12CellSmartTagPr(inherited Items[Index]);
end;

{ TXc12CellSmartTags }

function TXc12CellSmartTags.Add: TXc12CellSmartTag;
begin
  Result := TXc12CellSmartTag.Create;
  inherited Add(Result);
end;

procedure TXc12CellSmartTags.Clear;
begin
  inherited Clear;
  FRef.Clear;
end;

constructor TXc12CellSmartTags.Create;
begin
  inherited Create;
  FRef := TCellRef.Create;
end;

destructor TXc12CellSmartTags.Destroy;
begin
  FRef.Free;
  inherited;
end;

function TXc12CellSmartTags.GetItems(Index: integer): TXc12CellSmartTag;
begin
  Result := TXc12CellSmartTag(inherited Items[Index]);
end;

{ TXc12SmartTags }

function TXc12SmartTags.Add: TXc12CellSmartTags;
begin
  Result := TXc12CellSmartTags.Create;
  inherited Add(Result);
end;

constructor TXc12SmartTags.Create;
begin
  inherited Create;
end;

function TXc12SmartTags.GetItems(Index: integer): TXc12CellSmartTags;
begin
  Result := TXc12CellSmartTags(inherited Items[Index]);
end;

{ TXc12WebPublishItems }

function TXc12WebPublishItems.Add: TXc12WebPublishItem;
begin
  Result := TXc12WebPublishItem.Create;
  inherited Add(Result);
end;

constructor TXc12WebPublishItems.Create;
begin
  inherited Create;
end;

function TXc12WebPublishItems.GetItems(Index: integer): TXc12WebPublishItem;
begin
  Result := TXc12WebPublishItem(inherited Items[Index]);
end;

{ TXc12Controls }

function TXc12Controls.Add: TXc12Control;
begin
  Result := TXc12Control.Create;
  inherited Add(Result);
end;

constructor TXc12Controls.Create;
begin
  inherited Create;
end;

function TXc12Controls.GetItems(Index: integer): TXc12Control;
begin
  Result := TXc12Control(inherited Items[Index]);
end;

{ TXc12OleObjects }

function TXc12OleObjects.Add: TXc12OleObject;
begin
  Result := TXc12OleObject.Create;
  inherited Add(Result);
end;

constructor TXc12OleObjects.Create;
begin
  inherited Create;
end;

function TXc12OleObjects.GetItems(Index: integer): TXc12OleObject;
begin
  Result := TXc12OleObject(inherited Items[Index]);
end;

{ TXc12IgnoredErrors }

function TXc12IgnoredErrors.Add: TXc12IgnoredError;
begin
  Result := TXc12IgnoredError.Create;
  inherited Add(Result);
end;

constructor TXc12IgnoredErrors.Create;
begin
  inherited Create;
end;

function TXc12IgnoredErrors.GetItems(Index: integer): TXc12IgnoredError;
begin
  Result := TXc12IgnoredError(inherited Items[Index]);
end;

{ TXc12WebPublishItem }

procedure TXc12WebPublishItem.Clear;
begin
  FId := 0;
  FDivId := '';
  FSourceType := x12wstSheet;
  FSourceRef.Clear;
  FSourceObject := '';
  FDestinationFile := '';
  FTitle := '';
  FAutoRepublish := False;
end;

constructor TXc12WebPublishItem.Create;
begin
  FSourceRef := TCellRef.Create;
  Clear;
end;

destructor TXc12WebPublishItem.Destroy;
begin
  FSourceRef.Free;
  inherited;
end;

{ TXc12Control }

procedure TXc12Control.Clear;
begin
  FShapeId := 0;
  FRId := '';
  FName := '';
end;

constructor TXc12Control.Create;
begin
  Clear;
end;

{ TXc12OleObject }

procedure TXc12OleObject.Clear;
begin
  FProgId := '';
  FDvAspect := x12daDVASPECT_CONTENT;
  FLink := '';
  FOleUpdate := x12ouOLEUPDATE_ALWAYS;
  FAutoLoad := False;
  FShapeId := 0;
  FRId := '';
end;

constructor TXc12OleObject.Create;
begin
  Clear;
end;

{ TXc12IgnoredError }

procedure TXc12IgnoredError.Clear;
begin
  FEvalError := False;
  FTwoDigitTextYear := False;
  FNumberStoredAsText := False;
  FFormula := False;
  FFormulaRange := False;
  FUnlockedFormula := False;
  FEmptyCellReference := False;
  FListDataValidation := False;
  FCalculatedColumn := False;
  FSqref.Clear;
end;

constructor TXc12IgnoredError.Create;
begin
  FSqref := TCellAreas.Create;
  Clear;
end;

destructor TXc12IgnoredError.Destroy;
begin
  FSqref.Free;
  inherited;
end;

{ TXc12CustomProperty }

procedure TXc12CustomProperty.Clear;
begin
  FName := '';
  FRId := '';
end;

constructor TXc12CustomProperty.Create;
begin
  Clear;
end;

{ TXc12DataValidation }

procedure TXc12DataValidation.Assign(AValidation: TXc12DataValidation);
begin
  FType_ := AValidation.FType_;
  FErrorStyle := AValidation.FErrorStyle;
  FImeMode := AValidation.FImeMode;
  FOperator_ := AValidation.FOperator_;
  FAllowBlank := AValidation.FAllowBlank;
  FShowDropDown := AValidation.FShowDropDown;
  FShowInputMessage := AValidation.FShowInputMessage;
  FShowErrorMessage := AValidation.FShowErrorMessage;
  FErrorTitle := AValidation.FErrorTitle;
  FError := AValidation.FError;
  FPromptTitle := AValidation.FPromptTitle;
  FPrompt := AValidation.FPrompt;
  FSqref.Assign(AValidation.FSqref);
  FFormula1 := AValidation.FFormula1;
  FFormula2 := AValidation.FFormula2;
end;

procedure TXc12DataValidation.Clear;
begin
  FType_ := x12dvtNone;
  FErrorStyle := x12dvesStop;
  FImeMode := x12dvimNoControl;
  FOperator_ := x12dvoBetween;
  FAllowBlank := False;
  FShowDropDown := False;
  FShowInputMessage := False;
  FShowErrorMessage := False;
  FErrorTitle := '';
  FError := '';
  FPromptTitle := '';
  FPrompt := '';
  FSqref.Clear;
  FFormula1 := '';
  FFormula2 := '';
end;

constructor TXc12DataValidation.Create;
begin
  FSqref := TCellAreas.Create;
  Clear;
end;

destructor TXc12DataValidation.Destroy;
begin
  FSqref.Free;
  inherited;
end;

{ TXc12DataValidations }

function TXc12DataValidations.Add: TXLSMoveCopyItem;
begin
  Result := CreateMember;
  inherited Add(Result);
end;

function TXc12DataValidations.AddDV: TXc12DataValidation;
begin
  Result := CreateMember;
  inherited Add(Result);
end;

procedure TXc12DataValidations.Assign(AValidations: TXc12DataValidations);
var
  i: integer;
begin
  Clear;

  for i := 0 to AValidations.Count - 1 do
    AddDV.Assign(AValidations[i]);
end;

procedure TXc12DataValidations.Clear;
begin
  inherited Clear;
  FDisablePrompts := False;
  FXWindow := 0;
  FYWindow := 0;
end;

constructor TXc12DataValidations.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;

  FClassFactory := AClassFactory;
end;

function TXc12DataValidations.CreateMember: TXc12DataValidation;
begin
  Result := TXc12DataValidation(FClassFactory.CreateAClass(xcftDataValidationsMember));
end;

function TXc12DataValidations.GetItems(Index: integer): TXc12DataValidation;
begin
  Result := TXc12DataValidation(inherited Items[Index]);
end;

{ TXc12DataRef }

procedure TXc12DataRef.Clear;
begin
  ClearCellArea(FRef);
  FName := '';
  FSheet := '';
  FRId := '';
end;

constructor TXc12DataRef.Create;
begin
  Clear;
end;

{ TXc12DataRefs }

function TXc12DataRefs.Add: TXc12DataRef;
begin
  Result := TXc12DataRef.Create;
  inherited Add(Result);
end;

constructor TXc12DataRefs.Create;
begin
  inherited Create;
end;

function TXc12DataRefs.GetItems(Index: integer): TXc12DataRef;
begin
  Result := TXc12DataRef(inherited Items[Index]);
end;

{ TXc12InputCell }

procedure TXc12InputCell.Clear;
begin
  FRow := 0;
  FCol := 0;
  FDeleted := False;
  FUndone := False;
  FVal := '';
  FNumFmtId := 0;
end;

constructor TXc12InputCell.Create;
begin
  Clear;
end;

{ Tx12InputCells }

function Tx12InputCells.Add: TXc12InputCell;
begin
  Result := TXc12InputCell.Create;
  inherited Add(Result);
end;

function Tx12InputCells.Add(ACellRef: AxUCString): TXc12InputCell;
begin
  Result := Add;
  RefStrToColRow(ACellRef,Result.FCol,Result.FRow);
end;

constructor Tx12InputCells.Create;
begin
  inherited Create;
end;

function Tx12InputCells.GetItems(Index: integer): TXc12InputCell;
begin
  Result := TXc12InputCell(inherited Items[Index]);
end;

{ TXc12Scenario }

procedure TXc12Scenario.Clear;
begin
  FName := '';
  FLocked := False;
  FHidden := False;
  FUser := '';
  FComment := '';
  FInputCells.Clear;
end;

constructor TXc12Scenario.Create;
begin
  FInputCells := Tx12InputCells.Create;
  Clear;
end;

destructor TXc12Scenario.Destroy;
begin
  FInputCells.Free;
  inherited;
end;

{ TXc12ProtectedRange }

procedure TXc12ProtectedRange.Clear;
begin
  FPassword := 0;
  FSqref.Clear;
  FName := '';
  FSecurityDescriptor := '';
end;

constructor TXc12ProtectedRange.Create;
begin
  FSqref := TCellAreas.Create;
  Clear;
end;

destructor TXc12ProtectedRange.Destroy;
begin
  FSqref.Free;
  inherited;
end;

{ TXc12CustomProperties }

function TXc12CustomProperties.Add: TXc12CustomProperty;
begin
  Result := TXc12CustomProperty.Create;
  inherited Add(Result);
end;

constructor TXc12CustomProperties.Create;
begin
  inherited Create;
end;

function TXc12CustomProperties.GetItems(Index: integer): TXc12CustomProperty;
begin
  Result := TXc12CustomProperty(inherited Items[Index]);
end;

{ TXc12ProtectedRanges }

function TXc12ProtectedRanges.Add: TXc12ProtectedRange;
begin
  Result := TXc12ProtectedRange.Create;
  inherited Add(Result);
end;

constructor TXc12ProtectedRanges.Create;
begin
  inherited Create;
end;

function TXc12ProtectedRanges.GetItems(Index: integer): TXc12ProtectedRange;
begin
  Result := TXc12ProtectedRange(inherited Items[Index]);
end;

{ TXc12Scenarios }

function TXc12Scenarios.Add: TXc12Scenario;
begin
  Result := TXc12Scenario.Create;
  inherited Add(Result);
end;

constructor TXc12Scenarios.Create;
begin
  inherited Create;
  FSqref := TCellAreas.Create;
end;

destructor TXc12Scenarios.Destroy;
begin
  FSqref.Free;
  inherited;
end;

function TXc12Scenarios.GetItems(Index: integer): TXc12Scenario;
begin
  Result := TXc12Scenario(inherited Items[Index]);
end;

procedure TXc12Scenarios.Clear;
begin
  inherited Clear;
  FCurrent := 0;
  FShow := 0;
//{$ifdef _AXOLOT_DEBUG}
//{$message warn 'uncomment this causes an XLSRWException in FastMM when running in debug mode'}
//{$endif}
//  FSqref.Clear;
end;

{ TXc12OutlinePr }

procedure TXc12OutlinePr.Clear;
begin
  FApplyStyles := False;
  FSummaryBelow := True;
  FSummaryRight := True;
  FShowOutlineSymbols := True;
end;

constructor TXc12OutlinePr.Create;
begin
  Clear;
end;

{ TXc12PageSetupPr }

procedure TXc12PageSetupPr.Clear;
begin
  FAutoPageBreaks := True;
  FFitToPage := False;
end;

constructor TXc12PageSetupPr.Create;
begin
  Clear;
end;

{ TXc12SheetPr }

procedure TXc12SheetPr.Clear;
begin
  FSyncHorizontal := False;
  FSyncVertical := False;
  ClearCellArea(FSyncRef);
  FTransitionEvaluation := False;
  FTransitionEntry := False;
  FPublished_ := True;
  FCodeName := '';
  FFilterMode := False;
  FEnableFormatConditionsCalculation := True;
  FTabColor.ColorType := exctUnassigned;
  FOutlinePr.Clear;
  FPageSetupPr.Clear;
end;

constructor TXc12SheetPr.Create;
begin
  FOutlinePr := TXc12OutlinePr.Create;
  FPageSetupPr := TXc12PageSetupPr.Create;
  Clear;
end;

destructor TXc12SheetPr.Destroy;
begin
  FOutlinePr.Free;
  FPageSetupPr.Free;
  inherited;
end;

{ TXc12SheetFormatPr }

procedure TXc12SheetFormatPr.Assign(ASrc: TXc12SheetFormatPr);
begin
  FBaseColWidth := ASrc.FBaseColWidth;
  FDefaultColWidth := ASrc.FDefaultColWidth;
  FDefaultRowHeight := ASrc.FDefaultRowHeight;
  FCustomHeight := ASrc.FCustomHeight;
  FZeroHeight := ASrc.FZeroHeight;
  FThickTop := ASrc.FThickTop;
  FThickBottom := ASrc.FThickBottom;
  FOutlineLevelRow := ASrc.FOutlineLevelRow;
  FOutlineLevelCol := ASrc.FOutlineLevelCol;
end;

procedure TXc12SheetFormatPr.Clear;
begin
  FBaseColWidth := 8;
  FDefaultColWidth := 0;
  FDefaultRowHeight := 0;
  FCustomHeight := False;
  FZeroHeight := False;
  FThickTop := False;
  FThickBottom := False;
  FOutlineLevelRow := 0;
  FOutlineLevelCol := 0;
end;

constructor TXc12SheetFormatPr.Create;
begin
  Clear;
end;

{ TXc12SheetCalcPr }

procedure TXc12SheetCalcPr.Clear;
begin
  FFullCalcOnLoad := False;
end;

constructor TXc12SheetCalcPr.Create;
begin
  Clear;
end;

{ TXc12SheetProtection }

procedure TXc12SheetProtection.Clear;
begin
  FPassword := 0;
  FSheet := False;
  FObjects := False;
  FScenarios := False;
  FFormatCells := True;
  FFormatColumns := True;
  FFormatRows := True;
  FInsertColumns := True;
  FInsertRows := True;
  FInsertHyperlinks := True;
  FDeleteColumns := True;
  FDeleteRows := True;
  FSelectLockedCells := False;
  FSort := True;
  FAutoFilter := True;
  FPivotTables := True;
  FSelectUnlockedCells := False;
end;

constructor TXc12SheetProtection.Create;
begin
  Clear;
end;

function TXc12SheetProtection.GetPasswordAsString: AxUCString;
begin
  if FPassword <> 0 then
    Result := PasswordFromHash(FPassword)
  else
    Result := '';
end;

procedure TXc12SheetProtection.SetPasswordAsString(const Value: AxUCString);
begin
  FPassword := MakePasswordHash(AnsiString(Value));
end;

{ TXc12DataConsolidate }

procedure TXc12DataConsolidate.Clear;
begin
  FFunction := x12dcfSum;
  FLeftLabels := False;
  FTopLabels := False;
  FLink := False;
  FDataRefs.Clear;
end;

constructor TXc12DataConsolidate.Create;
begin
  FDataRefs := TXc12DataRefs.Create;
  Clear;
end;

destructor TXc12DataConsolidate.Destroy;
begin
  FDataRefs.Free;
  inherited;
end;

{ TXc12PhoneticPr }

procedure TXc12PhoneticPr.Clear;
begin
  FFontId := 0;
  FType := x12ptFullwidthKatakana;
  FAlignment := x12paLeft;
end;

constructor TXc12PhoneticPr.Create;
begin
  Clear;
end;

{ TXc12MergedCells }

function TXc12MergedCells.Add(const ARef: TXLSCellArea): TXc12MergedCell;
begin
  Result := CreateMember;
  Add(Result);
  Result.Ref := ARef;
end;

constructor TXc12MergedCells.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;
  FClassFactory := AClassFactory;
end;

function TXc12MergedCells.CreateMember: TXc12MergedCell;
begin
  Result := TXc12MergedCell(FClassFactory.CreateAClass(xcftMergedCellsMember));
end;

function TXc12MergedCells.CreateObject: TCellArea;
begin
  Result := TXc12MergedCell.Create;
end;

function TXc12MergedCells.GetItems(Index: integer): TXc12MergedCell;
begin
  Result := TXc12MergedCell(inherited Items[Index]);
end;

{ TXc12MergedCell }

procedure TXc12MergedCell.Assign(const AItem: TXc12MergedCell);
begin
  FRef := AItem.FRef;
end;

function TXc12MergedCell.GetCol1: integer;
begin
  Result := FRef.Col1;
end;

function TXc12MergedCell.GetCol2: integer;
begin
  Result := FRef.Col2;
end;

function TXc12MergedCell.GetRow1: integer;
begin
  Result := FRef.Row1;
end;

function TXc12MergedCell.GetRow2: integer;
begin
  Result := FRef.Row2;
end;

procedure TXc12MergedCell.SetCol1(const AValue: integer);
begin
  PXLSCellArea(@FRef).Col1 := AValue;
end;

procedure TXc12MergedCell.SetCol2(const AValue: integer);
begin
  PXLSCellArea(@FRef).Col2 := AValue;
end;

procedure TXc12MergedCell.SetRow1(const AValue: integer);
begin
  PXLSCellArea(@FRef).Row1 := AValue;
end;

procedure TXc12MergedCell.SetRow2(const AValue: integer);
begin
  PXLSCellArea(@FRef).Row2 := AValue;
end;

end.
