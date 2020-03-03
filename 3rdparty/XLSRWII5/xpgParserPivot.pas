unit xpgParserPivot;

// Copyright (c) 2010,2016 Axolot Data
// Web : http://www.axolot.com/xpg
// Mail: xpg@axolot.com
//
// X   X  PPP    GGG
//  X X   P  P  G
//   X    PPP   G  GG
//  X X   P     G   G
// X   X  P      GGG
//
// File generated with Axolot XPG, Xml Parser Generator.
// Version 0.00.90.
// File created on 2016-03-26 14:20:19

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, Math, xpgPUtils, xpgPLists, xpgPXMLUtils,
     xpgPXML,
     Xc12Utils5,
     XLSUtils5, XLSSharedItems5, XLSRelCells5;

const XPG_UNKNOWN_ENUM = $0000FFFF;

type TST_Axis =  (staAxisRow,staAxisCol,staAxisPage,staAxisValues);
const StrTST_Axis: array[0..3] of AxUCString = ('axisRow','axisCol','axisPage','axisValues');
type TST_PivotAreaType =  (stpatNone,stpatNormal,stpatData,stpatAll,stpatOrigin,stpatButton,stpatTopRight);
const StrTST_PivotAreaType: array[0..6] of AxUCString = ('none','normal','data','all','origin','button','topRight');
type TST_Type =  (sttNone,sttAll,sttRow,sttColumn);
const StrTST_Type: array[0..3] of AxUCString = ('none','all','row','column');
type TST_FieldSortType =  (stfstManual,stfstAscending,stfstDescending);
const StrTST_FieldSortType: array[0..2] of AxUCString = ('manual','ascending','descending');
type TST_GroupBy =  (stgbRange,stgbSeconds,stgbMinutes,stgbHours,stgbDays,stgbMonths,stgbQuarters,stgbYears);
const StrTST_GroupBy: array[0..7] of AxUCString = ('range','seconds','minutes','hours','days','months','quarters','years');
type TST_SortType =  (ststNone,ststAscending,ststDescending,ststAscendingAlpha,ststDescendingAlpha,ststAscendingNatural,ststDescendingNatural);
const StrTST_SortType: array[0..6] of AxUCString = ('none','ascending','descending','ascendingAlpha','descendingAlpha','ascendingNatural','descendingNatural');
type TST_ShowDataAs =  (stsdaNormal,stsdaDifference,stsdaPercent,stsdaPercentDiff,stsdaRunTotal,stsdaPercentOfRow,stsdaPercentOfCol,stsdaPercentOfTotal,stsdaIndex);
const StrTST_ShowDataAs: array[0..8] of AxUCString = ('normal','difference','percent','percentDiff','runTotal','percentOfRow','percentOfCol','percentOfTotal','index');
type TST_ItemType =  (stitData,stitDefault,stitSum,stitCountA,stitAvg,stitMax,stitMin,stitProduct,stitCount,stitStdDev,stitStdDevP,stitVar,stitVarP,stitGrand,stitBlank);
const StrTST_ItemType: array[0..14] of AxUCString = ('data','default','sum','countA','avg','max','min','product','count','stdDev','stdDevP','var','varP','grand','blank');
type TST_Scope =  (stsSelection,stsData,stsField);
const StrTST_Scope: array[0..2] of AxUCString = ('selection','data','field');
type TST_PivotFilterType =  (stpftUnknown,stpftCount,stpftPercent,stpftSum,stpftCaptionEqual,stpftCaptionNotEqual,stpftCaptionBeginsWith,stpftCaptionNotBeginsWith,stpftCaptionEndsWith,stpftCaptionNotEndsWith
,stpftCaptionContains,stpftCaptionNotContains,stpftCaptionGreaterThan,stpftCaptionGreaterThanOrEqual,stpftCaptionLessThan,stpftCaptionLessThanOrEqual,stpftCaptionBetween,stpftCaptionNotBetween,stpftValueEqual,
stpftValueNotEqual,stpftValueGreaterThan,stpftValueGreaterThanOrEqual,stpftValueLessThan,stpftValueLessThanOrEqual,stpftValueBetween,stpftValueNotBetween,stpftDateEqual,stpftDateNotEqual,stpftDateOlderThan,
stpftDateOlderThanOrEqual,stpftDateNewerThan,stpftDateNewerThanOrEqual,stpftDateBetween,stpftDateNotBetween,stpftTomorrow,stpftToday,stpftYesterday,stpftNextWeek,stpftThisWeek,stpftLastWeek,stpftNextMonth,
stpftThisMonth,stpftLastMonth,stpftNextQuarter,stpftThisQuarter,stpftLastQuarter,stpftNextYear,stpftThisYear,stpftLastYear,stpftYearToDate,stpftQ1,stpftQ2,stpftQ3,stpftQ4,stpftM1,stpftM2,stpftM3,stpftM4,stpftM5,
stpftM6,stpftM7,stpftM8,stpftM9,stpftM10,stpftM11,stpftM12);
const StrTST_PivotFilterType: array[0..65] of AxUCString = ('unknown','count','percent','sum','captionEqual','captionNotEqual','captionBeginsWith','captionNotBeginsWith','captionEndsWith','captionNotEndsWith',
'captionContains','captionNotContains','captionGreaterThan','captionGreaterThanOrEqual','captionLessThan','captionLessThanOrEqual','captionBetween','captionNotBetween','valueEqual','valueNotEqual',
'valueGreaterThan','valueGreaterThanOrEqual','valueLessThan','valueLessThanOrEqual','valueBetween','valueNotBetween','dateEqual','dateNotEqual','dateOlderThan','dateOlderThanOrEqual',
'dateNewerThan','dateNewerThanOrEqual','dateBetween','dateNotBetween','tomorrow','today','yesterday','nextWeek','thisWeek','lastWeek','nextMonth','thisMonth','lastMonth','nextQuarter','thisQuarter',
'lastQuarter','nextYear','thisYear','lastYear','yearToDate','Q1','Q2','Q3','Q4','M1','M2','M3','M4','M5','M6','M7','M8','M9','M10','M11','M12');
type TST_SourceType =  (ststWorksheet,ststExternal,ststConsolidation,ststScenario);
const StrTST_SourceType: array[0..3] of AxUCString = ('worksheet','external','consolidation','scenario');
type TST_FormatAction =  (stfaBlank,stfaFormatting,stfaDrill,stfaFormula);
const StrTST_FormatAction: array[0..3] of AxUCString = ('blank','formatting','drill','formula');

type TXPGDocBase = class(TObject)
protected
     FErrors: TXpgPErrors;
public
     property Errors: TXpgPErrors read FErrors;
     end;

     TXPGBase = class(TObject)
protected
     FOwner: TXPGDocBase;
     FElementCount: integer;
     FAttributeCount: integer;
     FAssigneds: TXpgAssigneds;
     function  CheckAssigned: integer; virtual;
     function  Assigned: boolean;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; virtual;
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); virtual;
     procedure AfterTag; virtual;

     class function  StrToEnum(AValue: AxUCString): integer;
     class function  StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
     class function  TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
public
     function  Available: boolean;

     property ElementCount: integer read FElementCount write FElementCount;
     property AttributeCount: integer read FAttributeCount write FAttributeCount;
     property Assigneds: TXpgAssigneds read FAssigneds write FAssigneds;
     end;

     TXPGAnyElement = class(TXPGBase)
protected
     FElementName: AxUCString;
     FContent: AxUCString;
     FAttributes: TXpgXMLAttributeList;
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create;
     destructor Destroy; override;

     property ElementName: AxUCString read FElementName write FElementName;
     property Content: AxUCString read FContent write FContent;
     property Attributes: TXpgXMLAttributeList read FAttributes;
     end;

     TXPGAnyElements = class(TObjectList)
protected
     function  GetItems(Index: integer): TXPGAnyElement;
public
     function  Add(AElementName: AxUCString; AContent: AxUCString): TXPGAnyElement;
     procedure Write(AWriter: TXpgWriteXML);

     property Items[Index: integer]: TXPGAnyElement read GetItems;
     end;

     TXPGBaseObjectList = class(TObjectList)
protected
     FOwner: TXPGDocBase;
     FAssigned: boolean;
     function  GetItems(Index: integer): TXPGBase;
public
     constructor Create(AOwner: TXPGDocBase);
     property Items[Index: integer]: TXPGBase read GetItems;
     end;

     TXPGReader = class(TXpgReadXML)
protected
     FCurrent: TXPGBase;
     FStack: TObjectStack;
     FErrors: TXpgPErrors;
public
     constructor Create(AErrors: TXpgPErrors; ARoot: TXPGBase);
     destructor Destroy; override;

     procedure BeginTag; override;
     procedure EndTag; override;
     end;

     TCT_Extension = class(TXPGBase)
protected
     FUri: AxUCString;
     FAnyElements: TXPGAnyElements;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Extension);
     procedure CopyTo(AItem: TCT_Extension);

     property Uri: AxUCString read FUri write FUri;
     property AnyElements: TXPGAnyElements read FAnyElements;
     end;

     TEG_ExtensionList = class(TXPGBase)
protected
     FExt: TCT_Extension;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TEG_ExtensionList);
     procedure CopyTo(AItem: TEG_ExtensionList);
     function  Create_Ext: TCT_Extension;

     property Ext: TCT_Extension read FExt;
     end;

     TCT_GroupMember = class(TXPGBase)
protected
     FUniqueName: AxUCString;
     FGroup: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupMember);
     procedure CopyTo(AItem: TCT_GroupMember);

     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Group: boolean read FGroup write FGroup;
     end;

     TCT_GroupMemberXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GroupMember;
public
     function  Add: TCT_GroupMember;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GroupMemberXpgList);
     procedure CopyTo(AItem: TCT_GroupMemberXpgList);
     property Items[Index: integer]: TCT_GroupMember read GetItems; default;
     end;

     TCT_Index = class(TXPGBase)
protected
     FV: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Index);
     procedure CopyTo(AItem: TCT_Index);

     property V: integer read FV write FV;
     end;

     TCT_IndexXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Index;
public
     function  Add: TCT_Index;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_IndexXpgList);
     procedure CopyTo(AItem: TCT_IndexXpgList);
     property Items[Index: integer]: TCT_Index read GetItems; default;
     end;

     TCT_ExtensionList = class(TXPGBase)
protected
     FEG_ExtensionList: TEG_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ExtensionList);
     procedure CopyTo(AItem: TCT_ExtensionList);

     property EG_ExtensionList: TEG_ExtensionList read FEG_ExtensionList;
     end;

     TCT_Tuple = class(TXPGBase)
protected
     FFld: integer;
     FHier: integer;
     FItem: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Tuple);
     procedure CopyTo(AItem: TCT_Tuple);

     property Fld: integer read FFld write FFld;
     property Hier: integer read FHier write FHier;
     property Item: integer read FItem write FItem;
     end;

     TCT_TupleXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Tuple;
public
     function  Add: TCT_Tuple;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TupleXpgList);
     procedure CopyTo(AItem: TCT_TupleXpgList);
     property Items[Index: integer]: TCT_Tuple read GetItems; default;
     end;

     TCT_GroupMembers = class(TXPGBase)
protected
     FCount: integer;
     FGroupMemberXpgList: TCT_GroupMemberXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupMembers);
     procedure CopyTo(AItem: TCT_GroupMembers);
     function  Create_GroupMemberXpgList: TCT_GroupMemberXpgList;

     property Count: integer read FCount write FCount;
     property GroupMemberXpgList: TCT_GroupMemberXpgList read FGroupMemberXpgList;
     end;

     TCT_PivotAreaReference = class(TXPGBase)
protected
     FField: AxUCString;
     FCount: integer;
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
     FXXpgList: TCT_IndexXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotAreaReference);
     procedure CopyTo(AItem: TCT_PivotAreaReference);
     function  Create_XXpgList: TCT_IndexXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     // Was integer but changed ti string as $FFFFFFFE (-2) is a popular value and it can't be converted to in t with StrToInt.
     property Field: AxUCString read FField write FField;
     property Count: integer read FCount write FCount;
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
     property XXpgList: TCT_IndexXpgList read FXXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PivotAreaReferenceXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotAreaReference;
public
     function  Add: TCT_PivotAreaReference;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotAreaReferenceXpgList);
     procedure CopyTo(AItem: TCT_PivotAreaReferenceXpgList);
     property Items[Index: integer]: TCT_PivotAreaReference read GetItems; default;
     end;

     TCT_Tuples = class(TXPGBase)
protected
     FC: integer;
     FTplXpgList: TCT_TupleXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Tuples);
     procedure CopyTo(AItem: TCT_Tuples);
     function  Create_TplXpgList: TCT_TupleXpgList;

     property C: integer read FC write FC;
     property TplXpgList: TCT_TupleXpgList read FTplXpgList;
     end;

     TCT_TuplesXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Tuples;
public
     function  Add: TCT_Tuples;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_TuplesXpgList);
     procedure CopyTo(AItem: TCT_TuplesXpgList);
     property Items[Index: integer]: TCT_Tuples read GetItems; default;
     end;

     TCT_X = class(TXPGBase)
protected
     FV: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_X);
     procedure CopyTo(AItem: TCT_X);

     property V: integer read FV write FV;
     end;

     TCT_XXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_X;
public
     function  Add: TCT_X;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_XXpgList);
     procedure CopyTo(AItem: TCT_XXpgList);
     property Items[Index: integer]: TCT_X read GetItems; default;
     end;

     TCT_LevelGroup = class(TXPGBase)
protected
     FName: AxUCString;
     FUniqueName: AxUCString;
     FCaption: AxUCString;
     FUniqueParent: AxUCString;
     FId: integer;
     FGroupMembers: TCT_GroupMembers;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_LevelGroup);
     procedure CopyTo(AItem: TCT_LevelGroup);
     function  Create_GroupMembers: TCT_GroupMembers;

     property Name: AxUCString read FName write FName;
     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Caption: AxUCString read FCaption write FCaption;
     property UniqueParent: AxUCString read FUniqueParent write FUniqueParent;
     property Id: integer read FId write FId;
     property GroupMembers: TCT_GroupMembers read FGroupMembers;
     end;

     TCT_LevelGroupXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_LevelGroup;
public
     function  Add: TCT_LevelGroup;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_LevelGroupXpgList);
     procedure CopyTo(AItem: TCT_LevelGroupXpgList);
     property Items[Index: integer]: TCT_LevelGroup read GetItems; default;
     end;

     TCT_PivotAreaReferences = class(TXPGBase)
protected
     FCount: integer;
     FReferenceXpgList: TCT_PivotAreaReferenceXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotAreaReferences);
     procedure CopyTo(AItem: TCT_PivotAreaReferences);
     function  Create_ReferenceXpgList: TCT_PivotAreaReferenceXpgList;

     property Count: integer read FCount write FCount;
     property ReferenceXpgList: TCT_PivotAreaReferenceXpgList read FReferenceXpgList;
     end;

     TCT_PageItem = class(TXPGBase)
protected
     FName: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PageItem);
     procedure CopyTo(AItem: TCT_PageItem);

     property Name: AxUCString read FName write FName;
     end;

     TCT_PageItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PageItem;
public
     function  Add: TCT_PageItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PageItemXpgList);
     procedure CopyTo(AItem: TCT_PageItemXpgList);
     property Items[Index: integer]: TCT_PageItem read GetItems; default;
     end;

     TCT_Missing = class(TXPGBase)
protected
     FU: boolean;
     FF: boolean;
     FC: AxUCString;
     FCp: integer;
     FIn: integer;
     FBc: integer;
     FFc: integer;
     FI: boolean;
     FUn: boolean;
     FSt: boolean;
     FB: boolean;
     FTplsXpgList: TCT_TuplesXpgList;
     FXXpgList: TCT_XXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Missing);
     procedure CopyTo(AItem: TCT_Missing);
     function  Create_TplsXpgList: TCT_TuplesXpgList;
     function  Create_XXpgList: TCT_XXpgList;

     property U: boolean read FU write FU;
     property F: boolean read FF write FF;
     property C: AxUCString read FC write FC;
     property Cp: integer read FCp write FCp;
     property In_: integer read FIn write FIn;
     property Bc: integer read FBc write FBc;
     property Fc_: integer read FFc write FFc;
     property I: boolean read FI write FI;
     property Un: boolean read FUn write FUn;
     property St: boolean read FSt write FSt;
     property B: boolean read FB write FB;
     property TplsXpgList: TCT_TuplesXpgList read FTplsXpgList;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_MissingXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Missing;
public
     function  Add: TCT_Missing;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_MissingXpgList);
     procedure CopyTo(AItem: TCT_MissingXpgList);
     property Items[Index: integer]: TCT_Missing read GetItems; default;
     end;

     TCT_Number = class(TXPGBase)
protected
     FV: double;
     FU: boolean;
     FF: boolean;
     FC: AxUCString;
     FCp: integer;
     FIn: integer;
     FBc: integer;
     FFc: integer;
     FI: boolean;
     FUn: boolean;
     FSt: boolean;
     FB: boolean;
     FTplsXpgList: TCT_TuplesXpgList;
     FXXpgList: TCT_XXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Number);
     procedure CopyTo(AItem: TCT_Number);
     function  Create_TplsXpgList: TCT_TuplesXpgList;
     function  Create_XXpgList: TCT_XXpgList;

     property V: double read FV write FV;
     property U: boolean read FU write FU;
     property F: boolean read FF write FF;
     property C: AxUCString read FC write FC;
     property Cp: integer read FCp write FCp;
     property In_: integer read FIn write FIn;
     property Bc: integer read FBc write FBc;
     property Fc_: integer read FFc write FFc;
     property I: boolean read FI write FI;
     property Un: boolean read FUn write FUn;
     property St: boolean read FSt write FSt;
     property B: boolean read FB write FB;
     property TplsXpgList: TCT_TuplesXpgList read FTplsXpgList;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_NumberXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Number;
public
     function  Add: TCT_Number;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_NumberXpgList);
     procedure CopyTo(AItem: TCT_NumberXpgList);
     property Items[Index: integer]: TCT_Number read GetItems; default;
     end;

     TCT_Boolean = class(TXPGBase)
protected
     FV: boolean;
     FU: boolean;
     FF: boolean;
     FC: AxUCString;
     FCp: integer;
     FXXpgList: TCT_XXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Boolean);
     procedure CopyTo(AItem: TCT_Boolean);
     function  Create_XXpgList: TCT_XXpgList;

     property V: boolean read FV write FV;
     property U: boolean read FU write FU;
     property F: boolean read FF write FF;
     property C: AxUCString read FC write FC;
     property Cp: integer read FCp write FCp;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_BooleanXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Boolean;
public
     function  Add: TCT_Boolean;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_BooleanXpgList);
     procedure CopyTo(AItem: TCT_BooleanXpgList);
     property Items[Index: integer]: TCT_Boolean read GetItems; default;
     end;

     TCT_Error = class(TXPGBase)
protected
     FV: AxUCString;
     FU: boolean;
     FF: boolean;
     FC: AxUCString;
     FCp: integer;
     FIn: integer;
     FBc: integer;
     FFc: integer;
     FI: boolean;
     FUn: boolean;
     FSt: boolean;
     FB: boolean;
     FTpls: TCT_Tuples;
     FXXpgList: TCT_XXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Error);
     procedure CopyTo(AItem: TCT_Error);
     function  Create_Tpls: TCT_Tuples;
     function  Create_XXpgList: TCT_XXpgList;

     property V: AxUCString read FV write FV;
     property U: boolean read FU write FU;
     property F: boolean read FF write FF;
     property C: AxUCString read FC write FC;
     property Cp: integer read FCp write FCp;
     property In_: integer read FIn write FIn;
     property Bc: integer read FBc write FBc;
     property Fc_: integer read FFc write FFc;
     property I: boolean read FI write FI;
     property Un: boolean read FUn write FUn;
     property St: boolean read FSt write FSt;
     property B: boolean read FB write FB;
     property Tpls: TCT_Tuples read FTpls;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_ErrorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Error;
public
     function  Add: TCT_Error;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ErrorXpgList);
     procedure CopyTo(AItem: TCT_ErrorXpgList);
     property Items[Index: integer]: TCT_Error read GetItems; default;
     end;

     TCT_String = class(TXPGBase)
protected
     FV: AxUCString;
     FU: boolean;
     FF: boolean;
     FC: AxUCString;
     FCp: integer;
     FIn: integer;
     FBc: integer;
     FFc: integer;
     FI: boolean;
     FUn: boolean;
     FSt: boolean;
     FB: boolean;
     FTplsXpgList: TCT_TuplesXpgList;
     FXXpgList: TCT_XXpgList;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_String);
     procedure CopyTo(AItem: TCT_String);

     function  Create_TplsXpgList: TCT_TuplesXpgList;
     function  Create_XXpgList: TCT_XXpgList;

     property V: AxUCString read FV write FV;
     property U: boolean read FU write FU;
     property F: boolean read FF write FF;
     property C: AxUCString read FC write FC;
     property Cp: integer read FCp write FCp;
     property In_: integer read FIn write FIn;
     property Bc: integer read FBc write FBc;
     property Fc_: integer read FFc write FFc;
     property I: boolean read FI write FI;
     property Un: boolean read FUn write FUn;
     property St: boolean read FSt write FSt;
     property B: boolean read FB write FB;
     property TplsXpgList: TCT_TuplesXpgList read FTplsXpgList;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_StringXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_String;
public
     function  Add: TCT_String;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_StringXpgList);
     procedure CopyTo(AItem: TCT_StringXpgList);
     property Items[Index: integer]: TCT_String read GetItems; default;
     end;

     TCT_DateTime = class(TXPGBase)
protected
     FV: TDateTime;
     FU: boolean;
     FF: boolean;
     FC: AxUCString;
     FCp: integer;
     FXXpgList: TCT_XXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DateTime);
     procedure CopyTo(AItem: TCT_DateTime);
     function  Create_XXpgList: TCT_XXpgList;

     property V: TDateTime read FV write FV;
     property U: boolean read FU write FU;
     property F: boolean read FF write FF;
     property C: AxUCString read FC write FC;
     property Cp: integer read FCp write FCp;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_DateTimeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DateTime;
public
     function  Add: TCT_DateTime;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DateTimeXpgList);
     procedure CopyTo(AItem: TCT_DateTimeXpgList);
     property Items[Index: integer]: TCT_DateTime read GetItems; default;
     end;

     TCT_Groups = class(TXPGBase)
protected
     FCount: integer;
     FGroupXpgList: TCT_LevelGroupXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Groups);
     procedure CopyTo(AItem: TCT_Groups);
     function  Create_GroupXpgList: TCT_LevelGroupXpgList;

     property Count: integer read FCount write FCount;
     property GroupXpgList: TCT_LevelGroupXpgList read FGroupXpgList;
     end;

     TCT_Item = class(TXPGBase)
protected
     FN: AxUCString;
     FT: TST_ItemType;
     FH: boolean;
     FS: boolean;
     FSd: boolean;
     FF: boolean;
     FM: boolean;
     FC: boolean;
     FX: integer;
     FD: boolean;
     FE: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Item);
     procedure CopyTo(AItem: TCT_Item);

     property N: AxUCString read FN write FN;
     property T: TST_ItemType read FT write FT;
     property H: boolean read FH write FH;
     property S: boolean read FS write FS;
     property Sd: boolean read FSd write FSd;
     property F: boolean read FF write FF;
     property M: boolean read FM write FM;
     property C: boolean read FC write FC;
     property X: integer read FX write FX;
     property D: boolean read FD write FD;
     property E: boolean read FE write FE;
     end;

     TCT_PivotArea = class(TXPGBase)
protected
     FField: integer;
     FType: TST_PivotAreaType;
     FDataOnly: boolean;
     FLabelOnly: boolean;
     FGrandRow: boolean;
     FGrandCol: boolean;
     FCacheIndex: boolean;
     FOutline: boolean;
     FOffset: AxUCString;
     FCollapsedLevelsAreSubtotals: boolean;
     FAxis: TST_Axis;
     FFieldPosition: integer;
     FReferences: TCT_PivotAreaReferences;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotArea);
     procedure CopyTo(AItem: TCT_PivotArea);
     function  Create_References: TCT_PivotAreaReferences;
     function  Create_ExtLst: TCT_ExtensionList;

     property Field: integer read FField write FField;
     property Type_: TST_PivotAreaType read FType write FType;
     property DataOnly: boolean read FDataOnly write FDataOnly;
     property LabelOnly: boolean read FLabelOnly write FLabelOnly;
     property GrandRow: boolean read FGrandRow write FGrandRow;
     property GrandCol: boolean read FGrandCol write FGrandCol;
     property CacheIndex: boolean read FCacheIndex write FCacheIndex;
     property Outline: boolean read FOutline write FOutline;
     property Offset: AxUCString read FOffset write FOffset;
     property CollapsedLevelsAreSubtotals: boolean read FCollapsedLevelsAreSubtotals write FCollapsedLevelsAreSubtotals;
     property Axis: TST_Axis read FAxis write FAxis;
     property FieldPosition: integer read FFieldPosition write FFieldPosition;
     property References: TCT_PivotAreaReferences read FReferences;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PivotAreaXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotArea;
public
     function  Add: TCT_PivotArea;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotAreaXpgList);
     procedure CopyTo(AItem: TCT_PivotAreaXpgList);
     property Items[Index: integer]: TCT_PivotArea read GetItems; default;
     end;

     TCT_MemberProperty = class(TXPGBase)
protected
     FName: AxUCString;
     FShowCell: boolean;
     FShowTip: boolean;
     FShowAsCaption: boolean;
     FNameLen: integer;
     FPPos: integer;
     FPLen: integer;
     FLevel: integer;
     FField: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MemberProperty);
     procedure CopyTo(AItem: TCT_MemberProperty);

     property Name: AxUCString read FName write FName;
     property ShowCell: boolean read FShowCell write FShowCell;
     property ShowTip: boolean read FShowTip write FShowTip;
     property ShowAsCaption: boolean read FShowAsCaption write FShowAsCaption;
     property NameLen: integer read FNameLen write FNameLen;
     property PPos: integer read FPPos write FPPos;
     property PLen: integer read FPLen write FPLen;
     property Level: integer read FLevel write FLevel;
     property Field: integer read FField write FField;
     end;

     TCT_MemberPropertyXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_MemberProperty;
public
     function  Add: TCT_MemberProperty;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_MemberPropertyXpgList);
     procedure CopyTo(AItem: TCT_MemberPropertyXpgList);
     property Items[Index: integer]: TCT_MemberProperty read GetItems; default;
     end;

     TCT_Member = class(TXPGBase)
protected
     FName: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Member);
     procedure CopyTo(AItem: TCT_Member);

     property Name: AxUCString read FName write FName;
     end;

     TCT_MemberXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Member;
public
     function  Add: TCT_Member;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_MemberXpgList);
     procedure CopyTo(AItem: TCT_MemberXpgList);
     property Items[Index: integer]: TCT_Member read GetItems; default;
     end;

     TCT_PCDSCPage = class(TXPGBase)
protected
     FCount: integer;
     FPageItemXpgList: TCT_PageItemXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PCDSCPage);
     procedure CopyTo(AItem: TCT_PCDSCPage);
     function  Create_PageItemXpgList: TCT_PageItemXpgList;

     property Count: integer read FCount write FCount;
     property PageItemXpgList: TCT_PageItemXpgList read FPageItemXpgList;
     end;

     TCT_PCDSCPageXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PCDSCPage;
public
     function  Add: TCT_PCDSCPage;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PCDSCPageXpgList);
     procedure CopyTo(AItem: TCT_PCDSCPageXpgList);
     property Items[Index: integer]: TCT_PCDSCPage read GetItems; default;
     end;

     TCT_RangeSet = class(TXPGBase)
protected
     FI1: integer;
     FI2: integer;
     FI3: integer;
     FI4: integer;
     FRef: AxUCString;
     FName: AxUCString;
     FSheet: AxUCString;
     FR_Id: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RangeSet);
     procedure CopyTo(AItem: TCT_RangeSet);

     property I1: integer read FI1 write FI1;
     property I2: integer read FI2 write FI2;
     property I3: integer read FI3 write FI3;
     property I4: integer read FI4 write FI4;
     property Ref: AxUCString read FRef write FRef;
     property Name: AxUCString read FName write FName;
     property Sheet: AxUCString read FSheet write FSheet;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_RangeSetXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RangeSet;
public
     function  Add: TCT_RangeSet;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_RangeSetXpgList);
     procedure CopyTo(AItem: TCT_RangeSetXpgList);
     property Items[Index: integer]: TCT_RangeSet read GetItems; default;
     end;

     TCT_RangePr = class(TXPGBase)
protected
     FAutoStart: boolean;
     FAutoEnd: boolean;
     FGroupBy: TST_GroupBy;
     FStartNum: double;
     FEndNum: double;
     FStartDate: TDateTime;
     FEndDate: TDateTime;
     FGroupInterval: double;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RangePr);
     procedure CopyTo(AItem: TCT_RangePr);

     property AutoStart: boolean read FAutoStart write FAutoStart;
     property AutoEnd: boolean read FAutoEnd write FAutoEnd;
     property GroupBy: TST_GroupBy read FGroupBy write FGroupBy;
     property StartNum: double read FStartNum write FStartNum;
     property EndNum: double read FEndNum write FEndNum;
     property StartDate: TDateTime read FStartDate write FStartDate;
     property EndDate: TDateTime read FEndDate write FEndDate;
     property GroupInterval: double read FGroupInterval write FGroupInterval;
     end;

     TCT_DiscretePr = class(TXPGBase)
protected
     FCount: integer;
     FXXpgList: TCT_IndexXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DiscretePr);
     procedure CopyTo(AItem: TCT_DiscretePr);
     function  Create_XXpgList: TCT_IndexXpgList;

     property Count: integer read FCount write FCount;
     property XXpgList: TCT_IndexXpgList read FXXpgList;
     end;

     TCT_GroupItems = class(TXPGBase)
protected
     FCount: integer;
     FMXpgList: TCT_MissingXpgList;
     FNXpgList: TCT_NumberXpgList;
     FBXpgList: TCT_BooleanXpgList;
     FEXpgList: TCT_ErrorXpgList;
     FSXpgList: TCT_StringXpgList;
     FDXpgList: TCT_DateTimeXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupItems);
     procedure CopyTo(AItem: TCT_GroupItems);
     function  Create_MXpgList: TCT_MissingXpgList;
     function  Create_NXpgList: TCT_NumberXpgList;
     function  Create_BXpgList: TCT_BooleanXpgList;
     function  Create_EXpgList: TCT_ErrorXpgList;
     function  Create_SXpgList: TCT_StringXpgList;
     function  Create_DXpgList: TCT_DateTimeXpgList;

     property Count: integer read FCount write FCount;
     property MXpgList: TCT_MissingXpgList read FMXpgList;
     property NXpgList: TCT_NumberXpgList read FNXpgList;
     property BXpgList: TCT_BooleanXpgList read FBXpgList;
     property EXpgList: TCT_ErrorXpgList read FEXpgList;
     property SXpgList: TCT_StringXpgList read FSXpgList;
     property DXpgList: TCT_DateTimeXpgList read FDXpgList;
     end;

     TCT_FieldUsage = class(TXPGBase)
protected
     FX: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FieldUsage);
     procedure CopyTo(AItem: TCT_FieldUsage);

     property X: integer read FX write FX;
     end;

     TCT_FieldUsageXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FieldUsage;
public
     function  Add: TCT_FieldUsage;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_FieldUsageXpgList);
     procedure CopyTo(AItem: TCT_FieldUsageXpgList);
     property Items[Index: integer]: TCT_FieldUsage read GetItems; default;
     end;

     TCT_GroupLevel = class(TXPGBase)
protected
     FUniqueName: AxUCString;
     FCaption: AxUCString;
     FUser: boolean;
     FCustomRollUp: boolean;
     FGroups: TCT_Groups;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupLevel);
     procedure CopyTo(AItem: TCT_GroupLevel);
     function  Create_Groups: TCT_Groups;
     function  Create_ExtLst: TCT_ExtensionList;

     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Caption: AxUCString read FCaption write FCaption;
     property User: boolean read FUser write FUser;
     property CustomRollUp: boolean read FCustomRollUp write FCustomRollUp;
     property Groups: TCT_Groups read FGroups;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_GroupLevelXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GroupLevel;
public
     function  Add: TCT_GroupLevel;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_GroupLevelXpgList);
     procedure CopyTo(AItem: TCT_GroupLevelXpgList);
     property Items[Index: integer]: TCT_GroupLevel read GetItems; default;
     end;

     TCT_Items = class(TXPGBase)
private
     function GetItems(Index: integer): TCT_Item;
protected
     FItems: TObjectList;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_Items);
     procedure CopyTo(AItem: TCT_Items);

     function  Add: TCT_Item;
     procedure Exchange(Aindex1,AIndex2: integer);

     function  Count: integer;

     property Items[Index: integer]: TCT_Item read GetItems; default;
     end;

     TCT_AutoSortScope = class(TXPGBase)
protected
     FPivotArea: TCT_PivotArea;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_AutoSortScope);
     procedure CopyTo(AItem: TCT_AutoSortScope);
     function  Create_PivotArea: TCT_PivotArea;

     property PivotArea: TCT_PivotArea read FPivotArea;
     end;

     TCT_PivotAreas = class(TXPGBase)
protected
     FCount: integer;
     FPivotAreaXpgList: TCT_PivotAreaXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotAreas);
     procedure CopyTo(AItem: TCT_PivotAreas);
     function  Create_PivotAreaXpgList: TCT_PivotAreaXpgList;

     property Count: integer read FCount write FCount;
     property PivotAreaXpgList: TCT_PivotAreaXpgList read FPivotAreaXpgList;
     end;

     TCT_MemberProperties = class(TXPGBase)
protected
     FCount: integer;
     FMpXpgList: TCT_MemberPropertyXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MemberProperties);
     procedure CopyTo(AItem: TCT_MemberProperties);
     function  Create_MpXpgList: TCT_MemberPropertyXpgList;

     property Count: integer read FCount write FCount;
     property MpXpgList: TCT_MemberPropertyXpgList read FMpXpgList;
     end;

     TCT_Members = class(TXPGBase)
protected
     FCount: integer;
     FLevel: integer;
     FMemberXpgList: TCT_MemberXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Members);
     procedure CopyTo(AItem: TCT_Members);
     function  Create_MemberXpgList: TCT_MemberXpgList;

     property Count: integer read FCount write FCount;
     property Level: integer read FLevel write FLevel;
     property MemberXpgList: TCT_MemberXpgList read FMemberXpgList;
     end;

     TCT_MembersXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Members;
public
     function  Add: TCT_Members;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_MembersXpgList);
     procedure CopyTo(AItem: TCT_MembersXpgList);
     property Items[Index: integer]: TCT_Members read GetItems; default;
     end;

     TCT_Pages = class(TXPGBase)
protected
     FCount: integer;
     FPageXpgList: TCT_PCDSCPageXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Pages);
     procedure CopyTo(AItem: TCT_Pages);
     function  Create_PageXpgList: TCT_PCDSCPageXpgList;

     property Count: integer read FCount write FCount;
     property PageXpgList: TCT_PCDSCPageXpgList read FPageXpgList;
     end;

     TCT_RangeSets = class(TXPGBase)
protected
     FCount: integer;
     FRangeSetXpgList: TCT_RangeSetXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RangeSets);
     procedure CopyTo(AItem: TCT_RangeSets);
     function  Create_RangeSetXpgList: TCT_RangeSetXpgList;

     property Count: integer read FCount write FCount;
     property RangeSetXpgList: TCT_RangeSetXpgList read FRangeSetXpgList;
     end;

     TXLSSharedItemsValuesHash = class(TXPGBase)
private
     function GetCount: integer;
protected
     FBooleanCount: integer;
     FDateCount   : integer;
     FErrorCount  : integer;
     FBlankCount  : integer;
     FNumericCount: integer;
     FIntegerCount: integer;
     FStringCount : integer;

     FMinValue    : double;
     FMaxValue    : double;
     FMinDate     : TDateTime;
     FMaxDate     : TDateTime;
     FLongText    : boolean;

     FCurrType    : TXLSSharedItemsValueType;

     FUsed        : boolean;

     FValueList   : TXLSUniqueSharedItemsValues;

     function  HashOf(AValue: TDateTime): longword; overload;
     function  HashOf(AValue: TXc12CellError): longword; overload;
     function  HashOf(AValue: double): longword; overload;
     function  HashOf(AValue: AxUCString): longword; overload;
public
     constructor Create(Size: Cardinal = 256);
     destructor Destroy; override;

     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;

     procedure AddBoolean(AValue: boolean);
     procedure AddDateTime(AValue: TDateTime);
     procedure AddError(AValue: TXc12CellError);
     procedure Add;
     procedure AddFloat(AValue: double);
     procedure AddString(AValue: AxUCString);

     function  CanSort: boolean;
     function  IsNumeric: boolean;
     function  IsMonthText: boolean;
     function  IsDayText: boolean;

     property Count       : integer read GetCount;
     property BooleanCount: integer read FBooleanCount;
     property DateCount   : integer read FDateCount;
     property ErrorCount  : integer read FErrorCount;
     property BlankCount  : integer read FBlankCount;
     property NumericCount: integer read FNumericCount;
     property IntegerCount: integer read FIntegerCount;
     property StringCount : integer read FStringCount;

     property MinValue    : double read FMinValue write FMinValue;
     property MaxValue    : double read FMaxValue write FMaxValue;
     property MinDate     : TDateTime read FMinDate write FMinDate;
     property MaxDate     : TDateTime read FMaxDate write FMaxDate;
     property LongText    : boolean read FLongText write FLongText;

     property Used        : boolean read FUsed write FUsed;

     property ValueList   : TXLSUniqueSharedItemsValues read FValueList;
     end;

     TCT_SharedItems = class(TXPGBase)
protected
     FContainsSemiMixedTypes: boolean;
     FContainsNonDate: boolean;
     FContainsMixedTypes: boolean;

     FValues: TXLSSharedItemsValuesHash;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_SharedItems);
     procedure CopyTo(AItem: TCT_SharedItems);

     procedure UpdateNum(AValue: double);
     procedure UpdateDate(AValue: TDateTime);
     procedure UpdateString(AValue: AxUCString);
     procedure UpdateBoolean(AValue: boolean);
     procedure UpdateBlank;

     property ContainsSemiMixedTypes: boolean read FContainsSemiMixedTypes write FContainsSemiMixedTypes;
     property ContainsNonDate: boolean read FContainsNonDate write FContainsNonDate;

     property Values: TXLSSharedItemsValuesHash read FValues;
     end;

     TCT_FieldGroup = class(TXPGBase)
protected
     FPar: integer;
     FBase: integer;
     FRangePr: TCT_RangePr;
     FDiscretePr: TCT_DiscretePr;
     FGroupItems: TCT_GroupItems;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FieldGroup);
     procedure CopyTo(AItem: TCT_FieldGroup);
     function  Create_RangePr: TCT_RangePr;
     function  Create_DiscretePr: TCT_DiscretePr;
     function  Create_GroupItems: TCT_GroupItems;

     property Par: integer read FPar write FPar;
     property Base: integer read FBase write FBase;
     property RangePr: TCT_RangePr read FRangePr;
     property DiscretePr: TCT_DiscretePr read FDiscretePr;
     property GroupItems: TCT_GroupItems read FGroupItems;
     end;

     TCT_FieldsUsage = class(TXPGBase)
protected
     FCount: integer;
     FFieldUsageXpgList: TCT_FieldUsageXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_FieldsUsage);
     procedure CopyTo(AItem: TCT_FieldsUsage);
     function  Create_FieldUsageXpgList: TCT_FieldUsageXpgList;

     property Count: integer read FCount write FCount;
     property FieldUsageXpgList: TCT_FieldUsageXpgList read FFieldUsageXpgList;
     end;

     TCT_GroupLevels = class(TXPGBase)
protected
     FCount: integer;
     FGroupLevelXpgList: TCT_GroupLevelXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_GroupLevels);
     procedure CopyTo(AItem: TCT_GroupLevels);
     function  Create_GroupLevelXpgList: TCT_GroupLevelXpgList;

     property Count: integer read FCount write FCount;
     property GroupLevelXpgList: TCT_GroupLevelXpgList read FGroupLevelXpgList;
     end;

     TCT_Set = class(TXPGBase)
protected
     FCount: integer;
     FMaxRank: integer;
     FSetDefinition: AxUCString;
     FSortType: TST_SortType;
     FQueryFailed: boolean;
     FTplsXpgList: TCT_TuplesXpgList;
     FSortByTuple: TCT_Tuples;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Set);
     procedure CopyTo(AItem: TCT_Set);
     function  Create_TplsXpgList: TCT_TuplesXpgList;
     function  Create_SortByTuple: TCT_Tuples;

     property Count: integer read FCount write FCount;
     property MaxRank: integer read FMaxRank write FMaxRank;
     property SetDefinition: AxUCString read FSetDefinition write FSetDefinition;
     property SortType: TST_SortType read FSortType write FSortType;
     property QueryFailed: boolean read FQueryFailed write FQueryFailed;
     property TplsXpgList: TCT_TuplesXpgList read FTplsXpgList;
     property SortByTuple: TCT_Tuples read FSortByTuple;
     end;

     TCT_SetXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Set;
public
     function  Add: TCT_Set;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_SetXpgList);
     procedure CopyTo(AItem: TCT_SetXpgList);
     property Items[Index: integer]: TCT_Set read GetItems; default;
     end;

     TCT_Query = class(TXPGBase)
protected
     FMdx: AxUCString;
     FTpls: TCT_Tuples;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Query);
     procedure CopyTo(AItem: TCT_Query);
     function  Create_Tpls: TCT_Tuples;

     property Mdx: AxUCString read FMdx write FMdx;
     property Tpls: TCT_Tuples read FTpls;
     end;

     TCT_QueryXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Query;
public
     function  Add: TCT_Query;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_QueryXpgList);
     procedure CopyTo(AItem: TCT_QueryXpgList);
     property Items[Index: integer]: TCT_Query read GetItems; default;
     end;

     TCT_ServerFormat = class(TXPGBase)
protected
     FCulture: AxUCString;
     FFormat: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ServerFormat);
     procedure CopyTo(AItem: TCT_ServerFormat);

     property Culture: AxUCString read FCulture write FCulture;
     property Format: AxUCString read FFormat write FFormat;
     end;

     TCT_ServerFormatXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ServerFormat;
public
     function  Add: TCT_ServerFormat;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ServerFormatXpgList);
     procedure CopyTo(AItem: TCT_ServerFormatXpgList);
     property Items[Index: integer]: TCT_ServerFormat read GetItems; default;
     end;

     TCT_PivotField = class(TXPGBase)
protected
     FName: AxUCString;
     FAxis: TST_Axis;
     FDataField: boolean;
     FSubtotalCaption: AxUCString;
     FShowDropDowns: boolean;
     FHiddenLevel: boolean;
     FUniqueMemberProperty: AxUCString;
     FCompact: boolean;
     FAllDrilled: boolean;
     FNumFmtId: AxUCString;
     FOutline: boolean;
     FSubtotalTop: boolean;
     FDragToRow: boolean;
     FDragToCol: boolean;
     FMultipleItemSelectionAllowed: boolean;
     FDragToPage: boolean;
     FDragToData: boolean;
     FDragOff: boolean;
     FShowAll: boolean;
     FInsertBlankRow: boolean;
     FServerField: boolean;
     FInsertPageBreak: boolean;
     FAutoShow: boolean;
     FTopAutoShow: boolean;
     FHideNewItems: boolean;
     FMeasureFilter: boolean;
     FIncludeNewItemsInFilter: boolean;
     FItemPageCount: integer;
     FSortType: TST_FieldSortType;
     FDataSourceSort: boolean;
     FNonAutoSortDefault: boolean;
     FRankBy: integer;
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
     FShowPropCell: boolean;
     FShowPropTip: boolean;
     FShowPropAsCaption: boolean;
     FDefaultAttributeDrillState: boolean;
     FItems: TCT_Items;
     FAutoSortScope: TCT_AutoSortScope;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotField);
     procedure CopyTo(AItem: TCT_PivotField);
     function  Create_Items: TCT_Items;
     function  Create_AutoSortScope: TCT_AutoSortScope;
     function  Create_ExtLst: TCT_ExtensionList;

     property Name: AxUCString read FName write FName;
     property Axis: TST_Axis read FAxis write FAxis;
     property DataField: boolean read FDataField write FDataField;
     property SubtotalCaption: AxUCString read FSubtotalCaption write FSubtotalCaption;
     property ShowDropDowns: boolean read FShowDropDowns write FShowDropDowns;
     property HiddenLevel: boolean read FHiddenLevel write FHiddenLevel;
     property UniqueMemberProperty: AxUCString read FUniqueMemberProperty write FUniqueMemberProperty;
     property Compact: boolean read FCompact write FCompact;
     property AllDrilled: boolean read FAllDrilled write FAllDrilled;
     property NumFmtId: AxUCString read FNumFmtId write FNumFmtId;
     property Outline: boolean read FOutline write FOutline;
     property SubtotalTop: boolean read FSubtotalTop write FSubtotalTop;
     property DragToRow: boolean read FDragToRow write FDragToRow;
     property DragToCol: boolean read FDragToCol write FDragToCol;
     property MultipleItemSelectionAllowed: boolean read FMultipleItemSelectionAllowed write FMultipleItemSelectionAllowed;
     property DragToPage: boolean read FDragToPage write FDragToPage;
     property DragToData: boolean read FDragToData write FDragToData;
     property DragOff: boolean read FDragOff write FDragOff;
     property ShowAll: boolean read FShowAll write FShowAll;
     property InsertBlankRow: boolean read FInsertBlankRow write FInsertBlankRow;
     property ServerField: boolean read FServerField write FServerField;
     property InsertPageBreak: boolean read FInsertPageBreak write FInsertPageBreak;
     property AutoShow: boolean read FAutoShow write FAutoShow;
     property TopAutoShow: boolean read FTopAutoShow write FTopAutoShow;
     property HideNewItems: boolean read FHideNewItems write FHideNewItems;
     property MeasureFilter: boolean read FMeasureFilter write FMeasureFilter;
     property IncludeNewItemsInFilter: boolean read FIncludeNewItemsInFilter write FIncludeNewItemsInFilter;
     property ItemPageCount: integer read FItemPageCount write FItemPageCount;
     property SortType: TST_FieldSortType read FSortType write FSortType;
     property DataSourceSort: boolean read FDataSourceSort write FDataSourceSort;
     property NonAutoSortDefault: boolean read FNonAutoSortDefault write FNonAutoSortDefault;
     property RankBy: integer read FRankBy write FRankBy;

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

     property ShowPropCell: boolean read FShowPropCell write FShowPropCell;
     property ShowPropTip: boolean read FShowPropTip write FShowPropTip;
     property ShowPropAsCaption: boolean read FShowPropAsCaption write FShowPropAsCaption;
     property DefaultAttributeDrillState: boolean read FDefaultAttributeDrillState write FDefaultAttributeDrillState;
     property Items: TCT_Items read FItems;
     property AutoSortScope: TCT_AutoSortScope read FAutoSortScope;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PivotFieldXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotField;
public
     function  Add: TCT_PivotField;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotFieldXpgList);
     procedure CopyTo(AItem: TCT_PivotFieldXpgList);
     property Items[Index: integer]: TCT_PivotField read GetItems; default;
     end;

     TCT_Field = class(TXPGBase)
protected
     FX: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Field);
     procedure CopyTo(AItem: TCT_Field);

     property X: integer read FX write FX;
     end;

     TCT_I = class(TXPGBase)
protected
     FT: TST_ItemType;
     FR: integer;
     FI: integer;
     FXXpgList: TCT_XXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_I);
     procedure CopyTo(AItem: TCT_I);
     function  Create_XXpgList: TCT_XXpgList;

     property T: TST_ItemType read FT write FT;
     property R: integer read FR write FR;
     property I: integer read FI write FI;
     property XXpgList: TCT_XXpgList read FXXpgList;
     end;

     TCT_PageField = class(TXPGBase)
protected
     FFld: integer;
     FItem: integer;
     FHier: integer;
     FName: AxUCString;
     FCap: AxUCString;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PageField);
     procedure CopyTo(AItem: TCT_PageField);
     function  Create_ExtLst: TCT_ExtensionList;

     property Fld: integer read FFld write FFld;
     property Item: integer read FItem write FItem;
     property Hier: integer read FHier write FHier;
     property Name: AxUCString read FName write FName;
     property Cap: AxUCString read FCap write FCap;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PageFieldXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PageField;
public
     function  Add: TCT_PageField;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PageFieldXpgList);
     procedure CopyTo(AItem: TCT_PageFieldXpgList);
     property Items[Index: integer]: TCT_PageField read GetItems; default;
     end;

     TCT_DataField = class(TXPGBase)
protected
     FName: AxUCString;
     FFld: integer;
     FSubtotal: AxUCString;
     FShowDataAs: TST_ShowDataAs;
     FBaseField: integer;
     FBaseItem: integer;
     FNumFmtId: AxUCString;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_DataField);
     procedure CopyTo(AItem: TCT_DataField);
     function  Create_ExtLst: TCT_ExtensionList;

     property Name: AxUCString read FName write FName;
     property Fld: integer read FFld write FFld;
     property Subtotal: AxUCString read FSubtotal write FSubtotal;
     property ShowDataAs: TST_ShowDataAs read FShowDataAs write FShowDataAs;
     property BaseField: integer read FBaseField write FBaseField;
     property BaseItem: integer read FBaseItem write FBaseItem;
     property NumFmtId: AxUCString read FNumFmtId write FNumFmtId;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DataFields = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DataField;
public
     function  Add: TCT_DataField;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_DataFields);
     procedure CopyTo(AItem: TCT_DataFields);

     property Items[Index: integer]: TCT_DataField read GetItems; default;
     end;

     TCT_Format = class(TXPGBase)
protected
     FAction: TST_FormatAction;
     FDxfId: AxUCString;
     FPivotArea: TCT_PivotArea;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Format);
     procedure CopyTo(AItem: TCT_Format);
     function  Create_PivotArea: TCT_PivotArea;
     function  Create_ExtLst: TCT_ExtensionList;

     property Action: TST_FormatAction read FAction write FAction;
     property DxfId: AxUCString read FDxfId write FDxfId;
     property PivotArea: TCT_PivotArea read FPivotArea;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_FormatXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Format;
public
     function  Add: TCT_Format;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_FormatXpgList);
     procedure CopyTo(AItem: TCT_FormatXpgList);
     property Items[Index: integer]: TCT_Format read GetItems; default;
     end;

     TCT_ConditionalFormat = class(TXPGBase)
protected
     FScope: TST_Scope;
     FType: TST_Type;
     FPriority: integer;
     FPivotAreas: TCT_PivotAreas;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ConditionalFormat);
     procedure CopyTo(AItem: TCT_ConditionalFormat);
     function  Create_PivotAreas: TCT_PivotAreas;
     function  Create_ExtLst: TCT_ExtensionList;

     property Scope: TST_Scope read FScope write FScope;
     property Type_: TST_Type read FType write FType;
     property Priority: integer read FPriority write FPriority;
     property PivotAreas: TCT_PivotAreas read FPivotAreas;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ConditionalFormatXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ConditionalFormat;
public
     function  Add: TCT_ConditionalFormat;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ConditionalFormatXpgList);
     procedure CopyTo(AItem: TCT_ConditionalFormatXpgList);
     property Items[Index: integer]: TCT_ConditionalFormat read GetItems; default;
     end;

     TCT_ChartFormat = class(TXPGBase)
protected
     FChart: integer;
     FFormat: integer;
     FSeries: boolean;
     FPivotArea: TCT_PivotArea;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ChartFormat);
     procedure CopyTo(AItem: TCT_ChartFormat);
     function  Create_PivotArea: TCT_PivotArea;

     property Chart: integer read FChart write FChart;
     property Format: integer read FFormat write FFormat;
     property Series: boolean read FSeries write FSeries;
     property PivotArea: TCT_PivotArea read FPivotArea;
     end;

     TCT_ChartFormatXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ChartFormat;
public
     function  Add: TCT_ChartFormat;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_ChartFormatXpgList);
     procedure CopyTo(AItem: TCT_ChartFormatXpgList);
     property Items[Index: integer]: TCT_ChartFormat read GetItems; default;
     end;

     TCT_PivotHierarchy = class(TXPGBase)
protected
     FOutline: boolean;
     FMultipleItemSelectionAllowed: boolean;
     FSubtotalTop: boolean;
     FShowInFieldList: boolean;
     FDragToRow: boolean;
     FDragToCol: boolean;
     FDragToPage: boolean;
     FDragToData: boolean;
     FDragOff: boolean;
     FIncludeNewItemsInFilter: boolean;
     FCaption: AxUCString;
     FMps: TCT_MemberProperties;
     FMembersXpgList: TCT_MembersXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotHierarchy);
     procedure CopyTo(AItem: TCT_PivotHierarchy);
     function  Create_Mps: TCT_MemberProperties;
     function  Create_MembersXpgList: TCT_MembersXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property Outline: boolean read FOutline write FOutline;
     property MultipleItemSelectionAllowed: boolean read FMultipleItemSelectionAllowed write FMultipleItemSelectionAllowed;
     property SubtotalTop: boolean read FSubtotalTop write FSubtotalTop;
     property ShowInFieldList: boolean read FShowInFieldList write FShowInFieldList;
     property DragToRow: boolean read FDragToRow write FDragToRow;
     property DragToCol: boolean read FDragToCol write FDragToCol;
     property DragToPage: boolean read FDragToPage write FDragToPage;
     property DragToData: boolean read FDragToData write FDragToData;
     property DragOff: boolean read FDragOff write FDragOff;
     property IncludeNewItemsInFilter: boolean read FIncludeNewItemsInFilter write FIncludeNewItemsInFilter;
     property Caption: AxUCString read FCaption write FCaption;
     property Mps: TCT_MemberProperties read FMps;
     property MembersXpgList: TCT_MembersXpgList read FMembersXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PivotHierarchyXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotHierarchy;
public
     function  Add: TCT_PivotHierarchy;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotHierarchyXpgList);
     procedure CopyTo(AItem: TCT_PivotHierarchyXpgList);
     property Items[Index: integer]: TCT_PivotHierarchy read GetItems; default;
     end;

     TCT_PivotFilter = class(TXPGBase)
protected
     FFld: integer;
     FMpFld: integer;
     FType: TST_PivotFilterType;
     FEvalOrder: integer;
     FId: integer;
     FIMeasureHier: integer;
     FIMeasureFld: integer;
     FName: AxUCString;
     FDescription: AxUCString;
     FStringValue1: AxUCString;
     FStringValue2: AxUCString;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotFilter);
     procedure CopyTo(AItem: TCT_PivotFilter);
     function  Create_ExtLst: TCT_ExtensionList;

     property Fld: integer read FFld write FFld;
     property MpFld: integer read FMpFld write FMpFld;
     property Type_: TST_PivotFilterType read FType write FType;
     property EvalOrder: integer read FEvalOrder write FEvalOrder;
     property Id: integer read FId write FId;
     property IMeasureHier: integer read FIMeasureHier write FIMeasureHier;
     property IMeasureFld: integer read FIMeasureFld write FIMeasureFld;
     property Name: AxUCString read FName write FName;
     property Description: AxUCString read FDescription write FDescription;
     property StringValue1: AxUCString read FStringValue1 write FStringValue1;
     property StringValue2: AxUCString read FStringValue2 write FStringValue2;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_PivotFilterXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotFilter;
public
     function  Add: TCT_PivotFilter;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotFilterXpgList);
     procedure CopyTo(AItem: TCT_PivotFilterXpgList);
     property Items[Index: integer]: TCT_PivotFilter read GetItems; default;
     end;

     TCT_HierarchyUsage = class(TXPGBase)
protected
     FHierarchyUsage: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_HierarchyUsage);
     procedure CopyTo(AItem: TCT_HierarchyUsage);

     property HierarchyUsage: integer read FHierarchyUsage write FHierarchyUsage;
     end;

     TCT_HierarchyUsageXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_HierarchyUsage;
public
     function  Add: TCT_HierarchyUsage;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_HierarchyUsageXpgList);
     procedure CopyTo(AItem: TCT_HierarchyUsageXpgList);
     property Items[Index: integer]: TCT_HierarchyUsage read GetItems; default;
     end;

     TCT_WorksheetSource = class(TXPGBase)
private
     procedure SetRCells(const Value: TXLSRelCells);
protected
     FRef: AxUCString;
     FRCells: TXLSRelCells;
     FName: AxUCString;
     FSheet: AxUCString;
     FR_Id: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_WorksheetSource);
     procedure CopyTo(AItem: TCT_WorksheetSource);

     property Ref: AxUCString read FRef write FRef;
     property RCells: TXLSRelCells read FRCells write SetRCells;
     property Name: AxUCString read FName write FName;
     property Sheet: AxUCString read FSheet write FSheet;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_Consolidation = class(TXPGBase)
protected
     FAutoPage: boolean;
     FPages: TCT_Pages;
     FRangeSets: TCT_RangeSets;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Consolidation);
     procedure CopyTo(AItem: TCT_Consolidation);
     function  Create_Pages: TCT_Pages;
     function  Create_RangeSets: TCT_RangeSets;

     property AutoPage: boolean read FAutoPage write FAutoPage;
     property Pages: TCT_Pages read FPages;
     property RangeSets: TCT_RangeSets read FRangeSets;
     end;

     TCT_CacheField = class(TXPGBase)
protected
     FName: AxUCString;
     FCaption: AxUCString;
     FPropertyName: AxUCString;
     FServerField: boolean;
     FUniqueList: boolean;
     FNumFmtId: AxUCString;
     FFormula: AxUCString;
     FSqlType: integer;
     FHierarchy: integer;
     FLevel: integer;
     FDatabaseField: boolean;
     FMappingCount: integer;
     FMemberPropertyField: boolean;
     FSharedItems: TCT_SharedItems;
     FFieldGroup: TCT_FieldGroup;
     FMpMapXpgList: TCT_XXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CacheField);
     procedure CopyTo(AItem: TCT_CacheField);
     function  Create_SharedItems: TCT_SharedItems;
     function  Create_FieldGroup: TCT_FieldGroup;
     function  Create_MpMapXpgList: TCT_XXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property Name: AxUCString read FName write FName;
     property Caption: AxUCString read FCaption write FCaption;
     property PropertyName: AxUCString read FPropertyName write FPropertyName;
     property ServerField: boolean read FServerField write FServerField;
     property UniqueList: boolean read FUniqueList write FUniqueList;
     property NumFmtId: AxUCString read FNumFmtId write FNumFmtId;
     property Formula: AxUCString read FFormula write FFormula;
     property SqlType: integer read FSqlType write FSqlType;
     property Hierarchy: integer read FHierarchy write FHierarchy;
     property Level: integer read FLevel write FLevel;
     property DatabaseField: boolean read FDatabaseField write FDatabaseField;
     property MappingCount: integer read FMappingCount write FMappingCount;
     property MemberPropertyField: boolean read FMemberPropertyField write FMemberPropertyField;
     property SharedItems: TCT_SharedItems read FSharedItems;
     property FieldGroup: TCT_FieldGroup read FFieldGroup;
     property MpMapXpgList: TCT_XXpgList read FMpMapXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CacheHierarchy = class(TXPGBase)
protected
     FUniqueName: AxUCString;
     FCaption: AxUCString;
     FMeasure: boolean;
     FSet: boolean;
     FParentSet: integer;
     FIconSet: integer;
     FAttribute: boolean;
     FTime: boolean;
     FKeyAttribute: boolean;
     FDefaultMemberUniqueName: AxUCString;
     FAllUniqueName: AxUCString;
     FAllCaption: AxUCString;
     FDimensionUniqueName: AxUCString;
     FDisplayFolder: AxUCString;
     FMeasureGroup: AxUCString;
     FMeasures: boolean;
     FCount: integer;
     FOneField: boolean;
     FMemberValueDatatype: integer;
     FUnbalanced: boolean;
     FUnbalancedGroup: boolean;
     FHidden: boolean;
     FFieldsUsage: TCT_FieldsUsage;
     FGroupLevels: TCT_GroupLevels;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CacheHierarchy);
     procedure CopyTo(AItem: TCT_CacheHierarchy);
     function  Create_FieldsUsage: TCT_FieldsUsage;
     function  Create_GroupLevels: TCT_GroupLevels;
     function  Create_ExtLst: TCT_ExtensionList;

     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Caption: AxUCString read FCaption write FCaption;
     property Measure: boolean read FMeasure write FMeasure;
     property Set_: boolean read FSet write FSet;
     property ParentSet: integer read FParentSet write FParentSet;
     property IconSet: integer read FIconSet write FIconSet;
     property Attribute: boolean read FAttribute write FAttribute;
     property Time: boolean read FTime write FTime;
     property KeyAttribute: boolean read FKeyAttribute write FKeyAttribute;
     property DefaultMemberUniqueName: AxUCString read FDefaultMemberUniqueName write FDefaultMemberUniqueName;
     property AllUniqueName: AxUCString read FAllUniqueName write FAllUniqueName;
     property AllCaption: AxUCString read FAllCaption write FAllCaption;
     property DimensionUniqueName: AxUCString read FDimensionUniqueName write FDimensionUniqueName;
     property DisplayFolder: AxUCString read FDisplayFolder write FDisplayFolder;
     property MeasureGroup: AxUCString read FMeasureGroup write FMeasureGroup;
     property Measures: boolean read FMeasures write FMeasures;
     property Count: integer read FCount write FCount;
     property OneField: boolean read FOneField write FOneField;
     property MemberValueDatatype: integer read FMemberValueDatatype write FMemberValueDatatype;
     property Unbalanced: boolean read FUnbalanced write FUnbalanced;
     property UnbalancedGroup: boolean read FUnbalancedGroup write FUnbalancedGroup;
     property Hidden: boolean read FHidden write FHidden;
     property FieldsUsage: TCT_FieldsUsage read FFieldsUsage;
     property GroupLevels: TCT_GroupLevels read FGroupLevels;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CacheHierarchyXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CacheHierarchy;
public
     function  Add: TCT_CacheHierarchy;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_CacheHierarchyXpgList);
     procedure CopyTo(AItem: TCT_CacheHierarchyXpgList);
     property Items[Index: integer]: TCT_CacheHierarchy read GetItems; default;
     end;

     TCT_PCDKPI = class(TXPGBase)
protected
     FUniqueName: AxUCString;
     FCaption: AxUCString;
     FDisplayFolder: AxUCString;
     FMeasureGroup: AxUCString;
     FParent: AxUCString;
     FValue: AxUCString;
     FGoal: AxUCString;
     FStatus: AxUCString;
     FTrend: AxUCString;
     FWeight: AxUCString;
     FTime: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PCDKPI);
     procedure CopyTo(AItem: TCT_PCDKPI);

     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Caption: AxUCString read FCaption write FCaption;
     property DisplayFolder: AxUCString read FDisplayFolder write FDisplayFolder;
     property MeasureGroup: AxUCString read FMeasureGroup write FMeasureGroup;
     property Parent: AxUCString read FParent write FParent;
     property Value: AxUCString read FValue write FValue;
     property Goal: AxUCString read FGoal write FGoal;
     property Status: AxUCString read FStatus write FStatus;
     property Trend: AxUCString read FTrend write FTrend;
     property Weight: AxUCString read FWeight write FWeight;
     property Time: AxUCString read FTime write FTime;
     end;

     TCT_PCDKPIXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PCDKPI;
public
     function  Add: TCT_PCDKPI;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PCDKPIXpgList);
     procedure CopyTo(AItem: TCT_PCDKPIXpgList);
     property Items[Index: integer]: TCT_PCDKPI read GetItems; default;
     end;

     TCT_PCDSDTCEntries = class(TXPGBase)
protected
     FCount: integer;
     FMXpgList: TCT_MissingXpgList;
     FNXpgList: TCT_NumberXpgList;
     FEXpgList: TCT_ErrorXpgList;
     FSXpgList: TCT_StringXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PCDSDTCEntries);
     procedure CopyTo(AItem: TCT_PCDSDTCEntries);
     function  Create_MXpgList: TCT_MissingXpgList;
     function  Create_NXpgList: TCT_NumberXpgList;
     function  Create_EXpgList: TCT_ErrorXpgList;
     function  Create_SXpgList: TCT_StringXpgList;

     property Count: integer read FCount write FCount;
     property MXpgList: TCT_MissingXpgList read FMXpgList;
     property NXpgList: TCT_NumberXpgList read FNXpgList;
     property EXpgList: TCT_ErrorXpgList read FEXpgList;
     property SXpgList: TCT_StringXpgList read FSXpgList;
     end;

     TCT_Sets = class(TXPGBase)
protected
     FCount: integer;
     FSetXpgList: TCT_SetXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Sets);
     procedure CopyTo(AItem: TCT_Sets);
     function  Create_SetXpgList: TCT_SetXpgList;

     property Count: integer read FCount write FCount;
     property SetXpgList: TCT_SetXpgList read FSetXpgList;
     end;

     TCT_QueryCache = class(TXPGBase)
protected
     FCount: integer;
     FQueryXpgList: TCT_QueryXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_QueryCache);
     procedure CopyTo(AItem: TCT_QueryCache);
     function  Create_QueryXpgList: TCT_QueryXpgList;

     property Count: integer read FCount write FCount;
     property QueryXpgList: TCT_QueryXpgList read FQueryXpgList;
     end;

     TCT_ServerFormats = class(TXPGBase)
protected
     FCount: integer;
     FServerFormatXpgList: TCT_ServerFormatXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ServerFormats);
     procedure CopyTo(AItem: TCT_ServerFormats);
     function  Create_ServerFormatXpgList: TCT_ServerFormatXpgList;

     property Count: integer read FCount write FCount;
     property ServerFormatXpgList: TCT_ServerFormatXpgList read FServerFormatXpgList;
     end;

     TCT_CalculatedItem = class(TXPGBase)
protected
     FField: integer;
     FFormula: AxUCString;
     FPivotArea: TCT_PivotArea;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CalculatedItem);
     procedure CopyTo(AItem: TCT_CalculatedItem);
     function  Create_PivotArea: TCT_PivotArea;
     function  Create_ExtLst: TCT_ExtensionList;

     property Field: integer read FField write FField;
     property Formula: AxUCString read FFormula write FFormula;
     property PivotArea: TCT_PivotArea read FPivotArea;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CalculatedItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CalculatedItem;
public
     function  Add: TCT_CalculatedItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_CalculatedItemXpgList);
     procedure CopyTo(AItem: TCT_CalculatedItemXpgList);
     property Items[Index: integer]: TCT_CalculatedItem read GetItems; default;
     end;

     TCT_CalculatedMember = class(TXPGBase)
protected
     FName: AxUCString;
     FMdx: AxUCString;
     FMemberName: AxUCString;
     FHierarchy: AxUCString;
     FParent: AxUCString;
     FSolveOrder: integer;
     FSet: boolean;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CalculatedMember);
     procedure CopyTo(AItem: TCT_CalculatedMember);
     function  Create_ExtLst: TCT_ExtensionList;

     property Name: AxUCString read FName write FName;
     property Mdx: AxUCString read FMdx write FMdx;
     property MemberName: AxUCString read FMemberName write FMemberName;
     property Hierarchy: AxUCString read FHierarchy write FHierarchy;
     property Parent: AxUCString read FParent write FParent;
     property SolveOrder: integer read FSolveOrder write FSolveOrder;
     property Set_: boolean read FSet write FSet;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CalculatedMemberXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CalculatedMember;
public
     function  Add: TCT_CalculatedMember;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_CalculatedMemberXpgList);
     procedure CopyTo(AItem: TCT_CalculatedMemberXpgList);
     property Items[Index: integer]: TCT_CalculatedMember read GetItems; default;
     end;

     TCT_PivotDimension = class(TXPGBase)
protected
     FMeasure: boolean;
     FName: AxUCString;
     FUniqueName: AxUCString;
     FCaption: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotDimension);
     procedure CopyTo(AItem: TCT_PivotDimension);

     property Measure: boolean read FMeasure write FMeasure;
     property Name: AxUCString read FName write FName;
     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Caption: AxUCString read FCaption write FCaption;
     end;

     TCT_PivotDimensionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotDimension;
public
     function  Add: TCT_PivotDimension;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_PivotDimensionXpgList);
     procedure CopyTo(AItem: TCT_PivotDimensionXpgList);
     property Items[Index: integer]: TCT_PivotDimension read GetItems; default;
     end;

     TCT_MeasureGroup = class(TXPGBase)
protected
     FName: AxUCString;
     FCaption: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MeasureGroup);
     procedure CopyTo(AItem: TCT_MeasureGroup);

     property Name: AxUCString read FName write FName;
     property Caption: AxUCString read FCaption write FCaption;
     end;

     TCT_MeasureGroupXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_MeasureGroup;
public
     function  Add: TCT_MeasureGroup;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_MeasureGroupXpgList);
     procedure CopyTo(AItem: TCT_MeasureGroupXpgList);
     property Items[Index: integer]: TCT_MeasureGroup read GetItems; default;
     end;

     TCT_MeasureDimensionMap = class(TXPGBase)
protected
     FMeasureGroup: integer;
     FDimension: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MeasureDimensionMap);
     procedure CopyTo(AItem: TCT_MeasureDimensionMap);

     property MeasureGroup: integer read FMeasureGroup write FMeasureGroup;
     property Dimension: integer read FDimension write FDimension;
     end;

     TCT_MeasureDimensionMapXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_MeasureDimensionMap;
public
     function  Add: TCT_MeasureDimensionMap;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_MeasureDimensionMapXpgList);
     procedure CopyTo(AItem: TCT_MeasureDimensionMapXpgList);
     property Items[Index: integer]: TCT_MeasureDimensionMap read GetItems; default;
     end;

     TCT_Record = class(TXPGBase)
protected
     FMXpgList: TCT_MissingXpgList;
     FNXpgList: TCT_NumberXpgList;
     FBXpgList: TCT_BooleanXpgList;
     FEXpgList: TCT_ErrorXpgList;
     FSXpgList: TCT_StringXpgList;
     FDXpgList: TCT_DateTimeXpgList;
     FXXpgList: TCT_IndexXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Record);
     procedure CopyTo(AItem: TCT_Record);
     function  Create_MXpgList: TCT_MissingXpgList;
     function  Create_NXpgList: TCT_NumberXpgList;
     function  Create_BXpgList: TCT_BooleanXpgList;
     function  Create_EXpgList: TCT_ErrorXpgList;
     function  Create_SXpgList: TCT_StringXpgList;
     function  Create_DXpgList: TCT_DateTimeXpgList;
     function  Create_XXpgList: TCT_IndexXpgList;

     property MXpgList: TCT_MissingXpgList read FMXpgList;
     property NXpgList: TCT_NumberXpgList read FNXpgList;
     property BXpgList: TCT_BooleanXpgList read FBXpgList;
     property EXpgList: TCT_ErrorXpgList read FEXpgList;
     property SXpgList: TCT_StringXpgList read FSXpgList;
     property DXpgList: TCT_DateTimeXpgList read FDXpgList;
     property XXpgList: TCT_IndexXpgList read FXXpgList;
     end;

     TCT_RecordXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Record;
public
     function  Add: TCT_Record;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     procedure Assign(AItem: TCT_RecordXpgList);
     procedure CopyTo(AItem: TCT_RecordXpgList);
     property Items[Index: integer]: TCT_Record read GetItems; default;
     end;

     TCT_Location = class(TXPGBase)
protected
     FRef: AxUCString;
     FRCells: TXLSRelCells;
     FFirstHeaderRow: integer;
     FFirstDataRow: integer;
     FFirstDataCol: integer;
     FRowPageCount: integer;
     FColPageCount: integer;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Location);
     procedure CopyTo(AItem: TCT_Location);

     property Ref: AxUCString read FRef write FRef;
     property RCells: TXLSRelCells read FRCells write FRCells;
     property FirstHeaderRow: integer read FFirstHeaderRow write FFirstHeaderRow;
     property FirstDataRow: integer read FFirstDataRow write FFirstDataRow;
     property FirstDataCol: integer read FFirstDataCol write FFirstDataCol;
     property RowPageCount: integer read FRowPageCount write FRowPageCount;
     property ColPageCount: integer read FColPageCount write FColPageCount;
     end;

     TCT_PivotFields = class(TXPGBase)
 private
     function GetItems(Index: integer): TCT_PivotField;
protected
     FItems: TObjectList;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_PivotFields);
     procedure CopyTo(AItem: TCT_PivotFields);

     function  Add: TCT_PivotField;

     function  Count: integer;

     property Items[Index: integer]: TCT_PivotField read GetItems; default;
     end;

     TCT_RowColFields = class(TXPGBase)
private
     function  GetItems(Index: integer): TCT_Field;
protected
     FItems: TObjectList;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  Add: TCT_Field;

     function  Count: integer;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;

     procedure Assign(AItem: TCT_RowColFields);
     procedure CopyTo(AItem: TCT_RowColFields);

     property Items[Index: integer]: TCT_Field read GetItems; default;
     end;

     TCT_RowColItems = class(TXPGBase)
private
     function  GetItems(Index: integer): TCT_I;
protected
     FItems: TObjectList;
     FIsRow: boolean;
public
     constructor Create(AOwner: TXPGDocBase; AIsRow: boolean);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;

     procedure Assign(AItem: TCT_RowColItems);
     procedure CopyTo(AItem: TCT_RowColItems);

     function  Add: TCT_I;
     function  Count:  integer;

     property IsRow: boolean read FIsRow write FIsRow;
     property Items[Index: integer]: TCT_I read GetItems; default;
     end;

     TCT_PageFields = class(TXPGBase)
protected
     FCount: integer;
     FPageFieldXpgList: TCT_PageFieldXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PageFields);
     procedure CopyTo(AItem: TCT_PageFields);
     function  Create_PageFieldXpgList: TCT_PageFieldXpgList;

     property Count: integer read FCount write FCount;
     property PageFieldXpgList: TCT_PageFieldXpgList read FPageFieldXpgList;
     end;

     TCT_Formats = class(TXPGBase)
protected
     FCount: integer;
     FFormatXpgList: TCT_FormatXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Formats);
     procedure CopyTo(AItem: TCT_Formats);
     function  Create_FormatXpgList: TCT_FormatXpgList;

     property Count: integer read FCount write FCount;
     property FormatXpgList: TCT_FormatXpgList read FFormatXpgList;
     end;

     TCT_ConditionalFormats = class(TXPGBase)
protected
     FCount: integer;
     FConditionalFormatXpgList: TCT_ConditionalFormatXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ConditionalFormats);
     procedure CopyTo(AItem: TCT_ConditionalFormats);
     function  Create_ConditionalFormatXpgList: TCT_ConditionalFormatXpgList;

     property Count: integer read FCount write FCount;
     property ConditionalFormatXpgList: TCT_ConditionalFormatXpgList read FConditionalFormatXpgList;
     end;

     TCT_ChartFormats = class(TXPGBase)
protected
     FCount: integer;
     FChartFormatXpgList: TCT_ChartFormatXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ChartFormats);
     procedure CopyTo(AItem: TCT_ChartFormats);
     function  Create_ChartFormatXpgList: TCT_ChartFormatXpgList;

     property Count: integer read FCount write FCount;
     property ChartFormatXpgList: TCT_ChartFormatXpgList read FChartFormatXpgList;
     end;

     TCT_PivotHierarchies = class(TXPGBase)
protected
     FCount: integer;
     FPivotHierarchyXpgList: TCT_PivotHierarchyXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotHierarchies);
     procedure CopyTo(AItem: TCT_PivotHierarchies);
     function  Create_PivotHierarchyXpgList: TCT_PivotHierarchyXpgList;

     property Count: integer read FCount write FCount;
     property PivotHierarchyXpgList: TCT_PivotHierarchyXpgList read FPivotHierarchyXpgList;
     end;

     TCT_PivotTableStyle = class(TXPGBase)
protected
     FName: AxUCString;
     FShowRowHeaders: boolean;
     FShowColHeaders: boolean;
     FShowRowStripes: boolean;
     FShowColStripes: boolean;
     FShowLastColumn: boolean;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotTableStyle);
     procedure CopyTo(AItem: TCT_PivotTableStyle);

     property Name: AxUCString read FName write FName;
     property ShowRowHeaders: boolean read FShowRowHeaders write FShowRowHeaders;
     property ShowColHeaders: boolean read FShowColHeaders write FShowColHeaders;
     property ShowRowStripes: boolean read FShowRowStripes write FShowRowStripes;
     property ShowColStripes: boolean read FShowColStripes write FShowColStripes;
     property ShowLastColumn: boolean read FShowLastColumn write FShowLastColumn;
     end;

     TCT_PivotFilters = class(TXPGBase)
protected
     FCount: integer;
     FFilterXpgList: TCT_PivotFilterXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PivotFilters);
     procedure CopyTo(AItem: TCT_PivotFilters);
     function  Create_FilterXpgList: TCT_PivotFilterXpgList;

     property Count: integer read FCount write FCount;
     property FilterXpgList: TCT_PivotFilterXpgList read FFilterXpgList;
     end;

     TCT_RowHierarchiesUsage = class(TXPGBase)
protected
     FCount: integer;
     FRowHierarchyUsageXpgList: TCT_HierarchyUsageXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_RowHierarchiesUsage);
     procedure CopyTo(AItem: TCT_RowHierarchiesUsage);
     function  Create_RowHierarchyUsageXpgList: TCT_HierarchyUsageXpgList;

     property Count: integer read FCount write FCount;
     property RowHierarchyUsageXpgList: TCT_HierarchyUsageXpgList read FRowHierarchyUsageXpgList;
     end;

     TCT_ColHierarchiesUsage = class(TXPGBase)
protected
     FCount: integer;
     FColHierarchyUsageXpgList: TCT_HierarchyUsageXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_ColHierarchiesUsage);
     procedure CopyTo(AItem: TCT_ColHierarchiesUsage);
     function  Create_ColHierarchyUsageXpgList: TCT_HierarchyUsageXpgList;

     property Count: integer read FCount write FCount;
     property ColHierarchyUsageXpgList: TCT_HierarchyUsageXpgList read FColHierarchyUsageXpgList;
     end;

     TCT_CacheSource = class(TXPGBase)
protected
     FType: TST_SourceType;
     FConnectionId: integer;
     FWorksheetSource: TCT_WorksheetSource;
     FConsolidation: TCT_Consolidation;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CacheSource);
     procedure CopyTo(AItem: TCT_CacheSource);
     function  Create_WorksheetSource: TCT_WorksheetSource;
     function  Create_Consolidation: TCT_Consolidation;
     function  Create_ExtLst: TCT_ExtensionList;

     property Type_: TST_SourceType read FType write FType;
     property ConnectionId: integer read FConnectionId write FConnectionId;
     property WorksheetSource: TCT_WorksheetSource read FWorksheetSource;
     property Consolidation: TCT_Consolidation read FConsolidation;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CacheFields = class(TXPGBase)
private
     function GetItems(Index: integer): TCT_CacheField;
protected
     FItems: TObjectList;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  Add: TCT_CacheField;

     function  Count: integer;

     function  Valid: boolean;

     function  Find(AName: AxUCString): TCT_CacheField;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_CacheFields);
     procedure CopyTo(AItem: TCT_CacheFields);

     property Items[Index: integer]: TCT_CacheField read GetItems; default;
     end;

     TCT_CacheHierarchies = class(TXPGBase)
protected
     FCount: integer;
     FCacheHierarchyXpgList: TCT_CacheHierarchyXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CacheHierarchies);
     procedure CopyTo(AItem: TCT_CacheHierarchies);
     function  Create_CacheHierarchyXpgList: TCT_CacheHierarchyXpgList;

     property Count: integer read FCount write FCount;
     property CacheHierarchyXpgList: TCT_CacheHierarchyXpgList read FCacheHierarchyXpgList;
     end;

     TCT_PCDKPIs = class(TXPGBase)
protected
     FCount: integer;
     FKpiXpgList: TCT_PCDKPIXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_PCDKPIs);
     procedure CopyTo(AItem: TCT_PCDKPIs);
     function  Create_KpiXpgList: TCT_PCDKPIXpgList;

     property Count: integer read FCount write FCount;
     property KpiXpgList: TCT_PCDKPIXpgList read FKpiXpgList;
     end;

     TCT_TupleCache = class(TXPGBase)
protected
     FEntries: TCT_PCDSDTCEntries;
     FSets: TCT_Sets;
     FQueryCache: TCT_QueryCache;
     FServerFormats: TCT_ServerFormats;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_TupleCache);
     procedure CopyTo(AItem: TCT_TupleCache);
     function  Create_Entries: TCT_PCDSDTCEntries;
     function  Create_Sets: TCT_Sets;
     function  Create_QueryCache: TCT_QueryCache;
     function  Create_ServerFormats: TCT_ServerFormats;
     function  Create_ExtLst: TCT_ExtensionList;

     property Entries: TCT_PCDSDTCEntries read FEntries;
     property Sets: TCT_Sets read FSets;
     property QueryCache: TCT_QueryCache read FQueryCache;
     property ServerFormats: TCT_ServerFormats read FServerFormats;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CalculatedItems = class(TXPGBase)
protected
     FCount: integer;
     FCalculatedItemXpgList: TCT_CalculatedItemXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CalculatedItems);
     procedure CopyTo(AItem: TCT_CalculatedItems);
     function  Create_CalculatedItemXpgList: TCT_CalculatedItemXpgList;

     property Count: integer read FCount write FCount;
     property CalculatedItemXpgList: TCT_CalculatedItemXpgList read FCalculatedItemXpgList;
     end;

     TCT_CalculatedMembers = class(TXPGBase)
protected
     FCount: integer;
     FCalculatedMemberXpgList: TCT_CalculatedMemberXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_CalculatedMembers);
     procedure CopyTo(AItem: TCT_CalculatedMembers);
     function  Create_CalculatedMemberXpgList: TCT_CalculatedMemberXpgList;

     property Count: integer read FCount write FCount;
     property CalculatedMemberXpgList: TCT_CalculatedMemberXpgList read FCalculatedMemberXpgList;
     end;

     TCT_Dimensions = class(TXPGBase)
protected
     FCount: integer;
     FDimensionXpgList: TCT_PivotDimensionXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_Dimensions);
     procedure CopyTo(AItem: TCT_Dimensions);
     function  Create_DimensionXpgList: TCT_PivotDimensionXpgList;

     property Count: integer read FCount write FCount;
     property DimensionXpgList: TCT_PivotDimensionXpgList read FDimensionXpgList;
     end;

     TCT_MeasureGroups = class(TXPGBase)
protected
     FCount: integer;
     FMeasureGroupXpgList: TCT_MeasureGroupXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MeasureGroups);
     procedure CopyTo(AItem: TCT_MeasureGroups);
     function  Create_MeasureGroupXpgList: TCT_MeasureGroupXpgList;

     property Count: integer read FCount write FCount;
     property MeasureGroupXpgList: TCT_MeasureGroupXpgList read FMeasureGroupXpgList;
     end;

     TCT_MeasureDimensionMaps = class(TXPGBase)
protected
     FCount: integer;
     FMapXpgList: TCT_MeasureDimensionMapXpgList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_MeasureDimensionMaps);
     procedure CopyTo(AItem: TCT_MeasureDimensionMaps);
     function  Create_MapXpgList: TCT_MeasureDimensionMapXpgList;

     property Count: integer read FCount write FCount;
     property MapXpgList: TCT_MeasureDimensionMapXpgList read FMapXpgList;
     end;

     TCT_PivotCacheRecords = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FCount: integer;
     FRXpgList: TCT_RecordXpgList;
     FExtLst: TCT_ExtensionList;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     procedure LoadFromStream(AStream: TStream);

     procedure Assign(AItem: TCT_PivotCacheRecords);
     procedure CopyTo(AItem: TCT_PivotCacheRecords);
     function  Create_RXpgList: TCT_RecordXpgList;
     function  Create_ExtLst: TCT_ExtensionList;

     property Count: integer read FCount write FCount;
     property RXpgList: TCT_RecordXpgList read FRXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_pivotCacheDefinition = class;

     TCT_pivotTableDefinition = class(TXPGBase)
  private
    function GetColItemValue(Index: integer): TXLSSharedItemsValue;
protected
     FName: AxUCString;
     FCacheId: integer;
     FDataOnRows: boolean;
     FDataPosition: integer;
     FDataCaption: AxUCString;
     FGrandTotalCaption: AxUCString;
     FErrorCaption: AxUCString;
     FShowError: boolean;
     FMissingCaption: AxUCString;
     FShowMissing: boolean;
     FPageStyle: AxUCString;
     FPivotTableStyle: AxUCString;
     FVacatedStyle: AxUCString;
     FTag: AxUCString;
     FUpdatedVersion: integer;
     FMinRefreshableVersion: integer;
     FAsteriskTotals: boolean;
     FShowItems: boolean;
     FEditData: boolean;
     FDisableFieldList: boolean;
     FShowCalcMbrs: boolean;
     FVisualTotals: boolean;
     FShowMultipleLabel: boolean;
     FShowDataDropDown: boolean;
     FShowDrill: boolean;
     FPrintDrill: boolean;
     FShowMemberPropertyTips: boolean;
     FShowDataTips: boolean;
     FEnableWizard: boolean;
     FEnableDrill: boolean;
     FEnableFieldProperties: boolean;
     FPreserveFormatting: boolean;
     FUseAutoFormatting: boolean;
     FPageWrap: integer;
     FPageOverThenDown: boolean;
     FSubtotalHiddenItems: boolean;
     FRowGrandTotals: boolean;
     FColGrandTotals: boolean;
     FFieldPrintTitles: boolean;
     FItemPrintTitles: boolean;
     FMergeItem: boolean;
     FShowDropZones: boolean;
     FCreatedVersion: integer;
     FIndent: integer;
     FShowEmptyRow: boolean;
     FShowEmptyCol: boolean;
     FShowHeaders: boolean;
     FCompact: boolean;
     FOutline: boolean;
     FOutlineData: boolean;
     FCompactData: boolean;
     FPublished: boolean;
     FGridDropZones: boolean;
     FImmersive: boolean;
     FMultipleFieldFilters: boolean;
     FChartFormat: integer;
     FRowHeaderCaption: AxUCString;
     FColHeaderCaption: AxUCString;
     FFieldListSortAscending: boolean;
     FMdxSubqueries: boolean;
     FCustomListSort: boolean;
     FLocation: TCT_Location;
     FPivotFields: TCT_PivotFields;
     FRowFields: TCT_RowColFields;
     FRowItems: TCT_RowColItems;
     FColFields: TCT_RowColFields;
     FColItems: TCT_RowColItems;
     FPageFields: TCT_PageFields;
     FDataFields: TCT_DataFields;
     FFormats: TCT_Formats;
     FConditionalFormats: TCT_ConditionalFormats;
     FChartFormats: TCT_ChartFormats;
     FPivotHierarchies: TCT_PivotHierarchies;
     FPivotTableStyleInfo: TCT_PivotTableStyle;
     FFilters: TCT_PivotFilters;
     FRowHierarchiesUsage: TCT_RowHierarchiesUsage;
     FColHierarchiesUsage: TCT_ColHierarchiesUsage;
     FExtLst: TCT_ExtensionList;

     FRootAttributes: TStringXpgList;
     FCache: TCT_pivotCacheDefinition;
     F_OPC: TObject;

     function GetRowItemValue(Index: integer): TXLSSharedItemsValue;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Assign(AItem: TCT_pivotTableDefinition);
     procedure CopyTo(AItem: TCT_pivotTableDefinition);

     function  Create_Location: TCT_Location;
     function  Create_PivotFields: TCT_PivotFields;
     function  Create_RowFields: TCT_RowColFields;
     function  Create_RowItems: TCT_RowColItems;
     function  Create_ColFields: TCT_RowColFields;
     function  Create_ColItems: TCT_RowColItems;
     function  Create_PageFields: TCT_PageFields;
     function  Create_DataFields: TCT_DataFields;
     function  Create_Formats: TCT_Formats;
     function  Create_ConditionalFormats: TCT_ConditionalFormats;
     function  Create_ChartFormats: TCT_ChartFormats;
     function  Create_PivotHierarchies: TCT_PivotHierarchies;
     function  Create_PivotTableStyleInfo: TCT_PivotTableStyle;
     function  Create_Filters: TCT_PivotFilters;
     function  Create_RowHierarchiesUsage: TCT_RowHierarchiesUsage;
     function  Create_ColHierarchiesUsage: TCT_ColHierarchiesUsage;
     function  Create_ExtLst: TCT_ExtensionList;

     procedure Delete_ColFields;
     procedure Delete_ColItems;

     property Name: AxUCString read FName write FName;
     property CacheId: integer read FCacheId write FCacheId;
     property DataOnRows: boolean read FDataOnRows write FDataOnRows;
     property DataPosition: integer read FDataPosition write FDataPosition;
     property DataCaption: AxUCString read FDataCaption write FDataCaption;
     property GrandTotalCaption: AxUCString read FGrandTotalCaption write FGrandTotalCaption;
     property ErrorCaption: AxUCString read FErrorCaption write FErrorCaption;
     property ShowError: boolean read FShowError write FShowError;
     property MissingCaption: AxUCString read FMissingCaption write FMissingCaption;
     property ShowMissing: boolean read FShowMissing write FShowMissing;
     property PageStyle: AxUCString read FPageStyle write FPageStyle;
     property PivotTableStyle: AxUCString read FPivotTableStyle write FPivotTableStyle;
     property VacatedStyle: AxUCString read FVacatedStyle write FVacatedStyle;
     property Tag: AxUCString read FTag write FTag;
     property UpdatedVersion: integer read FUpdatedVersion write FUpdatedVersion;
     property MinRefreshableVersion: integer read FMinRefreshableVersion write FMinRefreshableVersion;
     property AsteriskTotals: boolean read FAsteriskTotals write FAsteriskTotals;
     property ShowItems: boolean read FShowItems write FShowItems;
     property EditData: boolean read FEditData write FEditData;
     property DisableFieldList: boolean read FDisableFieldList write FDisableFieldList;
     property ShowCalcMbrs: boolean read FShowCalcMbrs write FShowCalcMbrs;
     property VisualTotals: boolean read FVisualTotals write FVisualTotals;
     property ShowMultipleLabel: boolean read FShowMultipleLabel write FShowMultipleLabel;
     property ShowDataDropDown: boolean read FShowDataDropDown write FShowDataDropDown;
     property ShowDrill: boolean read FShowDrill write FShowDrill;
     property PrintDrill: boolean read FPrintDrill write FPrintDrill;
     property ShowMemberPropertyTips: boolean read FShowMemberPropertyTips write FShowMemberPropertyTips;
     property ShowDataTips: boolean read FShowDataTips write FShowDataTips;
     property EnableWizard: boolean read FEnableWizard write FEnableWizard;
     property EnableDrill: boolean read FEnableDrill write FEnableDrill;
     property EnableFieldProperties: boolean read FEnableFieldProperties write FEnableFieldProperties;
     property PreserveFormatting: boolean read FPreserveFormatting write FPreserveFormatting;
     property UseAutoFormatting: boolean read FUseAutoFormatting write FUseAutoFormatting;
     property PageWrap: integer read FPageWrap write FPageWrap;
     property PageOverThenDown: boolean read FPageOverThenDown write FPageOverThenDown;
     property SubtotalHiddenItems: boolean read FSubtotalHiddenItems write FSubtotalHiddenItems;
     property RowGrandTotals: boolean read FRowGrandTotals write FRowGrandTotals;
     property ColGrandTotals: boolean read FColGrandTotals write FColGrandTotals;
     property FieldPrintTitles: boolean read FFieldPrintTitles write FFieldPrintTitles;
     property ItemPrintTitles: boolean read FItemPrintTitles write FItemPrintTitles;
     property MergeItem: boolean read FMergeItem write FMergeItem;
     property ShowDropZones: boolean read FShowDropZones write FShowDropZones;
     property CreatedVersion: integer read FCreatedVersion write FCreatedVersion;
     property Indent: integer read FIndent write FIndent;
     property ShowEmptyRow: boolean read FShowEmptyRow write FShowEmptyRow;
     property ShowEmptyCol: boolean read FShowEmptyCol write FShowEmptyCol;
     property ShowHeaders: boolean read FShowHeaders write FShowHeaders;
     property Compact: boolean read FCompact write FCompact;
     property Outline: boolean read FOutline write FOutline;
     property OutlineData: boolean read FOutlineData write FOutlineData;
     property CompactData: boolean read FCompactData write FCompactData;
     property Published_: boolean read FPublished write FPublished;
     property GridDropZones: boolean read FGridDropZones write FGridDropZones;
     property Immersive: boolean read FImmersive write FImmersive;
     property MultipleFieldFilters: boolean read FMultipleFieldFilters write FMultipleFieldFilters;
     property ChartFormat: integer read FChartFormat write FChartFormat;
     property RowHeaderCaption: AxUCString read FRowHeaderCaption write FRowHeaderCaption;
     property ColHeaderCaption: AxUCString read FColHeaderCaption write FColHeaderCaption;
     property FieldListSortAscending: boolean read FFieldListSortAscending write FFieldListSortAscending;
     property MdxSubqueries: boolean read FMdxSubqueries write FMdxSubqueries;
     property CustomListSort: boolean read FCustomListSort write FCustomListSort;
     property Location: TCT_Location read FLocation;
     property PivotFields: TCT_PivotFields read FPivotFields;
     property RowFields: TCT_RowColFields read FRowFields;
     property RowItems: TCT_RowColItems read FRowItems;
     property ColFields: TCT_RowColFields read FColFields;
     property ColItems: TCT_RowColItems read FColItems;
     property PageFields: TCT_PageFields read FPageFields;
     property DataFields: TCT_DataFields read FDataFields;
     property Formats: TCT_Formats read FFormats;
     property ConditionalFormats: TCT_ConditionalFormats read FConditionalFormats;
     property ChartFormats: TCT_ChartFormats read FChartFormats;
     property PivotHierarchies: TCT_PivotHierarchies read FPivotHierarchies;
     property PivotTableStyleInfo: TCT_PivotTableStyle read FPivotTableStyleInfo;
     property Filters: TCT_PivotFilters read FFilters;
     property RowHierarchiesUsage: TCT_RowHierarchiesUsage read FRowHierarchiesUsage;
     property ColHierarchiesUsage: TCT_ColHierarchiesUsage read FColHierarchiesUsage;
     property ExtLst: TCT_ExtensionList read FExtLst;

     property Cache: TCT_pivotCacheDefinition read FCache write FCache;
     property _OPC: TObject read F_OPC write F_OPC;
     property RowItemValue[Index: integer]: TXLSSharedItemsValue read GetRowItemValue;
     property ColItemValue[Index: integer]: TXLSSharedItemsValue read GetColItemValue;
     end;


     TCT_pivotTableDefinitions = class(TXPGBase)
private
     function GetItems(Index: integer): TCT_pivotTableDefinition;
protected
     FOwner         : TXPGDocBase;
     FItems         : TObjectList;
public
     constructor Create;
     destructor Destroy; override;

     function  LoadFromStream(AStream: TStream): TCT_pivotTableDefinition;
     procedure SaveToStream(AStream: TStream; AIndex: integer; AOPC: TObject);

     function  Add: TCT_pivotTableDefinition;
     function  Count: integer;

     function  FindByCache(ACache: TCT_pivotCacheDefinition): TCT_pivotTableDefinition;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Clear;

     property Items[Index: integer]: TCT_pivotTableDefinition read GetItems; default;
     end;

     TCT_PivotCacheDefinition = class(TXPGBase)
protected
     FR_Id: AxUCString;
     FInvalid: boolean;
     FSaveData: boolean;
     FRefreshOnLoad: boolean;
     FOptimizeMemory: boolean;
     FEnableRefresh: boolean;
     FRefreshedBy: AxUCString;
     FRefreshedDate: double;
     FBackgroundQuery: boolean;
     FMissingItemsLimit: integer;
     FCreatedVersion: integer;
     FRefreshedVersion: integer;
     FMinRefreshableVersion: integer;
     FRecordCount: integer;
     FUpgradeOnRefresh: boolean;
     FTupleCache: boolean;
     FSupportSubquery: boolean;
     FSupportAdvancedDrill: boolean;
     FCacheSource: TCT_CacheSource;
     FCacheFields: TCT_CacheFields;
     FCacheHierarchies: TCT_CacheHierarchies;
     FKpis: TCT_PCDKPIs;
// __TODO__ Field FTupleCache renamed to FTupleCache_Dup
     FTupleCache_Dup: TCT_TupleCache;
     FCalculatedItems: TCT_CalculatedItems;
     FCalculatedMembers: TCT_CalculatedMembers;
     FDimensions: TCT_Dimensions;
     FMeasureGroups: TCT_MeasureGroups;
     FMaps: TCT_MeasureDimensionMaps;
     FExtLst: TCT_ExtensionList;

     FRootAttributes: TStringXpgList;
     FUsageCount: integer;
     FCacheId: integer;
     FRecords: TCT_PivotCacheRecords;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     function  AddFields(ARef: TXLSRelCells): boolean;

     procedure CacheValues(ACol: integer); overload;
     procedure CacheValues; overload;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;

     procedure Use;
     procedure Release;

     procedure Assign(AItem: TCT_PivotCacheDefinition);
     procedure CopyTo(AItem: TCT_PivotCacheDefinition);

     function  Create_CacheSource: TCT_CacheSource;
     function  Create_CacheFields: TCT_CacheFields;
     function  Create_CacheHierarchies: TCT_CacheHierarchies;
     function  Create_Kpis: TCT_PCDKPIs;
     function  Create_TupleCache_Dup: TCT_TupleCache;
     function  Create_CalculatedItems: TCT_CalculatedItems;
     function  Create_CalculatedMembers: TCT_CalculatedMembers;
     function  Create_Dimensions: TCT_Dimensions;
     function  Create_MeasureGroups: TCT_MeasureGroups;
     function  Create_Maps: TCT_MeasureDimensionMaps;
     function  Create_ExtLst: TCT_ExtensionList;

     property R_Id: AxUCString read FR_Id write FR_Id;
     property Invalid: boolean read FInvalid write FInvalid;
     property SaveData: boolean read FSaveData write FSaveData;
     property RefreshOnLoad: boolean read FRefreshOnLoad write FRefreshOnLoad;
     property OptimizeMemory: boolean read FOptimizeMemory write FOptimizeMemory;
     property EnableRefresh: boolean read FEnableRefresh write FEnableRefresh;
     property RefreshedBy: AxUCString read FRefreshedBy write FRefreshedBy;
     property RefreshedDate: double read FRefreshedDate write FRefreshedDate;
     property BackgroundQuery: boolean read FBackgroundQuery write FBackgroundQuery;
     property MissingItemsLimit: integer read FMissingItemsLimit write FMissingItemsLimit;
     property CreatedVersion: integer read FCreatedVersion write FCreatedVersion;
     property RefreshedVersion: integer read FRefreshedVersion write FRefreshedVersion;
     property MinRefreshableVersion: integer read FMinRefreshableVersion write FMinRefreshableVersion;
     property RecordCount: integer read FRecordCount write FRecordCount;
     property UpgradeOnRefresh: boolean read FUpgradeOnRefresh write FUpgradeOnRefresh;
     property TupleCache: boolean read FTupleCache write FTupleCache;
     property SupportSubquery: boolean read FSupportSubquery write FSupportSubquery;
     property SupportAdvancedDrill: boolean read FSupportAdvancedDrill write FSupportAdvancedDrill;
     property CacheSource: TCT_CacheSource read FCacheSource;
     property CacheFields: TCT_CacheFields read FCacheFields;
     property CacheHierarchies: TCT_CacheHierarchies read FCacheHierarchies;
     property Kpis: TCT_PCDKPIs read FKpis;
     property TupleCache_Dup: TCT_TupleCache read FTupleCache_Dup;
     property CalculatedItems: TCT_CalculatedItems read FCalculatedItems;
     property CalculatedMembers: TCT_CalculatedMembers read FCalculatedMembers;
     property Dimensions: TCT_Dimensions read FDimensions;
     property MeasureGroups: TCT_MeasureGroups read FMeasureGroups;
     property Maps: TCT_MeasureDimensionMaps read FMaps;
     property ExtLst: TCT_ExtensionList read FExtLst;

     property CacheId: integer read FCacheId write FCacheId;
     property Records: TCT_PivotCacheRecords read FRecords;
     end;

     TCT_PivotCacheDefinitions = class(TXPGBase)
private
     function GetItems(Index: integer): TCT_pivotCacheDefinition;
protected
     FOwner         : TXPGDocBase;
     FItems         : TObjectList;
public
     constructor Create;
     destructor Destroy; override;

     function  LoadFromStream(AStream: TStream): TCT_pivotCacheDefinition;
     procedure SaveToStream(AStream: TStream; AIndex: integer);
     procedure RecordsSaveToStream(AStream: TStream; AIndex: integer);

     function  Add: TCT_pivotCacheDefinition;
     function  Count: integer;

     function  Find(const ACacheId: integer): TCT_pivotCacheDefinition; overload;

     procedure Enumerate;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure Clear;

     property Items[Index: integer]: TCT_pivotCacheDefinition read GetItems; default;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FCurrWriteClass: TClass;
     FPivotCacheRecords: TCT_PivotCacheRecords;
     FPivotTableDefinition: TCT_pivotTableDefinition;
     FPivotCacheDefinition: TCT_PivotCacheDefinition;

public
     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: T__ROOT__);
     procedure CopyTo(AItem: T__ROOT__);
     function  Create_PivotCacheRecords: TCT_PivotCacheRecords;
     function  Create_PivotTableDefinition: TCT_pivotTableDefinition;
     function  Create_PivotCacheDefinition: TCT_PivotCacheDefinition;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property PivotCacheRecords: TCT_PivotCacheRecords read FPivotCacheRecords;
     property PivotTableDefinition: TCT_pivotTableDefinition read FPivotTableDefinition;
     property PivotCacheDefinition: TCT_PivotCacheDefinition read FPivotCacheDefinition;
     end;

     TCT_XStringElement = class(TXPGBase)
protected
     FV: AxUCString;

public
     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;
     procedure Assign(AItem: TCT_XStringElement);
     procedure CopyTo(AItem: TCT_XStringElement);

     property V: AxUCString read FV write FV;
     end;

     TXPGPivotTable = class(TXPGDocBase)
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;

     function  GetPivotCacheRecords: TCT_PivotCacheRecords;
     function  GetPivotTableDefinition: TCT_pivotTableDefinition;
     function  GetPivotCacheDefinition: TCT_PivotCacheDefinition;
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString; AClassToWrite: TClass);

     procedure SaveToStream(AStream: TStream);

     property Root: T__ROOT__ read FRoot;
     property PivotCacheRecords: TCT_PivotCacheRecords read GetPivotCacheRecords;
     property PivotTableDefinition: TCT_pivotTableDefinition read GetPivotTableDefinition;
     property PivotCacheDefinition: TCT_PivotCacheDefinition read GetPivotCacheDefinition;
     end;


implementation

{$ifdef DELPHI_5}
var L_Enums: TStringList;
{$else}
var L_Enums: THashedStringList;
{$endif}

procedure AddEnums;
begin
  L_Enums.AddObject('staAxisRow',TObject(0));
  L_Enums.AddObject('staAxisCol',TObject(1));
  L_Enums.AddObject('staAxisPage',TObject(2));
  L_Enums.AddObject('staAxisValues',TObject(3));
  L_Enums.AddObject('stpatNone',TObject(0));
  L_Enums.AddObject('stpatNormal',TObject(1));
  L_Enums.AddObject('stpatData',TObject(2));
  L_Enums.AddObject('stpatAll',TObject(3));
  L_Enums.AddObject('stpatOrigin',TObject(4));
  L_Enums.AddObject('stpatButton',TObject(5));
  L_Enums.AddObject('stpatTopRight',TObject(6));
  L_Enums.AddObject('sttNone',TObject(0));
  L_Enums.AddObject('sttAll',TObject(1));
  L_Enums.AddObject('sttRow',TObject(2));
  L_Enums.AddObject('sttColumn',TObject(3));
  L_Enums.AddObject('stfstManual',TObject(0));
  L_Enums.AddObject('stfstAscending',TObject(1));
  L_Enums.AddObject('stfstDescending',TObject(2));
  L_Enums.AddObject('stgbRange',TObject(0));
  L_Enums.AddObject('stgbSeconds',TObject(1));
  L_Enums.AddObject('stgbMinutes',TObject(2));
  L_Enums.AddObject('stgbHours',TObject(3));
  L_Enums.AddObject('stgbDays',TObject(4));
  L_Enums.AddObject('stgbMonths',TObject(5));
  L_Enums.AddObject('stgbQuarters',TObject(6));
  L_Enums.AddObject('stgbYears',TObject(7));
  L_Enums.AddObject('ststNone',TObject(0));
  L_Enums.AddObject('ststAscending',TObject(1));
  L_Enums.AddObject('ststDescending',TObject(2));
  L_Enums.AddObject('ststAscendingAlpha',TObject(3));
  L_Enums.AddObject('ststDescendingAlpha',TObject(4));
  L_Enums.AddObject('ststAscendingNatural',TObject(5));
  L_Enums.AddObject('ststDescendingNatural',TObject(6));
  L_Enums.AddObject('stsdaNormal',TObject(0));
  L_Enums.AddObject('stsdaDifference',TObject(1));
  L_Enums.AddObject('stsdaPercent',TObject(2));
  L_Enums.AddObject('stsdaPercentDiff',TObject(3));
  L_Enums.AddObject('stsdaRunTotal',TObject(4));
  L_Enums.AddObject('stsdaPercentOfRow',TObject(5));
  L_Enums.AddObject('stsdaPercentOfCol',TObject(6));
  L_Enums.AddObject('stsdaPercentOfTotal',TObject(7));
  L_Enums.AddObject('stsdaIndex',TObject(8));
  L_Enums.AddObject('stitData',TObject(0));
  L_Enums.AddObject('stitDefault',TObject(1));
  L_Enums.AddObject('stitSum',TObject(2));
  L_Enums.AddObject('stitCountA',TObject(3));
  L_Enums.AddObject('stitAvg',TObject(4));
  L_Enums.AddObject('stitMax',TObject(5));
  L_Enums.AddObject('stitMin',TObject(6));
  L_Enums.AddObject('stitProduct',TObject(7));
  L_Enums.AddObject('stitCount',TObject(8));
  L_Enums.AddObject('stitStdDev',TObject(9));
  L_Enums.AddObject('stitStdDevP',TObject(10));
  L_Enums.AddObject('stitVar',TObject(11));
  L_Enums.AddObject('stitVarP',TObject(12));
  L_Enums.AddObject('stitGrand',TObject(13));
  L_Enums.AddObject('stitBlank',TObject(14));
  L_Enums.AddObject('stsSelection',TObject(0));
  L_Enums.AddObject('stsData',TObject(1));
  L_Enums.AddObject('stsField',TObject(2));
  L_Enums.AddObject('stpftUnknown',TObject(0));
  L_Enums.AddObject('stpftCount',TObject(1));
  L_Enums.AddObject('stpftPercent',TObject(2));
  L_Enums.AddObject('stpftSum',TObject(3));
  L_Enums.AddObject('stpftCaptionEqual',TObject(4));
  L_Enums.AddObject('stpftCaptionNotEqual',TObject(5));
  L_Enums.AddObject('stpftCaptionBeginsWith',TObject(6));
  L_Enums.AddObject('stpftCaptionNotBeginsWith',TObject(7));
  L_Enums.AddObject('stpftCaptionEndsWith',TObject(8));
  L_Enums.AddObject('stpftCaptionNotEndsWith',TObject(9));
  L_Enums.AddObject('stpftCaptionContains',TObject(10));
  L_Enums.AddObject('stpftCaptionNotContains',TObject(11));
  L_Enums.AddObject('stpftCaptionGreaterThan',TObject(12));
  L_Enums.AddObject('stpftCaptionGreaterThanOrEqual',TObject(13));
  L_Enums.AddObject('stpftCaptionLessThan',TObject(14));
  L_Enums.AddObject('stpftCaptionLessThanOrEqual',TObject(15));
  L_Enums.AddObject('stpftCaptionBetween',TObject(16));
  L_Enums.AddObject('stpftCaptionNotBetween',TObject(17));
  L_Enums.AddObject('stpftValueEqual',TObject(18));
  L_Enums.AddObject('stpftValueNotEqual',TObject(19));
  L_Enums.AddObject('stpftValueGreaterThan',TObject(20));
  L_Enums.AddObject('stpftValueGreaterThanOrEqual',TObject(21));
  L_Enums.AddObject('stpftValueLessThan',TObject(22));
  L_Enums.AddObject('stpftValueLessThanOrEqual',TObject(23));
  L_Enums.AddObject('stpftValueBetween',TObject(24));
  L_Enums.AddObject('stpftValueNotBetween',TObject(25));
  L_Enums.AddObject('stpftDateEqual',TObject(26));
  L_Enums.AddObject('stpftDateNotEqual',TObject(27));
  L_Enums.AddObject('stpftDateOlderThan',TObject(28));
  L_Enums.AddObject('stpftDateOlderThanOrEqual',TObject(29));
  L_Enums.AddObject('stpftDateNewerThan',TObject(30));
  L_Enums.AddObject('stpftDateNewerThanOrEqual',TObject(31));
  L_Enums.AddObject('stpftDateBetween',TObject(32));
  L_Enums.AddObject('stpftDateNotBetween',TObject(33));
  L_Enums.AddObject('stpftTomorrow',TObject(34));
  L_Enums.AddObject('stpftToday',TObject(35));
  L_Enums.AddObject('stpftYesterday',TObject(36));
  L_Enums.AddObject('stpftNextWeek',TObject(37));
  L_Enums.AddObject('stpftThisWeek',TObject(38));
  L_Enums.AddObject('stpftLastWeek',TObject(39));
  L_Enums.AddObject('stpftNextMonth',TObject(40));
  L_Enums.AddObject('stpftThisMonth',TObject(41));
  L_Enums.AddObject('stpftLastMonth',TObject(42));
  L_Enums.AddObject('stpftNextQuarter',TObject(43));
  L_Enums.AddObject('stpftThisQuarter',TObject(44));
  L_Enums.AddObject('stpftLastQuarter',TObject(45));
  L_Enums.AddObject('stpftNextYear',TObject(46));
  L_Enums.AddObject('stpftThisYear',TObject(47));
  L_Enums.AddObject('stpftLastYear',TObject(48));
  L_Enums.AddObject('stpftYearToDate',TObject(49));
  L_Enums.AddObject('stpftQ1',TObject(50));
  L_Enums.AddObject('stpftQ2',TObject(51));
  L_Enums.AddObject('stpftQ3',TObject(52));
  L_Enums.AddObject('stpftQ4',TObject(53));
  L_Enums.AddObject('stpftM1',TObject(54));
  L_Enums.AddObject('stpftM2',TObject(55));
  L_Enums.AddObject('stpftM3',TObject(56));
  L_Enums.AddObject('stpftM4',TObject(57));
  L_Enums.AddObject('stpftM5',TObject(58));
  L_Enums.AddObject('stpftM6',TObject(59));
  L_Enums.AddObject('stpftM7',TObject(60));
  L_Enums.AddObject('stpftM8',TObject(61));
  L_Enums.AddObject('stpftM9',TObject(62));
  L_Enums.AddObject('stpftM10',TObject(63));
  L_Enums.AddObject('stpftM11',TObject(64));
  L_Enums.AddObject('stpftM12',TObject(65));
  L_Enums.AddObject('ststWorksheet',TObject(0));
  L_Enums.AddObject('ststExternal',TObject(1));
  L_Enums.AddObject('ststConsolidation',TObject(2));
  L_Enums.AddObject('ststScenario',TObject(3));
  L_Enums.AddObject('stfaBlank',TObject(0));
  L_Enums.AddObject('stfaFormatting',TObject(1));
  L_Enums.AddObject('stfaDrill',TObject(2));
  L_Enums.AddObject('stfaFormula',TObject(3));
end;

{ TXPGBase }

function  TXPGBase.CheckAssigned: integer;
begin
  Result := 1;
end;

function  TXPGBase.Assigned: boolean;
begin
  Result := FAssigneds <> [];
end;

function  TXPGBase.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
end;

procedure TXPGBase.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
end;

procedure TXPGBase.AfterTag;
begin
end;

class function  TXPGBase.StrToEnum(AValue: AxUCString): integer;
var
  i: integer;
begin
  i := L_Enums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(L_Enums.Objects[i])
  else
    Result := 0;
end;

class function  TXPGBase.StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
var
  i: integer;
begin
  i := L_Enums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(L_Enums.Objects[i])
  else
    Result := ADefault;
end;

class function  TXPGBase.TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
var
  i: integer;
begin
  i := L_Enums.IndexOf(AValue);
  if i >= 0 then 
  begin
    i := Integer(L_Enums.Objects[i]);
    Result := (i <= High(AEnumNames)) and (AText = AEnumNames[i]);
    if Result then 
      APtrInt^ := i;
  end
  else 
    Result := False;
end;

function  TXPGBase.Available: boolean;
begin
  Result := xaRead in FAssigneds;
end;

{ TXPGAnyElement }

procedure TXPGAnyElement.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  FAttributes.Assign(AAttributes.FAttributes,AAttributes.Count);
end;

constructor TXPGAnyElement.Create;
begin
  FAttributes := TXpgXMLAttributeList.Create;
end;

destructor TXPGAnyElement.Destroy;
begin
  FAttributes.Free;
end;

{ TXPGAnyElements }

function  TXPGAnyElements.GetItems(Index: integer): TXPGAnyElement;
begin
  Result := TXPGAnyElement(inherited Items[Index]);
end;

function  TXPGAnyElements.Add(AElementName: AxUCString; AContent: AxUCString): TXPGAnyElement;
begin
  Result := TXPGAnyElement.Create;
  Result.FElementName := AElementName;
  Result.FContent := AContent;
  inherited Add(Result);
end;

procedure TXPGAnyElements.Write(AWriter: TXpgWriteXML);
var
  i: integer;
  j: integer;
  Elem: TXPGAnyElement;
begin
  for i := 0 to Count - 1 do 
  begin
    Elem := GetItems(i);
    for j := 0 to Elem.Attributes.Count - 1 do 
      AWriter.AddAttribute(Elem.Attributes.Attributes[j],Elem.Attributes.Values[j]);
    AWriter.Text := Elem.Content;
    AWriter.SimpleTag(Elem.ElementName);
  end
end;

{ TXPGBaseObjectList }

function  TXPGBaseObjectList.GetItems(Index: integer): TXPGBase;
begin
  Result := TXPGBase(inherited Items[Index]);
end;

constructor TXPGBaseObjectList.Create(AOwner: TXPGDocBase);
begin
  inherited Create;
  FOwner := AOwner;
end;

{ TXPGReader }

constructor TXPGReader.Create(AErrors: TXpgPErrors; ARoot: TXPGBase);
begin
  inherited Create;
  FErrors := AErrors;
  FErrors.NoDuplicates := True;
  FCurrent := ARoot;
  FStack := TObjectStack.Create;
end;

destructor TXPGReader.Destroy;
begin
  FStack.Free;
  inherited Destroy;
end;

procedure TXPGReader.BeginTag;
begin
  FStack.Push(FCurrent);
  FCurrent := FCurrent.HandleElement(Self);
  if Attributes.Count > 0 then
    FCurrent.AssignAttributes(Attributes);
end;

procedure TXPGReader.EndTag;
begin
  FCurrent := TXPGBase(FStack.Pop);
  FCurrent.AfterTag;
end;

{ TCT_Extension }

function  TCT_Extension.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FUri <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FAnyElements.Count);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Extension.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FAnyElements.Add(AReader.QName,AReader.Text);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Extension.Write(AWriter: TXpgWriteXML);
begin
  FAnyElements.Write(AWriter);
end;

procedure TCT_Extension.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FUri <> '' then 
    AWriter.AddAttribute('uri',FUri);
end;

procedure TCT_Extension.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'uri' then 
    FUri := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Extension.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FAnyElements := TXPGAnyElements.Create;
  FUri := '';
end;

destructor TCT_Extension.Destroy;
begin
  FAnyElements.Free;
end;

procedure TCT_Extension.Clear;
begin
  FAssigneds := [];
  FAnyElements.Clear;
  FUri := '';
end;

procedure TCT_Extension.Assign(AItem: TCT_Extension);
begin
end;

procedure TCT_Extension.CopyTo(AItem: TCT_Extension);
begin
end;

{ TEG_ExtensionList }

function  TEG_ExtensionList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FExt <> Nil then 
    Inc(ElemsAssigned,FExt.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TEG_ExtensionList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'ext' then 
  begin
    if FExt = Nil then 
      FExt := TCT_Extension.Create(FOwner);
    Result := FExt;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TEG_ExtensionList.Write(AWriter: TXpgWriteXML);
begin
  if (FExt <> Nil) and FExt.Assigned then 
  begin
    FExt.WriteAttributes(AWriter);
    if xaElements in FExt.FAssigneds then 
    begin
      AWriter.BeginTag('ext');
      FExt.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('ext');
  end;
end;

constructor TEG_ExtensionList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TEG_ExtensionList.Destroy;
begin
  if FExt <> Nil then 
    FExt.Free;
end;

procedure TEG_ExtensionList.Clear;
begin
  FAssigneds := [];
  if FExt <> Nil then 
    FreeAndNil(FExt);
end;

procedure TEG_ExtensionList.Assign(AItem: TEG_ExtensionList);
begin
end;

procedure TEG_ExtensionList.CopyTo(AItem: TEG_ExtensionList);
begin
end;

function  TEG_ExtensionList.Create_Ext: TCT_Extension;
begin
  if FExt = Nil then
    FExt := TCT_Extension.Create(FOwner);
  Result := FExt;
end;

{ TCT_GroupMember }

function  TCT_GroupMember.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_GroupMember.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_GroupMember.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('uniqueName',FUniqueName);
  if FGroup <> False then 
    AWriter.AddAttribute('group',XmlBoolToStr(FGroup));
end;

procedure TCT_GroupMember.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000418: FUniqueName := AAttributes.Values[i];
      $0000022D: FGroup := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GroupMember.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FGroup := False;
end;

destructor TCT_GroupMember.Destroy;
begin
end;

procedure TCT_GroupMember.Clear;
begin
  FAssigneds := [];
  FUniqueName := '';
  FGroup := False;
end;

procedure TCT_GroupMember.Assign(AItem: TCT_GroupMember);
begin
end;

procedure TCT_GroupMember.CopyTo(AItem: TCT_GroupMember);
begin
end;

{ TCT_GroupMemberXpgList }

function  TCT_GroupMemberXpgList.GetItems(Index: integer): TCT_GroupMember;
begin
  Result := TCT_GroupMember(inherited Items[Index]);
end;

function  TCT_GroupMemberXpgList.Add: TCT_GroupMember;
begin
  Result := TCT_GroupMember.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GroupMemberXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GroupMemberXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_GroupMemberXpgList.Assign(AItem: TCT_GroupMemberXpgList);
begin
end;

procedure TCT_GroupMemberXpgList.CopyTo(AItem: TCT_GroupMemberXpgList);
begin
end;

{ TCT_Index }

function  TCT_Index.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Index.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Index.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',XmlIntToStr(FV));
end;

procedure TCT_Index.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'v' then 
    FV := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Index.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FV := 2147483632;
end;

destructor TCT_Index.Destroy;
begin
end;

procedure TCT_Index.Clear;
begin
  FAssigneds := [];
  FV := 2147483632;
end;

procedure TCT_Index.Assign(AItem: TCT_Index);
begin
end;

procedure TCT_Index.CopyTo(AItem: TCT_Index);
begin
end;

{ TCT_IndexXpgList }

function  TCT_IndexXpgList.GetItems(Index: integer): TCT_Index;
begin
  Result := TCT_Index(inherited Items[Index]);
end;

function  TCT_IndexXpgList.Add: TCT_Index;
begin
  Result := TCT_Index.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_IndexXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_IndexXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_IndexXpgList.Assign(AItem: TCT_IndexXpgList);
begin
end;

procedure TCT_IndexXpgList.CopyTo(AItem: TCT_IndexXpgList);
begin
end;

{ TCT_ExtensionList }

function  TCT_ExtensionList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FEG_ExtensionList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExtensionList.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FEG_ExtensionList.HandleElement(AReader);
  if Result = Nil then 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ExtensionList.Write(AWriter: TXpgWriteXML);
begin
  FEG_ExtensionList.Write(AWriter);
end;

constructor TCT_ExtensionList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FEG_ExtensionList := TEG_ExtensionList.Create(FOwner);
end;

destructor TCT_ExtensionList.Destroy;
begin
  FEG_ExtensionList.Free;
end;

procedure TCT_ExtensionList.Clear;
begin
  FAssigneds := [];
  FEG_ExtensionList.Clear;
end;

procedure TCT_ExtensionList.Assign(AItem: TCT_ExtensionList);
begin
end;

procedure TCT_ExtensionList.CopyTo(AItem: TCT_ExtensionList);
begin
end;

{ TCT_Tuple }

function  TCT_Tuple.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Tuple.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Tuple.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FFld <> 2147483632 then 
    AWriter.AddAttribute('fld',XmlIntToStr(FFld));
  if FHier <> 2147483632 then 
    AWriter.AddAttribute('hier',XmlIntToStr(FHier));
  AWriter.AddAttribute('item',XmlIntToStr(FItem));
end;

procedure TCT_Tuple.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000136: FFld := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A8: FHier := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001AF: FItem := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Tuple.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FFld := 2147483632;
  FHier := 2147483632;
  FItem := 2147483632;
end;

destructor TCT_Tuple.Destroy;
begin
end;

procedure TCT_Tuple.Clear;
begin
  FAssigneds := [];
  FFld := 2147483632;
  FHier := 2147483632;
  FItem := 2147483632;
end;

procedure TCT_Tuple.Assign(AItem: TCT_Tuple);
begin
end;

procedure TCT_Tuple.CopyTo(AItem: TCT_Tuple);
begin
end;

{ TCT_TupleXpgList }

function  TCT_TupleXpgList.GetItems(Index: integer): TCT_Tuple;
begin
  Result := TCT_Tuple(inherited Items[Index]);
end;

function  TCT_TupleXpgList.Add: TCT_Tuple;
begin
  Result := TCT_Tuple.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TupleXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TupleXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_TupleXpgList.Assign(AItem: TCT_TupleXpgList);
begin
end;

procedure TCT_TupleXpgList.CopyTo(AItem: TCT_TupleXpgList);
begin
end;

{ TCT_GroupMembers }

function  TCT_GroupMembers.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FGroupMemberXpgList <> Nil then 
    Inc(ElemsAssigned,FGroupMemberXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GroupMembers.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'groupMember' then 
  begin
    if FGroupMemberXpgList = Nil then 
      FGroupMemberXpgList := TCT_GroupMemberXpgList.Create(FOwner);
    Result := FGroupMemberXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupMembers.Write(AWriter: TXpgWriteXML);
begin
  if FGroupMemberXpgList <> Nil then 
    FGroupMemberXpgList.Write(AWriter,'groupMember');
end;

procedure TCT_GroupMembers.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_GroupMembers.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GroupMembers.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_GroupMembers.Destroy;
begin
  if FGroupMemberXpgList <> Nil then 
    FGroupMemberXpgList.Free;
end;

procedure TCT_GroupMembers.Clear;
begin
  FAssigneds := [];
  if FGroupMemberXpgList <> Nil then 
    FreeAndNil(FGroupMemberXpgList);
  FCount := 2147483632;
end;

procedure TCT_GroupMembers.Assign(AItem: TCT_GroupMembers);
begin
end;

procedure TCT_GroupMembers.CopyTo(AItem: TCT_GroupMembers);
begin
end;

function  TCT_GroupMembers.Create_GroupMemberXpgList: TCT_GroupMemberXpgList;
begin
  if FGroupMemberXpgList = Nil then
      FGroupMemberXpgList := TCT_GroupMemberXpgList.Create(FOwner);
  Result := FGroupMemberXpgList ;
end;

{ TCT_PivotAreaReference }

function  TCT_PivotAreaReference.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FField <> '' then
    Inc(AttrsAssigned);
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FSelected <> True then 
    Inc(AttrsAssigned);
  if FByPosition <> False then 
    Inc(AttrsAssigned);
  if FRelative <> False then 
    Inc(AttrsAssigned);
  if FDefaultSubtotal <> False then 
    Inc(AttrsAssigned);
  if FSumSubtotal <> False then 
    Inc(AttrsAssigned);
  if FCountASubtotal <> False then
    Inc(AttrsAssigned);
  if FAvgSubtotal <> False then 
    Inc(AttrsAssigned);
  if FMaxSubtotal <> False then 
    Inc(AttrsAssigned);
  if FMinSubtotal <> False then 
    Inc(AttrsAssigned);
  if FProductSubtotal <> False then 
    Inc(AttrsAssigned);
  if FCountSubtotal <> False then 
    Inc(AttrsAssigned);
  if FStdDevSubtotal <> False then 
    Inc(AttrsAssigned);
  if FStdDevPSubtotal <> False then 
    Inc(AttrsAssigned);
  if FVarSubtotal <> False then 
    Inc(AttrsAssigned);
  if FVarPSubtotal <> False then 
    Inc(AttrsAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotAreaReference.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000078: begin
      if FXXpgList = Nil then 
        FXXpgList := TCT_IndexXpgList.Create(FOwner);
      Result := FXXpgList.Add;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotAreaReference.Write(AWriter: TXpgWriteXML);
begin
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
  if (FExtLst <> Nil) and FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotAreaReference.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FField <> '' then
    AWriter.AddAttribute('field',FField);
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if FSelected <> True then 
    AWriter.AddAttribute('selected',XmlBoolToStr(FSelected));
  if FByPosition <> False then 
    AWriter.AddAttribute('byPosition',XmlBoolToStr(FByPosition));
  if FRelative <> False then 
    AWriter.AddAttribute('relative',XmlBoolToStr(FRelative));
  if FDefaultSubtotal <> False then 
    AWriter.AddAttribute('defaultSubtotal',XmlBoolToStr(FDefaultSubtotal));
  if FSumSubtotal <> False then 
    AWriter.AddAttribute('sumSubtotal',XmlBoolToStr(FSumSubtotal));
  if FCountASubtotal <> False then 
    AWriter.AddAttribute('countASubtotal',XmlBoolToStr(FCountASubtotal));
  if FAvgSubtotal <> False then 
    AWriter.AddAttribute('avgSubtotal',XmlBoolToStr(FAvgSubtotal));
  if FMaxSubtotal <> False then 
    AWriter.AddAttribute('maxSubtotal',XmlBoolToStr(FMaxSubtotal));
  if FMinSubtotal <> False then 
    AWriter.AddAttribute('minSubtotal',XmlBoolToStr(FMinSubtotal));
  if FProductSubtotal <> False then 
    AWriter.AddAttribute('productSubtotal',XmlBoolToStr(FProductSubtotal));
  if FCountSubtotal <> False then 
    AWriter.AddAttribute('countSubtotal',XmlBoolToStr(FCountSubtotal));
  if FStdDevSubtotal <> False then 
    AWriter.AddAttribute('stdDevSubtotal',XmlBoolToStr(FStdDevSubtotal));
  if FStdDevPSubtotal <> False then 
    AWriter.AddAttribute('stdDevPSubtotal',XmlBoolToStr(FStdDevPSubtotal));
  if FVarSubtotal <> False then 
    AWriter.AddAttribute('varSubtotal',XmlBoolToStr(FVarSubtotal));
  if FVarPSubtotal <> False then 
    AWriter.AddAttribute('varPSubtotal',XmlBoolToStr(FVarPSubtotal));
end;

procedure TCT_PivotAreaReference.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashB[i] of
      $58197860: FField := AAttributes.Values[i];
      $7C8E2A59: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $CF21447D: FSelected := XmlStrToBoolDef(AAttributes.Values[i],True);
      $8FCBDD08: FByPosition := XmlStrToBoolDef(AAttributes.Values[i],False);
      $312FACFC: FRelative := XmlStrToBoolDef(AAttributes.Values[i],False);
      $63773259: FDefaultSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $9AF60007: FSumSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $08247C16: FCountASubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $7395335C: FAvgSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $17A64974: FMaxSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $C462FB72: FMinSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $5C8174E7: FProductSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $3C197E61: FCountSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $1CF9CC40: FStdDevSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $FEF300FA: FStdDevPSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $4BF9149D: FVarSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $65C456D9: FVarPSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotAreaReference.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 17;
  FField := '';
  FCount := 2147483632;
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
end;

destructor TCT_PivotAreaReference.Destroy;
begin
  if FXXpgList <> Nil then 
    FXXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PivotAreaReference.Clear;
begin
  FAssigneds := [];
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  if FExtLst <> Nil then
    FreeAndNil(FExtLst);
  FField := '';
  FCount := 2147483632;
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
end;

procedure TCT_PivotAreaReference.Assign(AItem: TCT_PivotAreaReference);
begin
end;

procedure TCT_PivotAreaReference.CopyTo(AItem: TCT_PivotAreaReference);
begin
end;

function  TCT_PivotAreaReference.Create_XXpgList: TCT_IndexXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_IndexXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

function  TCT_PivotAreaReference.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PivotAreaReferenceXpgList }

function  TCT_PivotAreaReferenceXpgList.GetItems(Index: integer): TCT_PivotAreaReference;
begin
  Result := TCT_PivotAreaReference(inherited Items[Index]);
end;

function  TCT_PivotAreaReferenceXpgList.Add: TCT_PivotAreaReference;
begin
  Result := TCT_PivotAreaReference.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotAreaReferenceXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotAreaReferenceXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PivotAreaReferenceXpgList.Assign(AItem: TCT_PivotAreaReferenceXpgList);
begin
end;

procedure TCT_PivotAreaReferenceXpgList.CopyTo(AItem: TCT_PivotAreaReferenceXpgList);
begin
end;

{ TCT_Tuples }

function  TCT_Tuples.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FC <> 2147483632 then 
    Inc(AttrsAssigned);
  if FTplXpgList <> Nil then 
    Inc(ElemsAssigned,FTplXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Tuples.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'tpl' then 
  begin
    if FTplXpgList = Nil then 
      FTplXpgList := TCT_TupleXpgList.Create(FOwner);
    Result := FTplXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Tuples.Write(AWriter: TXpgWriteXML);
begin
  if FTplXpgList <> Nil then 
    FTplXpgList.Write(AWriter,'tpl');
end;

procedure TCT_Tuples.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FC <> 2147483632 then 
    AWriter.AddAttribute('c',XmlIntToStr(FC));
end;

procedure TCT_Tuples.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'c' then 
    FC := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Tuples.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FC := 2147483632;
end;

destructor TCT_Tuples.Destroy;
begin
  if FTplXpgList <> Nil then 
    FTplXpgList.Free;
end;

procedure TCT_Tuples.Clear;
begin
  FAssigneds := [];
  if FTplXpgList <> Nil then 
    FreeAndNil(FTplXpgList);
  FC := 2147483632;
end;

procedure TCT_Tuples.Assign(AItem: TCT_Tuples);
begin
end;

procedure TCT_Tuples.CopyTo(AItem: TCT_Tuples);
begin
end;

function  TCT_Tuples.Create_TplXpgList: TCT_TupleXpgList;
begin
  if FTplXpgList = Nil then
      FTplXpgList := TCT_TupleXpgList.Create(FOwner);
  Result := FTplXpgList ;
end;

{ TCT_TuplesXpgList }

function  TCT_TuplesXpgList.GetItems(Index: integer): TCT_Tuples;
begin
  Result := TCT_Tuples(inherited Items[Index]);
end;

function  TCT_TuplesXpgList.Add: TCT_Tuples;
begin
  Result := TCT_Tuples.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TuplesXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TuplesXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_TuplesXpgList.Assign(AItem: TCT_TuplesXpgList);
begin
end;

procedure TCT_TuplesXpgList.CopyTo(AItem: TCT_TuplesXpgList);
begin
end;

{ TCT_X }

function  TCT_X.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FV <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_X.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_X.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FV <> 0 then 
    AWriter.AddAttribute('v',XmlIntToStr(FV));
end;

procedure TCT_X.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'v' then 
    FV := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_X.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FV := 0;
end;

destructor TCT_X.Destroy;
begin
end;

procedure TCT_X.Clear;
begin
  FAssigneds := [];
  FV := 0;
end;

procedure TCT_X.Assign(AItem: TCT_X);
begin
end;

procedure TCT_X.CopyTo(AItem: TCT_X);
begin
end;

{ TCT_XXpgList }

function  TCT_XXpgList.GetItems(Index: integer): TCT_X;
begin
  Result := TCT_X(inherited Items[Index]);
end;

function  TCT_XXpgList.Add: TCT_X;
begin
  Result := TCT_X.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_XXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_XXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_XXpgList.Assign(AItem: TCT_XXpgList);
begin
end;

procedure TCT_XXpgList.CopyTo(AItem: TCT_XXpgList);
begin
end;

{ TCT_LevelGroup }

function  TCT_LevelGroup.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FGroupMembers <> Nil then 
    Inc(ElemsAssigned,FGroupMembers.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_LevelGroup.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'groupMembers' then 
  begin
    if FGroupMembers = Nil then 
      FGroupMembers := TCT_GroupMembers.Create(FOwner);
    Result := FGroupMembers;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_LevelGroup.Write(AWriter: TXpgWriteXML);
begin
  if (FGroupMembers <> Nil) and FGroupMembers.Assigned then 
  begin
    FGroupMembers.WriteAttributes(AWriter);
    if xaElements in FGroupMembers.FAssigneds then 
    begin
      AWriter.BeginTag('groupMembers');
      FGroupMembers.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('groupMembers');
  end;
end;

procedure TCT_LevelGroup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('uniqueName',FUniqueName);
  AWriter.AddAttribute('caption',FCaption);
  if FUniqueParent <> '' then 
    AWriter.AddAttribute('uniqueParent',FUniqueParent);
  if FId <> 2147483632 then 
    AWriter.AddAttribute('id',XmlIntToStr(FId));
end;

procedure TCT_LevelGroup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000418: FUniqueName := AAttributes.Values[i];
      $000002EE: FCaption := AAttributes.Values[i];
      $00000501: FUniqueParent := AAttributes.Values[i];
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_LevelGroup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 5;
  FId := 2147483632;
end;

destructor TCT_LevelGroup.Destroy;
begin
  if FGroupMembers <> Nil then 
    FGroupMembers.Free;
end;

procedure TCT_LevelGroup.Clear;
begin
  FAssigneds := [];
  if FGroupMembers <> Nil then 
    FreeAndNil(FGroupMembers);
  FName := '';
  FUniqueName := '';
  FCaption := '';
  FUniqueParent := '';
  FId := 2147483632;
end;

procedure TCT_LevelGroup.Assign(AItem: TCT_LevelGroup);
begin
end;

procedure TCT_LevelGroup.CopyTo(AItem: TCT_LevelGroup);
begin
end;

function  TCT_LevelGroup.Create_GroupMembers: TCT_GroupMembers;
begin
  if FGroupMembers = Nil then
      FGroupMembers := TCT_GroupMembers.Create(FOwner);
  Result := FGroupMembers ;
end;

{ TCT_LevelGroupXpgList }

function  TCT_LevelGroupXpgList.GetItems(Index: integer): TCT_LevelGroup;
begin
  Result := TCT_LevelGroup(inherited Items[Index]);
end;

function  TCT_LevelGroupXpgList.Add: TCT_LevelGroup;
begin
  Result := TCT_LevelGroup.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_LevelGroupXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_LevelGroupXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_LevelGroupXpgList.Assign(AItem: TCT_LevelGroupXpgList);
begin
end;

procedure TCT_LevelGroupXpgList.CopyTo(AItem: TCT_LevelGroupXpgList);
begin
end;

{ TCT_PivotAreaReferences }

function  TCT_PivotAreaReferences.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FReferenceXpgList <> Nil then 
    Inc(ElemsAssigned,FReferenceXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotAreaReferences.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'reference' then 
  begin
    if FReferenceXpgList = Nil then 
      FReferenceXpgList := TCT_PivotAreaReferenceXpgList.Create(FOwner);
    Result := FReferenceXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotAreaReferences.Write(AWriter: TXpgWriteXML);
begin
  if FReferenceXpgList <> Nil then 
    FReferenceXpgList.Write(AWriter,'reference');
end;

procedure TCT_PivotAreaReferences.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PivotAreaReferences.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PivotAreaReferences.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PivotAreaReferences.Destroy;
begin
  if FReferenceXpgList <> Nil then 
    FReferenceXpgList.Free;
end;

procedure TCT_PivotAreaReferences.Clear;
begin
  FAssigneds := [];
  if FReferenceXpgList <> Nil then 
    FreeAndNil(FReferenceXpgList);
  FCount := 2147483632;
end;

procedure TCT_PivotAreaReferences.Assign(AItem: TCT_PivotAreaReferences);
begin
end;

procedure TCT_PivotAreaReferences.CopyTo(AItem: TCT_PivotAreaReferences);
begin
end;

function  TCT_PivotAreaReferences.Create_ReferenceXpgList: TCT_PivotAreaReferenceXpgList;
begin
  if FReferenceXpgList = Nil then
      FReferenceXpgList := TCT_PivotAreaReferenceXpgList.Create(FOwner);
  Result := FReferenceXpgList ;
end;

{ TCT_PageItem }

function  TCT_PageItem.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PageItem.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PageItem.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
end;

procedure TCT_PageItem.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PageItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_PageItem.Destroy;
begin
end;

procedure TCT_PageItem.Clear;
begin
  FAssigneds := [];
  FName := '';
end;

procedure TCT_PageItem.Assign(AItem: TCT_PageItem);
begin
end;

procedure TCT_PageItem.CopyTo(AItem: TCT_PageItem);
begin
end;

{ TCT_PageItemXpgList }

function  TCT_PageItemXpgList.GetItems(Index: integer): TCT_PageItem;
begin
  Result := TCT_PageItem(inherited Items[Index]);
end;

function  TCT_PageItemXpgList.Add: TCT_PageItem;
begin
  Result := TCT_PageItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PageItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PageItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PageItemXpgList.Assign(AItem: TCT_PageItemXpgList);
begin
end;

procedure TCT_PageItemXpgList.CopyTo(AItem: TCT_PageItemXpgList);
begin
end;

{ TCT_Missing }

function  TCT_Missing.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Byte(FU) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FF) <> 2 then 
    Inc(AttrsAssigned);
  if FC <> '' then 
    Inc(AttrsAssigned);
  if FCp <> 2147483632 then 
    Inc(AttrsAssigned);
  if FIn <> 2147483632 then 
    Inc(AttrsAssigned);
  if FBc <> -16 then 
    Inc(AttrsAssigned);
  if FFc <> -16 then 
    Inc(AttrsAssigned);
  if FI <> False then 
    Inc(AttrsAssigned);
  if FUn <> False then 
    Inc(AttrsAssigned);
  if FSt <> False then 
    Inc(AttrsAssigned);
  if FB <> False then 
    Inc(AttrsAssigned);
  if FTplsXpgList <> Nil then 
    Inc(ElemsAssigned,FTplsXpgList.CheckAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Missing.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001C3: begin
      if FTplsXpgList = Nil then 
        FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
      Result := FTplsXpgList.Add;
    end;
    $00000078: begin
      if FXXpgList = Nil then 
        FXXpgList := TCT_XXpgList.Create(FOwner);
      Result := FXXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Missing.Write(AWriter: TXpgWriteXML);
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Write(AWriter,'tpls');
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_Missing.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Byte(FU) <> 2 then 
    AWriter.AddAttribute('u',XmlBoolToStr(FU));
  if Byte(FF) <> 2 then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FC <> '' then 
    AWriter.AddAttribute('c',FC);
  if FCp <> 2147483632 then 
    AWriter.AddAttribute('cp',XmlIntToStr(FCp));
  if FIn <> 2147483632 then 
    AWriter.AddAttribute('in',XmlIntToStr(FIn));
  if FBc <> -16 then 
    AWriter.AddAttribute('bc',XmlIntToHexStr(FBc));
  if FFc <> -16 then 
    AWriter.AddAttribute('fc',XmlIntToHexStr(FFc));
  if FI <> False then 
    AWriter.AddAttribute('i',XmlBoolToStr(FI));
  if FUn <> False then 
    AWriter.AddAttribute('un',XmlBoolToStr(FUn));
  if FSt <> False then 
    AWriter.AddAttribute('st',XmlBoolToStr(FSt));
  if FB <> False then 
    AWriter.AddAttribute('b',XmlBoolToStr(FB));
end;

procedure TCT_Missing.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000075: FU := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := AAttributes.Values[i];
      $000000D3: FCp := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000D7: FIn := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000C5: FBc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $000000C9: FFc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000069: FI := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E3: FUn := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E7: FSt := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000062: FB := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Missing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 11;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

destructor TCT_Missing.Destroy;
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Free;
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_Missing.Clear;
begin
  FAssigneds := [];
  if FTplsXpgList <> Nil then 
    FreeAndNil(FTplsXpgList);
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  Byte(FU) := 2;
  Byte(FF) := 2;
  FC := '';
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

procedure TCT_Missing.Assign(AItem: TCT_Missing);
begin
end;

procedure TCT_Missing.CopyTo(AItem: TCT_Missing);
begin
end;

function  TCT_Missing.Create_TplsXpgList: TCT_TuplesXpgList;
begin
  if FTplsXpgList = Nil then
      FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
  Result := FTplsXpgList ;
end;

function  TCT_Missing.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_MissingXpgList }

function  TCT_MissingXpgList.GetItems(Index: integer): TCT_Missing;
begin
  Result := TCT_Missing(inherited Items[Index]);
end;

function  TCT_MissingXpgList.Add: TCT_Missing;
begin
  Result := TCT_Missing.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MissingXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MissingXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_MissingXpgList.Assign(AItem: TCT_MissingXpgList);
begin
end;

procedure TCT_MissingXpgList.CopyTo(AItem: TCT_MissingXpgList);
begin
end;

{ TCT_Number }

function  TCT_Number.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FTplsXpgList <> Nil then 
    Inc(ElemsAssigned,FTplsXpgList.CheckAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Number.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001C3: begin
      if FTplsXpgList = Nil then 
        FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
      Result := FTplsXpgList.Add;
    end;
    $00000078: begin
      if FXXpgList = Nil then 
        FXXpgList := TCT_XXpgList.Create(FOwner);
      Result := FXXpgList.Add;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Number.Write(AWriter: TXpgWriteXML);
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Write(AWriter,'tpls');
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_Number.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',XmlFloatToStr(FV));
  if Byte(FU) <> 2 then 
    AWriter.AddAttribute('u',XmlBoolToStr(FU));
  if Byte(FF) <> 2 then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FC <> '' then 
    AWriter.AddAttribute('c',FC);
  if FCp <> 2147483632 then 
    AWriter.AddAttribute('cp',XmlIntToStr(FCp));
  if FIn <> 2147483632 then 
    AWriter.AddAttribute('in',XmlIntToStr(FIn));
  if FBc <> -16 then 
    AWriter.AddAttribute('bc',XmlIntToHexStr(FBc));
  if FFc <> -16 then 
    AWriter.AddAttribute('fc',XmlIntToHexStr(FFc));
  if FI <> False then 
    AWriter.AddAttribute('i',XmlBoolToStr(FI));
  if FUn <> False then 
    AWriter.AddAttribute('un',XmlBoolToStr(FUn));
  if FSt <> False then 
    AWriter.AddAttribute('st',XmlBoolToStr(FSt));
  if FB <> False then 
    AWriter.AddAttribute('b',XmlBoolToStr(FB));
end;

procedure TCT_Number.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000076: FV := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000075: FU := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := AAttributes.Values[i];
      $000000D3: FCp := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000D7: FIn := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000C5: FBc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $000000C9: FFc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000069: FI := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E3: FUn := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E7: FSt := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000062: FB := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Number.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 12;
  FV := NaN;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

destructor TCT_Number.Destroy;
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Free;
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_Number.Clear;
begin
  FAssigneds := [];
  if FTplsXpgList <> Nil then 
    FreeAndNil(FTplsXpgList);
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  FV := NaN;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FC := '';
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

procedure TCT_Number.Assign(AItem: TCT_Number);
begin
end;

procedure TCT_Number.CopyTo(AItem: TCT_Number);
begin
end;

function  TCT_Number.Create_TplsXpgList: TCT_TuplesXpgList;
begin
  if FTplsXpgList = Nil then
      FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
  Result := FTplsXpgList ;
end;

function  TCT_Number.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_NumberXpgList }

function  TCT_NumberXpgList.GetItems(Index: integer): TCT_Number;
begin
  Result := TCT_Number(inherited Items[Index]);
end;

function  TCT_NumberXpgList.Add: TCT_Number;
begin
  Result := TCT_Number.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_NumberXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_NumberXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_NumberXpgList.Assign(AItem: TCT_NumberXpgList);
begin
end;

procedure TCT_NumberXpgList.CopyTo(AItem: TCT_NumberXpgList);
begin
end;

{ TCT_Boolean }

function  TCT_Boolean.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Boolean.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'x' then 
  begin
    if FXXpgList = Nil then 
      FXXpgList := TCT_XXpgList.Create(FOwner);
    Result := FXXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Boolean.Write(AWriter: TXpgWriteXML);
begin
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_Boolean.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',XmlBoolToStr(FV));
  if Byte(FU) <> 2 then 
    AWriter.AddAttribute('u',XmlBoolToStr(FU));
  if Byte(FF) <> 2 then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FC <> '' then 
    AWriter.AddAttribute('c',FC);
  if FCp <> 2147483632 then
    AWriter.AddAttribute('cp',XmlIntToStr(FCp));
end;

procedure TCT_Boolean.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000076: FV := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000075: FU := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := AAttributes.Values[i];
      $000000D3: FCp := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Boolean.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 5;
  Byte(FV) := 2;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FCp := 2147483632;
end;

destructor TCT_Boolean.Destroy;
begin
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_Boolean.Clear;
begin
  FAssigneds := [];
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  Byte(FV) := 2;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FC := '';
  FCp := 2147483632;
end;

procedure TCT_Boolean.Assign(AItem: TCT_Boolean);
begin
end;

procedure TCT_Boolean.CopyTo(AItem: TCT_Boolean);
begin
end;

function  TCT_Boolean.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_BooleanXpgList }

function  TCT_BooleanXpgList.GetItems(Index: integer): TCT_Boolean;
begin
  Result := TCT_Boolean(inherited Items[Index]);
end;

function  TCT_BooleanXpgList.Add: TCT_Boolean;
begin
  Result := TCT_Boolean.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BooleanXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BooleanXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_BooleanXpgList.Assign(AItem: TCT_BooleanXpgList);
begin
end;

procedure TCT_BooleanXpgList.CopyTo(AItem: TCT_BooleanXpgList);
begin
end;

{ TCT_Error }

function  TCT_Error.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FTpls <> Nil then 
    Inc(ElemsAssigned,FTpls.CheckAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Error.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001C3: begin
      if FTpls = Nil then 
        FTpls := TCT_Tuples.Create(FOwner);
      Result := FTpls;
    end;
    $00000078: begin
      if FXXpgList = Nil then 
        FXXpgList := TCT_XXpgList.Create(FOwner);
      Result := FXXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Error.Write(AWriter: TXpgWriteXML);
begin
  if (FTpls <> Nil) and FTpls.Assigned then 
  begin
    FTpls.WriteAttributes(AWriter);
    if xaElements in FTpls.FAssigneds then 
    begin
      AWriter.BeginTag('tpls');
      FTpls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('tpls');
  end;
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_Error.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',FV);
  if Byte(FU) <> 2 then 
    AWriter.AddAttribute('u',XmlBoolToStr(FU));
  if Byte(FF) <> 2 then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FC <> '' then 
    AWriter.AddAttribute('c',FC);
  if FCp <> 2147483632 then 
    AWriter.AddAttribute('cp',XmlIntToStr(FCp));
  if FIn <> 2147483632 then 
    AWriter.AddAttribute('in',XmlIntToStr(FIn));
  if FBc <> -16 then 
    AWriter.AddAttribute('bc',XmlIntToHexStr(FBc));
  if FFc <> -16 then 
    AWriter.AddAttribute('fc',XmlIntToHexStr(FFc));
  if FI <> False then 
    AWriter.AddAttribute('i',XmlBoolToStr(FI));
  if FUn <> False then 
    AWriter.AddAttribute('un',XmlBoolToStr(FUn));
  if FSt <> False then 
    AWriter.AddAttribute('st',XmlBoolToStr(FSt));
  if FB <> False then 
    AWriter.AddAttribute('b',XmlBoolToStr(FB));
end;

procedure TCT_Error.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000076: FV := AAttributes.Values[i];
      $00000075: FU := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := AAttributes.Values[i];
      $000000D3: FCp := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000D7: FIn := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000C5: FBc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $000000C9: FFc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000069: FI := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E3: FUn := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E7: FSt := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000062: FB := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Error.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 12;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

destructor TCT_Error.Destroy;
begin
  if FTpls <> Nil then 
    FTpls.Free;
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_Error.Clear;
begin
  FAssigneds := [];
  if FTpls <> Nil then 
    FreeAndNil(FTpls);
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  FV := '';
  Byte(FU) := 2;
  Byte(FF) := 2;
  FC := '';
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

procedure TCT_Error.Assign(AItem: TCT_Error);
begin
end;

procedure TCT_Error.CopyTo(AItem: TCT_Error);
begin
end;

function  TCT_Error.Create_Tpls: TCT_Tuples;
begin
  if FTpls = Nil then
      FTpls := TCT_Tuples.Create(FOwner);
  Result := FTpls ;
end;

function  TCT_Error.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_ErrorXpgList }

function  TCT_ErrorXpgList.GetItems(Index: integer): TCT_Error;
begin
  Result := TCT_Error(inherited Items[Index]);
end;

function  TCT_ErrorXpgList.Add: TCT_Error;
begin
  Result := TCT_Error.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ErrorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ErrorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ErrorXpgList.Assign(AItem: TCT_ErrorXpgList);
begin
end;

procedure TCT_ErrorXpgList.CopyTo(AItem: TCT_ErrorXpgList);
begin
end;

{ TCT_String }

function  TCT_String.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FTplsXpgList <> Nil then 
    Inc(ElemsAssigned,FTplsXpgList.CheckAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_String.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001C3: begin
      if FTplsXpgList = Nil then 
        FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
      Result := FTplsXpgList.Add;
    end;
    $00000078: begin
      if FXXpgList = Nil then 
        FXXpgList := TCT_XXpgList.Create(FOwner);
      Result := FXXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_String.Write(AWriter: TXpgWriteXML);
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Write(AWriter,'tpls');
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_String.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',FV);
  if Byte(FU) <> 2 then 
    AWriter.AddAttribute('u',XmlBoolToStr(FU));
  if Byte(FF) <> 2 then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FC <> '' then 
    AWriter.AddAttribute('c',FC);
  if FCp <> 2147483632 then 
    AWriter.AddAttribute('cp',XmlIntToStr(FCp));
  if FIn <> 2147483632 then 
    AWriter.AddAttribute('in',XmlIntToStr(FIn));
  if FBc <> -16 then 
    AWriter.AddAttribute('bc',XmlIntToHexStr(FBc));
  if FFc <> -16 then 
    AWriter.AddAttribute('fc',XmlIntToHexStr(FFc));
  if FI <> False then 
    AWriter.AddAttribute('i',XmlBoolToStr(FI));
  if FUn <> False then 
    AWriter.AddAttribute('un',XmlBoolToStr(FUn));
  if FSt <> False then 
    AWriter.AddAttribute('st',XmlBoolToStr(FSt));
  if FB <> False then 
    AWriter.AddAttribute('b',XmlBoolToStr(FB));
end;

procedure TCT_String.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000076: FV := AAttributes.Values[i];
      $00000075: FU := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := AAttributes.Values[i];
      $000000D3: FCp := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000D7: FIn := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000C5: FBc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $000000C9: FFc := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000069: FI := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E3: FUn := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E7: FSt := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000062: FB := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_String.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 12;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

destructor TCT_String.Destroy;
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Free;
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_String.Clear;
begin
  FAssigneds := [];
  if FTplsXpgList <> Nil then 
    FreeAndNil(FTplsXpgList);
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  FV := '';
  Byte(FU) := 2;
  Byte(FF) := 2;
  FC := '';
  FCp := 2147483632;
  FIn := 2147483632;
  FBc := -16;
  FFc := -16;
  FI := False;
  FUn := False;
  FSt := False;
  FB := False;
end;

procedure TCT_String.Assign(AItem: TCT_String);
begin
end;

procedure TCT_String.CopyTo(AItem: TCT_String);
begin
end;

function  TCT_String.Create_TplsXpgList: TCT_TuplesXpgList;
begin
  if FTplsXpgList = Nil then
      FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
  Result := FTplsXpgList ;
end;

function  TCT_String.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_StringXpgList }

function  TCT_StringXpgList.GetItems(Index: integer): TCT_String;
begin
  Result := TCT_String(inherited Items[Index]);
end;

function  TCT_StringXpgList.Add: TCT_String;
begin
  Result := TCT_String.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_StringXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_StringXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_StringXpgList.Assign(AItem: TCT_StringXpgList);
begin
end;

procedure TCT_StringXpgList.CopyTo(AItem: TCT_StringXpgList);
begin
end;

{ TCT_DateTime }

function  TCT_DateTime.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DateTime.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'x' then 
  begin
    if FXXpgList = Nil then 
      FXXpgList := TCT_XXpgList.Create(FOwner);
    Result := FXXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DateTime.Write(AWriter: TXpgWriteXML);
begin
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_DateTime.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',XmlDateTimeToStr(FV));
  if Byte(FU) <> 2 then 
    AWriter.AddAttribute('u',XmlBoolToStr(FU));
  if Byte(FF) <> 2 then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FC <> '' then 
    AWriter.AddAttribute('c',FC);
  if FCp <> 2147483632 then 
    AWriter.AddAttribute('cp',XmlIntToStr(FCp));
end;

procedure TCT_DateTime.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000076: FV := XmlStrToDateTime(AAttributes.Values[i]);
      $00000075: FU := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := AAttributes.Values[i];
      $000000D3: FCp := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_DateTime.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 5;
  FV := 0;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FCp := 2147483632;
end;

destructor TCT_DateTime.Destroy;
begin
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_DateTime.Clear;
begin
  FAssigneds := [];
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  FV := 0;
  Byte(FU) := 2;
  Byte(FF) := 2;
  FC := '';
  FCp := 2147483632;
end;

procedure TCT_DateTime.Assign(AItem: TCT_DateTime);
begin
end;

procedure TCT_DateTime.CopyTo(AItem: TCT_DateTime);
begin
end;

function  TCT_DateTime.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_DateTimeXpgList }

function  TCT_DateTimeXpgList.GetItems(Index: integer): TCT_DateTime;
begin
  Result := TCT_DateTime(inherited Items[Index]);
end;

function  TCT_DateTimeXpgList.Add: TCT_DateTime;
begin
  Result := TCT_DateTime.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DateTimeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DateTimeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_DateTimeXpgList.Assign(AItem: TCT_DateTimeXpgList);
begin
end;

procedure TCT_DateTimeXpgList.CopyTo(AItem: TCT_DateTimeXpgList);
begin
end;

{ TCT_Groups }

function  TCT_Groups.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FGroupXpgList <> Nil then 
    Inc(ElemsAssigned,FGroupXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Groups.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'group' then 
  begin
    if FGroupXpgList = Nil then 
      FGroupXpgList := TCT_LevelGroupXpgList.Create(FOwner);
    Result := FGroupXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Groups.Write(AWriter: TXpgWriteXML);
begin
  if FGroupXpgList <> Nil then 
    FGroupXpgList.Write(AWriter,'group');
end;

procedure TCT_Groups.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Groups.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Groups.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_Groups.Destroy;
begin
  if FGroupXpgList <> Nil then 
    FGroupXpgList.Free;
end;

procedure TCT_Groups.Clear;
begin
  FAssigneds := [];
  if FGroupXpgList <> Nil then 
    FreeAndNil(FGroupXpgList);
  FCount := 2147483632;
end;

procedure TCT_Groups.Assign(AItem: TCT_Groups);
begin
end;

procedure TCT_Groups.CopyTo(AItem: TCT_Groups);
begin
end;

function  TCT_Groups.Create_GroupXpgList: TCT_LevelGroupXpgList;
begin
  if FGroupXpgList = Nil then
      FGroupXpgList := TCT_LevelGroupXpgList.Create(FOwner);
  Result := FGroupXpgList ;
end;

{ TCT_Item }

function  TCT_Item.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FN <> '' then 
    Inc(AttrsAssigned);
  if FT <> stitData then 
    Inc(AttrsAssigned);
  if FH <> False then 
    Inc(AttrsAssigned);
  if FS <> False then 
    Inc(AttrsAssigned);
  if FSd <> True then 
    Inc(AttrsAssigned);
  if FF <> False then 
    Inc(AttrsAssigned);
  if FM <> False then 
    Inc(AttrsAssigned);
  if FC <> False then 
    Inc(AttrsAssigned);
  if FX <> 2147483632 then 
    Inc(AttrsAssigned);
  if FD <> False then 
    Inc(AttrsAssigned);
  if FE <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Item.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Item.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FN <> '' then 
    AWriter.AddAttribute('n',FN);
  if FT <> stitData then 
    AWriter.AddAttribute('t',StrTST_ItemType[Integer(FT)]);
  if FH <> False then 
    AWriter.AddAttribute('h',XmlBoolToStr(FH));
  if FS <> False then 
    AWriter.AddAttribute('s',XmlBoolToStr(FS));
  if FSd <> True then 
    AWriter.AddAttribute('sd',XmlBoolToStr(FSd));
  if FF <> False then 
    AWriter.AddAttribute('f',XmlBoolToStr(FF));
  if FM <> False then 
    AWriter.AddAttribute('m',XmlBoolToStr(FM));
  if FC <> False then 
    AWriter.AddAttribute('c',XmlBoolToStr(FC));
  if FX <> 2147483632 then 
    AWriter.AddAttribute('x',XmlIntToStr(FX));
  if FD <> False then 
    AWriter.AddAttribute('d',XmlBoolToStr(FD));
  if FE <> True then 
    AWriter.AddAttribute('e',XmlBoolToStr(FE));
end;

procedure TCT_Item.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000006E: FN := AAttributes.Values[i];
      $00000074: FT := TST_ItemType(StrToEnum('stit' + AAttributes.Values[i]));
      $00000068: FH := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000073: FS := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000D7: FSd := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000066: FF := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000006D: FM := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000063: FC := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000078: FX := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000064: FD := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000065: FE := XmlStrToBoolDef(AAttributes.Values[i],True);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Item.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 11;
  FT := TST_ItemType(XPG_UNKNOWN_ENUM);
  FT := stitData;
  FH := False;
  FS := False;
  FSd := True;
  FF := False;
  FM := False;
  FC := False;
  FX := 2147483632;
  FD := False;
  FE := True;
end;

destructor TCT_Item.Destroy;
begin
end;

procedure TCT_Item.Clear;
begin
  FAssigneds := [];
  FN := '';
  FT := stitData;
  FH := False;
  FS := False;
  FSd := True;
  FF := False;
  FM := False;
  FC := False;
  FX := 2147483632;
  FD := False;
  FE := True;
end;

procedure TCT_Item.Assign(AItem: TCT_Item);
begin
end;

procedure TCT_Item.CopyTo(AItem: TCT_Item);
begin
end;

{ TCT_PivotArea }

function  TCT_PivotArea.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FField <> 2147483632 then 
    Inc(AttrsAssigned);
  if FType <> stpatNormal then 
    Inc(AttrsAssigned);
  if FDataOnly <> True then 
    Inc(AttrsAssigned);
  if FLabelOnly <> False then 
    Inc(AttrsAssigned);
  if FGrandRow <> False then 
    Inc(AttrsAssigned);
  if FGrandCol <> False then 
    Inc(AttrsAssigned);
  if FCacheIndex <> False then 
    Inc(AttrsAssigned);
  if FOutline <> True then 
    Inc(AttrsAssigned);
  if FOffset <> '' then 
    Inc(AttrsAssigned);
  if FCollapsedLevelsAreSubtotals <> False then 
    Inc(AttrsAssigned);
  if Integer(FAxis) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FFieldPosition <> 2147483632 then 
    Inc(AttrsAssigned);
  if FReferences <> Nil then 
    Inc(ElemsAssigned,FReferences.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotArea.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000422: begin
      if FReferences = Nil then 
        FReferences := TCT_PivotAreaReferences.Create(FOwner);
      Result := FReferences;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotArea.Write(AWriter: TXpgWriteXML);
begin
  if (FReferences <> Nil) and FReferences.Assigned then 
  begin
    FReferences.WriteAttributes(AWriter);
    if xaElements in FReferences.FAssigneds then 
    begin
      AWriter.BeginTag('references');
      FReferences.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('references');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotArea.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FField <> 2147483632 then 
    AWriter.AddAttribute('field',XmlIntToStr(FField));
  if FType <> stpatNormal then 
    AWriter.AddAttribute('type',StrTST_PivotAreaType[Integer(FType)]);
  if FDataOnly <> True then 
    AWriter.AddAttribute('dataOnly',XmlBoolToStr(FDataOnly));
  if FLabelOnly <> False then 
    AWriter.AddAttribute('labelOnly',XmlBoolToStr(FLabelOnly));
  if FGrandRow <> False then 
    AWriter.AddAttribute('grandRow',XmlBoolToStr(FGrandRow));
  if FGrandCol <> False then 
    AWriter.AddAttribute('grandCol',XmlBoolToStr(FGrandCol));
  if FCacheIndex <> False then 
    AWriter.AddAttribute('cacheIndex',XmlBoolToStr(FCacheIndex));
  if FOutline <> True then 
    AWriter.AddAttribute('outline',XmlBoolToStr(FOutline));
  if FOffset <> '' then 
    AWriter.AddAttribute('offset',FOffset);
  if FCollapsedLevelsAreSubtotals <> False then 
    AWriter.AddAttribute('collapsedLevelsAreSubtotals',XmlBoolToStr(FCollapsedLevelsAreSubtotals));
  if Integer(FAxis) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('axis',StrTST_Axis[Integer(FAxis)]);
  if FFieldPosition <> 2147483632 then 
    AWriter.AddAttribute('fieldPosition',XmlIntToStr(FFieldPosition));
end;

procedure TCT_PivotArea.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000204: FField := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C2: FType := TST_PivotAreaType(StrToEnum('stpat' + AAttributes.Values[i]));
      $0000033C: FDataOnly := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000003A2: FLabelOnly := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000344: FGrandRow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000032A: FGrandCol := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003EC: FCacheIndex := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000300: FOutline := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000287: FOffset := AAttributes.Values[i];
      $00000AFB: FCollapsedLevelsAreSubtotals := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001B5: FAxis := TST_Axis(StrToEnum('sta' + AAttributes.Values[i]));
      $00000559: FFieldPosition := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotArea.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 12;
  FField := 2147483632;
  FType := TST_PivotAreaType(XPG_UNKNOWN_ENUM);
  FType := stpatNormal;
  FDataOnly := True;
  FLabelOnly := False;
  FGrandRow := False;
  FGrandCol := False;
  FCacheIndex := False;
  FOutline := True;
  FCollapsedLevelsAreSubtotals := False;
  FAxis := TST_Axis(XPG_UNKNOWN_ENUM);
  FFieldPosition := 2147483632;
end;

destructor TCT_PivotArea.Destroy;
begin
  if FReferences <> Nil then 
    FReferences.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PivotArea.Clear;
begin
  FAssigneds := [];
  if FReferences <> Nil then 
    FreeAndNil(FReferences);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FField := 2147483632;
  FType := stpatNormal;
  FDataOnly := True;
  FLabelOnly := False;
  FGrandRow := False;
  FGrandCol := False;
  FCacheIndex := False;
  FOutline := True;
  FOffset := '';
  FCollapsedLevelsAreSubtotals := False;
  FFieldPosition := 2147483632;
end;

procedure TCT_PivotArea.Assign(AItem: TCT_PivotArea);
begin
end;

procedure TCT_PivotArea.CopyTo(AItem: TCT_PivotArea);
begin
end;

function  TCT_PivotArea.Create_References: TCT_PivotAreaReferences;
begin
  if FReferences = Nil then
      FReferences := TCT_PivotAreaReferences.Create(FOwner);
  Result := FReferences ;
end;

function  TCT_PivotArea.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PivotAreaXpgList }

function  TCT_PivotAreaXpgList.GetItems(Index: integer): TCT_PivotArea;
begin
  Result := TCT_PivotArea(inherited Items[Index]);
end;

function  TCT_PivotAreaXpgList.Add: TCT_PivotArea;
begin
  Result := TCT_PivotArea.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotAreaXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotAreaXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PivotAreaXpgList.Assign(AItem: TCT_PivotAreaXpgList);
begin
end;

procedure TCT_PivotAreaXpgList.CopyTo(AItem: TCT_PivotAreaXpgList);
begin
end;

{ TCT_MemberProperty }

function  TCT_MemberProperty.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_MemberProperty.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_MemberProperty.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FShowCell <> False then 
    AWriter.AddAttribute('showCell',XmlBoolToStr(FShowCell));
  if FShowTip <> False then 
    AWriter.AddAttribute('showTip',XmlBoolToStr(FShowTip));
  if FShowAsCaption <> False then 
    AWriter.AddAttribute('showAsCaption',XmlBoolToStr(FShowAsCaption));
  if FNameLen <> 2147483632 then 
    AWriter.AddAttribute('nameLen',XmlIntToStr(FNameLen));
  if FPPos <> 2147483632 then 
    AWriter.AddAttribute('pPos',XmlIntToStr(FPPos));
  if FPLen <> 2147483632 then 
    AWriter.AddAttribute('pLen',XmlIntToStr(FPLen));
  if FLevel <> 2147483632 then 
    AWriter.AddAttribute('level',XmlIntToStr(FLevel));
  AWriter.AddAttribute('field',XmlIntToStr(FField));
end;

procedure TCT_MemberProperty.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000341: FShowCell := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002EE: FShowTip := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000543: FShowAsCaption := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002C0: FNameLen := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A2: FPPos := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000018F: FPLen := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000218: FLevel := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000204: FField := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_MemberProperty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 9;
  FShowCell := False;
  FShowTip := False;
  FShowAsCaption := False;
  FNameLen := 2147483632;
  FPPos := 2147483632;
  FPLen := 2147483632;
  FLevel := 2147483632;
  FField := 2147483632;
end;

destructor TCT_MemberProperty.Destroy;
begin
end;

procedure TCT_MemberProperty.Clear;
begin
  FAssigneds := [];
  FName := '';
  FShowCell := False;
  FShowTip := False;
  FShowAsCaption := False;
  FNameLen := 2147483632;
  FPPos := 2147483632;
  FPLen := 2147483632;
  FLevel := 2147483632;
  FField := 2147483632;
end;

procedure TCT_MemberProperty.Assign(AItem: TCT_MemberProperty);
begin
end;

procedure TCT_MemberProperty.CopyTo(AItem: TCT_MemberProperty);
begin
end;

{ TCT_MemberPropertyXpgList }

function  TCT_MemberPropertyXpgList.GetItems(Index: integer): TCT_MemberProperty;
begin
  Result := TCT_MemberProperty(inherited Items[Index]);
end;

function  TCT_MemberPropertyXpgList.Add: TCT_MemberProperty;
begin
  Result := TCT_MemberProperty.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MemberPropertyXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MemberPropertyXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_MemberPropertyXpgList.Assign(AItem: TCT_MemberPropertyXpgList);
begin
end;

procedure TCT_MemberPropertyXpgList.CopyTo(AItem: TCT_MemberPropertyXpgList);
begin
end;

{ TCT_Member }

function  TCT_Member.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Member.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Member.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
end;

procedure TCT_Member.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Member.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_Member.Destroy;
begin
end;

procedure TCT_Member.Clear;
begin
  FAssigneds := [];
  FName := '';
end;

procedure TCT_Member.Assign(AItem: TCT_Member);
begin
end;

procedure TCT_Member.CopyTo(AItem: TCT_Member);
begin
end;

{ TCT_MemberXpgList }

function  TCT_MemberXpgList.GetItems(Index: integer): TCT_Member;
begin
  Result := TCT_Member(inherited Items[Index]);
end;

function  TCT_MemberXpgList.Add: TCT_Member;
begin
  Result := TCT_Member.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MemberXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MemberXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_MemberXpgList.Assign(AItem: TCT_MemberXpgList);
begin
end;

procedure TCT_MemberXpgList.CopyTo(AItem: TCT_MemberXpgList);
begin
end;

{ TCT_PCDSCPage }

function  TCT_PCDSCPage.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FPageItemXpgList <> Nil then 
    Inc(ElemsAssigned,FPageItemXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PCDSCPage.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pageItem' then 
  begin
    if FPageItemXpgList = Nil then 
      FPageItemXpgList := TCT_PageItemXpgList.Create(FOwner);
    Result := FPageItemXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PCDSCPage.Write(AWriter: TXpgWriteXML);
begin
  if FPageItemXpgList <> Nil then 
    FPageItemXpgList.Write(AWriter,'pageItem');
end;

procedure TCT_PCDSCPage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PCDSCPage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PCDSCPage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PCDSCPage.Destroy;
begin
  if FPageItemXpgList <> Nil then 
    FPageItemXpgList.Free;
end;

procedure TCT_PCDSCPage.Clear;
begin
  FAssigneds := [];
  if FPageItemXpgList <> Nil then 
    FreeAndNil(FPageItemXpgList);
  FCount := 2147483632;
end;

procedure TCT_PCDSCPage.Assign(AItem: TCT_PCDSCPage);
begin
end;

procedure TCT_PCDSCPage.CopyTo(AItem: TCT_PCDSCPage);
begin
end;

function  TCT_PCDSCPage.Create_PageItemXpgList: TCT_PageItemXpgList;
begin
  if FPageItemXpgList = Nil then
      FPageItemXpgList := TCT_PageItemXpgList.Create(FOwner);
  Result := FPageItemXpgList ;
end;

{ TCT_PCDSCPageXpgList }

function  TCT_PCDSCPageXpgList.GetItems(Index: integer): TCT_PCDSCPage;
begin
  Result := TCT_PCDSCPage(inherited Items[Index]);
end;

function  TCT_PCDSCPageXpgList.Add: TCT_PCDSCPage;
begin
  Result := TCT_PCDSCPage.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PCDSCPageXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PCDSCPageXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PCDSCPageXpgList.Assign(AItem: TCT_PCDSCPageXpgList);
begin
end;

procedure TCT_PCDSCPageXpgList.CopyTo(AItem: TCT_PCDSCPageXpgList);
begin
end;

{ TCT_RangeSet }

function  TCT_RangeSet.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FI1 <> 2147483632 then 
    Inc(AttrsAssigned);
  if FI2 <> 2147483632 then 
    Inc(AttrsAssigned);
  if FI3 <> 2147483632 then 
    Inc(AttrsAssigned);
  if FI4 <> 2147483632 then 
    Inc(AttrsAssigned);
  if FRef <> '' then 
    Inc(AttrsAssigned);
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FSheet <> '' then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RangeSet.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RangeSet.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FI1 <> 2147483632 then 
    AWriter.AddAttribute('i1',XmlIntToStr(FI1));
  if FI2 <> 2147483632 then 
    AWriter.AddAttribute('i2',XmlIntToStr(FI2));
  if FI3 <> 2147483632 then 
    AWriter.AddAttribute('i3',XmlIntToStr(FI3));
  if FI4 <> 2147483632 then 
    AWriter.AddAttribute('i4',XmlIntToStr(FI4));
  if FRef <> '' then 
    AWriter.AddAttribute('ref',FRef);
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FSheet <> '' then 
    AWriter.AddAttribute('sheet',FSheet);
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_RangeSet.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000009A: FI1 := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000009B: FI2 := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000009C: FI3 := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000009D: FI4 := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013D: FRef := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $00000219: FSheet := AAttributes.Values[i];
      $00000179: FR_Id := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_RangeSet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 8;
  FI1 := 2147483632;
  FI2 := 2147483632;
  FI3 := 2147483632;
  FI4 := 2147483632;
  FR_Id := '';
end;

destructor TCT_RangeSet.Destroy;
begin
end;

procedure TCT_RangeSet.Clear;
begin
  FAssigneds := [];
  FI1 := 2147483632;
  FI2 := 2147483632;
  FI3 := 2147483632;
  FI4 := 2147483632;
  FRef := '';
  FName := '';
  FSheet := '';
  FR_Id := '';
end;

procedure TCT_RangeSet.Assign(AItem: TCT_RangeSet);
begin
end;

procedure TCT_RangeSet.CopyTo(AItem: TCT_RangeSet);
begin
end;

{ TCT_RangeSetXpgList }

function  TCT_RangeSetXpgList.GetItems(Index: integer): TCT_RangeSet;
begin
  Result := TCT_RangeSet(inherited Items[Index]);
end;

function  TCT_RangeSetXpgList.Add: TCT_RangeSet;
begin
  Result := TCT_RangeSet.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RangeSetXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RangeSetXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_RangeSetXpgList.Assign(AItem: TCT_RangeSetXpgList);
begin
end;

procedure TCT_RangeSetXpgList.CopyTo(AItem: TCT_RangeSetXpgList);
begin
end;

{ TCT_RangePr }

function  TCT_RangePr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAutoStart <> True then 
    Inc(AttrsAssigned);
  if FAutoEnd <> True then 
    Inc(AttrsAssigned);
  if FGroupBy <> stgbRange then 
    Inc(AttrsAssigned);
  if IsNaN(FStartNum) = False then 
    Inc(AttrsAssigned);
  if IsNaN(FEndNum) = False then 
    Inc(AttrsAssigned);
  if FStartDate <> 0 then 
    Inc(AttrsAssigned);
  if FEndDate <> 0 then 
    Inc(AttrsAssigned);
  if FGroupInterval <> 1 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RangePr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RangePr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAutoStart <> True then 
    AWriter.AddAttribute('autoStart',XmlBoolToStr(FAutoStart));
  if FAutoEnd <> True then 
    AWriter.AddAttribute('autoEnd',XmlBoolToStr(FAutoEnd));
  if FGroupBy <> stgbRange then 
    AWriter.AddAttribute('groupBy',StrTST_GroupBy[Integer(FGroupBy)]);
  if IsNaN(FStartNum) = False then 
    AWriter.AddAttribute('startNum',XmlFloatToStr(FStartNum));
  if IsNaN(FEndNum) = False then 
    AWriter.AddAttribute('endNum',XmlFloatToStr(FEndNum));
  if FStartDate <> 0 then 
    AWriter.AddAttribute('startDate',XmlDateTimeToStr(FStartDate));
  if FEndDate <> 0 then 
    AWriter.AddAttribute('endDate',XmlDateTimeToStr(FEndDate));
  if FGroupInterval <> 1 then 
    AWriter.AddAttribute('groupInterval',XmlFloatToStr(FGroupInterval));
end;

procedure TCT_RangePr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000003C7: FAutoStart := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000002D0: FAutoEnd := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000002E8: FGroupBy := TST_GroupBy(StrToEnum('stgb' + AAttributes.Values[i]));
      $0000035E: FStartNum := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000267: FEndNum := XmlStrToFloatDef(AAttributes.Values[i],0);
      $000003AC: FStartDate := XmlStrToDateTime(AAttributes.Values[i]);
      $000002B5: FEndDate := XmlStrToDateTime(AAttributes.Values[i]);
      $00000572: FGroupInterval := XmlStrToFloatDef(AAttributes.Values[i],1);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_RangePr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 8;
  FAutoStart := True;
  FAutoEnd := True;
  FGroupBy := TST_GroupBy(XPG_UNKNOWN_ENUM);
  FGroupBy := stgbRange;
  FStartNum := NaN;
  FEndNum := NaN;
  FStartDate := 0;
  FEndDate := 0;
  FGroupInterval := 1;
end;

destructor TCT_RangePr.Destroy;
begin
end;

procedure TCT_RangePr.Clear;
begin
  FAssigneds := [];
  FAutoStart := True;
  FAutoEnd := True;
  FGroupBy := stgbRange;
  FStartNum := NaN;
  FEndNum := NaN;
  FStartDate := 0;
  FEndDate := 0;
  FGroupInterval := 1;
end;

procedure TCT_RangePr.Assign(AItem: TCT_RangePr);
begin
end;

procedure TCT_RangePr.CopyTo(AItem: TCT_RangePr);
begin
end;

{ TCT_DiscretePr }

function  TCT_DiscretePr.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DiscretePr.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'x' then 
  begin
    if FXXpgList = Nil then 
      FXXpgList := TCT_IndexXpgList.Create(FOwner);
    Result := FXXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DiscretePr.Write(AWriter: TXpgWriteXML);
begin
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_DiscretePr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_DiscretePr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_DiscretePr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_DiscretePr.Destroy;
begin
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_DiscretePr.Clear;
begin
  FAssigneds := [];
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  FCount := 2147483632;
end;

procedure TCT_DiscretePr.Assign(AItem: TCT_DiscretePr);
begin
end;

procedure TCT_DiscretePr.CopyTo(AItem: TCT_DiscretePr);
begin
end;

function  TCT_DiscretePr.Create_XXpgList: TCT_IndexXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_IndexXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_GroupItems }

function  TCT_GroupItems.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMXpgList <> Nil then 
    Inc(ElemsAssigned,FMXpgList.CheckAssigned);
  if FNXpgList <> Nil then 
    Inc(ElemsAssigned,FNXpgList.CheckAssigned);
  if FBXpgList <> Nil then 
    Inc(ElemsAssigned,FBXpgList.CheckAssigned);
  if FEXpgList <> Nil then 
    Inc(ElemsAssigned,FEXpgList.CheckAssigned);
  if FSXpgList <> Nil then 
    Inc(ElemsAssigned,FSXpgList.CheckAssigned);
  if FDXpgList <> Nil then 
    Inc(ElemsAssigned,FDXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GroupItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000006D: begin
      if FMXpgList = Nil then
        FMXpgList := TCT_MissingXpgList.Create(FOwner);
      Result := FMXpgList.Add;
    end;
    $0000006E: begin
      if FNXpgList = Nil then
        FNXpgList := TCT_NumberXpgList.Create(FOwner);
      Result := FNXpgList.Add;
    end;
    $00000062: begin
      if FBXpgList = Nil then
        FBXpgList := TCT_BooleanXpgList.Create(FOwner);
      Result := FBXpgList.Add;
    end;
    $00000065: begin
      if FEXpgList = Nil then
        FEXpgList := TCT_ErrorXpgList.Create(FOwner);
      Result := FEXpgList.Add;
    end;
    $00000073: begin
      if FSXpgList = Nil then
        FSXpgList := TCT_StringXpgList.Create(FOwner);
      Result := FSXpgList.Add;
    end;
    $00000064: begin
      if FDXpgList = Nil then
        FDXpgList := TCT_DateTimeXpgList.Create(FOwner);
      Result := FDXpgList.Add;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupItems.Write(AWriter: TXpgWriteXML);
begin
  if FMXpgList <> Nil then 
    FMXpgList.Write(AWriter,'m');
  if FNXpgList <> Nil then 
    FNXpgList.Write(AWriter,'n');
  if FBXpgList <> Nil then 
    FBXpgList.Write(AWriter,'b');
  if FEXpgList <> Nil then 
    FEXpgList.Write(AWriter,'e');
  if FSXpgList <> Nil then 
    FSXpgList.Write(AWriter,'s');
  if FDXpgList <> Nil then 
    FDXpgList.Write(AWriter,'d');
end;

procedure TCT_GroupItems.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_GroupItems.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GroupItems.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_GroupItems.Destroy;
begin
  if FMXpgList <> Nil then 
    FMXpgList.Free;
  if FNXpgList <> Nil then 
    FNXpgList.Free;
  if FBXpgList <> Nil then 
    FBXpgList.Free;
  if FEXpgList <> Nil then 
    FEXpgList.Free;
  if FSXpgList <> Nil then 
    FSXpgList.Free;
  if FDXpgList <> Nil then 
    FDXpgList.Free;
end;

procedure TCT_GroupItems.Clear;
begin
  FAssigneds := [];
  if FMXpgList <> Nil then 
    FreeAndNil(FMXpgList);
  if FNXpgList <> Nil then 
    FreeAndNil(FNXpgList);
  if FBXpgList <> Nil then 
    FreeAndNil(FBXpgList);
  if FEXpgList <> Nil then 
    FreeAndNil(FEXpgList);
  if FSXpgList <> Nil then 
    FreeAndNil(FSXpgList);
  if FDXpgList <> Nil then 
    FreeAndNil(FDXpgList);
  FCount := 2147483632;
end;

procedure TCT_GroupItems.Assign(AItem: TCT_GroupItems);
begin
end;

procedure TCT_GroupItems.CopyTo(AItem: TCT_GroupItems);
begin
end;

function  TCT_GroupItems.Create_MXpgList: TCT_MissingXpgList;
begin
  if FMXpgList = Nil then
      FMXpgList := TCT_MissingXpgList.Create(FOwner);
  Result := FMXpgList ;
end;

function  TCT_GroupItems.Create_NXpgList: TCT_NumberXpgList;
begin
  if FNXpgList = Nil then
      FNXpgList := TCT_NumberXpgList.Create(FOwner);
  Result := FNXpgList ;
end;

function  TCT_GroupItems.Create_BXpgList: TCT_BooleanXpgList;
begin
  if FBXpgList = Nil then
      FBXpgList := TCT_BooleanXpgList.Create(FOwner);
  Result := FBXpgList ;
end;

function  TCT_GroupItems.Create_EXpgList: TCT_ErrorXpgList;
begin
  if FEXpgList = Nil then
      FEXpgList := TCT_ErrorXpgList.Create(FOwner);
  Result := FEXpgList ;
end;

function  TCT_GroupItems.Create_SXpgList: TCT_StringXpgList;
begin
  if FSXpgList = Nil then
      FSXpgList := TCT_StringXpgList.Create(FOwner);
  Result := FSXpgList ;
end;

function  TCT_GroupItems.Create_DXpgList: TCT_DateTimeXpgList;
begin
  if FDXpgList = Nil then
      FDXpgList := TCT_DateTimeXpgList.Create(FOwner);
  Result := FDXpgList ;
end;

{ TCT_FieldUsage }

function  TCT_FieldUsage.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_FieldUsage.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FieldUsage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('x',XmlIntToStr(FX));
end;

procedure TCT_FieldUsage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'x' then 
    FX := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FieldUsage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FX := 2147483632;
end;

destructor TCT_FieldUsage.Destroy;
begin
end;

procedure TCT_FieldUsage.Clear;
begin
  FAssigneds := [];
  FX := 2147483632;
end;

procedure TCT_FieldUsage.Assign(AItem: TCT_FieldUsage);
begin
end;

procedure TCT_FieldUsage.CopyTo(AItem: TCT_FieldUsage);
begin
end;

{ TCT_FieldUsageXpgList }

function  TCT_FieldUsageXpgList.GetItems(Index: integer): TCT_FieldUsage;
begin
  Result := TCT_FieldUsage(inherited Items[Index]);
end;

function  TCT_FieldUsageXpgList.Add: TCT_FieldUsage;
begin
  Result := TCT_FieldUsage.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FieldUsageXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FieldUsageXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_FieldUsageXpgList.Assign(AItem: TCT_FieldUsageXpgList);
begin
end;

procedure TCT_FieldUsageXpgList.CopyTo(AItem: TCT_FieldUsageXpgList);
begin
end;

{ TCT_GroupLevel }

function  TCT_GroupLevel.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FGroups <> Nil then 
    Inc(ElemsAssigned,FGroups.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_GroupLevel.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002A0: begin
      if FGroups = Nil then 
        FGroups := TCT_Groups.Create(FOwner);
      Result := FGroups;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupLevel.Write(AWriter: TXpgWriteXML);
begin
  if (FGroups <> Nil) and FGroups.Assigned then 
  begin
    FGroups.WriteAttributes(AWriter);
    if xaElements in FGroups.FAssigneds then 
    begin
      AWriter.BeginTag('groups');
      FGroups.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('groups');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_GroupLevel.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('uniqueName',FUniqueName);
  AWriter.AddAttribute('caption',FCaption);
  if FUser <> False then 
    AWriter.AddAttribute('user',XmlBoolToStr(FUser));
  if FCustomRollUp <> False then 
    AWriter.AddAttribute('customRollUp',XmlBoolToStr(FCustomRollUp));
end;

procedure TCT_GroupLevel.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000418: FUniqueName := AAttributes.Values[i];
      $000002EE: FCaption := AAttributes.Values[i];
      $000001BF: FUser := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F9: FCustomRollUp := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_GroupLevel.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 4;
  FUser := False;
  FCustomRollUp := False;
end;

destructor TCT_GroupLevel.Destroy;
begin
  if FGroups <> Nil then 
    FGroups.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_GroupLevel.Clear;
begin
  FAssigneds := [];
  if FGroups <> Nil then 
    FreeAndNil(FGroups);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FUniqueName := '';
  FCaption := '';
  FUser := False;
  FCustomRollUp := False;
end;

procedure TCT_GroupLevel.Assign(AItem: TCT_GroupLevel);
begin
end;

procedure TCT_GroupLevel.CopyTo(AItem: TCT_GroupLevel);
begin
end;

function  TCT_GroupLevel.Create_Groups: TCT_Groups;
begin
  if FGroups = Nil then
      FGroups := TCT_Groups.Create(FOwner);
  Result := FGroups ;
end;

function  TCT_GroupLevel.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_GroupLevelXpgList }

function  TCT_GroupLevelXpgList.GetItems(Index: integer): TCT_GroupLevel;
begin
  Result := TCT_GroupLevel(inherited Items[Index]);
end;

function  TCT_GroupLevelXpgList.Add: TCT_GroupLevel;
begin
  Result := TCT_GroupLevel.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GroupLevelXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GroupLevelXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_GroupLevelXpgList.Assign(AItem: TCT_GroupLevelXpgList);
begin
end;

procedure TCT_GroupLevelXpgList.CopyTo(AItem: TCT_GroupLevelXpgList);
begin
end;

{ TCT_Items }

function  TCT_Items.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;

  if Count > 0 then begin
    FAssigneds := [xaAttributes,xaElements];

    for i := 0 to FItems.Count - 1 do
      Inc(Result,Items[i].CheckAssigned);
  end;
end;

function  TCT_Items.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'item' then
    Result := Add
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_Items.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    Items[i].WriteAttributes(AWriter);
    AWriter.SimpleTag('item');
  end;
end;

procedure TCT_Items.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('count',XmlIntToStr(FItems.Count));
end;

procedure TCT_Items.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin

end;

constructor TCT_Items.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;

  FItems := TObjectList.Create;
end;

destructor TCT_Items.Destroy;
begin
  FItems.Free;
end;

procedure TCT_Items.Exchange(Aindex1, AIndex2: integer);
begin
  FItems.Exchange(AIndex1,AIndex2);
end;

function TCT_Items.GetItems(Index: integer): TCT_Item;
begin
  Result := TCT_Item(FItems[Index]);
end;

procedure TCT_Items.Clear;
begin
  FItems.Clear;
end;

function TCT_Items.Add: TCT_Item;
begin
  Result := TCT_Item.Create(FOwner);

  FItems.Add(Result);
end;

procedure TCT_Items.Assign(AItem: TCT_Items);
begin
end;

procedure TCT_Items.CopyTo(AItem: TCT_Items);
begin
end;

function TCT_Items.Count: integer;
begin
  Result := FItems.Count;
end;

{ TCT_AutoSortScope }

function  TCT_AutoSortScope.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPivotArea <> Nil then
    Inc(ElemsAssigned,FPivotArea.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_AutoSortScope.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pivotArea' then 
  begin
    if FPivotArea = Nil then 
      FPivotArea := TCT_PivotArea.Create(FOwner);
    Result := FPivotArea;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_AutoSortScope.Write(AWriter: TXpgWriteXML);
begin
  if (FPivotArea <> Nil) and FPivotArea.Assigned then 
  begin
    FPivotArea.WriteAttributes(AWriter);
    if xaElements in FPivotArea.FAssigneds then 
    begin
      AWriter.BeginTag('pivotArea');
      FPivotArea.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotArea');
  end
  else 
    AWriter.SimpleTag('pivotArea');
end;

constructor TCT_AutoSortScope.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
end;

destructor TCT_AutoSortScope.Destroy;
begin
  if FPivotArea <> Nil then 
    FPivotArea.Free;
end;

procedure TCT_AutoSortScope.Clear;
begin
  FAssigneds := [];
  if FPivotArea <> Nil then 
    FreeAndNil(FPivotArea);
end;

procedure TCT_AutoSortScope.Assign(AItem: TCT_AutoSortScope);
begin
end;

procedure TCT_AutoSortScope.CopyTo(AItem: TCT_AutoSortScope);
begin
end;

function  TCT_AutoSortScope.Create_PivotArea: TCT_PivotArea;
begin
  if FPivotArea = Nil then
      FPivotArea := TCT_PivotArea.Create(FOwner);
  Result := FPivotArea ;
end;

{ TCT_PivotAreas }

function  TCT_PivotAreas.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FPivotAreaXpgList <> Nil then 
    Inc(ElemsAssigned,FPivotAreaXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotAreas.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pivotArea' then 
  begin
    if FPivotAreaXpgList = Nil then 
      FPivotAreaXpgList := TCT_PivotAreaXpgList.Create(FOwner);
    Result := FPivotAreaXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotAreas.Write(AWriter: TXpgWriteXML);
begin
  if FPivotAreaXpgList <> Nil then 
    FPivotAreaXpgList.Write(AWriter,'pivotArea');
end;

procedure TCT_PivotAreas.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PivotAreas.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PivotAreas.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PivotAreas.Destroy;
begin
  if FPivotAreaXpgList <> Nil then 
    FPivotAreaXpgList.Free;
end;

procedure TCT_PivotAreas.Clear;
begin
  FAssigneds := [];
  if FPivotAreaXpgList <> Nil then 
    FreeAndNil(FPivotAreaXpgList);
  FCount := 2147483632;
end;

procedure TCT_PivotAreas.Assign(AItem: TCT_PivotAreas);
begin
end;

procedure TCT_PivotAreas.CopyTo(AItem: TCT_PivotAreas);
begin
end;

function  TCT_PivotAreas.Create_PivotAreaXpgList: TCT_PivotAreaXpgList;
begin
  if FPivotAreaXpgList = Nil then
      FPivotAreaXpgList := TCT_PivotAreaXpgList.Create(FOwner);
  Result := FPivotAreaXpgList ;
end;

{ TCT_MemberProperties }

function  TCT_MemberProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMpXpgList <> Nil then 
    Inc(ElemsAssigned,FMpXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_MemberProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'mp' then 
  begin
    if FMpXpgList = Nil then 
      FMpXpgList := TCT_MemberPropertyXpgList.Create(FOwner);
    Result := FMpXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_MemberProperties.Write(AWriter: TXpgWriteXML);
begin
  if FMpXpgList <> Nil then 
    FMpXpgList.Write(AWriter,'mp');
end;

procedure TCT_MemberProperties.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_MemberProperties.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_MemberProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_MemberProperties.Destroy;
begin
  if FMpXpgList <> Nil then 
    FMpXpgList.Free;
end;

procedure TCT_MemberProperties.Clear;
begin
  FAssigneds := [];
  if FMpXpgList <> Nil then 
    FreeAndNil(FMpXpgList);
  FCount := 2147483632;
end;

procedure TCT_MemberProperties.Assign(AItem: TCT_MemberProperties);
begin
end;

procedure TCT_MemberProperties.CopyTo(AItem: TCT_MemberProperties);
begin
end;

function  TCT_MemberProperties.Create_MpXpgList: TCT_MemberPropertyXpgList;
begin
  if FMpXpgList = Nil then
      FMpXpgList := TCT_MemberPropertyXpgList.Create(FOwner);
  Result := FMpXpgList ;
end;

{ TCT_Members }

function  TCT_Members.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FLevel <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMemberXpgList <> Nil then 
    Inc(ElemsAssigned,FMemberXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Members.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'member' then 
  begin
    if FMemberXpgList = Nil then 
      FMemberXpgList := TCT_MemberXpgList.Create(FOwner);
    Result := FMemberXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Members.Write(AWriter: TXpgWriteXML);
begin
  if FMemberXpgList <> Nil then 
    FMemberXpgList.Write(AWriter,'member');
end;

procedure TCT_Members.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if FLevel <> 2147483632 then 
    AWriter.AddAttribute('level',XmlIntToStr(FLevel));
end;

procedure TCT_Members.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000218: FLevel := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Members.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FCount := 2147483632;
  FLevel := 2147483632;
end;

destructor TCT_Members.Destroy;
begin
  if FMemberXpgList <> Nil then 
    FMemberXpgList.Free;
end;

procedure TCT_Members.Clear;
begin
  FAssigneds := [];
  if FMemberXpgList <> Nil then 
    FreeAndNil(FMemberXpgList);
  FCount := 2147483632;
  FLevel := 2147483632;
end;

procedure TCT_Members.Assign(AItem: TCT_Members);
begin
end;

procedure TCT_Members.CopyTo(AItem: TCT_Members);
begin
end;

function  TCT_Members.Create_MemberXpgList: TCT_MemberXpgList;
begin
  if FMemberXpgList = Nil then
      FMemberXpgList := TCT_MemberXpgList.Create(FOwner);
  Result := FMemberXpgList ;
end;

{ TCT_MembersXpgList }

function  TCT_MembersXpgList.GetItems(Index: integer): TCT_Members;
begin
  Result := TCT_Members(inherited Items[Index]);
end;

function  TCT_MembersXpgList.Add: TCT_Members;
begin
  Result := TCT_Members.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MembersXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MembersXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_MembersXpgList.Assign(AItem: TCT_MembersXpgList);
begin
end;

procedure TCT_MembersXpgList.CopyTo(AItem: TCT_MembersXpgList);
begin
end;

{ TCT_Pages }

function  TCT_Pages.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FPageXpgList <> Nil then 
    Inc(ElemsAssigned,FPageXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Pages.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'page' then 
  begin
    if FPageXpgList = Nil then 
      FPageXpgList := TCT_PCDSCPageXpgList.Create(FOwner);
    Result := FPageXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Pages.Write(AWriter: TXpgWriteXML);
begin
  if FPageXpgList <> Nil then 
    FPageXpgList.Write(AWriter,'page');
end;

procedure TCT_Pages.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Pages.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Pages.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_Pages.Destroy;
begin
  if FPageXpgList <> Nil then 
    FPageXpgList.Free;
end;

procedure TCT_Pages.Clear;
begin
  FAssigneds := [];
  if FPageXpgList <> Nil then 
    FreeAndNil(FPageXpgList);
  FCount := 2147483632;
end;

procedure TCT_Pages.Assign(AItem: TCT_Pages);
begin
end;

procedure TCT_Pages.CopyTo(AItem: TCT_Pages);
begin
end;

function  TCT_Pages.Create_PageXpgList: TCT_PCDSCPageXpgList;
begin
  if FPageXpgList = Nil then
      FPageXpgList := TCT_PCDSCPageXpgList.Create(FOwner);
  Result := FPageXpgList ;
end;

{ TCT_RangeSets }

function  TCT_RangeSets.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FRangeSetXpgList <> Nil then 
    Inc(ElemsAssigned,FRangeSetXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_RangeSets.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'rangeSet' then 
  begin
    if FRangeSetXpgList = Nil then 
      FRangeSetXpgList := TCT_RangeSetXpgList.Create(FOwner);
    Result := FRangeSetXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_RangeSets.Write(AWriter: TXpgWriteXML);
begin
  if FRangeSetXpgList <> Nil then 
    FRangeSetXpgList.Write(AWriter,'rangeSet');
end;

procedure TCT_RangeSets.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_RangeSets.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RangeSets.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_RangeSets.Destroy;
begin
  if FRangeSetXpgList <> Nil then 
    FRangeSetXpgList.Free;
end;

procedure TCT_RangeSets.Clear;
begin
  FAssigneds := [];
  if FRangeSetXpgList <> Nil then 
    FreeAndNil(FRangeSetXpgList);
  FCount := 2147483632;
end;

procedure TCT_RangeSets.Assign(AItem: TCT_RangeSets);
begin
end;

procedure TCT_RangeSets.CopyTo(AItem: TCT_RangeSets);
begin
end;

function  TCT_RangeSets.Create_RangeSetXpgList: TCT_RangeSetXpgList;
begin
  if FRangeSetXpgList = Nil then
      FRangeSetXpgList := TCT_RangeSetXpgList.Create(FOwner);
  Result := FRangeSetXpgList ;
end;

{ TCT_SharedItems }

function  TCT_SharedItems.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FContainsSemiMixedTypes <> False then
    Inc(AttrsAssigned);
  if FContainsNonDate <> True then
    Inc(AttrsAssigned);
  if FValues.DateCount > 0 then
    Inc(AttrsAssigned);
  if FValues.StringCount > 0 then
    Inc(AttrsAssigned);
  if FValues.BlankCount > 0 then
    Inc(AttrsAssigned);
  if FContainsMixedTypes <> False then
    Inc(AttrsAssigned);
  if FValues.NumericCount > 0 then
    Inc(AttrsAssigned);
  if FValues.IntegerCount > 0 then
    Inc(AttrsAssigned);

  Inc(ElemsAssigned,FValues.Count);

  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_SharedItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := FValues;

  FValues.Used := True;

  case AReader.QNameHashA of
    // Add blank here as blanks might not have any attributes.
    $0000006D: begin
      FValues.FCurrType := xsivtBlank;
      FValues.Add;
    end;
    $0000006E: FValues.FCurrType := xsivtNumeric;
    $00000062: FValues.FCurrType := xsivtBoolean;
    $00000065: FValues.FCurrType := xsivtError;
    $00000073: FValues.FCurrType := xsivtString;
    $00000064: FValues.FCurrType := xsivtDate;
  end;

  Result.Assigneds := [xaRead];
end;

procedure TCT_SharedItems.UpdateBlank;
begin
  FValues.Add;
end;

procedure TCT_SharedItems.UpdateBoolean(AValue: boolean);
begin
  FValues.AddBoolean(AValue);
end;

procedure TCT_SharedItems.UpdateDate(AValue: TDateTime);
begin
  FValues.AddDateTime(AValue);
end;

procedure TCT_SharedItems.UpdateNum(AValue: double);
begin
  FValues.AddFloat(AValue);
end;

procedure TCT_SharedItems.UpdateString(AValue: AxUCString);
begin
  FValues.AddString(AValue);
end;

procedure TCT_SharedItems.Write(AWriter: TXpgWriteXML);
begin
  FValues.Write(AWriter);
end;

procedure TCT_SharedItems.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FValues.DateCount > 0 then begin
//    AWriter.AddAttribute('containsString',XmlBoolToStr(FValues.StringCount > 0));
//    AWriter.AddAttribute('containsSemiMixedTypes',XmlBoolToStr(FContainsSemiMixedTypes));
    AWriter.AddAttribute('containsNonDate',XmlBoolToStr(FContainsNonDate));

    AWriter.AddAttribute('containsDate',XmlBoolToStr(True));
  end;

  if FValues.BlankCount > 0 then
    AWriter.AddAttribute('containsBlank',XmlBoolToStr(True));

  AWriter.AddAttribute('containsString',XmlBoolToStr(FValues.StringCount > 0));

  if (FValues.IntegerCount > 0) and (FValues.StringCount > 0) then
    AWriter.AddAttribute('containsMixedTypes',XmlBoolToStr(True));

  if FValues.NumericCount > 0 then begin
//    AWriter.AddAttribute('containsString',XmlBoolToStr(FValues.StringCount > 0));
//    AWriter.AddAttribute('containsSemiMixedTypes',XmlBoolToStr(FContainsSemiMixedTypes));

    AWriter.AddAttribute('containsNumber',XmlBoolToStr(True));
  end;
  if (FValues.IntegerCount > 0) and (FValues.IntegerCount = FValues.NumericCount) then
    AWriter.AddAttribute('containsInteger',XmlBoolToStr(True));

  if FValues.NumericCount > 0 then begin
    AWriter.AddAttribute('minValue',XmlFloatToStr(FValues.MinValue));
    AWriter.AddAttribute('maxValue',XmlFloatToStr(FValues.MaxValue));
  end;
  if FValues.DateCount > 0 then begin
    AWriter.AddAttribute('minDate',XmlDateTimeToStr(FValues.MinDate));
    AWriter.AddAttribute('maxDate',XmlDateTimeToStr(FValues.MaxDate));
  end;

  if FValues.Count > 0 then
    AWriter.AddAttribute('count',XmlIntToStr(FValues.Count));
  if FValues.LongText <> False then
    AWriter.AddAttribute('longText',XmlBoolToStr(FValues.LongText));
end;

procedure TCT_SharedItems.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000008F9: FContainsSemiMixedTypes := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000608: FContainsNonDate := XmlStrToBoolDef(AAttributes.Values[i],True);
//      $000004DD: FContainsDate := XmlStrToBoolDef(AAttributes.Values[i],False);
//      $000005D6: FContainsString := XmlStrToBoolDef(AAttributes.Values[i],True);
//      $00000547: FContainsBlank := XmlStrToBoolDef(AAttributes.Values[i],False);
//      $0000076B: FContainsMixedTypes := XmlStrToBoolDef(AAttributes.Values[i],False);
//      $000005C8: FContainsNumber := XmlStrToBoolDef(AAttributes.Values[i],False);
//      $0000062D: FContainsInteger := XmlStrToBoolDef(AAttributes.Values[i],False);
//      $00000341: FMinValue := XmlStrToFloatDef(AAttributes.Values[i],0);
//      $00000343: FMaxValue := XmlStrToFloatDef(AAttributes.Values[i],0);
//      $000002C2: FMinDate := XmlStrToDateTime(AAttributes.Values[i]);
//      $000002C4: FMaxDate := XmlStrToDateTime(AAttributes.Values[i]);
//      $00000355: FLongText := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_SharedItems.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 14;
  FContainsSemiMixedTypes := False;
  FContainsNonDate := False;

  FValues := TXLSSharedItemsValuesHash.Create;
end;

destructor TCT_SharedItems.Destroy;
begin
  FValues.Free;
end;

procedure TCT_SharedItems.Clear;
begin
  FAssigneds := [];
  FContainsSemiMixedTypes := False;
  FContainsNonDate := False;

  FValues.Clear;
end;

procedure TCT_SharedItems.Assign(AItem: TCT_SharedItems);
begin
end;

procedure TCT_SharedItems.CopyTo(AItem: TCT_SharedItems);
begin
end;

{ TCT_FieldGroup }

function  TCT_FieldGroup.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPar <> 2147483632 then
    Inc(AttrsAssigned);
  if FBase <> 2147483632 then
    Inc(AttrsAssigned);
  if FRangePr <> Nil then
    Inc(ElemsAssigned,FRangePr.CheckAssigned);
  if FDiscretePr <> Nil then
    Inc(ElemsAssigned,FDiscretePr.CheckAssigned);
  if FGroupItems <> Nil then
    Inc(ElemsAssigned,FGroupItems.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_FieldGroup.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002CF: begin
      if FRangePr = Nil then 
        FRangePr := TCT_RangePr.Create(FOwner);
      Result := FRangePr;
    end;
    $00000415: begin
      if FDiscretePr = Nil then 
        FDiscretePr := TCT_DiscretePr.Create(FOwner);
      Result := FDiscretePr;
    end;
    $0000042F: begin
      if FGroupItems = Nil then 
        FGroupItems := TCT_GroupItems.Create(FOwner);
      Result := FGroupItems;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FieldGroup.Write(AWriter: TXpgWriteXML);
begin
  if (FRangePr <> Nil) and FRangePr.Assigned then 
  begin
    FRangePr.WriteAttributes(AWriter);
    AWriter.SimpleTag('rangePr');
  end;
  if (FDiscretePr <> Nil) and FDiscretePr.Assigned then 
  begin
    FDiscretePr.WriteAttributes(AWriter);
    if xaElements in FDiscretePr.FAssigneds then 
    begin
      AWriter.BeginTag('discretePr');
      FDiscretePr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('discretePr');
  end;
  if (FGroupItems <> Nil) and FGroupItems.Assigned then 
  begin
    FGroupItems.WriteAttributes(AWriter);
    if xaElements in FGroupItems.FAssigneds then 
    begin
      AWriter.BeginTag('groupItems');
      FGroupItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('groupItems');
  end;
end;

procedure TCT_FieldGroup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPar <> 2147483632 then 
    AWriter.AddAttribute('par',XmlIntToStr(FPar));
  if FBase <> 2147483632 then 
    AWriter.AddAttribute('base',XmlIntToStr(FBase));
end;

procedure TCT_FieldGroup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000143: FPar := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000019B: FBase := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_FieldGroup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FPar := 2147483632;
  FBase := 2147483632;
end;

destructor TCT_FieldGroup.Destroy;
begin
  if FRangePr <> Nil then 
    FRangePr.Free;
  if FDiscretePr <> Nil then 
    FDiscretePr.Free;
  if FGroupItems <> Nil then 
    FGroupItems.Free;
end;

procedure TCT_FieldGroup.Clear;
begin
  FAssigneds := [];
  if FRangePr <> Nil then 
    FreeAndNil(FRangePr);
  if FDiscretePr <> Nil then 
    FreeAndNil(FDiscretePr);
  if FGroupItems <> Nil then 
    FreeAndNil(FGroupItems);
  FPar := 2147483632;
  FBase := 2147483632;
end;

procedure TCT_FieldGroup.Assign(AItem: TCT_FieldGroup);
begin
end;

procedure TCT_FieldGroup.CopyTo(AItem: TCT_FieldGroup);
begin
end;

function  TCT_FieldGroup.Create_RangePr: TCT_RangePr;
begin
  if FRangePr = Nil then
      FRangePr := TCT_RangePr.Create(FOwner);
  Result := FRangePr ;
end;

function  TCT_FieldGroup.Create_DiscretePr: TCT_DiscretePr;
begin
  if FDiscretePr = Nil then
      FDiscretePr := TCT_DiscretePr.Create(FOwner);
  Result := FDiscretePr ;
end;

function  TCT_FieldGroup.Create_GroupItems: TCT_GroupItems;
begin
  if FGroupItems = Nil then
      FGroupItems := TCT_GroupItems.Create(FOwner);
  Result := FGroupItems ;
end;

{ TCT_FieldsUsage }

function  TCT_FieldsUsage.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FFieldUsageXpgList <> Nil then 
    Inc(ElemsAssigned,FFieldUsageXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_FieldsUsage.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'fieldUsage' then 
  begin
    if FFieldUsageXpgList = Nil then 
      FFieldUsageXpgList := TCT_FieldUsageXpgList.Create(FOwner);
    Result := FFieldUsageXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_FieldsUsage.Write(AWriter: TXpgWriteXML);
begin
  if FFieldUsageXpgList <> Nil then 
    FFieldUsageXpgList.Write(AWriter,'fieldUsage');
end;

procedure TCT_FieldsUsage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_FieldsUsage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FieldsUsage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_FieldsUsage.Destroy;
begin
  if FFieldUsageXpgList <> Nil then 
    FFieldUsageXpgList.Free;
end;

procedure TCT_FieldsUsage.Clear;
begin
  FAssigneds := [];
  if FFieldUsageXpgList <> Nil then 
    FreeAndNil(FFieldUsageXpgList);
  FCount := 2147483632;
end;

procedure TCT_FieldsUsage.Assign(AItem: TCT_FieldsUsage);
begin
end;

procedure TCT_FieldsUsage.CopyTo(AItem: TCT_FieldsUsage);
begin
end;

function  TCT_FieldsUsage.Create_FieldUsageXpgList: TCT_FieldUsageXpgList;
begin
  if FFieldUsageXpgList = Nil then
      FFieldUsageXpgList := TCT_FieldUsageXpgList.Create(FOwner);
  Result := FFieldUsageXpgList ;
end;

{ TCT_GroupLevels }

function  TCT_GroupLevels.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FGroupLevelXpgList <> Nil then 
    Inc(ElemsAssigned,FGroupLevelXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GroupLevels.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'groupLevel' then 
  begin
    if FGroupLevelXpgList = Nil then 
      FGroupLevelXpgList := TCT_GroupLevelXpgList.Create(FOwner);
    Result := FGroupLevelXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_GroupLevels.Write(AWriter: TXpgWriteXML);
begin
  if FGroupLevelXpgList <> Nil then 
    FGroupLevelXpgList.Write(AWriter,'groupLevel');
end;

procedure TCT_GroupLevels.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_GroupLevels.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GroupLevels.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_GroupLevels.Destroy;
begin
  if FGroupLevelXpgList <> Nil then 
    FGroupLevelXpgList.Free;
end;

procedure TCT_GroupLevels.Clear;
begin
  FAssigneds := [];
  if FGroupLevelXpgList <> Nil then 
    FreeAndNil(FGroupLevelXpgList);
  FCount := 2147483632;
end;

procedure TCT_GroupLevels.Assign(AItem: TCT_GroupLevels);
begin
end;

procedure TCT_GroupLevels.CopyTo(AItem: TCT_GroupLevels);
begin
end;

function  TCT_GroupLevels.Create_GroupLevelXpgList: TCT_GroupLevelXpgList;
begin
  if FGroupLevelXpgList = Nil then
      FGroupLevelXpgList := TCT_GroupLevelXpgList.Create(FOwner);
  Result := FGroupLevelXpgList ;
end;

{ TCT_Set }

function  TCT_Set.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FTplsXpgList <> Nil then 
    Inc(ElemsAssigned,FTplsXpgList.CheckAssigned);
  if FSortByTuple <> Nil then 
    Inc(ElemsAssigned,FSortByTuple.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Set.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001C3: begin
      if FTplsXpgList = Nil then 
        FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
      Result := FTplsXpgList.Add;
    end;
    $0000048D: begin
      if FSortByTuple = Nil then 
        FSortByTuple := TCT_Tuples.Create(FOwner);
      Result := FSortByTuple;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Set.Write(AWriter: TXpgWriteXML);
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Write(AWriter,'tpls');
  if (FSortByTuple <> Nil) and FSortByTuple.Assigned then 
  begin
    FSortByTuple.WriteAttributes(AWriter);
    if xaElements in FSortByTuple.FAssigneds then 
    begin
      AWriter.BeginTag('sortByTuple');
      FSortByTuple.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sortByTuple');
  end;
end;

procedure TCT_Set.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  AWriter.AddAttribute('maxRank',XmlIntToStr(FMaxRank));
  AWriter.AddAttribute('setDefinition',FSetDefinition);
  if FSortType <> ststNone then 
    AWriter.AddAttribute('sortType',StrTST_SortType[Integer(FSortType)]);
  if FQueryFailed <> False then 
    AWriter.AddAttribute('queryFailed',XmlBoolToStr(FQueryFailed));
end;

procedure TCT_Set.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002D2: FMaxRank := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000555: FSetDefinition := AAttributes.Values[i];
      $0000036A: FSortType := TST_SortType(StrToEnum('stst' + AAttributes.Values[i]));
      $0000047B: FQueryFailed := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Set.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 5;
  FCount := 2147483632;
  FMaxRank := 2147483632;
  FSortType := TST_SortType(XPG_UNKNOWN_ENUM);
  FSortType := ststNone;
  FQueryFailed := False;
end;

destructor TCT_Set.Destroy;
begin
  if FTplsXpgList <> Nil then 
    FTplsXpgList.Free;
  if FSortByTuple <> Nil then 
    FSortByTuple.Free;
end;

procedure TCT_Set.Clear;
begin
  FAssigneds := [];
  if FTplsXpgList <> Nil then 
    FreeAndNil(FTplsXpgList);
  if FSortByTuple <> Nil then 
    FreeAndNil(FSortByTuple);
  FCount := 2147483632;
  FMaxRank := 2147483632;
  FSetDefinition := '';
  FSortType := ststNone;
  FQueryFailed := False;
end;

procedure TCT_Set.Assign(AItem: TCT_Set);
begin
end;

procedure TCT_Set.CopyTo(AItem: TCT_Set);
begin
end;

function  TCT_Set.Create_TplsXpgList: TCT_TuplesXpgList;
begin
  if FTplsXpgList = Nil then
      FTplsXpgList := TCT_TuplesXpgList.Create(FOwner);
  Result := FTplsXpgList ;
end;

function  TCT_Set.Create_SortByTuple: TCT_Tuples;
begin
  if FSortByTuple = Nil then
      FSortByTuple := TCT_Tuples.Create(FOwner);
  Result := FSortByTuple ;
end;

{ TCT_SetXpgList }

function  TCT_SetXpgList.GetItems(Index: integer): TCT_Set;
begin
  Result := TCT_Set(inherited Items[Index]);
end;

function  TCT_SetXpgList.Add: TCT_Set;
begin
  Result := TCT_Set.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SetXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SetXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_SetXpgList.Assign(AItem: TCT_SetXpgList);
begin
end;

procedure TCT_SetXpgList.CopyTo(AItem: TCT_SetXpgList);
begin
end;

{ TCT_Query }

function  TCT_Query.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FTpls <> Nil then 
    Inc(ElemsAssigned,FTpls.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Query.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'tpls' then 
  begin
    if FTpls = Nil then 
      FTpls := TCT_Tuples.Create(FOwner);
    Result := FTpls;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Query.Write(AWriter: TXpgWriteXML);
begin
  if (FTpls <> Nil) and FTpls.Assigned then 
  begin
    FTpls.WriteAttributes(AWriter);
    if xaElements in FTpls.FAssigneds then 
    begin
      AWriter.BeginTag('tpls');
      FTpls.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('tpls');
  end;
end;

procedure TCT_Query.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('mdx',FMdx);
end;

procedure TCT_Query.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'mdx' then 
    FMdx := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Query.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
end;

destructor TCT_Query.Destroy;
begin
  if FTpls <> Nil then 
    FTpls.Free;
end;

procedure TCT_Query.Clear;
begin
  FAssigneds := [];
  if FTpls <> Nil then 
    FreeAndNil(FTpls);
  FMdx := '';
end;

procedure TCT_Query.Assign(AItem: TCT_Query);
begin
end;

procedure TCT_Query.CopyTo(AItem: TCT_Query);
begin
end;

function  TCT_Query.Create_Tpls: TCT_Tuples;
begin
  if FTpls = Nil then
      FTpls := TCT_Tuples.Create(FOwner);
  Result := FTpls ;
end;

{ TCT_QueryXpgList }

function  TCT_QueryXpgList.GetItems(Index: integer): TCT_Query;
begin
  Result := TCT_Query(inherited Items[Index]);
end;

function  TCT_QueryXpgList.Add: TCT_Query;
begin
  Result := TCT_Query.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_QueryXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_QueryXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_QueryXpgList.Assign(AItem: TCT_QueryXpgList);
begin
end;

procedure TCT_QueryXpgList.CopyTo(AItem: TCT_QueryXpgList);
begin
end;

{ TCT_ServerFormat }

function  TCT_ServerFormat.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCulture <> '' then 
    Inc(AttrsAssigned);
  if FFormat <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ServerFormat.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ServerFormat.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCulture <> '' then 
    AWriter.AddAttribute('culture',FCulture);
  if FFormat <> '' then 
    AWriter.AddAttribute('format',FFormat);
end;

procedure TCT_ServerFormat.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000304: FCulture := AAttributes.Values[i];
      $00000289: FFormat := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ServerFormat.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_ServerFormat.Destroy;
begin
end;

procedure TCT_ServerFormat.Clear;
begin
  FAssigneds := [];
  FCulture := '';
  FFormat := '';
end;

procedure TCT_ServerFormat.Assign(AItem: TCT_ServerFormat);
begin
end;

procedure TCT_ServerFormat.CopyTo(AItem: TCT_ServerFormat);
begin
end;

{ TCT_ServerFormatXpgList }

function  TCT_ServerFormatXpgList.GetItems(Index: integer): TCT_ServerFormat;
begin
  Result := TCT_ServerFormat(inherited Items[Index]);
end;

function  TCT_ServerFormatXpgList.Add: TCT_ServerFormat;
begin
  Result := TCT_ServerFormat.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ServerFormatXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ServerFormatXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ServerFormatXpgList.Assign(AItem: TCT_ServerFormatXpgList);
begin
end;

procedure TCT_ServerFormatXpgList.CopyTo(AItem: TCT_ServerFormatXpgList);
begin
end;

{ TCT_PivotField }

function  TCT_PivotField.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if Integer(FAxis) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FDataField <> False then 
    Inc(AttrsAssigned);
  if FSubtotalCaption <> '' then 
    Inc(AttrsAssigned);
  if FShowDropDowns <> True then 
    Inc(AttrsAssigned);
  if FHiddenLevel <> False then 
    Inc(AttrsAssigned);
  if FUniqueMemberProperty <> '' then 
    Inc(AttrsAssigned);
  if FCompact <> True then 
    Inc(AttrsAssigned);
  if FAllDrilled <> False then 
    Inc(AttrsAssigned);
  if FNumFmtId <> '' then 
    Inc(AttrsAssigned);
  if FOutline <> True then 
    Inc(AttrsAssigned);
  if FSubtotalTop <> True then 
    Inc(AttrsAssigned);
  if FDragToRow <> True then 
    Inc(AttrsAssigned);
  if FDragToCol <> True then 
    Inc(AttrsAssigned);
  if FMultipleItemSelectionAllowed <> False then 
    Inc(AttrsAssigned);
  if FDragToPage <> True then 
    Inc(AttrsAssigned);
  if FDragToData <> True then 
    Inc(AttrsAssigned);
  if FDragOff <> True then 
    Inc(AttrsAssigned);
  if FShowAll <> True then 
    Inc(AttrsAssigned);
  if FInsertBlankRow <> False then 
    Inc(AttrsAssigned);
  if FServerField <> False then 
    Inc(AttrsAssigned);
  if FInsertPageBreak <> False then 
    Inc(AttrsAssigned);
  if FAutoShow <> False then 
    Inc(AttrsAssigned);
  if FTopAutoShow <> True then 
    Inc(AttrsAssigned);
  if FHideNewItems <> False then 
    Inc(AttrsAssigned);
  if FMeasureFilter <> False then 
    Inc(AttrsAssigned);
  if FIncludeNewItemsInFilter <> False then 
    Inc(AttrsAssigned);
  if FItemPageCount <> 10 then 
    Inc(AttrsAssigned);
  if FSortType <> stfstManual then 
    Inc(AttrsAssigned);
  if Byte(FDataSourceSort) <> 2 then 
    Inc(AttrsAssigned);
  if FNonAutoSortDefault <> False then 
    Inc(AttrsAssigned);
  if FRankBy <> 2147483632 then 
    Inc(AttrsAssigned);
  if FDefaultSubtotal <> True then 
    Inc(AttrsAssigned);
  if FSumSubtotal <> False then 
    Inc(AttrsAssigned);
  if FCountASubtotal <> False then 
    Inc(AttrsAssigned);
  if FAvgSubtotal <> False then 
    Inc(AttrsAssigned);
  if FMaxSubtotal <> False then 
    Inc(AttrsAssigned);
  if FMinSubtotal <> False then 
    Inc(AttrsAssigned);
  if FProductSubtotal <> False then 
    Inc(AttrsAssigned);
  if FCountSubtotal <> False then 
    Inc(AttrsAssigned);
  if FStdDevSubtotal <> False then 
    Inc(AttrsAssigned);
  if FStdDevPSubtotal <> False then 
    Inc(AttrsAssigned);
  if FVarSubtotal <> False then 
    Inc(AttrsAssigned);
  if FVarPSubtotal <> False then 
    Inc(AttrsAssigned);
  if FShowPropCell <> False then 
    Inc(AttrsAssigned);
  if FShowPropTip <> False then 
    Inc(AttrsAssigned);
  if FShowPropAsCaption <> False then 
    Inc(AttrsAssigned);
  if FDefaultAttributeDrillState <> False then 
    Inc(AttrsAssigned);
  if FItems <> Nil then 
    Inc(ElemsAssigned,FItems.CheckAssigned);
  if FAutoSortScope <> Nil then 
    Inc(ElemsAssigned,FAutoSortScope.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotField.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000222: begin
      if FItems = Nil then 
        FItems := TCT_Items.Create(FOwner);
      Result := FItems;
    end;
    $0000055B: begin
      if FAutoSortScope = Nil then 
        FAutoSortScope := TCT_AutoSortScope.Create(FOwner);
      Result := FAutoSortScope;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotField.Write(AWriter: TXpgWriteXML);
begin
  if (FItems <> Nil) and FItems.Assigned then 
  begin
    FItems.WriteAttributes(AWriter);
    if xaElements in FItems.FAssigneds then 
    begin
      AWriter.BeginTag('items');
      FItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('items');
  end;
  if (FAutoSortScope <> Nil) and FAutoSortScope.Assigned then 
    if xaElements in FAutoSortScope.FAssigneds then 
    begin
      AWriter.BeginTag('autoSortScope');
      FAutoSortScope.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('autoSortScope');
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotField.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if Integer(FAxis) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('axis',StrTST_Axis[Integer(FAxis)]);
  if FDataField <> False then 
    AWriter.AddAttribute('dataField',XmlBoolToStr(FDataField));
  if FSubtotalCaption <> '' then 
    AWriter.AddAttribute('subtotalCaption',FSubtotalCaption);
  if FShowDropDowns <> True then 
    AWriter.AddAttribute('showDropDowns',XmlBoolToStr(FShowDropDowns));
  if FHiddenLevel <> False then 
    AWriter.AddAttribute('hiddenLevel',XmlBoolToStr(FHiddenLevel));
  if FUniqueMemberProperty <> '' then 
    AWriter.AddAttribute('uniqueMemberProperty',FUniqueMemberProperty);
  if FCompact <> True then 
    AWriter.AddAttribute('compact',XmlBoolToStr(FCompact));
  if FAllDrilled <> False then 
    AWriter.AddAttribute('allDrilled',XmlBoolToStr(FAllDrilled));
  if FNumFmtId <> '' then 
    AWriter.AddAttribute('numFmtId',FNumFmtId);
  if FOutline <> True then 
    AWriter.AddAttribute('outline',XmlBoolToStr(FOutline));
  if FSubtotalTop <> True then 
    AWriter.AddAttribute('subtotalTop',XmlBoolToStr(FSubtotalTop));
  if FDragToRow <> True then 
    AWriter.AddAttribute('dragToRow',XmlBoolToStr(FDragToRow));
  if FDragToCol <> True then 
    AWriter.AddAttribute('dragToCol',XmlBoolToStr(FDragToCol));
  if FMultipleItemSelectionAllowed <> False then 
    AWriter.AddAttribute('multipleItemSelectionAllowed',XmlBoolToStr(FMultipleItemSelectionAllowed));
  if FDragToPage <> True then 
    AWriter.AddAttribute('dragToPage',XmlBoolToStr(FDragToPage));
  if FDragToData <> True then 
    AWriter.AddAttribute('dragToData',XmlBoolToStr(FDragToData));
  if FDragOff <> True then 
    AWriter.AddAttribute('dragOff',XmlBoolToStr(FDragOff));
  if FShowAll <> True then 
    AWriter.AddAttribute('showAll',XmlBoolToStr(FShowAll));
  if FInsertBlankRow <> False then 
    AWriter.AddAttribute('insertBlankRow',XmlBoolToStr(FInsertBlankRow));
  if FServerField <> False then 
    AWriter.AddAttribute('serverField',XmlBoolToStr(FServerField));
  if FInsertPageBreak <> False then 
    AWriter.AddAttribute('insertPageBreak',XmlBoolToStr(FInsertPageBreak));
  if FAutoShow <> False then 
    AWriter.AddAttribute('autoShow',XmlBoolToStr(FAutoShow));
  if FTopAutoShow <> True then 
    AWriter.AddAttribute('topAutoShow',XmlBoolToStr(FTopAutoShow));
  if FHideNewItems <> False then 
    AWriter.AddAttribute('hideNewItems',XmlBoolToStr(FHideNewItems));
  if FMeasureFilter <> False then 
    AWriter.AddAttribute('measureFilter',XmlBoolToStr(FMeasureFilter));
  if FIncludeNewItemsInFilter <> False then 
    AWriter.AddAttribute('includeNewItemsInFilter',XmlBoolToStr(FIncludeNewItemsInFilter));
  if FItemPageCount <> 10 then 
    AWriter.AddAttribute('itemPageCount',XmlIntToStr(FItemPageCount));
  if FSortType <> stfstManual then 
    AWriter.AddAttribute('sortType',StrTST_FieldSortType[Integer(FSortType)]);
  if Byte(FDataSourceSort) <> 2 then 
    AWriter.AddAttribute('dataSourceSort',XmlBoolToStr(FDataSourceSort));
  if FNonAutoSortDefault <> False then 
    AWriter.AddAttribute('nonAutoSortDefault',XmlBoolToStr(FNonAutoSortDefault));
  if FRankBy <> 2147483632 then 
    AWriter.AddAttribute('rankBy',XmlIntToStr(FRankBy));
  if FDefaultSubtotal <> True then 
    AWriter.AddAttribute('defaultSubtotal',XmlBoolToStr(FDefaultSubtotal));
  if FSumSubtotal <> False then 
    AWriter.AddAttribute('sumSubtotal',XmlBoolToStr(FSumSubtotal));
  if FCountASubtotal <> False then 
    AWriter.AddAttribute('countASubtotal',XmlBoolToStr(FCountASubtotal));
  if FAvgSubtotal <> False then 
    AWriter.AddAttribute('avgSubtotal',XmlBoolToStr(FAvgSubtotal));
  if FMaxSubtotal <> False then 
    AWriter.AddAttribute('maxSubtotal',XmlBoolToStr(FMaxSubtotal));
  if FMinSubtotal <> False then 
    AWriter.AddAttribute('minSubtotal',XmlBoolToStr(FMinSubtotal));
  if FProductSubtotal <> False then 
    AWriter.AddAttribute('productSubtotal',XmlBoolToStr(FProductSubtotal));
  if FCountSubtotal <> False then 
    AWriter.AddAttribute('countSubtotal',XmlBoolToStr(FCountSubtotal));
  if FStdDevSubtotal <> False then 
    AWriter.AddAttribute('stdDevSubtotal',XmlBoolToStr(FStdDevSubtotal));
  if FStdDevPSubtotal <> False then 
    AWriter.AddAttribute('stdDevPSubtotal',XmlBoolToStr(FStdDevPSubtotal));
  if FVarSubtotal <> False then 
    AWriter.AddAttribute('varSubtotal',XmlBoolToStr(FVarSubtotal));
  if FVarPSubtotal <> False then 
    AWriter.AddAttribute('varPSubtotal',XmlBoolToStr(FVarPSubtotal));
  if FShowPropCell <> False then 
    AWriter.AddAttribute('showPropCell',XmlBoolToStr(FShowPropCell));
  if FShowPropTip <> False then 
    AWriter.AddAttribute('showPropTip',XmlBoolToStr(FShowPropTip));
  if FShowPropAsCaption <> False then 
    AWriter.AddAttribute('showPropAsCaption',XmlBoolToStr(FShowPropAsCaption));
  if FDefaultAttributeDrillState <> False then 
    AWriter.AddAttribute('defaultAttributeDrillState',XmlBoolToStr(FDefaultAttributeDrillState));
end;

procedure TCT_PivotField.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashB[i] of
      $309B299D: FName := AAttributes.Values[i];
      $3DB9DC9B: FAxis := TST_Axis(StrToEnum('sta' + AAttributes.Values[i]));
      $6AE44638: FDataField := XmlStrToBoolDef(AAttributes.Values[i],False);
      $EFB42B20: FSubtotalCaption := AAttributes.Values[i];
      $11656625: FShowDropDowns := XmlStrToBoolDef(AAttributes.Values[i],True);
      $008B7D14: FHiddenLevel := XmlStrToBoolDef(AAttributes.Values[i],False);
      $239E33DE: FUniqueMemberProperty := AAttributes.Values[i];
      $49A1269F: FCompact := XmlStrToBoolDef(AAttributes.Values[i],True);
      $B1533CF9: FAllDrilled := XmlStrToBoolDef(AAttributes.Values[i],False);
      $FAA23D66: FNumFmtId := AAttributes.Values[i];
      $34662C20: FOutline := XmlStrToBoolDef(AAttributes.Values[i],True);
      $8CCB82A5: FSubtotalTop := XmlStrToBoolDef(AAttributes.Values[i],True);
      $564634A9: FDragToRow := XmlStrToBoolDef(AAttributes.Values[i],True);
      $50704265: FDragToCol := XmlStrToBoolDef(AAttributes.Values[i],True);
      $BCB0A5EF: FMultipleItemSelectionAllowed := XmlStrToBoolDef(AAttributes.Values[i],False);
      $572AFA68: FDragToPage := XmlStrToBoolDef(AAttributes.Values[i],True);
      $F346C5DB: FDragToData := XmlStrToBoolDef(AAttributes.Values[i],True);
      $5EE565ED: FDragOff := XmlStrToBoolDef(AAttributes.Values[i],True);
      $FE4BE514: FShowAll := XmlStrToBoolDef(AAttributes.Values[i],True);
      $412E38A1: FInsertBlankRow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $2A481663: FServerField := XmlStrToBoolDef(AAttributes.Values[i],False);
      $A54AD691: FInsertPageBreak := XmlStrToBoolDef(AAttributes.Values[i],False);
      $12891DF2: FAutoShow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $10E25E7B: FTopAutoShow := XmlStrToBoolDef(AAttributes.Values[i],True);
      $31214308: FHideNewItems := XmlStrToBoolDef(AAttributes.Values[i],False);
      $5D5658AE: FMeasureFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $100FE0FB: FIncludeNewItemsInFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $A3CC7EF3: FItemPageCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $E97BF2AE: FSortType := TST_FieldSortType(StrToEnum('stfst' + AAttributes.Values[i]));
      $8A493AA7: FDataSourceSort := XmlStrToBoolDef(AAttributes.Values[i],False);
      $3FF4EFA9: FNonAutoSortDefault := XmlStrToBoolDef(AAttributes.Values[i],False);
      $1D31DEA1: FRankBy := XmlStrToIntDef(AAttributes.Values[i],0);
      $63773259: FDefaultSubtotal := XmlStrToBoolDef(AAttributes.Values[i],True);
      $9AF60007: FSumSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $08247C16: FCountASubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $7395335C: FAvgSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $17A64974: FMaxSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $C462FB72: FMinSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $5C8174E7: FProductSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $3C197E61: FCountSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $1CF9CC40: FStdDevSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $FEF300FA: FStdDevPSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $4BF9149D: FVarSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $65C456D9: FVarPSubtotal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $3E2772A8: FShowPropCell := XmlStrToBoolDef(AAttributes.Values[i],False);
      $1EDFE4B3: FShowPropTip := XmlStrToBoolDef(AAttributes.Values[i],False);
      $D8C61D26: FShowPropAsCaption := XmlStrToBoolDef(AAttributes.Values[i],False);
      $310F5F01: FDefaultAttributeDrillState := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotField.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 48;
  FAxis := TST_Axis(XPG_UNKNOWN_ENUM);
  FDataField := False;
  FShowDropDowns := True;
  FHiddenLevel := False;
  FCompact := True;
  FAllDrilled := False;
  FOutline := True;
  FSubtotalTop := True;
  FDragToRow := True;
  FDragToCol := True;
  FMultipleItemSelectionAllowed := False;
  FDragToPage := True;
  FDragToData := True;
  FDragOff := True;
  FShowAll := True;
  FInsertBlankRow := False;
  FServerField := False;
  FInsertPageBreak := False;
  FAutoShow := False;
  FTopAutoShow := True;
  FHideNewItems := False;
  FMeasureFilter := False;
  FIncludeNewItemsInFilter := False;
  FItemPageCount := 10;
  FSortType := TST_FieldSortType(XPG_UNKNOWN_ENUM);
  FSortType := stfstManual;
  Byte(FDataSourceSort) := 2;
  FNonAutoSortDefault := False;
  FRankBy := 2147483632;
  FDefaultSubtotal := True;
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
  FShowPropCell := False;
  FShowPropTip := False;
  FShowPropAsCaption := False;
  FDefaultAttributeDrillState := False;
end;

destructor TCT_PivotField.Destroy;
begin
  if FItems <> Nil then 
    FItems.Free;
  if FAutoSortScope <> Nil then 
    FAutoSortScope.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PivotField.Clear;
begin
  FAssigneds := [];
  if FItems <> Nil then 
    FreeAndNil(FItems);
  if FAutoSortScope <> Nil then 
    FreeAndNil(FAutoSortScope);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FName := '';
  FDataField := False;
  FSubtotalCaption := '';
  FShowDropDowns := True;
  FHiddenLevel := False;
  FUniqueMemberProperty := '';
  FCompact := True;
  FAllDrilled := False;
  FNumFmtId := '';
  FOutline := True;
  FSubtotalTop := True;
  FDragToRow := True;
  FDragToCol := True;
  FMultipleItemSelectionAllowed := False;
  FDragToPage := True;
  FDragToData := True;
  FDragOff := True;
  FShowAll := True;
  FInsertBlankRow := False;
  FServerField := False;
  FInsertPageBreak := False;
  FAutoShow := False;
  FTopAutoShow := True;
  FHideNewItems := False;
  FMeasureFilter := False;
  FIncludeNewItemsInFilter := False;
  FItemPageCount := 10;
  FSortType := stfstManual;
  Byte(FDataSourceSort) := 2;
  FNonAutoSortDefault := False;
  FRankBy := 2147483632;
  FDefaultSubtotal := True;
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
  FShowPropCell := False;
  FShowPropTip := False;
  FShowPropAsCaption := False;
  FDefaultAttributeDrillState := False;
end;

procedure TCT_PivotField.Assign(AItem: TCT_PivotField);
begin
end;

procedure TCT_PivotField.CopyTo(AItem: TCT_PivotField);
begin
end;

function  TCT_PivotField.Create_Items: TCT_Items;
begin
  if FItems = Nil then
      FItems := TCT_Items.Create(FOwner);
  Result := FItems ;
end;

function  TCT_PivotField.Create_AutoSortScope: TCT_AutoSortScope;
begin
  if FAutoSortScope = Nil then
      FAutoSortScope := TCT_AutoSortScope.Create(FOwner);
  Result := FAutoSortScope ;
end;

function  TCT_PivotField.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PivotFieldXpgList }

function  TCT_PivotFieldXpgList.GetItems(Index: integer): TCT_PivotField;
begin
  Result := TCT_PivotField(inherited Items[Index]);
end;

function  TCT_PivotFieldXpgList.Add: TCT_PivotField;
begin
  Result := TCT_PivotField.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotFieldXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotFieldXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PivotFieldXpgList.Assign(AItem: TCT_PivotFieldXpgList);
begin
end;

procedure TCT_PivotFieldXpgList.CopyTo(AItem: TCT_PivotFieldXpgList);
begin
end;

{ TCT_Field }

function  TCT_Field.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Field.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Field.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('x',XmlIntToStr(FX));
end;

procedure TCT_Field.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'x' then 
    FX := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Field.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FX := 2147483632;
end;

destructor TCT_Field.Destroy;
begin
end;

procedure TCT_Field.Clear;
begin
  FAssigneds := [];
  FX := 2147483632;
end;

procedure TCT_Field.Assign(AItem: TCT_Field);
begin
end;

procedure TCT_Field.CopyTo(AItem: TCT_Field);
begin
end;

{ TCT_I }

function  TCT_I.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FT <> stitData then 
    Inc(AttrsAssigned);
  if FR <> 0 then 
    Inc(AttrsAssigned);
  if FI <> 0 then 
    Inc(AttrsAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_I.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'x' then 
  begin
    if FXXpgList = Nil then 
      FXXpgList := TCT_XXpgList.Create(FOwner);
    Result := FXXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_I.Write(AWriter: TXpgWriteXML);
begin
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

procedure TCT_I.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FT <> stitData then 
    AWriter.AddAttribute('t',StrTST_ItemType[Integer(FT)]);
  if FR <> 0 then 
    AWriter.AddAttribute('r',XmlIntToStr(FR));
  if FI <> 0 then 
    AWriter.AddAttribute('i',XmlIntToStr(FI));
end;

procedure TCT_I.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000074: FT := TST_ItemType(StrToEnum('stit' + AAttributes.Values[i]));
      $00000072: FR := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000069: FI := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_I.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FT := TST_ItemType(XPG_UNKNOWN_ENUM);
  FT := stitData;
  FR := 0;
  FI := 0;
end;

destructor TCT_I.Destroy;
begin
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_I.Clear;
begin
  FAssigneds := [];
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
  FT := stitData;
  FR := 0;
  FI := 0;
end;

procedure TCT_I.Assign(AItem: TCT_I);
begin
end;

procedure TCT_I.CopyTo(AItem: TCT_I);
begin
end;

function  TCT_I.Create_XXpgList: TCT_XXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_XXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_PageField }

function  TCT_PageField.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PageField.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'extLst' then 
  begin
    if FExtLst = Nil then 
      FExtLst := TCT_ExtensionList.Create(FOwner);
    Result := FExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PageField.Write(AWriter: TXpgWriteXML);
begin
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PageField.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('fld',XmlIntToStr(FFld));
  if FItem <> 2147483632 then 
    AWriter.AddAttribute('item',XmlIntToStr(FItem));
  if FHier <> 2147483632 then 
    AWriter.AddAttribute('hier',XmlIntToStr(FHier));
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FCap <> '' then 
    AWriter.AddAttribute('cap',FCap);
end;

procedure TCT_PageField.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000136: FFld := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001AF: FItem := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A8: FHier := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A1: FName := AAttributes.Values[i];
      $00000134: FCap := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PageField.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 5;
  FFld := 2147483632;
  FItem := 2147483632;
  FHier := 2147483632;
end;

destructor TCT_PageField.Destroy;
begin
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PageField.Clear;
begin
  FAssigneds := [];
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FFld := 2147483632;
  FItem := 2147483632;
  FHier := 2147483632;
  FName := '';
  FCap := '';
end;

procedure TCT_PageField.Assign(AItem: TCT_PageField);
begin
end;

procedure TCT_PageField.CopyTo(AItem: TCT_PageField);
begin
end;

function  TCT_PageField.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PageFieldXpgList }

function  TCT_PageFieldXpgList.GetItems(Index: integer): TCT_PageField;
begin
  Result := TCT_PageField(inherited Items[Index]);
end;

function  TCT_PageFieldXpgList.Add: TCT_PageField;
begin
  Result := TCT_PageField.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PageFieldXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PageFieldXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PageFieldXpgList.Assign(AItem: TCT_PageFieldXpgList);
begin
end;

procedure TCT_PageFieldXpgList.CopyTo(AItem: TCT_PageFieldXpgList);
begin
end;

{ TCT_DataField }

function  TCT_DataField.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DataField.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'extLst' then 
  begin
    if FExtLst = Nil then 
      FExtLst := TCT_ExtensionList.Create(FOwner);
    Result := FExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_DataField.Write(AWriter: TXpgWriteXML);
begin
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_DataField.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('fld',XmlIntToStr(FFld));
  if FSubtotal <> '' then 
    AWriter.AddAttribute('subtotal',FSubtotal);
  if FShowDataAs <> stsdaNormal then 
    AWriter.AddAttribute('showDataAs',StrTST_ShowDataAs[Integer(FShowDataAs)]);
  if FBaseField <> -1 then 
    AWriter.AddAttribute('baseField',XmlIntToStr(FBaseField));
  if FBaseItem <> 1048832 then 
    AWriter.AddAttribute('baseItem',XmlIntToStr(FBaseItem));
  if FNumFmtId <> '' then 
    AWriter.AddAttribute('numFmtId',FNumFmtId);
end;

procedure TCT_DataField.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000136: FFld := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000036E: FSubtotal := AAttributes.Values[i];
      $000003EF: FShowDataAs := TST_ShowDataAs(StrToEnum('stsda' + AAttributes.Values[i]));
      $0000037F: FBaseField := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000032A: FBaseItem := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000324: FNumFmtId := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_DataField.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 7;
  FFld := 2147483632;
  FSubtotal := '';
  FShowDataAs := TST_ShowDataAs(XPG_UNKNOWN_ENUM);
  FShowDataAs := stsdaNormal;
  FBaseField := -1;
  FBaseItem := 1048832;
end;

destructor TCT_DataField.Destroy;
begin
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_DataField.Clear;
begin
  FAssigneds := [];
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FName := '';
  FFld := 2147483632;
  FSubtotal := '';
  FShowDataAs := stsdaNormal;
  FBaseField := -1;
  FBaseItem := 1048832;
  FNumFmtId := '';
end;

procedure TCT_DataField.Assign(AItem: TCT_DataField);
begin
end;

procedure TCT_DataField.CopyTo(AItem: TCT_DataField);
begin
end;

function  TCT_DataField.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_DataFieldXpgList }

function  TCT_DataFields.GetItems(Index: integer): TCT_DataField;
begin
  Result := TCT_DataField(inherited Items[Index]);
end;

function  TCT_DataFields.Add: TCT_DataField;
begin
  Result := TCT_DataField.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DataFields.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DataFields.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
  begin
    if xaAttributes in Items[i].FAssigneds then
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_DataFields.Assign(AItem: TCT_DataFields);
begin
end;

procedure TCT_DataFields.CopyTo(AItem: TCT_DataFields);
begin
end;

{ TCT_Format }

function  TCT_Format.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAction <> stfaFormatting then 
    Inc(AttrsAssigned);
  if FDxfId <> '' then 
    Inc(AttrsAssigned);
  if FPivotArea <> Nil then 
    Inc(ElemsAssigned,FPivotArea.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Format.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003AB: begin
      if FPivotArea = Nil then 
        FPivotArea := TCT_PivotArea.Create(FOwner);
      Result := FPivotArea;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Format.Write(AWriter: TXpgWriteXML);
begin
  if (FPivotArea <> Nil) and FPivotArea.Assigned then 
  begin
    FPivotArea.WriteAttributes(AWriter);
    if xaElements in FPivotArea.FAssigneds then 
    begin
      AWriter.BeginTag('pivotArea');
      FPivotArea.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotArea');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Format.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAction <> stfaFormatting then 
    AWriter.AddAttribute('action',StrTST_FormatAction[Integer(FAction)]);
  if FDxfId <> '' then 
    AWriter.AddAttribute('dxfId',FDxfId);
end;

procedure TCT_Format.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000027E: FAction := TST_FormatAction(StrToEnum('stfa' + AAttributes.Values[i]));
      $000001EF: FDxfId := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Format.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 2;
  FAction := TST_FormatAction(XPG_UNKNOWN_ENUM);
  FAction := stfaFormatting;
end;

destructor TCT_Format.Destroy;
begin
  if FPivotArea <> Nil then 
    FPivotArea.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_Format.Clear;
begin
  FAssigneds := [];
  if FPivotArea <> Nil then 
    FreeAndNil(FPivotArea);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FAction := stfaFormatting;
  FDxfId := '';
end;

procedure TCT_Format.Assign(AItem: TCT_Format);
begin
end;

procedure TCT_Format.CopyTo(AItem: TCT_Format);
begin
end;

function  TCT_Format.Create_PivotArea: TCT_PivotArea;
begin
  if FPivotArea = Nil then
      FPivotArea := TCT_PivotArea.Create(FOwner);
  Result := FPivotArea ;
end;

function  TCT_Format.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_FormatXpgList }

function  TCT_FormatXpgList.GetItems(Index: integer): TCT_Format;
begin
  Result := TCT_Format(inherited Items[Index]);
end;

function  TCT_FormatXpgList.Add: TCT_Format;
begin
  Result := TCT_Format.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FormatXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FormatXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_FormatXpgList.Assign(AItem: TCT_FormatXpgList);
begin
end;

procedure TCT_FormatXpgList.CopyTo(AItem: TCT_FormatXpgList);
begin
end;

{ TCT_ConditionalFormat }

function  TCT_ConditionalFormat.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FPivotAreas <> Nil then 
    Inc(ElemsAssigned,FPivotAreas.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ConditionalFormat.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000041E: begin
      if FPivotAreas = Nil then 
        FPivotAreas := TCT_PivotAreas.Create(FOwner);
      Result := FPivotAreas;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ConditionalFormat.Write(AWriter: TXpgWriteXML);
begin
  if (FPivotAreas <> Nil) and FPivotAreas.Assigned then 
  begin
    FPivotAreas.WriteAttributes(AWriter);
    if xaElements in FPivotAreas.FAssigneds then 
    begin
      AWriter.BeginTag('pivotAreas');
      FPivotAreas.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotAreas');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_ConditionalFormat.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FScope <> stsSelection then 
    AWriter.AddAttribute('scope',StrTST_Scope[Integer(FScope)]);
  if FType <> sttNone then 
    AWriter.AddAttribute('type',StrTST_Type[Integer(FType)]);
  AWriter.AddAttribute('priority',XmlIntToStr(FPriority));
end;

procedure TCT_ConditionalFormat.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000021A: FScope := TST_Scope(StrToEnum('sts' + AAttributes.Values[i]));
      $000001C2: FType := TST_Type(StrToEnum('stt' + AAttributes.Values[i]));
      $00000382: FPriority := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ConditionalFormat.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 3;
  FScope := TST_Scope(XPG_UNKNOWN_ENUM);
  FScope := stsSelection;
  FType := TST_Type(XPG_UNKNOWN_ENUM);
  FType := sttNone;
  FPriority := 2147483632;
end;

destructor TCT_ConditionalFormat.Destroy;
begin
  if FPivotAreas <> Nil then 
    FPivotAreas.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_ConditionalFormat.Clear;
begin
  FAssigneds := [];
  if FPivotAreas <> Nil then 
    FreeAndNil(FPivotAreas);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FScope := stsSelection;
  FType := sttNone;
  FPriority := 2147483632;
end;

procedure TCT_ConditionalFormat.Assign(AItem: TCT_ConditionalFormat);
begin
end;

procedure TCT_ConditionalFormat.CopyTo(AItem: TCT_ConditionalFormat);
begin
end;

function  TCT_ConditionalFormat.Create_PivotAreas: TCT_PivotAreas;
begin
  if FPivotAreas = Nil then
      FPivotAreas := TCT_PivotAreas.Create(FOwner);
  Result := FPivotAreas ;
end;

function  TCT_ConditionalFormat.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_ConditionalFormatXpgList }

function  TCT_ConditionalFormatXpgList.GetItems(Index: integer): TCT_ConditionalFormat;
begin
  Result := TCT_ConditionalFormat(inherited Items[Index]);
end;

function  TCT_ConditionalFormatXpgList.Add: TCT_ConditionalFormat;
begin
  Result := TCT_ConditionalFormat.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ConditionalFormatXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ConditionalFormatXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ConditionalFormatXpgList.Assign(AItem: TCT_ConditionalFormatXpgList);
begin
end;

procedure TCT_ConditionalFormatXpgList.CopyTo(AItem: TCT_ConditionalFormatXpgList);
begin
end;

{ TCT_ChartFormat }

function  TCT_ChartFormat.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FPivotArea <> Nil then 
    Inc(ElemsAssigned,FPivotArea.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ChartFormat.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pivotArea' then 
  begin
    if FPivotArea = Nil then 
      FPivotArea := TCT_PivotArea.Create(FOwner);
    Result := FPivotArea;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ChartFormat.Write(AWriter: TXpgWriteXML);
begin
  if (FPivotArea <> Nil) and FPivotArea.Assigned then 
  begin
    FPivotArea.WriteAttributes(AWriter);
    if xaElements in FPivotArea.FAssigneds then 
    begin
      AWriter.BeginTag('pivotArea');
      FPivotArea.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotArea');
  end;
end;

procedure TCT_ChartFormat.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('chart',XmlIntToStr(FChart));
  AWriter.AddAttribute('format',XmlIntToStr(FFormat));
  if FSeries <> False then 
    AWriter.AddAttribute('series',XmlBoolToStr(FSeries));
end;

procedure TCT_ChartFormat.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000212: FChart := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000289: FFormat := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000028B: FSeries := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_ChartFormat.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FChart := 2147483632;
  FFormat := 2147483632;
  FSeries := False;
end;

destructor TCT_ChartFormat.Destroy;
begin
  if FPivotArea <> Nil then 
    FPivotArea.Free;
end;

procedure TCT_ChartFormat.Clear;
begin
  FAssigneds := [];
  if FPivotArea <> Nil then 
    FreeAndNil(FPivotArea);
  FChart := 2147483632;
  FFormat := 2147483632;
  FSeries := False;
end;

procedure TCT_ChartFormat.Assign(AItem: TCT_ChartFormat);
begin
end;

procedure TCT_ChartFormat.CopyTo(AItem: TCT_ChartFormat);
begin
end;

function  TCT_ChartFormat.Create_PivotArea: TCT_PivotArea;
begin
  if FPivotArea = Nil then
      FPivotArea := TCT_PivotArea.Create(FOwner);
  Result := FPivotArea ;
end;

{ TCT_ChartFormatXpgList }

function  TCT_ChartFormatXpgList.GetItems(Index: integer): TCT_ChartFormat;
begin
  Result := TCT_ChartFormat(inherited Items[Index]);
end;

function  TCT_ChartFormatXpgList.Add: TCT_ChartFormat;
begin
  Result := TCT_ChartFormat.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ChartFormatXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ChartFormatXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_ChartFormatXpgList.Assign(AItem: TCT_ChartFormatXpgList);
begin
end;

procedure TCT_ChartFormatXpgList.CopyTo(AItem: TCT_ChartFormatXpgList);
begin
end;

{ TCT_PivotHierarchy }

function  TCT_PivotHierarchy.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FOutline <> False then 
    Inc(AttrsAssigned);
  if FMultipleItemSelectionAllowed <> False then 
    Inc(AttrsAssigned);
  if FSubtotalTop <> False then 
    Inc(AttrsAssigned);
  if FShowInFieldList <> True then 
    Inc(AttrsAssigned);
  if FDragToRow <> True then 
    Inc(AttrsAssigned);
  if FDragToCol <> True then 
    Inc(AttrsAssigned);
  if FDragToPage <> True then 
    Inc(AttrsAssigned);
  if FDragToData <> False then 
    Inc(AttrsAssigned);
  if FDragOff <> True then 
    Inc(AttrsAssigned);
  if FIncludeNewItemsInFilter <> False then 
    Inc(AttrsAssigned);
  if FCaption <> '' then 
    Inc(AttrsAssigned);
  if FMps <> Nil then 
    Inc(ElemsAssigned,FMps.CheckAssigned);
  if FMembersXpgList <> Nil then 
    Inc(ElemsAssigned,FMembersXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotHierarchy.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000150: begin
      if FMps = Nil then 
        FMps := TCT_MemberProperties.Create(FOwner);
      Result := FMps;
    end;
    $000002EB: begin
      if FMembersXpgList = Nil then 
        FMembersXpgList := TCT_MembersXpgList.Create(FOwner);
      Result := FMembersXpgList.Add;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotHierarchy.Write(AWriter: TXpgWriteXML);
begin
  if (FMps <> Nil) and FMps.Assigned then 
  begin
    FMps.WriteAttributes(AWriter);
    if xaElements in FMps.FAssigneds then 
    begin
      AWriter.BeginTag('mps');
      FMps.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('mps');
  end;
  if FMembersXpgList <> Nil then 
    FMembersXpgList.Write(AWriter,'members');
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotHierarchy.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FOutline <> False then 
    AWriter.AddAttribute('outline',XmlBoolToStr(FOutline));
  if FMultipleItemSelectionAllowed <> False then 
    AWriter.AddAttribute('multipleItemSelectionAllowed',XmlBoolToStr(FMultipleItemSelectionAllowed));
  if FSubtotalTop <> False then 
    AWriter.AddAttribute('subtotalTop',XmlBoolToStr(FSubtotalTop));
  if FShowInFieldList <> True then 
    AWriter.AddAttribute('showInFieldList',XmlBoolToStr(FShowInFieldList));
  if FDragToRow <> True then 
    AWriter.AddAttribute('dragToRow',XmlBoolToStr(FDragToRow));
  if FDragToCol <> True then 
    AWriter.AddAttribute('dragToCol',XmlBoolToStr(FDragToCol));
  if FDragToPage <> True then 
    AWriter.AddAttribute('dragToPage',XmlBoolToStr(FDragToPage));
  if FDragToData <> False then 
    AWriter.AddAttribute('dragToData',XmlBoolToStr(FDragToData));
  if FDragOff <> True then 
    AWriter.AddAttribute('dragOff',XmlBoolToStr(FDragOff));
  if FIncludeNewItemsInFilter <> False then 
    AWriter.AddAttribute('includeNewItemsInFilter',XmlBoolToStr(FIncludeNewItemsInFilter));
  if FCaption <> '' then 
    AWriter.AddAttribute('caption',FCaption);
end;

procedure TCT_PivotHierarchy.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000300: FOutline := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000B69: FMultipleItemSelectionAllowed := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004A1: FSubtotalTop := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005F8: FShowInFieldList := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000399: FDragToRow := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000037F: FDragToCol := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000003DE: FDragToPage := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000003DB: FDragToData := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002B9: FDragOff := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000092D: FIncludeNewItemsInFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002EE: FCaption := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotHierarchy.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 11;
  FOutline := False;
  FMultipleItemSelectionAllowed := False;
  FSubtotalTop := False;
  FShowInFieldList := True;
  FDragToRow := True;
  FDragToCol := True;
  FDragToPage := True;
  FDragToData := False;
  FDragOff := True;
  FIncludeNewItemsInFilter := False;
end;

destructor TCT_PivotHierarchy.Destroy;
begin
  if FMps <> Nil then 
    FMps.Free;
  if FMembersXpgList <> Nil then 
    FMembersXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PivotHierarchy.Clear;
begin
  FAssigneds := [];
  if FMps <> Nil then 
    FreeAndNil(FMps);
  if FMembersXpgList <> Nil then 
    FreeAndNil(FMembersXpgList);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FOutline := False;
  FMultipleItemSelectionAllowed := False;
  FSubtotalTop := False;
  FShowInFieldList := True;
  FDragToRow := True;
  FDragToCol := True;
  FDragToPage := True;
  FDragToData := False;
  FDragOff := True;
  FIncludeNewItemsInFilter := False;
  FCaption := '';
end;

procedure TCT_PivotHierarchy.Assign(AItem: TCT_PivotHierarchy);
begin
end;

procedure TCT_PivotHierarchy.CopyTo(AItem: TCT_PivotHierarchy);
begin
end;

function  TCT_PivotHierarchy.Create_Mps: TCT_MemberProperties;
begin
  if FMps = Nil then
      FMps := TCT_MemberProperties.Create(FOwner);
  Result := FMps ;
end;

function  TCT_PivotHierarchy.Create_MembersXpgList: TCT_MembersXpgList;
begin
  if FMembersXpgList = Nil then
      FMembersXpgList := TCT_MembersXpgList.Create(FOwner);
  Result := FMembersXpgList ;
end;

function  TCT_PivotHierarchy.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PivotHierarchyXpgList }

function  TCT_PivotHierarchyXpgList.GetItems(Index: integer): TCT_PivotHierarchy;
begin
  Result := TCT_PivotHierarchy(inherited Items[Index]);
end;

function  TCT_PivotHierarchyXpgList.Add: TCT_PivotHierarchy;
begin
  Result := TCT_PivotHierarchy.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotHierarchyXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotHierarchyXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PivotHierarchyXpgList.Assign(AItem: TCT_PivotHierarchyXpgList);
begin
end;

procedure TCT_PivotHierarchyXpgList.CopyTo(AItem: TCT_PivotHierarchyXpgList);
begin
end;

{ TCT_PivotFilter }

function  TCT_PivotFilter.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PivotFilter.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'extLst' then 
  begin
    if FExtLst = Nil then 
      FExtLst := TCT_ExtensionList.Create(FOwner);
    Result := FExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotFilter.Write(AWriter: TXpgWriteXML);
begin
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotFilter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('fld',XmlIntToStr(FFld));
  if FMpFld <> 2147483632 then 
    AWriter.AddAttribute('mpFld',XmlIntToStr(FMpFld));
  AWriter.AddAttribute('type',StrTST_PivotFilterType[Integer(FType)]);
  if FEvalOrder <> 0 then 
    AWriter.AddAttribute('evalOrder',XmlIntToStr(FEvalOrder));
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  if FIMeasureHier <> 2147483632 then 
    AWriter.AddAttribute('iMeasureHier',XmlIntToStr(FIMeasureHier));
  if FIMeasureFld <> 2147483632 then 
    AWriter.AddAttribute('iMeasureFld',XmlIntToStr(FIMeasureFld));
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FDescription <> '' then 
    AWriter.AddAttribute('description',FDescription);
  if FStringValue1 <> '' then 
    AWriter.AddAttribute('stringValue1',FStringValue1);
  if FStringValue2 <> '' then 
    AWriter.AddAttribute('stringValue2',FStringValue2);
end;

procedure TCT_PivotFilter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000136: FFld := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001F3: FMpFld := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C2: FType := TST_PivotFilterType(StrToEnum('stpft' + AAttributes.Values[i]));
      $000003A4: FEvalOrder := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004C3: FIMeasureHier := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000451: FIMeasureFld := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A1: FName := AAttributes.Values[i];
      $000004A4: FDescription := AAttributes.Values[i];
      $000004C5: FStringValue1 := AAttributes.Values[i];
      $000004C6: FStringValue2 := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotFilter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 11;
  FFld := 2147483632;
  FMpFld := 2147483632;
  FType := TST_PivotFilterType(XPG_UNKNOWN_ENUM);
  FEvalOrder := 0;
  FId := 2147483632;
  FIMeasureHier := 2147483632;
  FIMeasureFld := 2147483632;
end;

destructor TCT_PivotFilter.Destroy;
begin
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_PivotFilter.Clear;
begin
  FAssigneds := [];
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FFld := 2147483632;
  FMpFld := 2147483632;
  FEvalOrder := 0;
  FId := 2147483632;
  FIMeasureHier := 2147483632;
  FIMeasureFld := 2147483632;
  FName := '';
  FDescription := '';
  FStringValue1 := '';
  FStringValue2 := '';
end;

procedure TCT_PivotFilter.Assign(AItem: TCT_PivotFilter);
begin
end;

procedure TCT_PivotFilter.CopyTo(AItem: TCT_PivotFilter);
begin
end;

function  TCT_PivotFilter.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PivotFilterXpgList }

function  TCT_PivotFilterXpgList.GetItems(Index: integer): TCT_PivotFilter;
begin
  Result := TCT_PivotFilter(inherited Items[Index]);
end;

function  TCT_PivotFilterXpgList.Add: TCT_PivotFilter;
begin
  Result := TCT_PivotFilter.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotFilterXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotFilterXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PivotFilterXpgList.Assign(AItem: TCT_PivotFilterXpgList);
begin
end;

procedure TCT_PivotFilterXpgList.CopyTo(AItem: TCT_PivotFilterXpgList);
begin
end;

{ TCT_HierarchyUsage }

function  TCT_HierarchyUsage.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_HierarchyUsage.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_HierarchyUsage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('hierarchyUsage',XmlIntToStr(FHierarchyUsage));
end;

procedure TCT_HierarchyUsage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'hierarchyUsage' then 
    FHierarchyUsage := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_HierarchyUsage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FHierarchyUsage := 2147483632;
end;

destructor TCT_HierarchyUsage.Destroy;
begin
end;

procedure TCT_HierarchyUsage.Clear;
begin
  FAssigneds := [];
  FHierarchyUsage := 2147483632;
end;

procedure TCT_HierarchyUsage.Assign(AItem: TCT_HierarchyUsage);
begin
end;

procedure TCT_HierarchyUsage.CopyTo(AItem: TCT_HierarchyUsage);
begin
end;

{ TCT_HierarchyUsageXpgList }

function  TCT_HierarchyUsageXpgList.GetItems(Index: integer): TCT_HierarchyUsage;
begin
  Result := TCT_HierarchyUsage(inherited Items[Index]);
end;

function  TCT_HierarchyUsageXpgList.Add: TCT_HierarchyUsage;
begin
  Result := TCT_HierarchyUsage.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_HierarchyUsageXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_HierarchyUsageXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_HierarchyUsageXpgList.Assign(AItem: TCT_HierarchyUsageXpgList);
begin
end;

procedure TCT_HierarchyUsageXpgList.CopyTo(AItem: TCT_HierarchyUsageXpgList);
begin
end;

{ TCT_WorksheetSource }

function  TCT_WorksheetSource.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];

  if (FRCells <> Nil) or (FRef <> '') then
    Inc(AttrsAssigned);
  if FName <> '' then
    Inc(AttrsAssigned);
  if FSheet <> '' then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_WorksheetSource.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_WorksheetSource.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRCells <> Nil then begin
    AWriter.AddAttribute('ref',FRCells.ShortRef);
    AWriter.AddAttribute('sheet',FRCells.SheetName);
  end
  else if FRef <> '' then begin
    AWriter.AddAttribute('ref',FRef);
    AWriter.AddAttribute('sheet',FSheet);
  end;

  if FName <> '' then
    AWriter.AddAttribute('name',FName);
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_WorksheetSource.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $0000013D: FRef := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $00000219: FSheet := AAttributes.Values[i];
      $00000179: FR_Id := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_WorksheetSource.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FR_Id := '';
end;

destructor TCT_WorksheetSource.Destroy;
begin
  if FRCells <> Nil then
    FRCells.Free;
end;

procedure TCT_WorksheetSource.SetRCells(const Value: TXLSRelCells);
begin
  if FRCells <> Nil then
    FRCells.Free;
  FRCells := Value;
end;

procedure TCT_WorksheetSource.Clear;
begin
  FAssigneds := [];
  FRef := '';
  FName := '';
  FSheet := '';
  FR_Id := '';
  if FRCells <> Nil then begin
    FRCells.Free;
    FRCells := Nil;
  end;
end;

procedure TCT_WorksheetSource.Assign(AItem: TCT_WorksheetSource);
begin
end;

procedure TCT_WorksheetSource.CopyTo(AItem: TCT_WorksheetSource);
begin
end;

{ TCT_Consolidation }

function  TCT_Consolidation.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAutoPage <> True then 
    Inc(AttrsAssigned);
  if FPages <> Nil then 
    Inc(ElemsAssigned,FPages.CheckAssigned);
  if FRangeSets <> Nil then 
    Inc(ElemsAssigned,FRangeSets.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Consolidation.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000210: begin
      if FPages = Nil then 
        FPages := TCT_Pages.Create(FOwner);
      Result := FPages;
    end;
    $000003AC: begin
      if FRangeSets = Nil then 
        FRangeSets := TCT_RangeSets.Create(FOwner);
      Result := FRangeSets;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Consolidation.Write(AWriter: TXpgWriteXML);
begin
  if (FPages <> Nil) and FPages.Assigned then 
  begin
    FPages.WriteAttributes(AWriter);
    if xaElements in FPages.FAssigneds then 
    begin
      AWriter.BeginTag('pages');
      FPages.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pages');
  end;
  if (FRangeSets <> Nil) and FRangeSets.Assigned then 
  begin
    FRangeSets.WriteAttributes(AWriter);
    if xaElements in FRangeSets.FAssigneds then 
    begin
      AWriter.BeginTag('rangeSets');
      FRangeSets.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rangeSets');
  end
  else 
    AWriter.SimpleTag('rangeSets');
end;

procedure TCT_Consolidation.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAutoPage <> True then 
    AWriter.AddAttribute('autoPage',XmlBoolToStr(FAutoPage));
end;

procedure TCT_Consolidation.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'autoPage' then 
    FAutoPage := XmlStrToBoolDef(AAttributes.Values[0],True)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Consolidation.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FAutoPage := True;
end;

destructor TCT_Consolidation.Destroy;
begin
  if FPages <> Nil then 
    FPages.Free;
  if FRangeSets <> Nil then 
    FRangeSets.Free;
end;

procedure TCT_Consolidation.Clear;
begin
  FAssigneds := [];
  if FPages <> Nil then 
    FreeAndNil(FPages);
  if FRangeSets <> Nil then 
    FreeAndNil(FRangeSets);
  FAutoPage := True;
end;

procedure TCT_Consolidation.Assign(AItem: TCT_Consolidation);
begin
end;

procedure TCT_Consolidation.CopyTo(AItem: TCT_Consolidation);
begin
end;

function  TCT_Consolidation.Create_Pages: TCT_Pages;
begin
  if FPages = Nil then
      FPages := TCT_Pages.Create(FOwner);
  Result := FPages ;
end;

function  TCT_Consolidation.Create_RangeSets: TCT_RangeSets;
begin
  if FRangeSets = Nil then
      FRangeSets := TCT_RangeSets.Create(FOwner);
  Result := FRangeSets ;
end;

{ TCT_CacheField }

function  TCT_CacheField.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FSharedItems <> Nil then 
    Inc(ElemsAssigned,FSharedItems.CheckAssigned);
  if FFieldGroup <> Nil then 
    Inc(ElemsAssigned,FFieldGroup.CheckAssigned);
  if FMpMapXpgList <> Nil then 
    Inc(ElemsAssigned,FMpMapXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaAttributes,xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CacheField.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000479: begin
      if FSharedItems = Nil then 
        FSharedItems := TCT_SharedItems.Create(FOwner);
      Result := FSharedItems;
    end;
    $00000411: begin
      if FFieldGroup = Nil then 
        FFieldGroup := TCT_FieldGroup.Create(FOwner);
      Result := FFieldGroup;
    end;
    $000001FB: begin
      if FMpMapXpgList = Nil then 
        FMpMapXpgList := TCT_XXpgList.Create(FOwner);
      Result := FMpMapXpgList.Add;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CacheField.Write(AWriter: TXpgWriteXML);
begin
  if (FSharedItems <> Nil) and FSharedItems.Assigned then 
  begin
    FSharedItems.WriteAttributes(AWriter);
    if xaElements in FSharedItems.FAssigneds then 
    begin
      AWriter.BeginTag('sharedItems');
      FSharedItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sharedItems');
  end;
  if (FFieldGroup <> Nil) and FFieldGroup.Assigned then 
  begin
    FFieldGroup.WriteAttributes(AWriter);
    if xaElements in FFieldGroup.FAssigneds then 
    begin
      AWriter.BeginTag('fieldGroup');
      FFieldGroup.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('fieldGroup');
  end;
  if FMpMapXpgList <> Nil then 
    FMpMapXpgList.Write(AWriter,'mpMap');
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CacheField.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  if FCaption <> '' then 
    AWriter.AddAttribute('caption',FCaption);
  if FPropertyName <> '' then 
    AWriter.AddAttribute('propertyName',FPropertyName);
  if FServerField <> False then 
    AWriter.AddAttribute('serverField',XmlBoolToStr(FServerField));
  if FUniqueList <> True then 
    AWriter.AddAttribute('uniqueList',XmlBoolToStr(FUniqueList));

  if FNumFmtId <> '' then
    AWriter.AddAttribute('numFmtId',FNumFmtId);

  if FFormula <> '' then 
    AWriter.AddAttribute('formula',FFormula);
  if FSqlType <> 0 then 
    AWriter.AddAttribute('sqlType',XmlIntToStr(FSqlType));
  if FHierarchy <> 0 then 
    AWriter.AddAttribute('hierarchy',XmlIntToStr(FHierarchy));
  if FLevel <> 0 then 
    AWriter.AddAttribute('level',XmlIntToStr(FLevel));
  if FDatabaseField <> True then 
    AWriter.AddAttribute('databaseField',XmlBoolToStr(FDatabaseField));
  if FMappingCount <> 2147483632 then 
    AWriter.AddAttribute('mappingCount',XmlIntToStr(FMappingCount));
  if FMemberPropertyField <> False then 
    AWriter.AddAttribute('memberPropertyField',XmlBoolToStr(FMemberPropertyField));
end;

procedure TCT_CacheField.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000002EE: FCaption := AAttributes.Values[i];
      $00000506: FPropertyName := AAttributes.Values[i];
      $0000047B: FServerField := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000433: FUniqueList := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000324: FNumFmtId := AAttributes.Values[i];
      $000002F6: FFormula := AAttributes.Values[i];
      $000002F2: FSqlType := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003BF: FHierarchy := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000218: FLevel := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000519: FDatabaseField := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000004F5: FMappingCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000007C1: FMemberPropertyField := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_CacheField.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 13;
  FServerField := False;
  FUniqueList := True;
  FSqlType := 0;
  FHierarchy := 0;
  FLevel := 0;
  FNumFmtId := '0';
  FDatabaseField := True;
  FMappingCount := 2147483632;
  FMemberPropertyField := False;
end;

destructor TCT_CacheField.Destroy;
begin
  if FSharedItems <> Nil then 
    FSharedItems.Free;
  if FFieldGroup <> Nil then 
    FFieldGroup.Free;
  if FMpMapXpgList <> Nil then 
    FMpMapXpgList.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_CacheField.Clear;
begin
  FAssigneds := [];
  if FSharedItems <> Nil then
    FreeAndNil(FSharedItems);
  if FFieldGroup <> Nil then
    FreeAndNil(FFieldGroup);
  if FMpMapXpgList <> Nil then
    FreeAndNil(FMpMapXpgList);
  if FExtLst <> Nil then
    FreeAndNil(FExtLst);
  FName := '';
  FCaption := '';
  FPropertyName := '';
  FServerField := False;
  FUniqueList := True;
  FNumFmtId := '0';
  FFormula := '';
  FSqlType := 0;
  FHierarchy := 0;
  FLevel := 0;
  FDatabaseField := True;
  FMappingCount := 2147483632;
  FMemberPropertyField := False;
end;

procedure TCT_CacheField.Assign(AItem: TCT_CacheField);
begin
end;

procedure TCT_CacheField.CopyTo(AItem: TCT_CacheField);
begin
end;

function  TCT_CacheField.Create_SharedItems: TCT_SharedItems;
begin
  if FSharedItems = Nil then
    FSharedItems := TCT_SharedItems.Create(FOwner);
  Result := FSharedItems ;
end;

function  TCT_CacheField.Create_FieldGroup: TCT_FieldGroup;
begin
  if FFieldGroup = Nil then
      FFieldGroup := TCT_FieldGroup.Create(FOwner);
  Result := FFieldGroup ;
end;

function  TCT_CacheField.Create_MpMapXpgList: TCT_XXpgList;
begin
  if FMpMapXpgList = Nil then
      FMpMapXpgList := TCT_XXpgList.Create(FOwner);
  Result := FMpMapXpgList ;
end;

function  TCT_CacheField.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_CacheFieldXpgList }

{ TCT_CacheHierarchy }

function  TCT_CacheHierarchy.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FFieldsUsage <> Nil then 
    Inc(ElemsAssigned,FFieldsUsage.CheckAssigned);
  if FGroupLevels <> Nil then 
    Inc(ElemsAssigned,FGroupLevels.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CacheHierarchy.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000046C: begin
      if FFieldsUsage = Nil then 
        FFieldsUsage := TCT_FieldsUsage.Create(FOwner);
      Result := FFieldsUsage;
    end;
    $00000498: begin
      if FGroupLevels = Nil then 
        FGroupLevels := TCT_GroupLevels.Create(FOwner);
      Result := FGroupLevels;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CacheHierarchy.Write(AWriter: TXpgWriteXML);
begin
  if (FFieldsUsage <> Nil) and FFieldsUsage.Assigned then 
  begin
    FFieldsUsage.WriteAttributes(AWriter);
    if xaElements in FFieldsUsage.FAssigneds then 
    begin
      AWriter.BeginTag('fieldsUsage');
      FFieldsUsage.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('fieldsUsage');
  end;
  if (FGroupLevels <> Nil) and FGroupLevels.Assigned then 
  begin
    FGroupLevels.WriteAttributes(AWriter);
    if xaElements in FGroupLevels.FAssigneds then 
    begin
      AWriter.BeginTag('groupLevels');
      FGroupLevels.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('groupLevels');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CacheHierarchy.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('uniqueName',FUniqueName);
  if FCaption <> '' then 
    AWriter.AddAttribute('caption',FCaption);
  if FMeasure <> False then 
    AWriter.AddAttribute('measure',XmlBoolToStr(FMeasure));
  if FSet <> False then 
    AWriter.AddAttribute('set',XmlBoolToStr(FSet));
  if FParentSet <> 2147483632 then 
    AWriter.AddAttribute('parentSet',XmlIntToStr(FParentSet));
  if FIconSet <> 0 then 
    AWriter.AddAttribute('iconSet',XmlIntToStr(FIconSet));
  if FAttribute <> False then 
    AWriter.AddAttribute('attribute',XmlBoolToStr(FAttribute));
  if FTime <> False then 
    AWriter.AddAttribute('time',XmlBoolToStr(FTime));
  if FKeyAttribute <> False then 
    AWriter.AddAttribute('keyAttribute',XmlBoolToStr(FKeyAttribute));
  if FDefaultMemberUniqueName <> '' then 
    AWriter.AddAttribute('defaultMemberUniqueName',FDefaultMemberUniqueName);
  if FAllUniqueName <> '' then 
    AWriter.AddAttribute('allUniqueName',FAllUniqueName);
  if FAllCaption <> '' then 
    AWriter.AddAttribute('allCaption',FAllCaption);
  if FDimensionUniqueName <> '' then 
    AWriter.AddAttribute('dimensionUniqueName',FDimensionUniqueName);
  if FDisplayFolder <> '' then 
    AWriter.AddAttribute('displayFolder',FDisplayFolder);
  if FMeasureGroup <> '' then 
    AWriter.AddAttribute('measureGroup',FMeasureGroup);
  if FMeasures <> False then 
    AWriter.AddAttribute('measures',XmlBoolToStr(FMeasures));
  AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if FOneField <> False then 
    AWriter.AddAttribute('oneField',XmlBoolToStr(FOneField));
  if FMemberValueDatatype <> 2147483632 then 
    AWriter.AddAttribute('memberValueDatatype',XmlIntToStr(FMemberValueDatatype));
  if Byte(FUnbalanced) <> 2 then 
    AWriter.AddAttribute('unbalanced',XmlBoolToStr(FUnbalanced));
  if Byte(FUnbalancedGroup) <> 2 then 
    AWriter.AddAttribute('unbalancedGroup',XmlBoolToStr(FUnbalancedGroup));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
end;

procedure TCT_CacheHierarchy.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000418: FUniqueName := AAttributes.Values[i];
      $000002EE: FCaption := AAttributes.Values[i];
      $000002F2: FMeasure := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000014C: FSet := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003B6: FParentSet := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002D5: FIconSet := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003D4: FAttribute := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001AF: FTime := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004FD: FKeyAttribute := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000935: FDefaultMemberUniqueName := AAttributes.Values[i];
      $00000531: FAllUniqueName := AAttributes.Values[i];
      $00000407: FAllCaption := AAttributes.Values[i];
      $000007BE: FDimensionUniqueName := AAttributes.Values[i];
      $00000552: FDisplayFolder := AAttributes.Values[i];
      $000004FF: FMeasureGroup := AAttributes.Values[i];
      $00000365: FMeasures := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000326: FOneField := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000007B1: FMemberValueDatatype := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000040D: FUnbalanced := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000061A: FUnbalancedGroup := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_CacheHierarchy.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 22;
  FMeasure := False;
  FSet := False;
  FParentSet := 2147483632;
  FIconSet := 0;
  FAttribute := False;
  FTime := False;
  FKeyAttribute := False;
  FMeasures := False;
  FCount := 2147483632;
  FOneField := False;
  FMemberValueDatatype := 2147483632;
  Byte(FUnbalanced) := 2;
  Byte(FUnbalancedGroup) := 2;
  FHidden := False;
end;

destructor TCT_CacheHierarchy.Destroy;
begin
  if FFieldsUsage <> Nil then 
    FFieldsUsage.Free;
  if FGroupLevels <> Nil then 
    FGroupLevels.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_CacheHierarchy.Clear;
begin
  FAssigneds := [];
  if FFieldsUsage <> Nil then 
    FreeAndNil(FFieldsUsage);
  if FGroupLevels <> Nil then 
    FreeAndNil(FGroupLevels);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FUniqueName := '';
  FCaption := '';
  FMeasure := False;
  FSet := False;
  FParentSet := 2147483632;
  FIconSet := 0;
  FAttribute := False;
  FTime := False;
  FKeyAttribute := False;
  FDefaultMemberUniqueName := '';
  FAllUniqueName := '';
  FAllCaption := '';
  FDimensionUniqueName := '';
  FDisplayFolder := '';
  FMeasureGroup := '';
  FMeasures := False;
  FCount := 2147483632;
  FOneField := False;
  FMemberValueDatatype := 2147483632;
  Byte(FUnbalanced) := 2;
  Byte(FUnbalancedGroup) := 2;
  FHidden := False;
end;

procedure TCT_CacheHierarchy.Assign(AItem: TCT_CacheHierarchy);
begin
end;

procedure TCT_CacheHierarchy.CopyTo(AItem: TCT_CacheHierarchy);
begin
end;

function  TCT_CacheHierarchy.Create_FieldsUsage: TCT_FieldsUsage;
begin
  if FFieldsUsage = Nil then
      FFieldsUsage := TCT_FieldsUsage.Create(FOwner);
  Result := FFieldsUsage ;
end;

function  TCT_CacheHierarchy.Create_GroupLevels: TCT_GroupLevels;
begin
  if FGroupLevels = Nil then
      FGroupLevels := TCT_GroupLevels.Create(FOwner);
  Result := FGroupLevels ;
end;

function  TCT_CacheHierarchy.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_CacheHierarchyXpgList }

function  TCT_CacheHierarchyXpgList.GetItems(Index: integer): TCT_CacheHierarchy;
begin
  Result := TCT_CacheHierarchy(inherited Items[Index]);
end;

function  TCT_CacheHierarchyXpgList.Add: TCT_CacheHierarchy;
begin
  Result := TCT_CacheHierarchy.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CacheHierarchyXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CacheHierarchyXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_CacheHierarchyXpgList.Assign(AItem: TCT_CacheHierarchyXpgList);
begin
end;

procedure TCT_CacheHierarchyXpgList.CopyTo(AItem: TCT_CacheHierarchyXpgList);
begin
end;

{ TCT_PCDKPI }

function  TCT_PCDKPI.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PCDKPI.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PCDKPI.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('uniqueName',FUniqueName);
  if FCaption <> '' then 
    AWriter.AddAttribute('caption',FCaption);
  if FDisplayFolder <> '' then 
    AWriter.AddAttribute('displayFolder',FDisplayFolder);
  if FMeasureGroup <> '' then 
    AWriter.AddAttribute('measureGroup',FMeasureGroup);
  if FParent <> '' then 
    AWriter.AddAttribute('parent',FParent);
  AWriter.AddAttribute('value',FValue);
  if FGoal <> '' then 
    AWriter.AddAttribute('goal',FGoal);
  if FStatus <> '' then 
    AWriter.AddAttribute('status',FStatus);
  if FTrend <> '' then 
    AWriter.AddAttribute('trend',FTrend);
  if FWeight <> '' then 
    AWriter.AddAttribute('weight',FWeight);
  if FTime <> '' then 
    AWriter.AddAttribute('time',FTime);
end;

procedure TCT_PCDKPI.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashB[i] of
      $160ACBEE: FUniqueName := AAttributes.Values[i];
      $A0F59B0E: FCaption := AAttributes.Values[i];
      $425D04AE: FDisplayFolder := AAttributes.Values[i];
      $4645497D: FMeasureGroup := AAttributes.Values[i];
      $75589F6C: FParent := AAttributes.Values[i];
      $97985D6B: FValue := AAttributes.Values[i];
      $1BC058D3: FGoal := AAttributes.Values[i];
      $800303B4: FStatus := AAttributes.Values[i];
      $04E16977: FTrend := AAttributes.Values[i];
      $BCD0247C: FWeight := AAttributes.Values[i];
      $5069420B: FTime := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PCDKPI.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 11;
end;

destructor TCT_PCDKPI.Destroy;
begin
end;

procedure TCT_PCDKPI.Clear;
begin
  FAssigneds := [];
  FUniqueName := '';
  FCaption := '';
  FDisplayFolder := '';
  FMeasureGroup := '';
  FParent := '';
  FValue := '';
  FGoal := '';
  FStatus := '';
  FTrend := '';
  FWeight := '';
  FTime := '';
end;

procedure TCT_PCDKPI.Assign(AItem: TCT_PCDKPI);
begin
end;

procedure TCT_PCDKPI.CopyTo(AItem: TCT_PCDKPI);
begin
end;

{ TCT_PCDKPIXpgList }

function  TCT_PCDKPIXpgList.GetItems(Index: integer): TCT_PCDKPI;
begin
  Result := TCT_PCDKPI(inherited Items[Index]);
end;

function  TCT_PCDKPIXpgList.Add: TCT_PCDKPI;
begin
  Result := TCT_PCDKPI.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PCDKPIXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PCDKPIXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PCDKPIXpgList.Assign(AItem: TCT_PCDKPIXpgList);
begin
end;

procedure TCT_PCDKPIXpgList.CopyTo(AItem: TCT_PCDKPIXpgList);
begin
end;

{ TCT_PCDSDTCEntries }

function  TCT_PCDSDTCEntries.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMXpgList <> Nil then 
    Inc(ElemsAssigned,FMXpgList.CheckAssigned);
  if FNXpgList <> Nil then 
    Inc(ElemsAssigned,FNXpgList.CheckAssigned);
  if FEXpgList <> Nil then 
    Inc(ElemsAssigned,FEXpgList.CheckAssigned);
  if FSXpgList <> Nil then 
    Inc(ElemsAssigned,FSXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PCDSDTCEntries.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000006D: begin
      if FMXpgList = Nil then 
        FMXpgList := TCT_MissingXpgList.Create(FOwner);
      Result := FMXpgList.Add;
    end;
    $0000006E: begin
      if FNXpgList = Nil then 
        FNXpgList := TCT_NumberXpgList.Create(FOwner);
      Result := FNXpgList.Add;
    end;
    $00000065: begin
      if FEXpgList = Nil then 
        FEXpgList := TCT_ErrorXpgList.Create(FOwner);
      Result := FEXpgList.Add;
    end;
    $00000073: begin
      if FSXpgList = Nil then 
        FSXpgList := TCT_StringXpgList.Create(FOwner);
      Result := FSXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PCDSDTCEntries.Write(AWriter: TXpgWriteXML);
begin
  if FMXpgList <> Nil then 
    FMXpgList.Write(AWriter,'m');
  if FNXpgList <> Nil then 
    FNXpgList.Write(AWriter,'n');
  if FEXpgList <> Nil then 
    FEXpgList.Write(AWriter,'e');
  if FSXpgList <> Nil then 
    FSXpgList.Write(AWriter,'s');
end;

procedure TCT_PCDSDTCEntries.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PCDSDTCEntries.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PCDSDTCEntries.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PCDSDTCEntries.Destroy;
begin
  if FMXpgList <> Nil then 
    FMXpgList.Free;
  if FNXpgList <> Nil then 
    FNXpgList.Free;
  if FEXpgList <> Nil then 
    FEXpgList.Free;
  if FSXpgList <> Nil then 
    FSXpgList.Free;
end;

procedure TCT_PCDSDTCEntries.Clear;
begin
  FAssigneds := [];
  if FMXpgList <> Nil then 
    FreeAndNil(FMXpgList);
  if FNXpgList <> Nil then 
    FreeAndNil(FNXpgList);
  if FEXpgList <> Nil then 
    FreeAndNil(FEXpgList);
  if FSXpgList <> Nil then 
    FreeAndNil(FSXpgList);
  FCount := 2147483632;
end;

procedure TCT_PCDSDTCEntries.Assign(AItem: TCT_PCDSDTCEntries);
begin
end;

procedure TCT_PCDSDTCEntries.CopyTo(AItem: TCT_PCDSDTCEntries);
begin
end;

function  TCT_PCDSDTCEntries.Create_MXpgList: TCT_MissingXpgList;
begin
  if FMXpgList = Nil then
      FMXpgList := TCT_MissingXpgList.Create(FOwner);
  Result := FMXpgList ;
end;

function  TCT_PCDSDTCEntries.Create_NXpgList: TCT_NumberXpgList;
begin
  if FNXpgList = Nil then
      FNXpgList := TCT_NumberXpgList.Create(FOwner);
  Result := FNXpgList ;
end;

function  TCT_PCDSDTCEntries.Create_EXpgList: TCT_ErrorXpgList;
begin
  if FEXpgList = Nil then
      FEXpgList := TCT_ErrorXpgList.Create(FOwner);
  Result := FEXpgList ;
end;

function  TCT_PCDSDTCEntries.Create_SXpgList: TCT_StringXpgList;
begin
  if FSXpgList = Nil then
      FSXpgList := TCT_StringXpgList.Create(FOwner);
  Result := FSXpgList ;
end;

{ TCT_Sets }

function  TCT_Sets.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FSetXpgList <> Nil then 
    Inc(ElemsAssigned,FSetXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Sets.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'set' then 
  begin
    if FSetXpgList = Nil then 
      FSetXpgList := TCT_SetXpgList.Create(FOwner);
    Result := FSetXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Sets.Write(AWriter: TXpgWriteXML);
begin
  if FSetXpgList <> Nil then 
    FSetXpgList.Write(AWriter,'set');
end;

procedure TCT_Sets.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Sets.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Sets.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_Sets.Destroy;
begin
  if FSetXpgList <> Nil then 
    FSetXpgList.Free;
end;

procedure TCT_Sets.Clear;
begin
  FAssigneds := [];
  if FSetXpgList <> Nil then 
    FreeAndNil(FSetXpgList);
  FCount := 2147483632;
end;

procedure TCT_Sets.Assign(AItem: TCT_Sets);
begin
end;

procedure TCT_Sets.CopyTo(AItem: TCT_Sets);
begin
end;

function  TCT_Sets.Create_SetXpgList: TCT_SetXpgList;
begin
  if FSetXpgList = Nil then
      FSetXpgList := TCT_SetXpgList.Create(FOwner);
  Result := FSetXpgList ;
end;

{ TCT_QueryCache }

function  TCT_QueryCache.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FQueryXpgList <> Nil then 
    Inc(ElemsAssigned,FQueryXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_QueryCache.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'query' then 
  begin
    if FQueryXpgList = Nil then 
      FQueryXpgList := TCT_QueryXpgList.Create(FOwner);
    Result := FQueryXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_QueryCache.Write(AWriter: TXpgWriteXML);
begin
  if FQueryXpgList <> Nil then 
    FQueryXpgList.Write(AWriter,'query');
end;

procedure TCT_QueryCache.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_QueryCache.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_QueryCache.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_QueryCache.Destroy;
begin
  if FQueryXpgList <> Nil then 
    FQueryXpgList.Free;
end;

procedure TCT_QueryCache.Clear;
begin
  FAssigneds := [];
  if FQueryXpgList <> Nil then 
    FreeAndNil(FQueryXpgList);
  FCount := 2147483632;
end;

procedure TCT_QueryCache.Assign(AItem: TCT_QueryCache);
begin
end;

procedure TCT_QueryCache.CopyTo(AItem: TCT_QueryCache);
begin
end;

function  TCT_QueryCache.Create_QueryXpgList: TCT_QueryXpgList;
begin
  if FQueryXpgList = Nil then
      FQueryXpgList := TCT_QueryXpgList.Create(FOwner);
  Result := FQueryXpgList ;
end;

{ TCT_ServerFormats }

function  TCT_ServerFormats.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FServerFormatXpgList <> Nil then 
    Inc(ElemsAssigned,FServerFormatXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ServerFormats.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'serverFormat' then 
  begin
    if FServerFormatXpgList = Nil then 
      FServerFormatXpgList := TCT_ServerFormatXpgList.Create(FOwner);
    Result := FServerFormatXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ServerFormats.Write(AWriter: TXpgWriteXML);
begin
  if FServerFormatXpgList <> Nil then 
    FServerFormatXpgList.Write(AWriter,'serverFormat');
end;

procedure TCT_ServerFormats.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_ServerFormats.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ServerFormats.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_ServerFormats.Destroy;
begin
  if FServerFormatXpgList <> Nil then 
    FServerFormatXpgList.Free;
end;

procedure TCT_ServerFormats.Clear;
begin
  FAssigneds := [];
  if FServerFormatXpgList <> Nil then 
    FreeAndNil(FServerFormatXpgList);
  FCount := 2147483632;
end;

procedure TCT_ServerFormats.Assign(AItem: TCT_ServerFormats);
begin
end;

procedure TCT_ServerFormats.CopyTo(AItem: TCT_ServerFormats);
begin
end;

function  TCT_ServerFormats.Create_ServerFormatXpgList: TCT_ServerFormatXpgList;
begin
  if FServerFormatXpgList = Nil then
      FServerFormatXpgList := TCT_ServerFormatXpgList.Create(FOwner);
  Result := FServerFormatXpgList ;
end;

{ TCT_CalculatedItem }

function  TCT_CalculatedItem.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FField <> 2147483632 then 
    Inc(AttrsAssigned);
  if FFormula <> '' then 
    Inc(AttrsAssigned);
  if FPivotArea <> Nil then 
    Inc(ElemsAssigned,FPivotArea.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CalculatedItem.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003AB: begin
      if FPivotArea = Nil then 
        FPivotArea := TCT_PivotArea.Create(FOwner);
      Result := FPivotArea;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CalculatedItem.Write(AWriter: TXpgWriteXML);
begin
  if (FPivotArea <> Nil) and FPivotArea.Assigned then 
  begin
    FPivotArea.WriteAttributes(AWriter);
    if xaElements in FPivotArea.FAssigneds then 
    begin
      AWriter.BeginTag('pivotArea');
      FPivotArea.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotArea');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CalculatedItem.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FField <> 2147483632 then 
    AWriter.AddAttribute('field',XmlIntToStr(FField));
  if FFormula <> '' then 
    AWriter.AddAttribute('formula',FFormula);
end;

procedure TCT_CalculatedItem.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $00000204: FField := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002F6: FFormula := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_CalculatedItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 2;
  FField := 2147483632;
end;

destructor TCT_CalculatedItem.Destroy;
begin
  if FPivotArea <> Nil then 
    FPivotArea.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_CalculatedItem.Clear;
begin
  FAssigneds := [];
  if FPivotArea <> Nil then 
    FreeAndNil(FPivotArea);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FField := 2147483632;
  FFormula := '';
end;

procedure TCT_CalculatedItem.Assign(AItem: TCT_CalculatedItem);
begin
end;

procedure TCT_CalculatedItem.CopyTo(AItem: TCT_CalculatedItem);
begin
end;

function  TCT_CalculatedItem.Create_PivotArea: TCT_PivotArea;
begin
  if FPivotArea = Nil then
      FPivotArea := TCT_PivotArea.Create(FOwner);
  Result := FPivotArea ;
end;

function  TCT_CalculatedItem.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_CalculatedItemXpgList }

function  TCT_CalculatedItemXpgList.GetItems(Index: integer): TCT_CalculatedItem;
begin
  Result := TCT_CalculatedItem(inherited Items[Index]);
end;

function  TCT_CalculatedItemXpgList.Add: TCT_CalculatedItem;
begin
  Result := TCT_CalculatedItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CalculatedItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CalculatedItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_CalculatedItemXpgList.Assign(AItem: TCT_CalculatedItemXpgList);
begin
end;

procedure TCT_CalculatedItemXpgList.CopyTo(AItem: TCT_CalculatedItemXpgList);
begin
end;

{ TCT_CalculatedMember }

function  TCT_CalculatedMember.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CalculatedMember.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'extLst' then 
  begin
    if FExtLst = Nil then 
      FExtLst := TCT_ExtensionList.Create(FOwner);
    Result := FExtLst;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CalculatedMember.Write(AWriter: TXpgWriteXML);
begin
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CalculatedMember.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('mdx',FMdx);
  if FMemberName <> '' then 
    AWriter.AddAttribute('memberName',FMemberName);
  if FHierarchy <> '' then 
    AWriter.AddAttribute('hierarchy',FHierarchy);
  if FParent <> '' then 
    AWriter.AddAttribute('parent',FParent);
  if FSolveOrder <> 0 then 
    AWriter.AddAttribute('solveOrder',XmlIntToStr(FSolveOrder));
  if FSet <> False then 
    AWriter.AddAttribute('set',XmlBoolToStr(FSet));
end;

procedure TCT_CalculatedMember.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000149: FMdx := AAttributes.Values[i];
      $000003F9: FMemberName := AAttributes.Values[i];
      $000003BF: FHierarchy := AAttributes.Values[i];
      $0000028A: FParent := AAttributes.Values[i];
      $00000425: FSolveOrder := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000014C: FSet := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_CalculatedMember.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 7;
  FSolveOrder := 0;
  FSet := False;
end;

destructor TCT_CalculatedMember.Destroy;
begin
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_CalculatedMember.Clear;
begin
  FAssigneds := [];
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FName := '';
  FMdx := '';
  FMemberName := '';
  FHierarchy := '';
  FParent := '';
  FSolveOrder := 0;
  FSet := False;
end;

procedure TCT_CalculatedMember.Assign(AItem: TCT_CalculatedMember);
begin
end;

procedure TCT_CalculatedMember.CopyTo(AItem: TCT_CalculatedMember);
begin
end;

function  TCT_CalculatedMember.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_CalculatedMemberXpgList }

function  TCT_CalculatedMemberXpgList.GetItems(Index: integer): TCT_CalculatedMember;
begin
  Result := TCT_CalculatedMember(inherited Items[Index]);
end;

function  TCT_CalculatedMemberXpgList.Add: TCT_CalculatedMember;
begin
  Result := TCT_CalculatedMember.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CalculatedMemberXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CalculatedMemberXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_CalculatedMemberXpgList.Assign(AItem: TCT_CalculatedMemberXpgList);
begin
end;

procedure TCT_CalculatedMemberXpgList.CopyTo(AItem: TCT_CalculatedMemberXpgList);
begin
end;

{ TCT_PivotDimension }

function  TCT_PivotDimension.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_PivotDimension.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PivotDimension.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMeasure <> False then 
    AWriter.AddAttribute('measure',XmlBoolToStr(FMeasure));
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('uniqueName',FUniqueName);
  AWriter.AddAttribute('caption',FCaption);
end;

procedure TCT_PivotDimension.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000002F2: FMeasure := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001A1: FName := AAttributes.Values[i];
      $00000418: FUniqueName := AAttributes.Values[i];
      $000002EE: FCaption := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotDimension.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FMeasure := False;
end;

destructor TCT_PivotDimension.Destroy;
begin
end;

procedure TCT_PivotDimension.Clear;
begin
  FAssigneds := [];
  FMeasure := False;
  FName := '';
  FUniqueName := '';
  FCaption := '';
end;

procedure TCT_PivotDimension.Assign(AItem: TCT_PivotDimension);
begin
end;

procedure TCT_PivotDimension.CopyTo(AItem: TCT_PivotDimension);
begin
end;

{ TCT_PivotDimensionXpgList }

function  TCT_PivotDimensionXpgList.GetItems(Index: integer): TCT_PivotDimension;
begin
  Result := TCT_PivotDimension(inherited Items[Index]);
end;

function  TCT_PivotDimensionXpgList.Add: TCT_PivotDimension;
begin
  Result := TCT_PivotDimension.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotDimensionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotDimensionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_PivotDimensionXpgList.Assign(AItem: TCT_PivotDimensionXpgList);
begin
end;

procedure TCT_PivotDimensionXpgList.CopyTo(AItem: TCT_PivotDimensionXpgList);
begin
end;

{ TCT_MeasureGroup }

function  TCT_MeasureGroup.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_MeasureGroup.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_MeasureGroup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('caption',FCaption);
end;

procedure TCT_MeasureGroup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000002EE: FCaption := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_MeasureGroup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_MeasureGroup.Destroy;
begin
end;

procedure TCT_MeasureGroup.Clear;
begin
  FAssigneds := [];
  FName := '';
  FCaption := '';
end;

procedure TCT_MeasureGroup.Assign(AItem: TCT_MeasureGroup);
begin
end;

procedure TCT_MeasureGroup.CopyTo(AItem: TCT_MeasureGroup);
begin
end;

{ TCT_MeasureGroupXpgList }

function  TCT_MeasureGroupXpgList.GetItems(Index: integer): TCT_MeasureGroup;
begin
  Result := TCT_MeasureGroup(inherited Items[Index]);
end;

function  TCT_MeasureGroupXpgList.Add: TCT_MeasureGroup;
begin
  Result := TCT_MeasureGroup.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MeasureGroupXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MeasureGroupXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_MeasureGroupXpgList.Assign(AItem: TCT_MeasureGroupXpgList);
begin
end;

procedure TCT_MeasureGroupXpgList.CopyTo(AItem: TCT_MeasureGroupXpgList);
begin
end;

{ TCT_MeasureDimensionMap }

function  TCT_MeasureDimensionMap.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMeasureGroup <> 2147483632 then 
    Inc(AttrsAssigned);
  if FDimension <> 2147483632 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_MeasureDimensionMap.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_MeasureDimensionMap.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMeasureGroup <> 2147483632 then 
    AWriter.AddAttribute('measureGroup',XmlIntToStr(FMeasureGroup));
  if FDimension <> 2147483632 then 
    AWriter.AddAttribute('dimension',XmlIntToStr(FDimension));
end;

procedure TCT_MeasureDimensionMap.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000004FF: FMeasureGroup := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003C6: FDimension := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_MeasureDimensionMap.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FMeasureGroup := 2147483632;
  FDimension := 2147483632;
end;

destructor TCT_MeasureDimensionMap.Destroy;
begin
end;

procedure TCT_MeasureDimensionMap.Clear;
begin
  FAssigneds := [];
  FMeasureGroup := 2147483632;
  FDimension := 2147483632;
end;

procedure TCT_MeasureDimensionMap.Assign(AItem: TCT_MeasureDimensionMap);
begin
end;

procedure TCT_MeasureDimensionMap.CopyTo(AItem: TCT_MeasureDimensionMap);
begin
end;

{ TCT_MeasureDimensionMapXpgList }

function  TCT_MeasureDimensionMapXpgList.GetItems(Index: integer): TCT_MeasureDimensionMap;
begin
  Result := TCT_MeasureDimensionMap(inherited Items[Index]);
end;

function  TCT_MeasureDimensionMapXpgList.Add: TCT_MeasureDimensionMap;
begin
  Result := TCT_MeasureDimensionMap.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MeasureDimensionMapXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MeasureDimensionMapXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

procedure TCT_MeasureDimensionMapXpgList.Assign(AItem: TCT_MeasureDimensionMapXpgList);
begin
end;

procedure TCT_MeasureDimensionMapXpgList.CopyTo(AItem: TCT_MeasureDimensionMapXpgList);
begin
end;

{ TCT_Record }

function  TCT_Record.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FMXpgList <> Nil then 
    Inc(ElemsAssigned,FMXpgList.CheckAssigned);
  if FNXpgList <> Nil then 
    Inc(ElemsAssigned,FNXpgList.CheckAssigned);
  if FBXpgList <> Nil then 
    Inc(ElemsAssigned,FBXpgList.CheckAssigned);
  if FEXpgList <> Nil then 
    Inc(ElemsAssigned,FEXpgList.CheckAssigned);
  if FSXpgList <> Nil then 
    Inc(ElemsAssigned,FSXpgList.CheckAssigned);
  if FDXpgList <> Nil then 
    Inc(ElemsAssigned,FDXpgList.CheckAssigned);
  if FXXpgList <> Nil then 
    Inc(ElemsAssigned,FXXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Record.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000006D: begin
      if FMXpgList = Nil then 
        FMXpgList := TCT_MissingXpgList.Create(FOwner);
      Result := FMXpgList.Add;
    end;
    $0000006E: begin
      if FNXpgList = Nil then 
        FNXpgList := TCT_NumberXpgList.Create(FOwner);
      Result := FNXpgList.Add;
    end;
    $00000062: begin
      if FBXpgList = Nil then 
        FBXpgList := TCT_BooleanXpgList.Create(FOwner);
      Result := FBXpgList.Add;
    end;
    $00000065: begin
      if FEXpgList = Nil then 
        FEXpgList := TCT_ErrorXpgList.Create(FOwner);
      Result := FEXpgList.Add;
    end;
    $00000073: begin
      if FSXpgList = Nil then 
        FSXpgList := TCT_StringXpgList.Create(FOwner);
      Result := FSXpgList.Add;
    end;
    $00000064: begin
      if FDXpgList = Nil then 
        FDXpgList := TCT_DateTimeXpgList.Create(FOwner);
      Result := FDXpgList.Add;
    end;
    $00000078: begin
      if FXXpgList = Nil then 
        FXXpgList := TCT_IndexXpgList.Create(FOwner);
      Result := FXXpgList.Add;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Record.Write(AWriter: TXpgWriteXML);
begin
  if FMXpgList <> Nil then 
    FMXpgList.Write(AWriter,'m');
  if FNXpgList <> Nil then 
    FNXpgList.Write(AWriter,'n');
  if FBXpgList <> Nil then 
    FBXpgList.Write(AWriter,'b');
  if FEXpgList <> Nil then 
    FEXpgList.Write(AWriter,'e');
  if FSXpgList <> Nil then 
    FSXpgList.Write(AWriter,'s');
  if FDXpgList <> Nil then 
    FDXpgList.Write(AWriter,'d');
  if FXXpgList <> Nil then 
    FXXpgList.Write(AWriter,'x');
end;

constructor TCT_Record.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
end;

destructor TCT_Record.Destroy;
begin
  if FMXpgList <> Nil then 
    FMXpgList.Free;
  if FNXpgList <> Nil then 
    FNXpgList.Free;
  if FBXpgList <> Nil then 
    FBXpgList.Free;
  if FEXpgList <> Nil then 
    FEXpgList.Free;
  if FSXpgList <> Nil then 
    FSXpgList.Free;
  if FDXpgList <> Nil then 
    FDXpgList.Free;
  if FXXpgList <> Nil then 
    FXXpgList.Free;
end;

procedure TCT_Record.Clear;
begin
  FAssigneds := [];
  if FMXpgList <> Nil then 
    FreeAndNil(FMXpgList);
  if FNXpgList <> Nil then 
    FreeAndNil(FNXpgList);
  if FBXpgList <> Nil then 
    FreeAndNil(FBXpgList);
  if FEXpgList <> Nil then 
    FreeAndNil(FEXpgList);
  if FSXpgList <> Nil then 
    FreeAndNil(FSXpgList);
  if FDXpgList <> Nil then 
    FreeAndNil(FDXpgList);
  if FXXpgList <> Nil then 
    FreeAndNil(FXXpgList);
end;

procedure TCT_Record.Assign(AItem: TCT_Record);
begin
end;

procedure TCT_Record.CopyTo(AItem: TCT_Record);
begin
end;

function  TCT_Record.Create_MXpgList: TCT_MissingXpgList;
begin
  if FMXpgList = Nil then
      FMXpgList := TCT_MissingXpgList.Create(FOwner);
  Result := FMXpgList ;
end;

function  TCT_Record.Create_NXpgList: TCT_NumberXpgList;
begin
  if FNXpgList = Nil then
      FNXpgList := TCT_NumberXpgList.Create(FOwner);
  Result := FNXpgList ;
end;

function  TCT_Record.Create_BXpgList: TCT_BooleanXpgList;
begin
  if FBXpgList = Nil then
      FBXpgList := TCT_BooleanXpgList.Create(FOwner);
  Result := FBXpgList ;
end;

function  TCT_Record.Create_EXpgList: TCT_ErrorXpgList;
begin
  if FEXpgList = Nil then
      FEXpgList := TCT_ErrorXpgList.Create(FOwner);
  Result := FEXpgList ;
end;

function  TCT_Record.Create_SXpgList: TCT_StringXpgList;
begin
  if FSXpgList = Nil then
      FSXpgList := TCT_StringXpgList.Create(FOwner);
  Result := FSXpgList ;
end;

function  TCT_Record.Create_DXpgList: TCT_DateTimeXpgList;
begin
  if FDXpgList = Nil then
      FDXpgList := TCT_DateTimeXpgList.Create(FOwner);
  Result := FDXpgList ;
end;

function  TCT_Record.Create_XXpgList: TCT_IndexXpgList;
begin
  if FXXpgList = Nil then
      FXXpgList := TCT_IndexXpgList.Create(FOwner);
  Result := FXXpgList ;
end;

{ TCT_RecordXpgList }

function  TCT_RecordXpgList.GetItems(Index: integer): TCT_Record;
begin
  Result := TCT_Record(inherited Items[Index]);
end;

function  TCT_RecordXpgList.Add: TCT_Record;
begin
  Result := TCT_Record.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RecordXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RecordXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    if xaElements in Items[i].FAssigneds then 
    begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
end;

procedure TCT_RecordXpgList.Assign(AItem: TCT_RecordXpgList);
begin
end;

procedure TCT_RecordXpgList.CopyTo(AItem: TCT_RecordXpgList);
begin
end;

{ TCT_Location }

function  TCT_Location.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_Location.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Location.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRCells <> Nil then
    AWriter.AddAttribute('ref',FRCells.ShortRef)
  else if FRef <> '' then
    AWriter.AddAttribute('ref',FRef);

  AWriter.AddAttribute('firstHeaderRow',XmlIntToStr(FFirstHeaderRow));
  AWriter.AddAttribute('firstDataRow',XmlIntToStr(FFirstDataRow));
  AWriter.AddAttribute('firstDataCol',XmlIntToStr(FFirstDataCol));
  if FRowPageCount <> 0 then 
    AWriter.AddAttribute('rowPageCount',XmlIntToStr(FRowPageCount));
  if FColPageCount <> 0 then 
    AWriter.AddAttribute('colPageCount',XmlIntToStr(FColPageCount));
end;

procedure TCT_Location.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $0000013D: FRef := AAttributes.Values[i];
      $000005A9: FFirstHeaderRow := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004DA: FFirstDataRow := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004C0: FFirstDataCol := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004DE: FRowPageCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004C4: FColPageCount := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_Location.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
  FFirstHeaderRow := 2147483632;
  FFirstDataRow := 2147483632;
  FFirstDataCol := 2147483632;
  FRowPageCount := 0;
  FColPageCount := 0;
end;

destructor TCT_Location.Destroy;
begin
end;

procedure TCT_Location.Clear;
begin
  FAssigneds := [];
  FRef := '';
  FFirstHeaderRow := 2147483632;
  FFirstDataRow := 2147483632;
  FFirstDataCol := 2147483632;
  FRowPageCount := 0;
  FColPageCount := 0;
end;

procedure TCT_Location.Assign(AItem: TCT_Location);
begin
end;

procedure TCT_Location.CopyTo(AItem: TCT_Location);
begin
end;

{ TCT_PivotFields }

function  TCT_PivotFields.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;

  FAssigneds := [];

  if FItems.Count > 0 then begin
    Inc(Result,FItems.Count);

    FAssigneds := [xaElements];

    for i := 0 to FItems.Count - 1 do
      Inc(Result,Items[i].CheckAssigned);
  end;
end;

function  TCT_PivotFields.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pivotField' then
    Result := Add
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotFields.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    Items[i].WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then begin
      AWriter.BeginTag('pivotField');
      Items[i].Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('pivotField');
  end;
end;

procedure TCT_PivotFields.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FItems.Count > 0 then
    AWriter.AddAttribute('count',XmlIntToStr(FItems.Count));
end;

procedure TCT_PivotFields.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin

end;

constructor TCT_PivotFields.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;

  FItems := TObjectList.Create;
end;

destructor TCT_PivotFields.Destroy;
begin
  FItems.Free;
end;

function TCT_PivotFields.GetItems(Index: integer): TCT_PivotField;
begin
  Result := TCT_PivotField(FItems[Index]);
end;

procedure TCT_PivotFields.Clear;
begin
  FAssigneds := [];

  FItems.Clear;
end;

function TCT_PivotFields.Add: TCT_PivotField;
begin
  Result := TCT_PivotField.Create(FOwner);

  FItems.Add(Result);
end;

procedure TCT_PivotFields.Assign(AItem: TCT_PivotFields);
begin
end;

procedure TCT_PivotFields.CopyTo(AItem: TCT_PivotFields);
begin
end;

function TCT_PivotFields.Count: integer;
begin
  Result := FItems.Count;
end;

{ TCT_RowFields }

function  TCT_RowColFields.CheckAssigned: integer;
var
  i: integer;
begin
  FAssigneds := [];
  if FItems.Count > 0 then begin
    Result := 1;
    FAssigneds := [xaAttributes,xaElements];

    for i := 0 to FItems.Count - 1 do
      Inc(Result,Items[i].CheckAssigned);
  end
  else
    Result := 0;
end;

function  TCT_RowColFields.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;

  if AReader.QName = 'field' then
    Result := Add;

  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_RowColFields.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    Items[i].WriteAttributes(AWriter);
    if xaElements in Items[i].Assigneds then begin
      AWriter.BeginTag('field');
      Items[i].Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('field');
  end;
end;

procedure TCT_RowColFields.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('count',XmlIntToStr(FItems.Count));
end;

procedure TCT_RowColFields.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin

end;

constructor TCT_RowColFields.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FItems := TObjectList.Create;
end;

destructor TCT_RowColFields.Destroy;
begin
  FItems.Free;
end;

function TCT_RowColFields.GetItems(Index: integer): TCT_Field;
begin
  Result := TCT_Field(FItems[Index]);
end;

procedure TCT_RowColFields.Clear;
begin
  FAssigneds := [];

  FItems.Clear;
end;

function TCT_RowColFields.Add: TCT_Field;
begin
  Result := TCT_Field.Create(FOwner);

  FItems.Add(Result);
end;

procedure TCT_RowColFields.Assign(AItem: TCT_RowColFields);
begin
end;

procedure TCT_RowColFields.CopyTo(AItem: TCT_RowColFields);
begin
end;

function TCT_RowColFields.Count: integer;
begin
  Result := FItems.Count;
end;

{ TCT_colItems }

function  TCT_RowColItems.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  FAssigneds := [];
  if FItems.Count > 0 then begin
    FAssigneds := [xaElements];
    Inc(Result,FItems.Count);
    for i := 0 to FItems.Count - 1 do
      Inc(Result,Items[i].CheckAssigned);
  end;
end;

function  TCT_RowColItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'i' then
    Result := Add;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_RowColItems.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  if FItems.Count > 0 then begin
    for i := 0 to FItems.Count - 1 do begin
      Items[i].WriteAttributes(AWriter);
      AWriter.BeginTag('i');
      Items[i].Write(AWriter);
      AWriter.EndTag;
    end;
  end;
end;

procedure TCT_RowColItems.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('count',XmlIntToStr(FItems.Count));
end;

procedure TCT_RowColItems.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin

end;

constructor TCT_RowColItems.Create(AOwner: TXPGDocBase; AIsRow: boolean);
begin
  FOwner := AOwner;
  FIsRow := AIsRow;
  FElementCount := 1;
  FAttributeCount := 1;

  FItems := TObjectList.Create;
end;

destructor TCT_RowColItems.Destroy;
begin
  FItems.Free;
end;

function TCT_RowColItems.GetItems(Index: integer): TCT_I;
begin
  Result := TCT_I(FItems[Index]);
end;

procedure TCT_RowColItems.Clear;
begin
  FAssigneds := [];
  FItems.Clear;
end;

function TCT_RowColItems.Add: TCT_I;
begin
  Result := TCT_I.Create(FOwner);

  FItems.Add(Result)
end;

procedure TCT_RowColItems.Assign(AItem: TCT_RowColItems);
begin
end;

procedure TCT_RowColItems.CopyTo(AItem: TCT_RowColItems);
begin
end;

function TCT_RowColItems.Count: integer;
begin
  Result := FItems.Count;
end;

{ TCT_PageFields }

function  TCT_PageFields.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FPageFieldXpgList <> Nil then 
    Inc(ElemsAssigned,FPageFieldXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PageFields.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pageField' then 
  begin
    if FPageFieldXpgList = Nil then 
      FPageFieldXpgList := TCT_PageFieldXpgList.Create(FOwner);
    Result := FPageFieldXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PageFields.Write(AWriter: TXpgWriteXML);
begin
  if FPageFieldXpgList <> Nil then 
    FPageFieldXpgList.Write(AWriter,'pageField');
end;

procedure TCT_PageFields.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PageFields.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PageFields.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PageFields.Destroy;
begin
  if FPageFieldXpgList <> Nil then 
    FPageFieldXpgList.Free;
end;

procedure TCT_PageFields.Clear;
begin
  FAssigneds := [];
  if FPageFieldXpgList <> Nil then 
    FreeAndNil(FPageFieldXpgList);
  FCount := 2147483632;
end;

procedure TCT_PageFields.Assign(AItem: TCT_PageFields);
begin
end;

procedure TCT_PageFields.CopyTo(AItem: TCT_PageFields);
begin
end;

function  TCT_PageFields.Create_PageFieldXpgList: TCT_PageFieldXpgList;
begin
  if FPageFieldXpgList = Nil then
      FPageFieldXpgList := TCT_PageFieldXpgList.Create(FOwner);
  Result := FPageFieldXpgList ;
end;

{ TCT_Formats }

function  TCT_Formats.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FFormatXpgList <> Nil then 
    Inc(ElemsAssigned,FFormatXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Formats.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'format' then 
  begin
    if FFormatXpgList = Nil then 
      FFormatXpgList := TCT_FormatXpgList.Create(FOwner);
    Result := FFormatXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Formats.Write(AWriter: TXpgWriteXML);
begin
  if FFormatXpgList <> Nil then 
    FFormatXpgList.Write(AWriter,'format');
end;

procedure TCT_Formats.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Formats.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Formats.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 0;
end;

destructor TCT_Formats.Destroy;
begin
  if FFormatXpgList <> Nil then 
    FFormatXpgList.Free;
end;

procedure TCT_Formats.Clear;
begin
  FAssigneds := [];
  if FFormatXpgList <> Nil then 
    FreeAndNil(FFormatXpgList);
  FCount := 0;
end;

procedure TCT_Formats.Assign(AItem: TCT_Formats);
begin
end;

procedure TCT_Formats.CopyTo(AItem: TCT_Formats);
begin
end;

function  TCT_Formats.Create_FormatXpgList: TCT_FormatXpgList;
begin
  if FFormatXpgList = Nil then
      FFormatXpgList := TCT_FormatXpgList.Create(FOwner);
  Result := FFormatXpgList ;
end;

{ TCT_ConditionalFormats }

function  TCT_ConditionalFormats.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FConditionalFormatXpgList <> Nil then 
    Inc(ElemsAssigned,FConditionalFormatXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ConditionalFormats.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'conditionalFormat' then 
  begin
    if FConditionalFormatXpgList = Nil then 
      FConditionalFormatXpgList := TCT_ConditionalFormatXpgList.Create(FOwner);
    Result := FConditionalFormatXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ConditionalFormats.Write(AWriter: TXpgWriteXML);
begin
  if FConditionalFormatXpgList <> Nil then 
    FConditionalFormatXpgList.Write(AWriter,'conditionalFormat');
end;

procedure TCT_ConditionalFormats.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_ConditionalFormats.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ConditionalFormats.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 0;
end;

destructor TCT_ConditionalFormats.Destroy;
begin
  if FConditionalFormatXpgList <> Nil then 
    FConditionalFormatXpgList.Free;
end;

procedure TCT_ConditionalFormats.Clear;
begin
  FAssigneds := [];
  if FConditionalFormatXpgList <> Nil then 
    FreeAndNil(FConditionalFormatXpgList);
  FCount := 0;
end;

procedure TCT_ConditionalFormats.Assign(AItem: TCT_ConditionalFormats);
begin
end;

procedure TCT_ConditionalFormats.CopyTo(AItem: TCT_ConditionalFormats);
begin
end;

function  TCT_ConditionalFormats.Create_ConditionalFormatXpgList: TCT_ConditionalFormatXpgList;
begin
  if FConditionalFormatXpgList = Nil then
      FConditionalFormatXpgList := TCT_ConditionalFormatXpgList.Create(FOwner);
  Result := FConditionalFormatXpgList ;
end;

{ TCT_ChartFormats }

function  TCT_ChartFormats.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FChartFormatXpgList <> Nil then 
    Inc(ElemsAssigned,FChartFormatXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ChartFormats.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'chartFormat' then 
  begin
    if FChartFormatXpgList = Nil then 
      FChartFormatXpgList := TCT_ChartFormatXpgList.Create(FOwner);
    Result := FChartFormatXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ChartFormats.Write(AWriter: TXpgWriteXML);
begin
  if FChartFormatXpgList <> Nil then 
    FChartFormatXpgList.Write(AWriter,'chartFormat');
end;

procedure TCT_ChartFormats.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_ChartFormats.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ChartFormats.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 0;
end;

destructor TCT_ChartFormats.Destroy;
begin
  if FChartFormatXpgList <> Nil then 
    FChartFormatXpgList.Free;
end;

procedure TCT_ChartFormats.Clear;
begin
  FAssigneds := [];
  if FChartFormatXpgList <> Nil then 
    FreeAndNil(FChartFormatXpgList);
  FCount := 0;
end;

procedure TCT_ChartFormats.Assign(AItem: TCT_ChartFormats);
begin
end;

procedure TCT_ChartFormats.CopyTo(AItem: TCT_ChartFormats);
begin
end;

function  TCT_ChartFormats.Create_ChartFormatXpgList: TCT_ChartFormatXpgList;
begin
  if FChartFormatXpgList = Nil then
      FChartFormatXpgList := TCT_ChartFormatXpgList.Create(FOwner);
  Result := FChartFormatXpgList ;
end;

{ TCT_PivotHierarchies }

function  TCT_PivotHierarchies.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FPivotHierarchyXpgList <> Nil then 
    Inc(ElemsAssigned,FPivotHierarchyXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotHierarchies.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'pivotHierarchy' then 
  begin
    if FPivotHierarchyXpgList = Nil then 
      FPivotHierarchyXpgList := TCT_PivotHierarchyXpgList.Create(FOwner);
    Result := FPivotHierarchyXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotHierarchies.Write(AWriter: TXpgWriteXML);
begin
  if FPivotHierarchyXpgList <> Nil then 
    FPivotHierarchyXpgList.Write(AWriter,'pivotHierarchy');
end;

procedure TCT_PivotHierarchies.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PivotHierarchies.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PivotHierarchies.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PivotHierarchies.Destroy;
begin
  if FPivotHierarchyXpgList <> Nil then 
    FPivotHierarchyXpgList.Free;
end;

procedure TCT_PivotHierarchies.Clear;
begin
  FAssigneds := [];
  if FPivotHierarchyXpgList <> Nil then 
    FreeAndNil(FPivotHierarchyXpgList);
  FCount := 2147483632;
end;

procedure TCT_PivotHierarchies.Assign(AItem: TCT_PivotHierarchies);
begin
end;

procedure TCT_PivotHierarchies.CopyTo(AItem: TCT_PivotHierarchies);
begin
end;

function  TCT_PivotHierarchies.Create_PivotHierarchyXpgList: TCT_PivotHierarchyXpgList;
begin
  if FPivotHierarchyXpgList = Nil then
      FPivotHierarchyXpgList := TCT_PivotHierarchyXpgList.Create(FOwner);
  Result := FPivotHierarchyXpgList ;
end;

{ TCT_PivotTableStyle }

function  TCT_PivotTableStyle.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if Byte(FShowRowHeaders) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FShowColHeaders) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FShowRowStripes) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FShowColStripes) <> 2 then 
    Inc(AttrsAssigned);
  if Byte(FShowLastColumn) <> 2 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PivotTableStyle.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PivotTableStyle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if Byte(FShowRowHeaders) <> 2 then 
    AWriter.AddAttribute('showRowHeaders',XmlBoolToStr(FShowRowHeaders));
  if Byte(FShowColHeaders) <> 2 then 
    AWriter.AddAttribute('showColHeaders',XmlBoolToStr(FShowColHeaders));
  if Byte(FShowRowStripes) <> 2 then 
    AWriter.AddAttribute('showRowStripes',XmlBoolToStr(FShowRowStripes));
  if Byte(FShowColStripes) <> 2 then 
    AWriter.AddAttribute('showColStripes',XmlBoolToStr(FShowColStripes));
  if Byte(FShowLastColumn) <> 2 then 
    AWriter.AddAttribute('showLastColumn',XmlBoolToStr(FShowLastColumn));
end;

procedure TCT_PivotTableStyle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000005B5: FShowRowHeaders := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000059B: FShowColHeaders := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005E3: FShowRowStripes := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005C9: FShowColStripes := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005C3: FShowLastColumn := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_PivotTableStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
  Byte(FShowRowHeaders) := 2;
  Byte(FShowColHeaders) := 2;
  Byte(FShowRowStripes) := 2;
  Byte(FShowColStripes) := 2;
  Byte(FShowLastColumn) := 2;
end;

destructor TCT_PivotTableStyle.Destroy;
begin
end;

procedure TCT_PivotTableStyle.Clear;
begin
  FAssigneds := [];
  FName := '';
  Byte(FShowRowHeaders) := 2;
  Byte(FShowColHeaders) := 2;
  Byte(FShowRowStripes) := 2;
  Byte(FShowColStripes) := 2;
  Byte(FShowLastColumn) := 2;
end;

procedure TCT_PivotTableStyle.Assign(AItem: TCT_PivotTableStyle);
begin
end;

procedure TCT_PivotTableStyle.CopyTo(AItem: TCT_PivotTableStyle);
begin
end;

{ TCT_PivotFilters }

function  TCT_PivotFilters.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FFilterXpgList <> Nil then 
    Inc(ElemsAssigned,FFilterXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotFilters.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'filter' then 
  begin
    if FFilterXpgList = Nil then 
      FFilterXpgList := TCT_PivotFilterXpgList.Create(FOwner);
    Result := FFilterXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotFilters.Write(AWriter: TXpgWriteXML);
begin
  if FFilterXpgList <> Nil then 
    FFilterXpgList.Write(AWriter,'filter');
end;

procedure TCT_PivotFilters.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PivotFilters.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PivotFilters.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 0;
end;

destructor TCT_PivotFilters.Destroy;
begin
  if FFilterXpgList <> Nil then 
    FFilterXpgList.Free;
end;

procedure TCT_PivotFilters.Clear;
begin
  FAssigneds := [];
  if FFilterXpgList <> Nil then 
    FreeAndNil(FFilterXpgList);
  FCount := 0;
end;

procedure TCT_PivotFilters.Assign(AItem: TCT_PivotFilters);
begin
end;

procedure TCT_PivotFilters.CopyTo(AItem: TCT_PivotFilters);
begin
end;

function  TCT_PivotFilters.Create_FilterXpgList: TCT_PivotFilterXpgList;
begin
  if FFilterXpgList = Nil then
      FFilterXpgList := TCT_PivotFilterXpgList.Create(FOwner);
  Result := FFilterXpgList ;
end;

{ TCT_RowHierarchiesUsage }

function  TCT_RowHierarchiesUsage.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FRowHierarchyUsageXpgList <> Nil then 
    Inc(ElemsAssigned,FRowHierarchyUsageXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_RowHierarchiesUsage.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'rowHierarchyUsage' then 
  begin
    if FRowHierarchyUsageXpgList = Nil then 
      FRowHierarchyUsageXpgList := TCT_HierarchyUsageXpgList.Create(FOwner);
    Result := FRowHierarchyUsageXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_RowHierarchiesUsage.Write(AWriter: TXpgWriteXML);
begin
  if FRowHierarchyUsageXpgList <> Nil then 
    FRowHierarchyUsageXpgList.Write(AWriter,'rowHierarchyUsage');
end;

procedure TCT_RowHierarchiesUsage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_RowHierarchiesUsage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RowHierarchiesUsage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_RowHierarchiesUsage.Destroy;
begin
  if FRowHierarchyUsageXpgList <> Nil then 
    FRowHierarchyUsageXpgList.Free;
end;

procedure TCT_RowHierarchiesUsage.Clear;
begin
  FAssigneds := [];
  if FRowHierarchyUsageXpgList <> Nil then 
    FreeAndNil(FRowHierarchyUsageXpgList);
  FCount := 2147483632;
end;

procedure TCT_RowHierarchiesUsage.Assign(AItem: TCT_RowHierarchiesUsage);
begin
end;

procedure TCT_RowHierarchiesUsage.CopyTo(AItem: TCT_RowHierarchiesUsage);
begin
end;

function  TCT_RowHierarchiesUsage.Create_RowHierarchyUsageXpgList: TCT_HierarchyUsageXpgList;
begin
  if FRowHierarchyUsageXpgList = Nil then
      FRowHierarchyUsageXpgList := TCT_HierarchyUsageXpgList.Create(FOwner);
  Result := FRowHierarchyUsageXpgList ;
end;

{ TCT_ColHierarchiesUsage }

function  TCT_ColHierarchiesUsage.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FColHierarchyUsageXpgList <> Nil then 
    Inc(ElemsAssigned,FColHierarchyUsageXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ColHierarchiesUsage.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'colHierarchyUsage' then 
  begin
    if FColHierarchyUsageXpgList = Nil then 
      FColHierarchyUsageXpgList := TCT_HierarchyUsageXpgList.Create(FOwner);
    Result := FColHierarchyUsageXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_ColHierarchiesUsage.Write(AWriter: TXpgWriteXML);
begin
  if FColHierarchyUsageXpgList <> Nil then 
    FColHierarchyUsageXpgList.Write(AWriter,'colHierarchyUsage');
end;

procedure TCT_ColHierarchiesUsage.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_ColHierarchiesUsage.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ColHierarchiesUsage.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_ColHierarchiesUsage.Destroy;
begin
  if FColHierarchyUsageXpgList <> Nil then 
    FColHierarchyUsageXpgList.Free;
end;

procedure TCT_ColHierarchiesUsage.Clear;
begin
  FAssigneds := [];
  if FColHierarchyUsageXpgList <> Nil then 
    FreeAndNil(FColHierarchyUsageXpgList);
  FCount := 2147483632;
end;

procedure TCT_ColHierarchiesUsage.Assign(AItem: TCT_ColHierarchiesUsage);
begin
end;

procedure TCT_ColHierarchiesUsage.CopyTo(AItem: TCT_ColHierarchiesUsage);
begin
end;

function  TCT_ColHierarchiesUsage.Create_ColHierarchyUsageXpgList: TCT_HierarchyUsageXpgList;
begin
  if FColHierarchyUsageXpgList = Nil then
      FColHierarchyUsageXpgList := TCT_HierarchyUsageXpgList.Create(FOwner);
  Result := FColHierarchyUsageXpgList ;
end;

{ TCT_CacheSource }

function  TCT_CacheSource.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FWorksheetSource <> Nil then 
    Inc(ElemsAssigned,FWorksheetSource.CheckAssigned);
  if FConsolidation <> Nil then 
    Inc(ElemsAssigned,FConsolidation.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CacheSource.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000064D: begin
      if FWorksheetSource = Nil then 
        FWorksheetSource := TCT_WorksheetSource.Create(FOwner);
      Result := FWorksheetSource;
    end;
    $00000576: begin
      if FConsolidation = Nil then 
        FConsolidation := TCT_Consolidation.Create(FOwner);
      Result := FConsolidation;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CacheSource.Write(AWriter: TXpgWriteXML);
begin
  if (FWorksheetSource <> Nil) and FWorksheetSource.Assigned then 
  begin
    FWorksheetSource.WriteAttributes(AWriter);
    AWriter.SimpleTag('worksheetSource');
  end;
  if (FConsolidation <> Nil) and FConsolidation.Assigned then 
  begin
    FConsolidation.WriteAttributes(AWriter);
    if xaElements in FConsolidation.FAssigneds then 
    begin
      AWriter.BeginTag('consolidation');
      FConsolidation.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('consolidation');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CacheSource.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('type',StrTST_SourceType[Integer(FType)]);
  if FConnectionId <> 0 then 
    AWriter.AddAttribute('connectionId',XmlIntToStr(FConnectionId));
end;

procedure TCT_CacheSource.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do 
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_SourceType(StrToEnum('stst' + AAttributes.Values[i]));
      $000004DD: FConnectionId := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
end;

constructor TCT_CacheSource.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 2;
  FType := TST_SourceType(XPG_UNKNOWN_ENUM);
  FConnectionId := 0;
end;

destructor TCT_CacheSource.Destroy;
begin
  if FWorksheetSource <> Nil then 
    FWorksheetSource.Free;
  if FConsolidation <> Nil then 
    FConsolidation.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_CacheSource.Clear;
begin
  FAssigneds := [];
  if FWorksheetSource <> Nil then 
    FreeAndNil(FWorksheetSource);
  if FConsolidation <> Nil then 
    FreeAndNil(FConsolidation);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FConnectionId := 0;
end;

procedure TCT_CacheSource.Assign(AItem: TCT_CacheSource);
begin
end;

procedure TCT_CacheSource.CopyTo(AItem: TCT_CacheSource);
begin
end;

function  TCT_CacheSource.Create_WorksheetSource: TCT_WorksheetSource;
begin
  if FWorksheetSource = Nil then
    FWorksheetSource := TCT_WorksheetSource.Create(FOwner);
  Result := FWorksheetSource;
end;

function  TCT_CacheSource.Create_Consolidation: TCT_Consolidation;
begin
  if FConsolidation = Nil then
      FConsolidation := TCT_Consolidation.Create(FOwner);
  Result := FConsolidation ;
end;

function  TCT_CacheSource.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_CacheFields }

function  TCT_CacheFields.CheckAssigned: integer;
var
  i: integer;
begin
  FAssigneds := [];
  Result := FItems.Count;
  if Result > 0 then begin
    FAssigneds := [xaAttributes,xaElements];

    for i := 0 to FItems.Count - 1 do
      Inc(Result,Items[i].CheckAssigned);
  end;
end;

function  TCT_CacheFields.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'cacheField' then
    Result := Add
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TCT_CacheFields.Valid: boolean;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Trim(Items[i].Name) = '' then begin
      Result := False;
      Exit;
    end;
  end;
  Result := True;
end;

procedure TCT_CacheFields.Write(AWriter: TXpgWriteXML);
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    Items[i].WriteAttributes(AWriter);

    AWriter.BeginTag('cacheField');
    Items[i].Write(AWriter);
    AWriter.EndTag;
  end;
end;

procedure TCT_CacheFields.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('count',XmlIntToStr(Count));
end;

procedure TCT_CacheFields.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
end;

constructor TCT_CacheFields.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;

  FItems := TObjectList.Create;
end;

destructor TCT_CacheFields.Destroy;
begin
  FItems.Free;
end;

function TCT_CacheFields.Find(AName: AxUCString): TCT_CacheField;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    if Items[i].Name = AName then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TCT_CacheFields.GetItems(Index: integer): TCT_CacheField;
begin
  Result := TCT_CacheField(Fitems[Index]);
end;

procedure TCT_CacheFields.Clear;
begin
  FAssigneds := [];
  FItems.Clear;
end;

function TCT_CacheFields.Add: TCT_CacheField;
begin
  Result := TCT_CacheField.Create(FOwner);
  FItems.Add(Result)
end;

procedure TCT_CacheFields.Assign(AItem: TCT_CacheFields);
begin
end;

procedure TCT_CacheFields.CopyTo(AItem: TCT_CacheFields);
begin
end;

function TCT_CacheFields.Count: integer;
begin
  Result := FItems.Count;
end;

{ TCT_CacheHierarchies }

function  TCT_CacheHierarchies.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FCacheHierarchyXpgList <> Nil then 
    Inc(ElemsAssigned,FCacheHierarchyXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CacheHierarchies.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'cacheHierarchy' then 
  begin
    if FCacheHierarchyXpgList = Nil then 
      FCacheHierarchyXpgList := TCT_CacheHierarchyXpgList.Create(FOwner);
    Result := FCacheHierarchyXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CacheHierarchies.Write(AWriter: TXpgWriteXML);
begin
  if FCacheHierarchyXpgList <> Nil then 
    FCacheHierarchyXpgList.Write(AWriter,'cacheHierarchy');
end;

procedure TCT_CacheHierarchies.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_CacheHierarchies.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CacheHierarchies.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_CacheHierarchies.Destroy;
begin
  if FCacheHierarchyXpgList <> Nil then 
    FCacheHierarchyXpgList.Free;
end;

procedure TCT_CacheHierarchies.Clear;
begin
  FAssigneds := [];
  if FCacheHierarchyXpgList <> Nil then 
    FreeAndNil(FCacheHierarchyXpgList);
  FCount := 2147483632;
end;

procedure TCT_CacheHierarchies.Assign(AItem: TCT_CacheHierarchies);
begin
end;

procedure TCT_CacheHierarchies.CopyTo(AItem: TCT_CacheHierarchies);
begin
end;

function  TCT_CacheHierarchies.Create_CacheHierarchyXpgList: TCT_CacheHierarchyXpgList;
begin
  if FCacheHierarchyXpgList = Nil then
      FCacheHierarchyXpgList := TCT_CacheHierarchyXpgList.Create(FOwner);
  Result := FCacheHierarchyXpgList ;
end;

{ TCT_PCDKPIs }

function  TCT_PCDKPIs.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FKpiXpgList <> Nil then 
    Inc(ElemsAssigned,FKpiXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PCDKPIs.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'kpi' then 
  begin
    if FKpiXpgList = Nil then 
      FKpiXpgList := TCT_PCDKPIXpgList.Create(FOwner);
    Result := FKpiXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PCDKPIs.Write(AWriter: TXpgWriteXML);
begin
  if FKpiXpgList <> Nil then 
    FKpiXpgList.Write(AWriter,'kpi');
end;

procedure TCT_PCDKPIs.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PCDKPIs.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PCDKPIs.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_PCDKPIs.Destroy;
begin
  if FKpiXpgList <> Nil then 
    FKpiXpgList.Free;
end;

procedure TCT_PCDKPIs.Clear;
begin
  FAssigneds := [];
  if FKpiXpgList <> Nil then 
    FreeAndNil(FKpiXpgList);
  FCount := 2147483632;
end;

procedure TCT_PCDKPIs.Assign(AItem: TCT_PCDKPIs);
begin
end;

procedure TCT_PCDKPIs.CopyTo(AItem: TCT_PCDKPIs);
begin
end;

function  TCT_PCDKPIs.Create_KpiXpgList: TCT_PCDKPIXpgList;
begin
  if FKpiXpgList = Nil then
      FKpiXpgList := TCT_PCDKPIXpgList.Create(FOwner);
  Result := FKpiXpgList ;
end;

{ TCT_TupleCache }

function  TCT_TupleCache.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FEntries <> Nil then 
    Inc(ElemsAssigned,FEntries.CheckAssigned);
  if FSets <> Nil then 
    Inc(ElemsAssigned,FSets.CheckAssigned);
  if FQueryCache <> Nil then 
    Inc(ElemsAssigned,FQueryCache.CheckAssigned);
  if FServerFormats <> Nil then 
    Inc(ElemsAssigned,FServerFormats.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_TupleCache.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002FA: begin
      if FEntries = Nil then 
        FEntries := TCT_PCDSDTCEntries.Create(FOwner);
      Result := FEntries;
    end;
    $000001BF: begin
      if FSets = Nil then 
        FSets := TCT_Sets.Create(FOwner);
      Result := FSets;
    end;
    $0000040A: begin
      if FQueryCache = Nil then 
        FQueryCache := TCT_QueryCache.Create(FOwner);
      Result := FQueryCache;
    end;
    $00000573: begin
      if FServerFormats = Nil then 
        FServerFormats := TCT_ServerFormats.Create(FOwner);
      Result := FServerFormats;
    end;
    $00000284: begin
      if FExtLst = Nil then 
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_TupleCache.Write(AWriter: TXpgWriteXML);
begin
  if (FEntries <> Nil) and FEntries.Assigned then 
  begin
    FEntries.WriteAttributes(AWriter);
    if xaElements in FEntries.FAssigneds then 
    begin
      AWriter.BeginTag('entries');
      FEntries.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('entries');
  end;
  if (FSets <> Nil) and FSets.Assigned then 
  begin
    FSets.WriteAttributes(AWriter);
    if xaElements in FSets.FAssigneds then 
    begin
      AWriter.BeginTag('sets');
      FSets.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sets');
  end;
  if (FQueryCache <> Nil) and FQueryCache.Assigned then 
  begin
    FQueryCache.WriteAttributes(AWriter);
    if xaElements in FQueryCache.FAssigneds then 
    begin
      AWriter.BeginTag('queryCache');
      FQueryCache.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('queryCache');
  end;
  if (FServerFormats <> Nil) and FServerFormats.Assigned then 
  begin
    FServerFormats.WriteAttributes(AWriter);
    if xaElements in FServerFormats.FAssigneds then 
    begin
      AWriter.BeginTag('serverFormats');
      FServerFormats.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('serverFormats');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_TupleCache.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 0;
end;

destructor TCT_TupleCache.Destroy;
begin
  if FEntries <> Nil then 
    FEntries.Free;
  if FSets <> Nil then 
    FSets.Free;
  if FQueryCache <> Nil then 
    FQueryCache.Free;
  if FServerFormats <> Nil then 
    FServerFormats.Free;
  if FExtLst <> Nil then 
    FExtLst.Free;
end;

procedure TCT_TupleCache.Clear;
begin
  FAssigneds := [];
  if FEntries <> Nil then 
    FreeAndNil(FEntries);
  if FSets <> Nil then 
    FreeAndNil(FSets);
  if FQueryCache <> Nil then 
    FreeAndNil(FQueryCache);
  if FServerFormats <> Nil then 
    FreeAndNil(FServerFormats);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
end;

procedure TCT_TupleCache.Assign(AItem: TCT_TupleCache);
begin
end;

procedure TCT_TupleCache.CopyTo(AItem: TCT_TupleCache);
begin
end;

function  TCT_TupleCache.Create_Entries: TCT_PCDSDTCEntries;
begin
  if FEntries = Nil then
      FEntries := TCT_PCDSDTCEntries.Create(FOwner);
  Result := FEntries ;
end;

function  TCT_TupleCache.Create_Sets: TCT_Sets;
begin
  if FSets = Nil then
      FSets := TCT_Sets.Create(FOwner);
  Result := FSets ;
end;

function  TCT_TupleCache.Create_QueryCache: TCT_QueryCache;
begin
  if FQueryCache = Nil then
      FQueryCache := TCT_QueryCache.Create(FOwner);
  Result := FQueryCache ;
end;

function  TCT_TupleCache.Create_ServerFormats: TCT_ServerFormats;
begin
  if FServerFormats = Nil then
      FServerFormats := TCT_ServerFormats.Create(FOwner);
  Result := FServerFormats ;
end;

function  TCT_TupleCache.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_CalculatedItems }

function  TCT_CalculatedItems.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FCalculatedItemXpgList <> Nil then 
    Inc(ElemsAssigned,FCalculatedItemXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CalculatedItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'calculatedItem' then 
  begin
    if FCalculatedItemXpgList = Nil then 
      FCalculatedItemXpgList := TCT_CalculatedItemXpgList.Create(FOwner);
    Result := FCalculatedItemXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CalculatedItems.Write(AWriter: TXpgWriteXML);
begin
  if FCalculatedItemXpgList <> Nil then 
    FCalculatedItemXpgList.Write(AWriter,'calculatedItem');
end;

procedure TCT_CalculatedItems.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_CalculatedItems.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CalculatedItems.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_CalculatedItems.Destroy;
begin
  if FCalculatedItemXpgList <> Nil then 
    FCalculatedItemXpgList.Free;
end;

procedure TCT_CalculatedItems.Clear;
begin
  FAssigneds := [];
  if FCalculatedItemXpgList <> Nil then 
    FreeAndNil(FCalculatedItemXpgList);
  FCount := 2147483632;
end;

procedure TCT_CalculatedItems.Assign(AItem: TCT_CalculatedItems);
begin
end;

procedure TCT_CalculatedItems.CopyTo(AItem: TCT_CalculatedItems);
begin
end;

function  TCT_CalculatedItems.Create_CalculatedItemXpgList: TCT_CalculatedItemXpgList;
begin
  if FCalculatedItemXpgList = Nil then
      FCalculatedItemXpgList := TCT_CalculatedItemXpgList.Create(FOwner);
  Result := FCalculatedItemXpgList ;
end;

{ TCT_CalculatedMembers }

function  TCT_CalculatedMembers.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FCalculatedMemberXpgList <> Nil then 
    Inc(ElemsAssigned,FCalculatedMemberXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CalculatedMembers.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'calculatedMember' then 
  begin
    if FCalculatedMemberXpgList = Nil then 
      FCalculatedMemberXpgList := TCT_CalculatedMemberXpgList.Create(FOwner);
    Result := FCalculatedMemberXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_CalculatedMembers.Write(AWriter: TXpgWriteXML);
begin
  if FCalculatedMemberXpgList <> Nil then 
    FCalculatedMemberXpgList.Write(AWriter,'calculatedMember');
end;

procedure TCT_CalculatedMembers.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_CalculatedMembers.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CalculatedMembers.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_CalculatedMembers.Destroy;
begin
  if FCalculatedMemberXpgList <> Nil then 
    FCalculatedMemberXpgList.Free;
end;

procedure TCT_CalculatedMembers.Clear;
begin
  FAssigneds := [];
  if FCalculatedMemberXpgList <> Nil then 
    FreeAndNil(FCalculatedMemberXpgList);
  FCount := 2147483632;
end;

procedure TCT_CalculatedMembers.Assign(AItem: TCT_CalculatedMembers);
begin
end;

procedure TCT_CalculatedMembers.CopyTo(AItem: TCT_CalculatedMembers);
begin
end;

function  TCT_CalculatedMembers.Create_CalculatedMemberXpgList: TCT_CalculatedMemberXpgList;
begin
  if FCalculatedMemberXpgList = Nil then
      FCalculatedMemberXpgList := TCT_CalculatedMemberXpgList.Create(FOwner);
  Result := FCalculatedMemberXpgList ;
end;

{ TCT_Dimensions }

function  TCT_Dimensions.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FDimensionXpgList <> Nil then 
    Inc(ElemsAssigned,FDimensionXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Dimensions.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'dimension' then 
  begin
    if FDimensionXpgList = Nil then 
      FDimensionXpgList := TCT_PivotDimensionXpgList.Create(FOwner);
    Result := FDimensionXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_Dimensions.Write(AWriter: TXpgWriteXML);
begin
  if FDimensionXpgList <> Nil then 
    FDimensionXpgList.Write(AWriter,'dimension');
end;

procedure TCT_Dimensions.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Dimensions.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Dimensions.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_Dimensions.Destroy;
begin
  if FDimensionXpgList <> Nil then 
    FDimensionXpgList.Free;
end;

procedure TCT_Dimensions.Clear;
begin
  FAssigneds := [];
  if FDimensionXpgList <> Nil then 
    FreeAndNil(FDimensionXpgList);
  FCount := 2147483632;
end;

procedure TCT_Dimensions.Assign(AItem: TCT_Dimensions);
begin
end;

procedure TCT_Dimensions.CopyTo(AItem: TCT_Dimensions);
begin
end;

function  TCT_Dimensions.Create_DimensionXpgList: TCT_PivotDimensionXpgList;
begin
  if FDimensionXpgList = Nil then
      FDimensionXpgList := TCT_PivotDimensionXpgList.Create(FOwner);
  Result := FDimensionXpgList ;
end;

{ TCT_MeasureGroups }

function  TCT_MeasureGroups.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMeasureGroupXpgList <> Nil then 
    Inc(ElemsAssigned,FMeasureGroupXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_MeasureGroups.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'measureGroup' then 
  begin
    if FMeasureGroupXpgList = Nil then 
      FMeasureGroupXpgList := TCT_MeasureGroupXpgList.Create(FOwner);
    Result := FMeasureGroupXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_MeasureGroups.Write(AWriter: TXpgWriteXML);
begin
  if FMeasureGroupXpgList <> Nil then 
    FMeasureGroupXpgList.Write(AWriter,'measureGroup');
end;

procedure TCT_MeasureGroups.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_MeasureGroups.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_MeasureGroups.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_MeasureGroups.Destroy;
begin
  if FMeasureGroupXpgList <> Nil then 
    FMeasureGroupXpgList.Free;
end;

procedure TCT_MeasureGroups.Clear;
begin
  FAssigneds := [];
  if FMeasureGroupXpgList <> Nil then 
    FreeAndNil(FMeasureGroupXpgList);
  FCount := 2147483632;
end;

procedure TCT_MeasureGroups.Assign(AItem: TCT_MeasureGroups);
begin
end;

procedure TCT_MeasureGroups.CopyTo(AItem: TCT_MeasureGroups);
begin
end;

function  TCT_MeasureGroups.Create_MeasureGroupXpgList: TCT_MeasureGroupXpgList;
begin
  if FMeasureGroupXpgList = Nil then
      FMeasureGroupXpgList := TCT_MeasureGroupXpgList.Create(FOwner);
  Result := FMeasureGroupXpgList ;
end;

{ TCT_MeasureDimensionMaps }

function  TCT_MeasureDimensionMaps.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FMapXpgList <> Nil then 
    Inc(ElemsAssigned,FMapXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_MeasureDimensionMaps.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QName = 'map' then 
  begin
    if FMapXpgList = Nil then 
      FMapXpgList := TCT_MeasureDimensionMapXpgList.Create(FOwner);
    Result := FMapXpgList.Add;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_MeasureDimensionMaps.Write(AWriter: TXpgWriteXML);
begin
  if FMapXpgList <> Nil then
    FMapXpgList.Write(AWriter,'map');
end;

procedure TCT_MeasureDimensionMaps.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 2147483632 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_MeasureDimensionMaps.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_MeasureDimensionMaps.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCount := 2147483632;
end;

destructor TCT_MeasureDimensionMaps.Destroy;
begin
  if FMapXpgList <> Nil then 
    FMapXpgList.Free;
end;

procedure TCT_MeasureDimensionMaps.Clear;
begin
  FAssigneds := [];
  if FMapXpgList <> Nil then 
    FreeAndNil(FMapXpgList);
  FCount := 2147483632;
end;

procedure TCT_MeasureDimensionMaps.Assign(AItem: TCT_MeasureDimensionMaps);
begin
end;

procedure TCT_MeasureDimensionMaps.CopyTo(AItem: TCT_MeasureDimensionMaps);
begin
end;

function  TCT_MeasureDimensionMaps.Create_MapXpgList: TCT_MeasureDimensionMapXpgList;
begin
  if FMapXpgList = Nil then
      FMapXpgList := TCT_MeasureDimensionMapXpgList.Create(FOwner);
  Result := FMapXpgList ;
end;

{ TCT_PivotCacheRecords }

function  TCT_PivotCacheRecords.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 1;
  FAssigneds := [];
  if FCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FRXpgList <> Nil then 
    Inc(ElemsAssigned,FRXpgList.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotCacheRecords.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000072: begin
      if FRXpgList = Nil then
        FRXpgList := TCT_RecordXpgList.Create(FOwner);
      Result := FRXpgList.Add;
    end;
    $00000284: begin
      if FExtLst = Nil then
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.Assigneds := [xaRead];
end;

procedure TCT_PivotCacheRecords.LoadFromStream(AStream: TStream);
var
  Reader: TXPGReader;
begin
  FOwner.FErrors := TXpgPErrors.Create;
  Reader := TXPGReader.Create(FOwner.FErrors,Self);
  try
    Reader.LoadFromStream(AStream);
  finally
    Reader.Free;
    FOwner.FErrors.Free;
  end;
end;

procedure TCT_PivotCacheRecords.Write(AWriter: TXpgWriteXML);
begin
  if FRXpgList <> Nil then
    FRXpgList.Write(AWriter,'r');
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotCacheRecords.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRootAttributes.Text = '' then
    AWriter.Attributes := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  else
    AWriter.Attributes := FRootAttributes.Text;

  AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_PivotCacheRecords.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do begin
    if AAttributes[i] = 'count' then
      FCount := XmlStrToIntDef(AAttributes.Values[0],0)
    else
      FRootAttributes.Add(AAttributes.AsXmlText2(i));
  end;
end;

constructor TCT_PivotCacheRecords.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FCount := 0;
  FRootAttributes := TStringXpgList.Create;
end;

destructor TCT_PivotCacheRecords.Destroy;
begin
  if FRXpgList <> Nil then
    FRXpgList.Free;
  if FExtLst <> Nil then
    FExtLst.Free;

  FRootAttributes.Free;
end;

procedure TCT_PivotCacheRecords.Clear;
begin
  FAssigneds := [];
  if FRXpgList <> Nil then
    FreeAndNil(FRXpgList);
  if FExtLst <> Nil then
    FreeAndNil(FExtLst);
  FCount := 0;
end;

procedure TCT_PivotCacheRecords.Assign(AItem: TCT_PivotCacheRecords);
begin
end;

procedure TCT_PivotCacheRecords.CopyTo(AItem: TCT_PivotCacheRecords);
begin
end;

function  TCT_PivotCacheRecords.Create_RXpgList: TCT_RecordXpgList;
begin
  if FRXpgList = Nil then
      FRXpgList := TCT_RecordXpgList.Create(FOwner);
  Result := FRXpgList ;
end;

function  TCT_PivotCacheRecords.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_pivotTableDefinition }

function  TCT_pivotTableDefinition.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [xaAttributes];
  if FLocation <> Nil then 
    Inc(ElemsAssigned,FLocation.CheckAssigned);
  if FPivotFields <> Nil then 
    Inc(ElemsAssigned,FPivotFields.CheckAssigned);
  if FRowFields <> Nil then 
    Inc(ElemsAssigned,FRowFields.CheckAssigned);
  if FRowItems <> Nil then 
    Inc(ElemsAssigned,FRowItems.CheckAssigned);
  if FColFields <> Nil then 
    Inc(ElemsAssigned,FColFields.CheckAssigned);
  if FColItems <> Nil then 
    Inc(ElemsAssigned,FColItems.CheckAssigned);
  if FPageFields <> Nil then 
    Inc(ElemsAssigned,FPageFields.CheckAssigned);
  if FDataFields <> Nil then 
    Inc(ElemsAssigned,FDataFields.CheckAssigned);
  if FFormats <> Nil then 
    Inc(ElemsAssigned,FFormats.CheckAssigned);
  if FConditionalFormats <> Nil then 
    Inc(ElemsAssigned,FConditionalFormats.CheckAssigned);
  if FChartFormats <> Nil then 
    Inc(ElemsAssigned,FChartFormats.CheckAssigned);
  if FPivotHierarchies <> Nil then 
    Inc(ElemsAssigned,FPivotHierarchies.CheckAssigned);
  if FPivotTableStyleInfo <> Nil then 
    Inc(ElemsAssigned,FPivotTableStyleInfo.CheckAssigned);
  if FFilters <> Nil then 
    Inc(ElemsAssigned,FFilters.CheckAssigned);
  if FRowHierarchiesUsage <> Nil then 
    Inc(ElemsAssigned,FRowHierarchiesUsage.CheckAssigned);
  if FColHierarchiesUsage <> Nil then 
    Inc(ElemsAssigned,FColHierarchiesUsage.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 1;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_pivotTableDefinition.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000359: begin
      if FLocation = Nil then 
        FLocation := TCT_Location.Create(FOwner);
      Result := FLocation;
    end;
    $00000489: begin
      if FPivotFields = Nil then 
        FPivotFields := TCT_PivotFields.Create(FOwner);
      Result := FPivotFields;
    end;
    $000003AF: begin
      if FRowFields = Nil then 
        FRowFields := TCT_RowColFields.Create(FOwner);
      Result := FRowFields;
    end;
    $0000035A: begin
      if FRowItems = Nil then
        FRowItems := TCT_RowColItems.Create(FOwner,True);
      Result := FRowItems;
    end;
    $00000395: begin
      if FColFields = Nil then
        FColFields := TCT_RowColFields.Create(FOwner);
      Result := FColFields;
    end;
    $00000340: begin
      if FColItems = Nil then 
        FColItems := TCT_RowColItems.Create(FOwner,False);
      Result := FColItems;
    end;
    $000003F4: begin
      if FPageFields = Nil then 
        FPageFields := TCT_PageFields.Create(FOwner);
      Result := FPageFields;
    end;
    $000003F1: begin
//      if FDataFields = Nil then
//        FDataFields := TCT_DataFields.Create(FOwner);
//      Result := FDataFields.Add;
    end;
    $000002FC: begin
      if FFormats = Nil then
        FFormats := TCT_Formats.Create(FOwner);
      Result := FFormats;
    end;
    $00000770: begin
      if FConditionalFormats = Nil then
        FConditionalFormats := TCT_ConditionalFormats.Create(FOwner);
      Result := FConditionalFormats;
    end;
    $000004EE: begin
      if FChartFormats = Nil then
        FChartFormats := TCT_ChartFormats.Create(FOwner);
      Result := FChartFormats;
    end;
    $00000699: begin
      if FPivotHierarchies = Nil then
        FPivotHierarchies := TCT_PivotHierarchies.Create(FOwner);
      Result := FPivotHierarchies;
    end;
    $000007B7: begin
      if FPivotTableStyleInfo = Nil then
        FPivotTableStyleInfo := TCT_PivotTableStyle.Create(FOwner);
      Result := FPivotTableStyleInfo;
    end;
    $000002F9: begin
      if FFilters = Nil then
        FFilters := TCT_PivotFilters.Create(FOwner);
      Result := FFilters;
    end;
    $000007B4: begin
      if FRowHierarchiesUsage = Nil then
        FRowHierarchiesUsage := TCT_RowHierarchiesUsage.Create(FOwner);
      Result := FRowHierarchiesUsage;
    end;
    $0000079A: begin
      if FColHierarchiesUsage = Nil then
        FColHierarchiesUsage := TCT_ColHierarchiesUsage.Create(FOwner);
      Result := FColHierarchiesUsage;
    end;
    $00000284: begin
      if FExtLst = Nil then
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else begin
      if AReader.QName = 'dataField' then begin
        if FDataFields = Nil then
          FDataFields := TCT_DataFields.Create(FOwner);
        Result := FDataFields.Add;
      end
      else
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end;
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure TCT_pivotTableDefinition.Write(AWriter: TXpgWriteXML);
begin
  if (FLocation <> Nil) and FLocation.Assigned then
  begin
    FLocation.WriteAttributes(AWriter);
    AWriter.SimpleTag('location');
  end;
  if (FPivotFields <> Nil) and FPivotFields.Assigned then 
  begin
    FPivotFields.WriteAttributes(AWriter);
    if xaElements in FPivotFields.FAssigneds then 
    begin
      AWriter.BeginTag('pivotFields');
      FPivotFields.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotFields');
  end;
  if (FRowFields <> Nil) and FRowFields.Assigned then 
  begin
    FRowFields.WriteAttributes(AWriter);
    if xaElements in FRowFields.FAssigneds then 
    begin
      AWriter.BeginTag('rowFields');
      FRowFields.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rowFields');
  end;
  if (FRowItems <> Nil) and FRowItems.Assigned then 
  begin
    FRowItems.WriteAttributes(AWriter);
    if xaElements in FRowItems.FAssigneds then 
    begin
      AWriter.BeginTag('rowItems');
      FRowItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rowItems');
  end;
  if (FColFields <> Nil) and FColFields.Assigned then 
  begin
    FColFields.WriteAttributes(AWriter);
    if xaElements in FColFields.FAssigneds then 
    begin
      AWriter.BeginTag('colFields');
      FColFields.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colFields');
  end;
  if (FColItems <> Nil) and FColItems.Assigned then 
  begin
    FColItems.WriteAttributes(AWriter);
    if xaElements in FColItems.FAssigneds then 
    begin
      AWriter.BeginTag('colItems');
      FColItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colItems');
  end;
  if (FPageFields <> Nil) and FPageFields.Assigned then 
  begin
    FPageFields.WriteAttributes(AWriter);
    if xaElements in FPageFields.FAssigneds then 
    begin
      AWriter.BeginTag('pageFields');
      FPageFields.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pageFields');
  end;
  if (FDataFields <> Nil) then begin
    AWriter.AddAttribute('count',IntToStr(FDataFields.Count));
    AWriter.BeginTag('dataFields');
    FDataFields.Write(AWriter,'dataField');
    AWriter.EndTag;
  end;
  if (FFormats <> Nil) and FFormats.Assigned then 
  begin
    FFormats.WriteAttributes(AWriter);
    if xaElements in FFormats.FAssigneds then 
    begin
      AWriter.BeginTag('formats');
      FFormats.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('formats');
  end;
  if (FConditionalFormats <> Nil) and FConditionalFormats.Assigned then 
  begin
    FConditionalFormats.WriteAttributes(AWriter);
    if xaElements in FConditionalFormats.FAssigneds then 
    begin
      AWriter.BeginTag('conditionalFormats');
      FConditionalFormats.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('conditionalFormats');
  end;
  if (FChartFormats <> Nil) and FChartFormats.Assigned then 
  begin
    FChartFormats.WriteAttributes(AWriter);
    if xaElements in FChartFormats.FAssigneds then 
    begin
      AWriter.BeginTag('chartFormats');
      FChartFormats.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('chartFormats');
  end;
  if (FPivotHierarchies <> Nil) and FPivotHierarchies.Assigned then 
  begin
    FPivotHierarchies.WriteAttributes(AWriter);
    if xaElements in FPivotHierarchies.FAssigneds then 
    begin
      AWriter.BeginTag('pivotHierarchies');
      FPivotHierarchies.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotHierarchies');
  end;
  if (FPivotTableStyleInfo <> Nil) and FPivotTableStyleInfo.Assigned then 
  begin
    FPivotTableStyleInfo.WriteAttributes(AWriter);
    AWriter.SimpleTag('pivotTableStyleInfo');
  end;
  if (FFilters <> Nil) and FFilters.Assigned then 
  begin
    FFilters.WriteAttributes(AWriter);
    if xaElements in FFilters.FAssigneds then 
    begin
      AWriter.BeginTag('filters');
      FFilters.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('filters');
  end;
  if (FRowHierarchiesUsage <> Nil) and FRowHierarchiesUsage.Assigned then 
  begin
    FRowHierarchiesUsage.WriteAttributes(AWriter);
    if xaElements in FRowHierarchiesUsage.FAssigneds then 
    begin
      AWriter.BeginTag('rowHierarchiesUsage');
      FRowHierarchiesUsage.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rowHierarchiesUsage');
  end;
  if (FColHierarchiesUsage <> Nil) and FColHierarchiesUsage.Assigned then 
  begin
    FColHierarchiesUsage.WriteAttributes(AWriter);
    if xaElements in FColHierarchiesUsage.FAssigneds then 
    begin
      AWriter.BeginTag('colHierarchiesUsage');
      FColHierarchiesUsage.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colHierarchiesUsage');
  end;
  if (FExtLst <> Nil) and FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_pivotTableDefinition.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRootAttributes.Text = '' then
    AWriter.Attributes := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
  else
    AWriter.Attributes := FRootAttributes.Text;

  AWriter.AddAttribute('name',FName);
  if FCache <> Nil then
    AWriter.AddAttribute('cacheId',XmlIntToStr(FCache.CacheId))
  else
    AWriter.AddAttribute('cacheId',XmlIntToStr(0));
  if FDataOnRows <> False then
    AWriter.AddAttribute('dataOnRows',XmlBoolToStr(FDataOnRows));
  if FDataPosition <> 2147483632 then 
    AWriter.AddAttribute('dataPosition',XmlIntToStr(FDataPosition));
  AWriter.AddAttribute('dataCaption',FDataCaption);
  if FGrandTotalCaption <> '' then 
    AWriter.AddAttribute('grandTotalCaption',FGrandTotalCaption);
  if FErrorCaption <> '' then 
    AWriter.AddAttribute('errorCaption',FErrorCaption);
  if FShowError <> False then 
    AWriter.AddAttribute('showError',XmlBoolToStr(FShowError));
  if FMissingCaption <> '' then 
    AWriter.AddAttribute('missingCaption',FMissingCaption);
  if FShowMissing <> True then 
    AWriter.AddAttribute('showMissing',XmlBoolToStr(FShowMissing));
  if FPageStyle <> '' then 
    AWriter.AddAttribute('pageStyle',FPageStyle);
  if FPivotTableStyle <> '' then 
    AWriter.AddAttribute('pivotTableStyle',FPivotTableStyle);
  if FVacatedStyle <> '' then 
    AWriter.AddAttribute('vacatedStyle',FVacatedStyle);
  if FTag <> '' then 
    AWriter.AddAttribute('tag',FTag);
  if FUpdatedVersion <> 0 then 
    AWriter.AddAttribute('updatedVersion',XmlIntToStr(FUpdatedVersion));
  if FMinRefreshableVersion <> 0 then 
    AWriter.AddAttribute('minRefreshableVersion',XmlIntToStr(FMinRefreshableVersion));
  if FAsteriskTotals <> False then 
    AWriter.AddAttribute('asteriskTotals',XmlBoolToStr(FAsteriskTotals));
  if FShowItems <> True then 
    AWriter.AddAttribute('showItems',XmlBoolToStr(FShowItems));
  if FEditData <> False then 
    AWriter.AddAttribute('editData',XmlBoolToStr(FEditData));
  if FDisableFieldList <> False then 
    AWriter.AddAttribute('disableFieldList',XmlBoolToStr(FDisableFieldList));
  if FShowCalcMbrs <> True then 
    AWriter.AddAttribute('showCalcMbrs',XmlBoolToStr(FShowCalcMbrs));
  if FVisualTotals <> True then 
    AWriter.AddAttribute('visualTotals',XmlBoolToStr(FVisualTotals));
  if FShowMultipleLabel <> True then 
    AWriter.AddAttribute('showMultipleLabel',XmlBoolToStr(FShowMultipleLabel));
  if FShowDataDropDown <> True then 
    AWriter.AddAttribute('showDataDropDown',XmlBoolToStr(FShowDataDropDown));
  if FShowDrill <> True then 
    AWriter.AddAttribute('showDrill',XmlBoolToStr(FShowDrill));
  if FPrintDrill <> False then 
    AWriter.AddAttribute('printDrill',XmlBoolToStr(FPrintDrill));
  if FShowMemberPropertyTips <> True then 
    AWriter.AddAttribute('showMemberPropertyTips',XmlBoolToStr(FShowMemberPropertyTips));
  if FShowDataTips <> True then 
    AWriter.AddAttribute('showDataTips',XmlBoolToStr(FShowDataTips));
  if FEnableWizard <> True then 
    AWriter.AddAttribute('enableWizard',XmlBoolToStr(FEnableWizard));
  if FEnableDrill <> True then 
    AWriter.AddAttribute('enableDrill',XmlBoolToStr(FEnableDrill));
  if FEnableFieldProperties <> True then 
    AWriter.AddAttribute('enableFieldProperties',XmlBoolToStr(FEnableFieldProperties));
  if FPreserveFormatting <> True then 
    AWriter.AddAttribute('preserveFormatting',XmlBoolToStr(FPreserveFormatting));
  if FUseAutoFormatting <> False then 
    AWriter.AddAttribute('useAutoFormatting',XmlBoolToStr(FUseAutoFormatting));
  if FPageWrap <> 0 then 
    AWriter.AddAttribute('pageWrap',XmlIntToStr(FPageWrap));
  if FPageOverThenDown <> False then 
    AWriter.AddAttribute('pageOverThenDown',XmlBoolToStr(FPageOverThenDown));
  if FSubtotalHiddenItems <> False then 
    AWriter.AddAttribute('subtotalHiddenItems',XmlBoolToStr(FSubtotalHiddenItems));
  if FRowGrandTotals <> True then 
    AWriter.AddAttribute('rowGrandTotals',XmlBoolToStr(FRowGrandTotals));
  if FColGrandTotals <> True then 
    AWriter.AddAttribute('colGrandTotals',XmlBoolToStr(FColGrandTotals));
  if FFieldPrintTitles <> False then 
    AWriter.AddAttribute('fieldPrintTitles',XmlBoolToStr(FFieldPrintTitles));
  if FItemPrintTitles <> False then 
    AWriter.AddAttribute('itemPrintTitles',XmlBoolToStr(FItemPrintTitles));
  if FMergeItem <> False then 
    AWriter.AddAttribute('mergeItem',XmlBoolToStr(FMergeItem));
  if FShowDropZones <> True then 
    AWriter.AddAttribute('showDropZones',XmlBoolToStr(FShowDropZones));
  if FCreatedVersion <> 0 then 
    AWriter.AddAttribute('createdVersion',XmlIntToStr(FCreatedVersion));
  if FIndent <> 1 then 
    AWriter.AddAttribute('indent',XmlIntToStr(FIndent));
  if FShowEmptyRow <> False then 
    AWriter.AddAttribute('showEmptyRow',XmlBoolToStr(FShowEmptyRow));
  if FShowEmptyCol <> False then 
    AWriter.AddAttribute('showEmptyCol',XmlBoolToStr(FShowEmptyCol));
  if FShowHeaders <> True then 
    AWriter.AddAttribute('showHeaders',XmlBoolToStr(FShowHeaders));
  if FCompact <> True then 
    AWriter.AddAttribute('compact',XmlBoolToStr(FCompact));
  if FOutline <> False then 
    AWriter.AddAttribute('outline',XmlBoolToStr(FOutline));
  if FOutlineData <> False then 
    AWriter.AddAttribute('outlineData',XmlBoolToStr(FOutlineData));
  if FCompactData <> True then 
    AWriter.AddAttribute('compactData',XmlBoolToStr(FCompactData));
  if FPublished <> False then 
    AWriter.AddAttribute('published',XmlBoolToStr(FPublished));
  if FGridDropZones <> False then 
    AWriter.AddAttribute('gridDropZones',XmlBoolToStr(FGridDropZones));
  if FImmersive <> True then 
    AWriter.AddAttribute('immersive',XmlBoolToStr(FImmersive));
  if FMultipleFieldFilters <> True then 
    AWriter.AddAttribute('multipleFieldFilters',XmlBoolToStr(FMultipleFieldFilters));
  if FChartFormat <> 0 then 
    AWriter.AddAttribute('chartFormat',XmlIntToStr(FChartFormat));
  if FRowHeaderCaption <> '' then 
    AWriter.AddAttribute('rowHeaderCaption',FRowHeaderCaption);
  if FColHeaderCaption <> '' then 
    AWriter.AddAttribute('colHeaderCaption',FColHeaderCaption);
  if FFieldListSortAscending <> False then 
    AWriter.AddAttribute('fieldListSortAscending',XmlBoolToStr(FFieldListSortAscending));
  if FMdxSubqueries <> False then 
    AWriter.AddAttribute('mdxSubqueries',XmlBoolToStr(FMdxSubqueries));
  if FCustomListSort <> True then 
    AWriter.AddAttribute('customListSort',XmlBoolToStr(FCustomListSort));
end;

procedure TCT_pivotTableDefinition.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000002A1: FCacheId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000402: FDataOnRows := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004EF: FDataPosition := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000468: FDataCaption := AAttributes.Values[i];
      $000006DE: FGrandTotalCaption := AAttributes.Values[i];
      $000004F8: FErrorCaption := AAttributes.Values[i];
      $000003CB: FShowError := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005C8: FMissingCaption := AAttributes.Values[i];
      $0000049B: FShowMissing := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000003AE: FPageStyle := AAttributes.Values[i];
      $0000062B: FPivotTableStyle := AAttributes.Values[i];
      $000004E9: FVacatedStyle := AAttributes.Values[i];
      $0000013C: FTag := AAttributes.Values[i];
      $000005CD: FUpdatedVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000088D: FMinRefreshableVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005DD: FAsteriskTotals := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003C3: FShowItems := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000320: FEditData := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000654: FDisableFieldList := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004C8: FShowCalcMbrs := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000050B: FVisualTotals := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000006ED: FShowMultipleLabel := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000668: FShowDataDropDown := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000003B8: FShowDrill := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000424: FPrintDrill := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000091E: FShowMemberPropertyTips := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000004DB: FShowDataTips := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000004D8: FEnableWizard := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000045E: FEnableDrill := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000878: FEnableFieldProperties := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000787: FPreserveFormatting := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000701: FUseAutoFormatting := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000337: FPageWrap := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000660: FPageOverThenDown := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000007BC: FSubtotalHiddenItems := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005BB: FRowGrandTotals := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000005A1: FColGrandTotals := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000686: FFieldPrintTitles := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000631: FItemPrintTitles := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000039F: FMergeItem := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000565: FShowDropZones := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000005BE: FCreatedVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000282: FIndent := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000508: FShowEmptyRow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004EE: FShowEmptyCol := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000047D: FShowHeaders := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000002E7: FCompact := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000300: FOutline := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000047A: FOutlineData := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000461: FCompactData := XmlStrToBoolDef(AAttributes.Values[i],True);
      $000003C0: FPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000054A: FGridDropZones := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003D1: FImmersive := XmlStrToBoolDef(AAttributes.Values[i],True);
      $00000829: FMultipleFieldFilters := XmlStrToBoolDef(AAttributes.Values[i],True);
      $0000047B: FChartFormat := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000066F: FRowHeaderCaption := AAttributes.Values[i];
      $00000655: FColHeaderCaption := AAttributes.Values[i];
      $000008D4: FFieldListSortAscending := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000571: FMdxSubqueries := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005DF: FCustomListSort := XmlStrToBoolDef(AAttributes.Values[i],True);
      else
        FRootAttributes.Add(AAttributes.AsXmlText2(i));
    end;
end;

constructor TCT_pivotTableDefinition.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 17;
  FAttributeCount := 61;
  FCacheId := 2147483632;
  FDataOnRows := False;
  FDataPosition := 2147483632;
  FShowError := False;
  FShowMissing := True;
  FUpdatedVersion := 0;
  FMinRefreshableVersion := 0;
  FAsteriskTotals := False;
  FShowItems := True;
  FEditData := False;
  FDisableFieldList := False;
  FShowCalcMbrs := True;
  FVisualTotals := True;
  FShowMultipleLabel := True;
  FShowDataDropDown := True;
  FShowDrill := True;
  FPrintDrill := False;
  FShowMemberPropertyTips := True;
  FShowDataTips := True;
  FEnableWizard := True;
  FEnableDrill := True;
  FEnableFieldProperties := True;
  FPreserveFormatting := True;
  FUseAutoFormatting := False;
  FPageWrap := 0;
  FPageOverThenDown := False;
  FSubtotalHiddenItems := False;
  FRowGrandTotals := True;
  FColGrandTotals := True;
  FFieldPrintTitles := False;
  FItemPrintTitles := False;
  FMergeItem := False;
  FShowDropZones := True;
  FCreatedVersion := 0;
  FIndent := 1;
  FShowEmptyRow := False;
  FShowEmptyCol := False;
  FShowHeaders := True;
  FCompact := True;
  FOutline := False;
  FOutlineData := False;
  FCompactData := True;
  FPublished := False;
  FGridDropZones := False;
  FImmersive := True;
  FMultipleFieldFilters := True;
  FChartFormat := 0;
  FFieldListSortAscending := False;
  FMdxSubqueries := False;
  FCustomListSort := True;

  FRootAttributes := TStringXpgList.Create;
end;

procedure TCT_pivotTableDefinition.Delete_ColFields;
begin
  if FColFields <> Nil then begin
    FColFields.Free;
    FColFields := Nil;
  end;
end;

procedure TCT_pivotTableDefinition.Delete_ColItems;
begin
  if FColItems <> Nil then begin
    FColItems.Free;
    FColItems := Nil;
  end;
end;

destructor TCT_pivotTableDefinition.Destroy;
begin
  if FLocation <> Nil then
    FLocation.Free;
  if FPivotFields <> Nil then
    FPivotFields.Free;
  if FRowFields <> Nil then
    FRowFields.Free;
  if FRowItems <> Nil then
    FRowItems.Free;
  if FColFields <> Nil then
    FColFields.Free;
  if FColItems <> Nil then
    FColItems.Free;
  if FPageFields <> Nil then
    FPageFields.Free;
  if FDataFields <> Nil then
    FDataFields.Free;
  if FFormats <> Nil then
    FFormats.Free;
  if FConditionalFormats <> Nil then
    FConditionalFormats.Free;
  if FChartFormats <> Nil then
    FChartFormats.Free;
  if FPivotHierarchies <> Nil then
    FPivotHierarchies.Free;
  if FPivotTableStyleInfo <> Nil then
    FPivotTableStyleInfo.Free;
  if FFilters <> Nil then
    FFilters.Free;
  if FRowHierarchiesUsage <> Nil then
    FRowHierarchiesUsage.Free;
  if FColHierarchiesUsage <> Nil then
    FColHierarchiesUsage.Free;
  if FExtLst <> Nil then
    FExtLst.Free;

  FRootAttributes.Free;
end;

function TCT_pivotTableDefinition.GetColItemValue(Index: integer): TXLSSharedItemsValue;
var
  I         : TCT_I;
  Field     : TCT_Field;
  PivField  : TCT_PivotField;
  Item      : TCT_Item;
  CacheField: TCT_CacheField;
begin
  Result := Nil;

  if Index >= ColItems.Count then
    Exit;

  I := ColItems[Index];
  if I.R >= ColFields.Count then
    Exit;

  Field := ColFields[I.R];
  if Field.X >= PivotFields.Count then
    Exit;

  PivField := PivotFields[Field.X];
  if (I.XXpgList.Count <= 0) or (I.XXpgList[0].V >= PivField.Items.Count) then
    Exit;

  Item := PivField.Items[I.XXpgList[0].V];
  CacheField := FCache.CacheFields[Field.X];

  if Item.X >= CacheField.SharedItems.Values.ValueList.Count then
    Exit;

  Result := CacheField.SharedItems.Values.ValueList[Item.X];
end;

function TCT_pivotTableDefinition.GetRowItemValue(Index: integer): TXLSSharedItemsValue;
var
  I         : TCT_I;
  Field     : TCT_Field;
  PivField  : TCT_PivotField;
  Item      : TCT_Item;
  CacheField: TCT_CacheField;
begin
  Result := Nil;

  if Index >= RowItems.Count then
    Exit;

  I := RowItems[Index];
  if I.R >= RowFields.Count then
    Exit;

  Field := RowFields[I.R];
  if Field.X >= PivotFields.Count then
    Exit;

  PivField := PivotFields[Field.X];
  if (I.XXpgList.Count <= 0) or (I.XXpgList[0].V >= PivField.Items.Count) then
    Exit;

  Item := PivField.Items[I.XXpgList[0].V];
  CacheField := FCache.CacheFields[Field.X];

  if Item.X >= CacheField.SharedItems.Values.ValueList.Count then
    Exit;

  Result := CacheField.SharedItems.Values.ValueList[Item.X];
end;

procedure TCT_pivotTableDefinition.Clear;
begin
  FAssigneds := [];
  if FLocation <> Nil then 
    FreeAndNil(FLocation);
  if FPivotFields <> Nil then 
    FreeAndNil(FPivotFields);
  if FRowFields <> Nil then 
    FreeAndNil(FRowFields);
  if FRowItems <> Nil then 
    FreeAndNil(FRowItems);
  if FColFields <> Nil then 
    FreeAndNil(FColFields);
  if FColItems <> Nil then 
    FreeAndNil(FColItems);
  if FPageFields <> Nil then 
    FreeAndNil(FPageFields);
  if FDataFields <> Nil then 
    FreeAndNil(FDataFields);
  if FFormats <> Nil then 
    FreeAndNil(FFormats);
  if FConditionalFormats <> Nil then 
    FreeAndNil(FConditionalFormats);
  if FChartFormats <> Nil then 
    FreeAndNil(FChartFormats);
  if FPivotHierarchies <> Nil then 
    FreeAndNil(FPivotHierarchies);
  if FPivotTableStyleInfo <> Nil then 
    FreeAndNil(FPivotTableStyleInfo);
  if FFilters <> Nil then 
    FreeAndNil(FFilters);
  if FRowHierarchiesUsage <> Nil then 
    FreeAndNil(FRowHierarchiesUsage);
  if FColHierarchiesUsage <> Nil then 
    FreeAndNil(FColHierarchiesUsage);
  if FExtLst <> Nil then 
    FreeAndNil(FExtLst);
  FName := '';
  FCacheId := 2147483632;
  FDataOnRows := False;
  FDataPosition := 2147483632;
  FDataCaption := '';
  FGrandTotalCaption := '';
  FErrorCaption := '';
  FShowError := False;
  FMissingCaption := '';
  FShowMissing := True;
  FPageStyle := '';
  FPivotTableStyle := '';
  FVacatedStyle := '';
  FTag := '';
  FUpdatedVersion := 0;
  FMinRefreshableVersion := 0;
  FAsteriskTotals := False;
  FShowItems := True;
  FEditData := False;
  FDisableFieldList := False;
  FShowCalcMbrs := True;
  FVisualTotals := True;
  FShowMultipleLabel := True;
  FShowDataDropDown := True;
  FShowDrill := True;
  FPrintDrill := False;
  FShowMemberPropertyTips := True;
  FShowDataTips := True;
  FEnableWizard := True;
  FEnableDrill := True;
  FEnableFieldProperties := True;
  FPreserveFormatting := True;
  FUseAutoFormatting := False;
  FPageWrap := 0;
  FPageOverThenDown := False;
  FSubtotalHiddenItems := False;
  FRowGrandTotals := True;
  FColGrandTotals := True;
  FFieldPrintTitles := False;
  FItemPrintTitles := False;
  FMergeItem := False;
  FShowDropZones := True;
  FCreatedVersion := 0;
  FIndent := 1;
  FShowEmptyRow := False;
  FShowEmptyCol := False;
  FShowHeaders := True;
  FCompact := True;
  FOutline := False;
  FOutlineData := False;
  FCompactData := True;
  FPublished := False;
  FGridDropZones := False;
  FImmersive := True;
  FMultipleFieldFilters := True;
  FChartFormat := 0;
  FRowHeaderCaption := '';
  FColHeaderCaption := '';
  FFieldListSortAscending := False;
  FMdxSubqueries := False;
  FCustomListSort := True;
end;

procedure TCT_pivotTableDefinition.Assign(AItem: TCT_pivotTableDefinition);
begin
end;

procedure TCT_pivotTableDefinition.CopyTo(AItem: TCT_pivotTableDefinition);
begin
end;

function  TCT_pivotTableDefinition.Create_Location: TCT_Location;
begin
  if FLocation = Nil then
      FLocation := TCT_Location.Create(FOwner);
  Result := FLocation ;
end;

function  TCT_pivotTableDefinition.Create_PivotFields: TCT_PivotFields;
begin
  if FPivotFields = Nil then
      FPivotFields := TCT_PivotFields.Create(FOwner);
  Result := FPivotFields ;
end;

function  TCT_pivotTableDefinition.Create_RowFields: TCT_RowColFields;
begin
  if FRowFields = Nil then
    FRowFields := TCT_RowColFields.Create(FOwner);
  Result := FRowFields ;
end;

function  TCT_pivotTableDefinition.Create_RowItems: TCT_RowColItems;
begin
  if FRowItems = Nil then
    FRowItems := TCT_RowColItems.Create(FOwner,True);
  Result := FRowItems ;
end;

function  TCT_pivotTableDefinition.Create_ColFields: TCT_RowColFields;
begin
  if FColFields = Nil then
    FColFields := TCT_RowColFields.Create(FOwner);
  Result := FColFields ;
end;

function  TCT_pivotTableDefinition.Create_ColItems: TCT_RowColItems;
begin
  if FColItems = Nil then
      FColItems := TCT_RowColItems.Create(FOwner,False);
  Result := FColItems ;
end;

function  TCT_pivotTableDefinition.Create_PageFields: TCT_PageFields;
begin
  if FPageFields = Nil then
      FPageFields := TCT_PageFields.Create(FOwner);
  Result := FPageFields ;
end;

function  TCT_pivotTableDefinition.Create_DataFields: TCT_DataFields;
begin
  if FDataFields = Nil then
      FDataFields := TCT_DataFields.Create(FOwner);
  Result := FDataFields ;
end;

function  TCT_pivotTableDefinition.Create_Formats: TCT_Formats;
begin
  if FFormats = Nil then
      FFormats := TCT_Formats.Create(FOwner);
  Result := FFormats ;
end;

function  TCT_pivotTableDefinition.Create_ConditionalFormats: TCT_ConditionalFormats;
begin
  if FConditionalFormats = Nil then
      FConditionalFormats := TCT_ConditionalFormats.Create(FOwner);
  Result := FConditionalFormats ;
end;

function  TCT_pivotTableDefinition.Create_ChartFormats: TCT_ChartFormats;
begin
  if FChartFormats = Nil then
      FChartFormats := TCT_ChartFormats.Create(FOwner);
  Result := FChartFormats ;
end;

function  TCT_pivotTableDefinition.Create_PivotHierarchies: TCT_PivotHierarchies;
begin
  if FPivotHierarchies = Nil then
      FPivotHierarchies := TCT_PivotHierarchies.Create(FOwner);
  Result := FPivotHierarchies ;
end;

function  TCT_pivotTableDefinition.Create_PivotTableStyleInfo: TCT_PivotTableStyle;
begin
  if FPivotTableStyleInfo = Nil then
      FPivotTableStyleInfo := TCT_PivotTableStyle.Create(FOwner);
  Result := FPivotTableStyleInfo ;
end;

function  TCT_pivotTableDefinition.Create_Filters: TCT_PivotFilters;
begin
  if FFilters = Nil then
      FFilters := TCT_PivotFilters.Create(FOwner);
  Result := FFilters ;
end;

function  TCT_pivotTableDefinition.Create_RowHierarchiesUsage: TCT_RowHierarchiesUsage;
begin
  if FRowHierarchiesUsage = Nil then
      FRowHierarchiesUsage := TCT_RowHierarchiesUsage.Create(FOwner);
  Result := FRowHierarchiesUsage ;
end;

function  TCT_pivotTableDefinition.Create_ColHierarchiesUsage: TCT_ColHierarchiesUsage;
begin
  if FColHierarchiesUsage = Nil then
      FColHierarchiesUsage := TCT_ColHierarchiesUsage.Create(FOwner);
  Result := FColHierarchiesUsage ;
end;

function  TCT_pivotTableDefinition.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
      FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst ;
end;

{ TCT_PivotCacheDefinition }

procedure TCT_PivotCacheDefinition.CacheValues;
var
  i  : integer;
  Ref: TXLSRelCells;
begin
  Ref := FCacheSource.WorksheetSource.RCells;

  for i := 0 to Ref.ColCount - 1 do begin
    if not FCacheFields[i].SharedItems.Values.Used then
      FCacheFields[i].SharedItems.Values.Clear;

    CacheValues(i);
  end;
end;

procedure TCT_PivotCacheDefinition.CacheValues(ACol: integer);
var
  r  : integer;
  Ref: TXLSRelCells;
begin
  FCacheFields[ACol].SharedItems.Values.Clear;

  Ref := FCacheSource.WorksheetSource.RCells;

  FCacheFields[ACol].SharedItems.Values.ValueList.Resize(Max(2,Ref.RowCount div 16));

  for r := Ref.Row1 + 1 to Ref.Row2 do begin
    case Ref.CellTypeRel[ACol,r] of
      xctNone,
      xctBlank          : FCacheFields[ACol].SharedItems.UpdateBlank;
      xctStringFormula,
      xctString         : FCacheFields[ACol].SharedItems.UpdateString(Ref.AsStringRel[ACol,r]);
      xctFloatFormula,
      xctFloat          : begin
        if Ref.IsDateTime[ACol,r] then
          FCacheFields[ACol].SharedItems.UpdateDate(Ref.AsFloatRel[ACol,r])
        else
          FCacheFields[ACol].SharedItems.UpdateNum(Ref.AsFloatRel[ACol,r]);
      end;
      xctBooleanFormula,
      xctBoolean        : FCacheFields[ACol].SharedItems.UpdateBoolean(Ref.AsBooleanRel[ACol,r]);
    end;
  end;
end;

function  TCT_PivotCacheDefinition.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  if FInvalid <> False then 
    Inc(AttrsAssigned);
  if FSaveData <> True then 
    Inc(AttrsAssigned);
  if FRefreshOnLoad <> False then 
    Inc(AttrsAssigned);
  if FOptimizeMemory <> False then 
    Inc(AttrsAssigned);
  if FEnableRefresh <> True then 
    Inc(AttrsAssigned);
  if FRefreshedBy <> '' then 
    Inc(AttrsAssigned);
  if IsNaN(FRefreshedDate) = False then 
    Inc(AttrsAssigned);
  if FBackgroundQuery <> False then 
    Inc(AttrsAssigned);
  if FMissingItemsLimit <> 2147483632 then 
    Inc(AttrsAssigned);
  if FCreatedVersion <> 0 then 
    Inc(AttrsAssigned);
  if FRefreshedVersion <> 0 then 
    Inc(AttrsAssigned);
  if FMinRefreshableVersion <> 0 then 
    Inc(AttrsAssigned);
  if FRecordCount <> 2147483632 then 
    Inc(AttrsAssigned);
  if FUpgradeOnRefresh <> False then 
    Inc(AttrsAssigned);
  if FTupleCache <> False then 
    Inc(AttrsAssigned);
  if FSupportSubquery <> False then 
    Inc(AttrsAssigned);
  if FSupportAdvancedDrill <> False then 
    Inc(AttrsAssigned);
  if FCacheSource <> Nil then 
    Inc(ElemsAssigned,FCacheSource.CheckAssigned);
  if FCacheFields <> Nil then 
    Inc(ElemsAssigned,FCacheFields.CheckAssigned);
  if FCacheHierarchies <> Nil then 
    Inc(ElemsAssigned,FCacheHierarchies.CheckAssigned);
  if FKpis <> Nil then 
    Inc(ElemsAssigned,FKpis.CheckAssigned);
  if FTupleCache_Dup <> Nil then 
    Inc(ElemsAssigned,FTupleCache_Dup.CheckAssigned);
  if FCalculatedItems <> Nil then 
    Inc(ElemsAssigned,FCalculatedItems.CheckAssigned);
  if FCalculatedMembers <> Nil then 
    Inc(ElemsAssigned,FCalculatedMembers.CheckAssigned);
  if FDimensions <> Nil then 
    Inc(ElemsAssigned,FDimensions.CheckAssigned);
  if FMeasureGroups <> Nil then 
    Inc(ElemsAssigned,FMeasureGroups.CheckAssigned);
  if FMaps <> Nil then 
    Inc(ElemsAssigned,FMaps.CheckAssigned);
  if FExtLst <> Nil then 
    Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotCacheDefinition.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000465: begin
      if FCacheSource = Nil then 
        FCacheSource := TCT_CacheSource.Create(FOwner);
      Result := FCacheSource;
    end;
    $0000044B: begin
      if FCacheFields = Nil then 
        FCacheFields := TCT_CacheFields.Create(FOwner);
      Result := FCacheFields;
    end;
    $0000065B: begin
      if FCacheHierarchies = Nil then 
        FCacheHierarchies := TCT_CacheHierarchies.Create(FOwner);
      Result := FCacheHierarchies;
    end;
    $000001B7: begin
      if FKpis = Nil then 
        FKpis := TCT_PCDKPIs.Create(FOwner);
      Result := FKpis;
    end;
    $000003FE: begin
      if FTupleCache_Dup = Nil then 
        FTupleCache_Dup := TCT_TupleCache.Create(FOwner);
      Result := FTupleCache_Dup;
    end;
    $00000614: begin
      if FCalculatedItems = Nil then 
        FCalculatedItems := TCT_CalculatedItems.Create(FOwner);
      Result := FCalculatedItems;
    end;
    $000006DD: begin
      if FCalculatedMembers = Nil then 
        FCalculatedMembers := TCT_CalculatedMembers.Create(FOwner);
      Result := FCalculatedMembers;
    end;
    $00000439: begin
      if FDimensions = Nil then 
        FDimensions := TCT_Dimensions.Create(FOwner);
      Result := FDimensions;
    end;
    $00000572: begin
      if FMeasureGroups = Nil then 
        FMeasureGroups := TCT_MeasureGroups.Create(FOwner);
      Result := FMeasureGroups;
    end;
    $000001B1: begin
      if FMaps = Nil then 
        FMaps := TCT_MeasureDimensionMaps.Create(FOwner);
      Result := FMaps;
    end;
    $00000284: begin
      if FExtLst = Nil then
        FExtLst := TCT_ExtensionList.Create(FOwner);
      Result := FExtLst;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TCT_PivotCacheDefinition.AddFields(ARef: TXLSRelCells): boolean;
var
  c  : integer;
  n  : integer;
  S  : AxUCString;
  Fld: TCT_CacheField;
begin
  Create_CacheSource;
  FCacheSource.Type_ := ststWorksheet;
  FCacheSource.Create_WorksheetSource;
  FCacheSource.WorksheetSource.RCells := ARef;

  Create_CacheFields;

  for c := ARef.Col1 to ARef.Col2 do begin
    Fld := FCacheFields.Add;
    S := ARef.AsString[c,ARef.Row1];
    n := 1;
    while FCacheFields.Find(S) <> Nil do begin
      Inc(n);
      S := ARef.AsString[c,ARef.Row1] + IntToStr(n);
    end;
    Fld.Name := S;
    Fld.Create_SharedItems;
  end;

  Result := FCacheFields.Valid;
  if not Result then begin
    Clear;
    Exit;
  end;

  FRecordCount := ARef.RowCount - 1;
end;

procedure TCT_PivotCacheDefinition.Release;
begin
  Dec(FUsageCount);
end;

procedure TCT_PivotCacheDefinition.Use;
begin
  Inc(FUsageCount);
end;

procedure TCT_PivotCacheDefinition.Write(AWriter: TXpgWriteXML);
begin
  if (FCacheSource <> Nil) and FCacheSource.Assigned then
  begin
    FCacheSource.WriteAttributes(AWriter);
    if xaElements in FCacheSource.FAssigneds then 
    begin
      AWriter.BeginTag('cacheSource');
      FCacheSource.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cacheSource');
  end;
  if (FCacheFields <> Nil) and FCacheFields.Assigned then 
  begin
    FCacheFields.WriteAttributes(AWriter);
    if xaElements in FCacheFields.FAssigneds then 
    begin
      AWriter.BeginTag('cacheFields');
      FCacheFields.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cacheFields');
  end;
  if (FCacheHierarchies <> Nil) and FCacheHierarchies.Assigned then 
  begin
    FCacheHierarchies.WriteAttributes(AWriter);
    if xaElements in FCacheHierarchies.FAssigneds then 
    begin
      AWriter.BeginTag('cacheHierarchies');
      FCacheHierarchies.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cacheHierarchies');
  end;
  if (FKpis <> Nil) and FKpis.Assigned then 
  begin
    FKpis.WriteAttributes(AWriter);
    if xaElements in FKpis.FAssigneds then 
    begin
      AWriter.BeginTag('kpis');
      FKpis.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('kpis');
  end;
  if (FTupleCache_Dup <> Nil) and FTupleCache_Dup.Assigned then 
    if xaElements in FTupleCache_Dup.FAssigneds then 
    begin
      AWriter.BeginTag('TupleCache_Dup');
      FTupleCache_Dup.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('TupleCache_Dup');
  if (FCalculatedItems <> Nil) and FCalculatedItems.Assigned then 
  begin
    FCalculatedItems.WriteAttributes(AWriter);
    if xaElements in FCalculatedItems.FAssigneds then 
    begin
      AWriter.BeginTag('calculatedItems');
      FCalculatedItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('calculatedItems');
  end;
  if (FCalculatedMembers <> Nil) and FCalculatedMembers.Assigned then 
  begin
    FCalculatedMembers.WriteAttributes(AWriter);
    if xaElements in FCalculatedMembers.FAssigneds then 
    begin
      AWriter.BeginTag('calculatedMembers');
      FCalculatedMembers.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('calculatedMembers');
  end;
  if (FDimensions <> Nil) and FDimensions.Assigned then 
  begin
    FDimensions.WriteAttributes(AWriter);
    if xaElements in FDimensions.FAssigneds then 
    begin
      AWriter.BeginTag('dimensions');
      FDimensions.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dimensions');
  end;
  if (FMeasureGroups <> Nil) and FMeasureGroups.Assigned then 
  begin
    FMeasureGroups.WriteAttributes(AWriter);
    if xaElements in FMeasureGroups.FAssigneds then 
    begin
      AWriter.BeginTag('measureGroups');
      FMeasureGroups.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('measureGroups');
  end;
  if (FMaps <> Nil) and FMaps.Assigned then 
  begin
    FMaps.WriteAttributes(AWriter);
    if xaElements in FMaps.FAssigneds then 
    begin
      AWriter.BeginTag('maps');
      FMaps.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('maps');
  end;
//  if (FExtLst <> Nil) and FExtLst.Assigned then
//    if xaElements in FExtLst.FAssigneds then
//    begin
//      AWriter.BeginTag('extLst');
//      FExtLst.Write(AWriter);
//      AWriter.EndTag;
//    end
//    else
//      AWriter.SimpleTag('extLst');
end;

procedure TCT_PivotCacheDefinition.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRootAttributes.Text = '' then
    AWriter.Attributes := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
  else
    AWriter.Attributes := FRootAttributes.Text;

  if FR_Id <> '' then
    AWriter.AddAttribute('r:id',FR_Id);
  if FInvalid <> False then 
    AWriter.AddAttribute('invalid',XmlBoolToStr(FInvalid));
  if FSaveData <> True then 
    AWriter.AddAttribute('saveData',XmlBoolToStr(FSaveData));
  if FRefreshOnLoad <> False then
    AWriter.AddAttribute('refreshOnLoad',XmlBoolToStr(FRefreshOnLoad));
  if FOptimizeMemory <> False then 
    AWriter.AddAttribute('optimizeMemory',XmlBoolToStr(FOptimizeMemory));
  if FEnableRefresh <> True then 
    AWriter.AddAttribute('enableRefresh',XmlBoolToStr(FEnableRefresh));
  if FRefreshedBy <> '' then 
    AWriter.AddAttribute('refreshedBy',FRefreshedBy);
  if IsNaN(FRefreshedDate) = False then 
    AWriter.AddAttribute('refreshedDate',XmlFloatToStr(FRefreshedDate));
  if FBackgroundQuery <> False then 
    AWriter.AddAttribute('backgroundQuery',XmlBoolToStr(FBackgroundQuery));
  if FMissingItemsLimit <> 2147483632 then 
    AWriter.AddAttribute('missingItemsLimit',XmlIntToStr(FMissingItemsLimit));
  if FCreatedVersion <> 0 then 
    AWriter.AddAttribute('createdVersion',XmlIntToStr(FCreatedVersion));
  if FRefreshedVersion <> 0 then 
    AWriter.AddAttribute('refreshedVersion',XmlIntToStr(FRefreshedVersion));
  if FMinRefreshableVersion <> 0 then 
    AWriter.AddAttribute('minRefreshableVersion',XmlIntToStr(FMinRefreshableVersion));
  if FRecordCount <> 2147483632 then 
    AWriter.AddAttribute('recordCount',XmlIntToStr(FRecordCount));
  if FUpgradeOnRefresh <> False then 
    AWriter.AddAttribute('upgradeOnRefresh',XmlBoolToStr(FUpgradeOnRefresh));
  if FTupleCache <> False then 
    AWriter.AddAttribute('tupleCache',XmlBoolToStr(FTupleCache));
  if FSupportSubquery <> False then 
    AWriter.AddAttribute('supportSubquery',XmlBoolToStr(FSupportSubquery));
  if FSupportAdvancedDrill <> False then 
    AWriter.AddAttribute('supportAdvancedDrill',XmlBoolToStr(FSupportAdvancedDrill));
end;

procedure TCT_PivotCacheDefinition.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
    case AAttributes.HashB[i] of
      $A4781FF3: FR_Id := AAttributes.Values[i];
      $61C4363B: FInvalid := XmlStrToBoolDef(AAttributes.Values[i],False);
      $8F9BB329: FSaveData := XmlStrToBoolDef(AAttributes.Values[i],True);
      $49D6F752: FRefreshOnLoad := XmlStrToBoolDef(AAttributes.Values[i],False);
      $117345FA: FOptimizeMemory := XmlStrToBoolDef(AAttributes.Values[i],False);
      $F017239C: FEnableRefresh := XmlStrToBoolDef(AAttributes.Values[i],True);
      $CE801F77: FRefreshedBy := AAttributes.Values[i];
      $D1AF3C60: FRefreshedDate := XmlStrToFloatDef(AAttributes.Values[i],0);
      $E47B60BC: FBackgroundQuery := XmlStrToBoolDef(AAttributes.Values[i],False);
      $B64FC859: FMissingItemsLimit := XmlStrToIntDef(AAttributes.Values[i],0);
      $C83C8758: FCreatedVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      $72D20174: FRefreshedVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      $0ECFE487: FMinRefreshableVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      $7ABD97FA: FRecordCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $079DEAEE: FUpgradeOnRefresh := XmlStrToBoolDef(AAttributes.Values[i],False);
      $1122EB1E: FTupleCache := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0BAA1F6D: FSupportSubquery := XmlStrToBoolDef(AAttributes.Values[i],False);
      $4D16119A: FSupportAdvancedDrill := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FRootAttributes.Add(AAttributes.AsXmlText2(i));
    end;
end;

constructor TCT_PivotCacheDefinition.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 11;
  FAttributeCount := 18;
  FR_Id := '';
  FInvalid := False;
  FSaveData := True;
  FRefreshOnLoad := False;
  FOptimizeMemory := False;
  FEnableRefresh := True;
  FRefreshedDate := NaN;
  FBackgroundQuery := False;
  FMissingItemsLimit := 2147483632;
  FCreatedVersion := 3;
  FRefreshedVersion := 3;
  FMinRefreshableVersion := 0;
  FRecordCount := 2147483632;
  FUpgradeOnRefresh := False;
  FTupleCache := False;
  FSupportSubquery := False;
  FSupportAdvancedDrill := False;

  FRootAttributes := TStringXpgList.Create;
  FRecords := TCT_PivotCacheRecords.Create(FOwner);
end;

destructor TCT_PivotCacheDefinition.Destroy;
begin
  if FCacheSource <> Nil then
    FCacheSource.Free;
  if FCacheFields <> Nil then
    FCacheFields.Free;
  if FCacheHierarchies <> Nil then
    FCacheHierarchies.Free;
  if FKpis <> Nil then
    FKpis.Free;
  if FTupleCache_Dup <> Nil then
    FTupleCache_Dup.Free;
  if FCalculatedItems <> Nil then
    FCalculatedItems.Free;
  if FCalculatedMembers <> Nil then
    FCalculatedMembers.Free;
  if FDimensions <> Nil then
    FDimensions.Free;
  if FMeasureGroups <> Nil then
    FMeasureGroups.Free;
  if FMaps <> Nil then
    FMaps.Free;
  if FExtLst <> Nil then
    FExtLst.Free;

  FRootAttributes.Free;
  FRecords.Free;
end;

procedure TCT_PivotCacheDefinition.Clear;
begin
  FAssigneds := [];
  if FCacheSource <> Nil then
    FreeAndNil(FCacheSource);
  if FCacheFields <> Nil then
    FreeAndNil(FCacheFields);
  if FCacheHierarchies <> Nil then
    FreeAndNil(FCacheHierarchies);
  if FKpis <> Nil then
    FreeAndNil(FKpis);
  if FTupleCache_Dup <> Nil then
    FreeAndNil(FTupleCache_Dup);
  if FCalculatedItems <> Nil then
    FreeAndNil(FCalculatedItems);
  if FCalculatedMembers <> Nil then
    FreeAndNil(FCalculatedMembers);
  if FDimensions <> Nil then
    FreeAndNil(FDimensions);
  if FMeasureGroups <> Nil then
    FreeAndNil(FMeasureGroups);
  if FMaps <> Nil then
    FreeAndNil(FMaps);
  if FExtLst <> Nil then
    FreeAndNil(FExtLst);
  FR_Id := '';
  FInvalid := False;
  FSaveData := True;
  FRefreshOnLoad := False;
  FOptimizeMemory := False;
  FEnableRefresh := True;
  FRefreshedBy := '';
  FRefreshedDate := NaN;
  FBackgroundQuery := False;
  FMissingItemsLimit := 2147483632;
  FCreatedVersion := 0;
  FRefreshedVersion := 0;
  FMinRefreshableVersion := 0;
  FRecordCount := 2147483632;
  FUpgradeOnRefresh := False;
  FTupleCache := False;
  FSupportSubquery := False;
  FSupportAdvancedDrill := False;

  FRecords.Clear;
end;

procedure TCT_PivotCacheDefinition.Assign(AItem: TCT_PivotCacheDefinition);
begin
end;

procedure TCT_PivotCacheDefinition.CopyTo(AItem: TCT_PivotCacheDefinition);
begin
end;

function  TCT_PivotCacheDefinition.Create_CacheSource: TCT_CacheSource;
begin
  if FCacheSource = Nil then
      FCacheSource := TCT_CacheSource.Create(FOwner);
  Result := FCacheSource ;
end;

function  TCT_PivotCacheDefinition.Create_CacheFields: TCT_CacheFields;
begin
  if FCacheFields = Nil then
      FCacheFields := TCT_CacheFields.Create(FOwner);
  Result := FCacheFields ;
end;

function  TCT_PivotCacheDefinition.Create_CacheHierarchies: TCT_CacheHierarchies;
begin
  if FCacheHierarchies = Nil then
      FCacheHierarchies := TCT_CacheHierarchies.Create(FOwner);
  Result := FCacheHierarchies ;
end;

function  TCT_PivotCacheDefinition.Create_Kpis: TCT_PCDKPIs;
begin
  if FKpis = Nil then
      FKpis := TCT_PCDKPIs.Create(FOwner);
  Result := FKpis ;
end;

function  TCT_PivotCacheDefinition.Create_TupleCache_Dup: TCT_TupleCache;
begin
  if FTupleCache_Dup = Nil then
      FTupleCache_Dup := TCT_TupleCache.Create(FOwner);
  Result := FTupleCache_Dup ;
end;

function  TCT_PivotCacheDefinition.Create_CalculatedItems: TCT_CalculatedItems;
begin
  if FCalculatedItems = Nil then
    FCalculatedItems := TCT_CalculatedItems.Create(FOwner);
  Result := FCalculatedItems ;
end;

function  TCT_PivotCacheDefinition.Create_CalculatedMembers: TCT_CalculatedMembers;
begin
  if FCalculatedMembers = Nil then
    FCalculatedMembers := TCT_CalculatedMembers.Create(FOwner);
  Result := FCalculatedMembers;
end;

function  TCT_PivotCacheDefinition.Create_Dimensions: TCT_Dimensions;
begin
  if FDimensions = Nil then
    FDimensions := TCT_Dimensions.Create(FOwner);
  Result := FDimensions;
end;

function  TCT_PivotCacheDefinition.Create_MeasureGroups: TCT_MeasureGroups;
begin
  if FMeasureGroups = Nil then
    FMeasureGroups := TCT_MeasureGroups.Create(FOwner);
  Result := FMeasureGroups;
end;

function  TCT_PivotCacheDefinition.Create_Maps: TCT_MeasureDimensionMaps;
begin
  if FMaps = Nil then
    FMaps := TCT_MeasureDimensionMaps.Create(FOwner);
  Result := FMaps;
end;

function  TCT_PivotCacheDefinition.Create_ExtLst: TCT_ExtensionList;
begin
  if FExtLst = Nil then
    FExtLst := TCT_ExtensionList.Create(FOwner);
  Result := FExtLst;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if FPivotCacheRecords.ClassName = FCurrWriteClass.ClassName then
    if FPivotCacheRecords <> Nil then
      Inc(ElemsAssigned,FPivotCacheRecords.CheckAssigned);
  if FPivotTableDefinition.ClassName = FCurrWriteClass.ClassName then
    if FPivotTableDefinition <> Nil then
      Inc(ElemsAssigned,FPivotTableDefinition.CheckAssigned);
  if FPivotCacheDefinition.ClassName = FCurrWriteClass.ClassName then
    if FPivotCacheDefinition <> Nil then 
      Inc(ElemsAssigned,FPivotCacheDefinition.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  T__ROOT__.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  i: integer;
begin
  for i := 0 to AReader.Attributes.Count - 1 do
    FRootAttributes.Add(AReader.Attributes.AsXmlText2(i));
  Result := Self;
  case AReader.QNameHashA of
    $000006D8: begin
      if FPivotCacheRecords = Nil then
        FPivotCacheRecords := TCT_PivotCacheRecords.Create(FOwner);
      Result := FPivotCacheRecords;
    end;
    $00000823: begin
      if FPivotTableDefinition = Nil then
        FPivotTableDefinition := TCT_pivotTableDefinition.Create(FOwner);
      Result := FPivotTableDefinition;
    end;
    $0000080F: begin
      if FPivotCacheDefinition = Nil then
        FPivotCacheDefinition := TCT_PivotCacheDefinition.Create(FOwner);
      Result := FPivotCacheDefinition;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;
  if (FCurrWriteClass = Nil) or (FPivotCacheRecords.ClassName = FCurrWriteClass.ClassName) then
    if (FPivotCacheRecords <> Nil) and FPivotCacheRecords.Assigned then
    begin
      FPivotCacheRecords.WriteAttributes(AWriter);
      if xaElements in FPivotCacheRecords.FAssigneds then
      begin
        AWriter.BeginTag('pivotCacheRecords');
        FPivotCacheRecords.Write(AWriter);
        AWriter.EndTag;
      end
      else
        AWriter.SimpleTag('pivotCacheRecords');
    end
    else
      AWriter.SimpleTag('pivotCacheRecords');
  if (FCurrWriteClass = Nil) or (FPivotTableDefinition.ClassName = FCurrWriteClass.ClassName) then
    if (FPivotTableDefinition <> Nil) and FPivotTableDefinition.Assigned then
    begin
      FPivotTableDefinition.WriteAttributes(AWriter);
      if xaElements in FPivotTableDefinition.FAssigneds then
      begin
        AWriter.BeginTag('pivotTableDefinition');
        FPivotTableDefinition.Write(AWriter);
        AWriter.EndTag;
      end
      else
        AWriter.SimpleTag('pivotTableDefinition');
    end
    else 
      AWriter.SimpleTag('pivotTableDefinition');
  if (FCurrWriteClass = Nil) or (FPivotCacheDefinition.ClassName = FCurrWriteClass.ClassName) then 
    if (FPivotCacheDefinition <> Nil) and FPivotCacheDefinition.Assigned then 
    begin
      FPivotCacheDefinition.WriteAttributes(AWriter);
      if xaElements in FPivotCacheDefinition.FAssigneds then 
      begin
        AWriter.BeginTag('pivotCacheDefinition');
        FPivotCacheDefinition.Write(AWriter);
        AWriter.EndTag;
      end
      else 
        AWriter.SimpleTag('pivotCacheDefinition');
    end
    else 
      AWriter.SimpleTag('pivotCacheDefinition');
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 3;
  FAttributeCount := 0;
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  if FPivotCacheRecords <> Nil then 
    FPivotCacheRecords.Free;
  if FPivotTableDefinition <> Nil then 
    FPivotTableDefinition.Free;
  if FPivotCacheDefinition <> Nil then 
    FPivotCacheDefinition.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  if FPivotCacheRecords <> Nil then 
    FreeAndNil(FPivotCacheRecords);
  if FPivotTableDefinition <> Nil then 
    FreeAndNil(FPivotTableDefinition);
  if FPivotCacheDefinition <> Nil then 
    FreeAndNil(FPivotCacheDefinition);
end;

procedure T__ROOT__.Assign(AItem: T__ROOT__);
begin
end;

procedure T__ROOT__.CopyTo(AItem: T__ROOT__);
begin
end;

function  T__ROOT__.Create_PivotCacheRecords: TCT_PivotCacheRecords;
begin
  Result := TCT_PivotCacheRecords.Create(FOwner);
  FPivotCacheRecords := Result;
end;

function  T__ROOT__.Create_PivotTableDefinition: TCT_pivotTableDefinition;
begin
  Result := TCT_pivotTableDefinition.Create(FOwner);
  FPivotTableDefinition := Result;
end;

function  T__ROOT__.Create_PivotCacheDefinition: TCT_PivotCacheDefinition;
begin
  Result := TCT_PivotCacheDefinition.Create(FOwner);
  FPivotCacheDefinition := Result;
end;

{ TCT_XStringElement }

function  TCT_XStringElement.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];
  Result := 1;
end;

procedure TCT_XStringElement.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_XStringElement.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('v',FV);
end;

procedure TCT_XStringElement.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'v' then 
    FV := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_XStringElement.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_XStringElement.Destroy;
begin
end;

procedure TCT_XStringElement.Clear;
begin
  FAssigneds := [];
  FV := '';
end;

procedure TCT_XStringElement.Assign(AItem: TCT_XStringElement);
begin
end;

procedure TCT_XStringElement.CopyTo(AItem: TCT_XStringElement);
begin
end;

{ TXPGDocument }

function  TXPGPivotTable.GetPivotCacheRecords: TCT_PivotCacheRecords;
begin
  Result := FRoot.PivotCacheRecords;
end;

function  TXPGPivotTable.GetPivotTableDefinition: TCT_pivotTableDefinition;
begin
  Result := FRoot.PivotTableDefinition;
end;

function  TXPGPivotTable.GetPivotCacheDefinition: TCT_PivotCacheDefinition;
begin
  Result := FRoot.PivotCacheDefinition;
end;

constructor TXPGPivotTable.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGPivotTable.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGPivotTable.LoadFromFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead);
  try
    FReader.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXPGPivotTable.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGPivotTable.SaveToFile(AFilename: AxUCString; AClassToWrite: TClass);
begin
  FRoot.FCurrWriteClass := AClassToWrite;
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGPivotTable.SaveToStream(AStream: TStream);
begin
  FWriter.SaveToStream(AStream);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

{ TCT_pivotTableDefinitions }

function TCT_pivotTableDefinitions.Add: TCT_pivotTableDefinition;
begin
  Result := TCT_pivotTableDefinition.Create(FOwner);
  FItems.Add(Result)
end;

procedure TCT_pivotTableDefinitions.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  inherited;

  raise Exception.Create('TODO');
end;

function TCT_pivotTableDefinitions.CheckAssigned: integer;
begin
  raise Exception.Create('TODO');
end;

procedure TCT_pivotTableDefinitions.Clear;
begin
  FItems.Clear;
end;

function TCT_pivotTableDefinitions.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TCT_pivotTableDefinitions.Create;
begin
  FOwner := TXPGDocBase.Create;
  FItems := TObjectList.Create;
end;

destructor TCT_pivotTableDefinitions.Destroy;
begin
  FItems.Free;
  FOwner.Free;

  inherited;
end;

function TCT_pivotTableDefinitions.FindByCache(ACache: TCT_pivotCacheDefinition): TCT_pivotTableDefinition;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do begin
    if Items[i].Cache = ACache then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TCT_pivotTableDefinitions.GetItems(Index: integer): TCT_pivotTableDefinition;
begin
  Result := TCT_pivotTableDefinition(FItems[Index]);
end;

function TCT_pivotTableDefinitions.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000823: Result := Add;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TCT_pivotTableDefinitions.LoadFromStream(AStream: TStream): TCT_pivotTableDefinition;
var
  n     : integer;
  Reader: TXPGReader;
begin
  n := FItems.Count;

  FOwner.FErrors := TXpgPErrors.Create;
  Reader := TXPGReader.Create(FOwner.FErrors,Self);
  try
    Reader.LoadFromStream(AStream);

    if FItems.Count > n then
      Result := Items[n]
    else
      Result := Nil;
  finally
    Reader.Free;
    FOwner.FErrors.Free;
  end;
end;

procedure TCT_pivotTableDefinitions.SaveToStream(AStream: TStream; AIndex: integer; AOPC: TObject);
var
  Writer: TXpgWriteXML;
  Def   : TCT_pivotTableDefinition;
begin
  Def := Items[AIndex];

  Def._OPC := AOPC;

  Writer := TXpgWriteXML.Create;
  try
    Writer.SaveToStream(AStream);

    Def.CheckAssigned;

    Def.WriteAttributes(Writer);
    if xaElements in Def.FAssigneds then
    begin
      Writer.BeginTag('pivotTableDefinition');
      Def.Write(Writer);
      Writer.EndTag;
    end
    else
      Writer.SimpleTag('pivotTableDefinition');
  finally
    Writer.Free;
  end;
end;

procedure TCT_pivotTableDefinitions.Write(AWriter: TXpgWriteXML);
begin
  raise Exception.Create('TODO');
end;

procedure TCT_pivotTableDefinitions.WriteAttributes(AWriter: TXpgWriteXML);
begin
  raise Exception.Create('TODO');
end;

{ TCT_PivotCacheDefinitions }

function TCT_PivotCacheDefinitions.Add: TCT_pivotCacheDefinition;
begin
  Result := TCT_pivotCacheDefinition.Create(FOwner);
  FItems.Add(Result)
end;

procedure TCT_PivotCacheDefinitions.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  inherited;

  raise Exception.Create('TODO');
end;

function TCT_PivotCacheDefinitions.CheckAssigned: integer;
begin
  raise Exception.Create('TODO');
end;

procedure TCT_PivotCacheDefinitions.Clear;
begin
  FItems.Clear;
end;

function TCT_PivotCacheDefinitions.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TCT_PivotCacheDefinitions.Create;
begin
  FOwner := TXPGDocBase.Create;
  FItems := TObjectList.Create;
end;

destructor TCT_PivotCacheDefinitions.Destroy;
begin
  FItems.Free;
  FOwner.Free;

  inherited;
end;

procedure TCT_PivotCacheDefinitions.Enumerate;
var
  i: integer;
begin
  for i := 0 to FItems.Count - 1 do
    Items[i].CacheId := i;
end;

function TCT_PivotCacheDefinitions.Find(const ACacheId: integer): TCT_pivotCacheDefinition;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].CacheId = ACacheId then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TCT_PivotCacheDefinitions.GetItems(Index: integer): TCT_pivotCacheDefinition;
begin
  Result := TCT_pivotCacheDefinition(FItems[Index]);
end;

function TCT_PivotCacheDefinitions.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000080F: Result := Add;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.Assigneds := [xaRead];
end;

function TCT_PivotCacheDefinitions.LoadFromStream(AStream: TStream): TCT_pivotCacheDefinition;
var
  n     : integer;
  Reader: TXPGReader;
begin
  n := FItems.Count;

  FOwner.FErrors := TXpgPErrors.Create;
  Reader := TXPGReader.Create(FOwner.FErrors,Self);
  try
    Reader.LoadFromStream(AStream);

    if FItems.Count > n then
      Result := Items[n]
    else
      Result := Nil;
  finally
    Reader.Free;
    FOwner.FErrors.Free;
  end;
end;

procedure TCT_PivotCacheDefinitions.RecordsSaveToStream(AStream: TStream; AIndex: integer);
var
  Writer: TXpgWriteXML;
  Def   : TCT_PivotCacheDefinition;
begin
  Def := Items[AIndex];

  Writer := TXpgWriteXML.Create;
  try
    Writer.SaveToStream(AStream);

    Def.Records.CheckAssigned;

    Def.Records.WriteAttributes(Writer);
    if xaElements in Def.Records.FAssigneds then
    begin
      Writer.BeginTag('pivotCacheRecords');
      Def.Records.Write(Writer);
      Writer.EndTag;
    end
    else
      Writer.SimpleTag('pivotCacheRecords');
  finally
    Writer.Free;
  end;
end;

procedure TCT_PivotCacheDefinitions.SaveToStream(AStream: TStream; AIndex: integer);
var
  Writer: TXpgWriteXML;
  Def   : TCT_PivotCacheDefinition;
begin
  Def := Items[AIndex];

  Writer := TXpgWriteXML.Create;
  try
    Writer.SaveToStream(AStream);

    Def.CheckAssigned;

    Def.WriteAttributes(Writer);
    if xaElements in Def.FAssigneds then
    begin
      Writer.BeginTag('pivotCacheDefinition');
      Def.Write(Writer);
      Writer.EndTag;
    end
    else
      Writer.SimpleTag('pivotCacheDefinition');
  finally
    Writer.Free;
  end;
end;

procedure TCT_PivotCacheDefinitions.Write(AWriter: TXpgWriteXML);
begin
  raise Exception.Create('TODO');
end;

procedure TCT_PivotCacheDefinitions.WriteAttributes(AWriter: TXpgWriteXML);
begin
  raise Exception.Create('TODO');
end;

{ TXLSSharedItemsValuesHash }

procedure TXLSSharedItemsValuesHash.AddBoolean(AValue: boolean);
begin
  if FUsed and FValueList.Add(AValue) then
    Inc(FBooleanCount);
end;

procedure TXLSSharedItemsValuesHash.AddDateTime(AValue: TDateTime);
begin
  if FUsed then begin
    if FValueList.Add(AValue) then
      Inc(FDateCount);
  end
  else
    Inc(FDateCount);

  FMinDate := Min(FMinDate,AValue);
  FMaxDate := Max(FMaxDate,AValue);
end;

procedure TXLSSharedItemsValuesHash.Add;
begin
  if FUsed and FValueList.Add then
    Inc(FBlankCount);
end;

procedure TXLSSharedItemsValuesHash.AddError(AValue: TXc12CellError);
begin
  if FUsed and FValueList.Add(AValue) then
    Inc(FErrorCount);
end;

procedure TXLSSharedItemsValuesHash.AddString(AValue: AxUCString);
begin
  if FUsed and FValueList.Add(AValue) then begin
    Inc(FStringCount);
    FLongText := not FLongText and (Length(AValue) > 255);
  end;
end;

procedure TXLSSharedItemsValuesHash.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  d: double;
  v: AxUCString;
  u: boolean;
begin
  u := False;

  for i := 0 to AAttributes.Count - 1 do begin
    case AAttributes[i][1] of
      'v': v := AAttributes.Values[i];
      'u': u := XmlStrToBool(AAttributes.Values[i]);
    end;
  end;

  case FCurrType of
    xsivtBoolean: AddBoolean(XmlStrToBool(v));
    xsivtDate   : AddDateTime(XmlStrToDateTime(v));
    xsivtError  : AddError(ErrorTextToCellError(v));
    // Blank is already added.
    xsivtBlank  : ;
    xsivtNumeric: begin
      d := XmlStrToFloat(v);
      AddFloat(d); // n
    end;
    xsivtString : AddString(v);
  end;

  FValueList[FValueList.Count - 1].Unused := u;
end;

procedure TXLSSharedItemsValuesHash.AddFloat(AValue: double);
begin
  if FUsed then begin
    if FValueList.Add(AValue) then begin
      if (Frac(AValue) = 0) and (AValue >= -MAXINT) and (AValue <= MAXINT) then
        Inc(FIntegerCount);

      Inc(FNumericCount);
    end;
  end
  else begin
    if (Frac(AValue) = 0) and (AValue >= -MAXINT) and (AValue <= MAXINT) then
      Inc(FIntegerCount);

    Inc(FNumericCount);
  end;

  FMinValue := Min(FMinValue,AValue);
  FMaxValue := Max(FMaxValue,AValue);
end;

function TXLSSharedItemsValuesHash.CanSort: boolean;
begin
  Result := (FStringCount > 1) and ((FStringCount / Count) >= 0.75);
end;

function TXLSSharedItemsValuesHash.CheckAssigned: integer;
begin
  FAssigneds := [xaAttributes];

  Result := 1;
end;

procedure TXLSSharedItemsValuesHash.Clear;
begin
  FBooleanCount := 0;
  FDateCount    := 0;
  FErrorCount   := 0;
  FBlankCount   := 0;
  FNumericCount := 0;
  FIntegerCount := 0;
  FStringCount  := 0;

  FMinValue := MAXDOUBLE;
  FMaxValue := MINDOUBLE;
  FMinDate := MAXDOUBLE;
  FMaxDate := MINDOUBLE;
  FLongText := False;

// Don't clear used.
//  FUsed := False;

  FValueList.Clear;
end;

constructor TXLSSharedItemsValuesHash.Create(Size: Cardinal);
begin
  FMinValue := MAXDOUBLE;
  FMaxValue := MINDOUBLE;
  FMinDate := MAXDOUBLE;
  FMaxDate := MINDOUBLE;
  FLongText := False;

  FValueList := TXLSUniqueSharedItemsValues.Create;
end;

destructor TXLSSharedItemsValuesHash.Destroy;
begin
  Clear;

  FValueList.Free;

  inherited;
end;

function TXLSSharedItemsValuesHash.GetCount: integer;
begin
  if FUsed then
    Result := FBooleanCount + FDateCount + FErrorCount + FBlankCount + FNumericCount + FStringCount
  else
    Result := 0;
end;

function TXLSSharedItemsValuesHash.HashOf(AValue: TDateTime): longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @AValue;

  Result := 0;
  for i := 0 to SizeOf(AValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

function TXLSSharedItemsValuesHash.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
end;

function TXLSSharedItemsValuesHash.HashOf(AValue: AxUCString): longword;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(AValue) do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(AValue[i]);
end;

function TXLSSharedItemsValuesHash.IsDayText: boolean;
var
  i,j: integer;
  S  : AxUCString;
  Ok : boolean;
begin
  Result := False;

  for i := 0 to Min(FValueList.Count - 1,25) do begin
    S := Lowercase(FValueList[i].AsText);

    for j := Low(FormatSettings.ShortDayNames) to High(FormatSettings.ShortDayNames) do begin
      Ok := S = Lowercase(FormatSettings.ShortDayNames[j]);
      if not Ok then
        Ok := S = Lowercase(FormatSettings.LongDayNames[j]);

      if Ok then
        Exit;
    end;
  end;

  Result := True;
end;

function TXLSSharedItemsValuesHash.IsMonthText: boolean;
var
  i,j: integer;
  S  : AxUCString;
  Ok : boolean;
begin
  Result := False;

  for i := 0 to Min(FValueList.Count - 1,25) do begin
    S := Lowercase(FValueList[i].AsText);

    for j := Low(FormatSettings.ShortMonthNames) to High(FormatSettings.ShortMonthNames) do begin
      Ok := S = Lowercase(FormatSettings.ShortMonthNames[j]);
      if not Ok then
        Ok := S = Lowercase(FormatSettings.LongMonthNames[j]);

      if Ok then
        Exit;
    end;
  end;

  Result := True;
end;

function TXLSSharedItemsValuesHash.IsNumeric: boolean;
begin
  Result := True;

       if FNumericCount < FBooleanCount then Result := False
  else if FNumericCount < FDateCount    then Result := False
  else if FNumericCount < FErrorCount   then Result := False
  else if FNumericCount < FBlankCount   then Result := False
  else if FNumericCount < FIntegerCount then Result := False
  else if FNumericCount < FStringCount  then Result := False;
end;

procedure TXLSSharedItemsValuesHash.Write(AWriter: TXpgWriteXML);
var
  i: integer;
  e: AxUCChar;
  v: AxUCString;
begin
  for i := 0 to FValueList.Count - 1 do begin
    e := 'n';

    case FValueList[i].Type_ of
      xsivtBoolean : begin
        e := 'b';
        v := XmlBoolToStr(FValueList[i].AsBoolean);
      end;
      xsivtDate   : begin
        e := 'd';
        v := XmlDateTimeToStr(FValueList[i].AsDate);
      end;
      xsivtError  : begin
        e := 'b';
        v := Xc12CellErrorNames[FValueList[i].AsError];
      end;
      xsivtBlank  : begin
        e := 'm';
      end;
      xsivtNumeric: begin
        e := 'n';
        v := XmlFloatToStr(FValueList[i].AsNumeric);
      end;
      xsivtString : begin
        e := 's';
        v := FValueList[i].AsString;
      end;
    end;
    if (e <> '') and (e <> 'm') then
      AWriter.AddAttribute('v',v);
    if FValueList[i].Unused then
      AWriter.AddAttribute('u','1');
    AWriter.SimpleTag(e);
  end;
end;

procedure TXLSSharedItemsValuesHash.WriteAttributes(AWriter: TXpgWriteXML);
begin

end;

function TXLSSharedItemsValuesHash.HashOf(AValue: double): longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @AValue;

  Result := 0;
  for i := 0 to SizeOf(AValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

function TXLSSharedItemsValuesHash.HashOf(AValue: TXc12CellError): longword;
var
  i: Integer;
  P: PByteArray;
begin
  P := @AValue;

  Result := 0;
  for i := 0 to SizeOf(AValue) - 1 do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(P[i]);
end;

initialization
{$ifdef DELPHI_5}
  L_Enums := TStringList.Create;
{$else}
  L_Enums := THashedStringList.Create;
{$endif}
  AddEnums;

finalization
  L_Enums.Free;

end.
