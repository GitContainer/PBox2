unit xpgParserXLSX;

// Copyright (c) 2010,2011 Axolot Data
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
// File created on 2011-12-03 17:14:41

{$I AxCompilers.inc}

{.$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

// TODO xmlns:x in more attributes than root. Example externalLink.

// TODO Consider sequence. If sequence contains elements A B C then C can't be written without first writing A and B.

interface

uses Classes, SysUtils, Contnrs, IniFiles, Math, xpgPUtils, xpgPLists, xpgPXMLUtils,
XLSUtils5,
xpgPXML;

type TST_UnderlineValues =  (stuvSingle,stuvDouble,stuvSingleAccounting,stuvDoubleAccounting,stuvNone);
const StrTST_UnderlineValues: array[0..4] of AxUCString = ('single','double','singleAccounting','doubleAccounting','none');
type TST_VerticalAlignRun =  (stvarBaseline,stvarSuperscript,stvarSubscript);
const StrTST_VerticalAlignRun: array[0..2] of AxUCString = ('baseline','superscript','subscript');
type TST_VerticalAlignment =  (stvaTop,stvaCenter,stvaBottom,stvaJustify,stvaDistributed);
const StrTST_VerticalAlignment: array[0..4] of AxUCString = ('top','center','bottom','justify','distributed');
type TST_BorderStyle =  (stbsNone,stbsThin,stbsMedium,stbsDashed,stbsDotted,stbsThick,stbsDouble,stbsHair,stbsMediumDashed,stbsDashDot,stbsMediumDashDot,stbsDashDotDot,stbsMediumDashDotDot,stbsSlantDashDot);
const StrTST_BorderStyle: array[0..13] of AxUCString = ('none','thin','medium','dashed','dotted','thick','double','hair','mediumDashed','dashDot','mediumDashDot','dashDotDot','mediumDashDotDot','slantDashDot');
type TST_PatternType =  (stptNone,stptSolid,stptMediumGray,stptDarkGray,stptLightGray,stptDarkHorizontal,stptDarkVertical,stptDarkDown,stptDarkUp,stptDarkGrid,stptDarkTrellis,stptLightHorizontal,stptLightVertical,stptLightDown,stptLightUp,stptLightGrid,stptLightTrellis,stptGray125,stptGray0625);
const StrTST_PatternType: array[0..18] of AxUCString = ('none','solid','mediumGray','darkGray','lightGray','darkHorizontal','darkVertical','darkDown','darkUp','darkGrid','darkTrellis','lightHorizontal','lightVertical','lightDown','lightUp','lightGrid','lightTrellis','gray125','gray0625');
type TST_FontScheme =  (stfsNone,stfsMajor,stfsMinor);
const StrTST_FontScheme: array[0..2] of AxUCString = ('none','major','minor');
type TST_GradientType =  (stgtLinear,stgtPath);
const StrTST_GradientType: array[0..1] of AxUCString = ('linear','path');
type TST_TableStyleType =  (sttstWholeTable,sttstHeaderRow,sttstTotalRow,sttstFirstColumn,sttstLastColumn,sttstFirstRowStripe,sttstSecondRowStripe,sttstFirstColumnStripe,sttstSecondColumnStripe,sttstFirstHeaderCell,sttstLastHeaderCell,sttstFirstTotalCell,sttstLastTotalCell,sttstFirstSubtotalColumn,sttstSecondSubtotalColumn,sttstThirdSubtotalColumn,sttstFirstSubtotalRow,sttstSecondSubtotalRow,sttstThirdSubtotalRow,sttstBlankRow,sttstFirstColumnSubheading,sttstSecondColumnSubheading,sttstThirdColumnSubheading,sttstFirstRowSubheading,sttstSecondRowSubheading,sttstThirdRowSubheading,sttstPageFieldLabels,sttstPageFieldValues);
const StrTST_TableStyleType: array[0..27] of AxUCString = ('wholeTable','headerRow','totalRow','firstColumn','lastColumn','firstRowStripe','secondRowStripe','firstColumnStripe','secondColumnStripe','firstHeaderCell','lastHeaderCell','firstTotalCell','lastTotalCell','firstSubtotalColumn','secondSubtotalColumn','thirdSubtotalColumn','firstSubtotalRow','secondSubtotalRow','thirdSubtotalRow','blankRow','firstColumnSubheading','secondColumnSubheading','thirdColumnSubheading','firstRowSubheading','secondRowSubheading','thirdRowSubheading','pageFieldLabels','pageFieldValues');
type TST_HorizontalAlignment =  (sthaGeneral,sthaLeft,sthaCenter,sthaRight,sthaFill,sthaJustify,sthaCenterContinuous,sthaDistributed);
const StrTST_HorizontalAlignment: array[0..7] of AxUCString = ('general','left','center','right','fill','justify','centerContinuous','distributed');
type TST_PhoneticAlignment =  (stpaNoControl,stpaLeft,stpaCenter,stpaDistributed);
const StrTST_PhoneticAlignment: array[0..3] of AxUCString = ('noControl','left','center','distributed');
type TST_PhoneticType =  (stptHalfwidthKatakana,stptFullwidthKatakana,stptHiragana,stptNoConversion);
const StrTST_PhoneticType: array[0..3] of AxUCString = ('halfwidthKatakana','fullwidthKatakana','Hiragana','noConversion');
type TST_DateTimeGrouping =  (stdtgYear,stdtgMonth,stdtgDay,stdtgHour,stdtgMinute,stdtgSecond);
const StrTST_DateTimeGrouping: array[0..5] of AxUCString = ('year','month','day','hour','minute','second');
type TST_SortMethod =  (stsmStroke,stsmPinYin,stsmNone);
const StrTST_SortMethod: array[0..2] of AxUCString = ('stroke','pinYin','none');
type TST_CalendarType =  (stctNone,stctGregorian,stctGregorianUs,stctJapan,stctTaiwan,stctKorea,stctHijri,stctThai,stctHebrew,stctGregorianMeFrench,stctGregorianArabic,stctGregorianXlitEnglish,stctGregorianXlitFrench);
const StrTST_CalendarType: array[0..12] of AxUCString = ('none','gregorian','gregorianUs','japan','taiwan','korea','hijri','thai','hebrew','gregorianMeFrench','gregorianArabic','gregorianXlitEnglish','gregorianXlitFrench');
type TST_SortBy =  (stsbValue,stsbCellColor,stsbFontColor,stsbIcon);
const StrTST_SortBy: array[0..3] of AxUCString = ('value','cellColor','fontColor','icon');
type TST_IconSetType =  (stist3Arrows,stist3ArrowsGray,stist3Flags,stist3TrafficLights1,stist3TrafficLights2,stist3Signs,stist3Symbols,stist3Symbols2,stist4Arrows,stist4ArrowsGray,stist4RedToBlack,stist4Rating,stist4TrafficLights,stist5Arrows,stist5ArrowsGray,stist5Rating,stist5Quarters);
const StrTST_IconSetType: array[0..16] of AxUCString = ('3Arrows','3ArrowsGray','3Flags','3TrafficLights1','3TrafficLights2','3Signs','3Symbols','3Symbols2','4Arrows','4ArrowsGray','4RedToBlack','4Rating','4TrafficLights','5Arrows','5ArrowsGray','5Rating','5Quarters');
type TST_DynamicFilterType =  (stdftNull,stdftAboveAverage,stdftBelowAverage,stdftTomorrow,stdftToday,stdftYesterday,stdftNextWeek,stdftThisWeek,stdftLastWeek,stdftNextMonth,stdftThisMonth,stdftLastMonth,stdftNextQuarter,stdftThisQuarter,stdftLastQuarter,stdftNextYear,stdftThisYear,stdftLastYear,stdftYearToDate,stdftQ1,stdftQ2,stdftQ3,stdftQ4,stdftM1,stdftM2,stdftM3,stdftM4,stdftM5,stdftM6,stdftM7,stdftM8,stdftM9,stdftM10,stdftM11,stdftM12);
const StrTST_DynamicFilterType: array[0..34] of AxUCString = ('null','aboveAverage','belowAverage','tomorrow','today','yesterday','nextWeek','thisWeek','lastWeek','nextMonth','thisMonth','lastMonth','nextQuarter','thisQuarter','lastQuarter','nextYear','thisYear','lastYear','yearToDate','Q1','Q2','Q3','Q4','M1','M2','M3','M4','M5','M6','M7','M8','M9','M10','M11','M12');
type TST_FilterOperator =  (stfoEqual,stfoLessThan,stfoLessThanOrEqual,stfoNotEqual,stfoGreaterThanOrEqual,stfoGreaterThan);
const StrTST_FilterOperator: array[0..5] of AxUCString = ('equal','lessThan','lessThanOrEqual','notEqual','greaterThanOrEqual','greaterThan');
type TST_Axis =  (staAxisRow,staAxisCol,staAxisPage,staAxisValues);
const StrTST_Axis: array[0..3] of AxUCString = ('axisRow','axisCol','axisPage','axisValues');
type TST_PivotAreaType =  (stpatNone,stpatNormal,stpatData,stpatAll,stpatOrigin,stpatButton,stpatTopRight);
const StrTST_PivotAreaType: array[0..6] of AxUCString = ('none','normal','data','all','origin','button','topRight');
type TST_Visibility =  (stvVisible,stvHidden,stvVeryHidden);
const StrTST_Visibility: array[0..2] of AxUCString = ('visible','hidden','veryHidden');
type TST_CalcMode =  (stcmManual,stcmAuto,stcmAutoNoTable);
const StrTST_CalcMode: array[0..2] of AxUCString = ('manual','auto','autoNoTable');
type TST_UpdateLinks =  (stulUserSet,stulNever,stulAlways);
const StrTST_UpdateLinks: array[0..2] of AxUCString = ('userSet','never','always');
type TST_TargetScreenSize =  (sttss544x376,sttss640x480,sttss720x512,sttss800x600,sttss1024x768,sttss1152x882,sttss1152x900,sttss1280x1024,sttss1600x1200,sttss1800x1440,sttss1920x1200);
const StrTST_TargetScreenSize: array[0..10] of AxUCString = ('544x376','640x480','720x512','800x600','1024x768','1152x882','1152x900','1280x1024','1600x1200','1800x1440','1920x1200');
type TST_Comments =  (stcCommNone,stcCommIndicator,stcCommIndAndComment);
const StrTST_Comments: array[0..2] of AxUCString = ('commNone','commIndicator','commIndAndComment');
type TST_Objects =  (stoAll,stoPlaceholders,stoNone);
const StrTST_Objects: array[0..2] of AxUCString = ('all','placeholders','none');
type TST_SmartTagShow =  (ststsAll,ststsNone,ststsNoIndicator);
const StrTST_SmartTagShow: array[0..2] of AxUCString = ('all','none','noIndicator');
type TST_RefMode =  (strmA1,strmR1C1);
const StrTST_RefMode: array[0..1] of AxUCString = ('A1','R1C1');
type TST_SheetState =  (stssVisible,stssHidden,stssVeryHidden);
const StrTST_SheetState: array[0..2] of AxUCString = ('visible','hidden','veryHidden');
type TST_DvAspect =  (stdaDVASPECT_CONTENT,stdaDVASPECT_ICON);
const StrTST_DvAspect: array[0..1] of AxUCString = ('DVASPECT_CONTENT','DVASPECT_ICON');
type TST_CfvoType =  (stctNum,stctPercent,stctMax,stctMin,stctFormula,stctPercentile);
const StrTST_CfvoType: array[0..5] of AxUCString = ('num','percent','max','min','formula','percentile');
type TST_PageOrder =  (stpoDownThenOver,stpoOverThenDown);
const StrTST_PageOrder: array[0..1] of AxUCString = ('downThenOver','overThenDown');
type TST_CellType =  (stctB,stctN,stctD,stctE,stctS,stctStr,stctInlineStr);
const StrTST_CellType: array[0..6] of AxUCString = ('b','n','d','e','s','str','inlineStr');
type TST_PaneState =  (stpsSplit,stpsFrozen,stpsFrozenSplit);
const StrTST_PaneState: array[0..2] of AxUCString = ('split','frozen','frozenSplit');
type TST_CellFormulaType =  (stcftNormal,stcftArray,stcftDataTable,stcftShared);
const StrTST_CellFormulaType: array[0..3] of AxUCString = ('normal','array','dataTable','shared');
type TST_CellComments =  (stccNone,stccAsDisplayed,stccAtEnd);
const StrTST_CellComments: array[0..2] of AxUCString = ('none','asDisplayed','atEnd');
type TST_DataValidationOperator =  (stdvoBetween,stdvoNotBetween,stdvoEqual,stdvoNotEqual,stdvoLessThan,stdvoLessThanOrEqual,stdvoGreaterThan,stdvoGreaterThanOrEqual);
const StrTST_DataValidationOperator: array[0..7] of AxUCString = ('between','notBetween','equal','notEqual','lessThan','lessThanOrEqual','greaterThan','greaterThanOrEqual');
type TST_CfType =  (stctExpression,stctCellIs,stctColorScale,stctDataBar,stctIconSet,stctTop10,stctUniqueValues,stctDuplicateValues,stctContainsText,stctNotContainsText,stctBeginsWith,stctEndsWith,stctContainsBlanks,stctNotContainsBlanks,stctContainsErrors,stctNotContainsErrors,stctTimePeriod,stctAboveAverage);
const StrTST_CfType: array[0..17] of AxUCString = ('expression','cellIs','colorScale','dataBar','iconSet','top10','uniqueValues','duplicateValues','containsText','notContainsText','beginsWith','endsWith','containsBlanks','notContainsBlanks','containsErrors','notContainsErrors','timePeriod','aboveAverage');
type TST_PrintError =  (stpeDisplayed,stpeBlank,stpeDash,stpeNA);
const StrTST_PrintError: array[0..3] of AxUCString = ('displayed','blank','dash','NA');
type TST_OleUpdate =  (stouOLEUPDATE_ALWAYS,stouOLEUPDATE_ONCALL);
const StrTST_OleUpdate: array[0..1] of AxUCString = ('OLEUPDATE_ALWAYS','OLEUPDATE_ONCALL');
type TST_WebSourceType =  (stwstSheet,stwstPrintArea,stwstAutoFilter,stwstRange,stwstChart,stwstPivotTable,stwstQuery,stwstLabel);
const StrTST_WebSourceType: array[0..7] of AxUCString = ('sheet','printArea','autoFilter','range','chart','pivotTable','query','label');
type TST_DataValidationErrorStyle =  (stdvesStop,stdvesWarning,stdvesInformation);
const StrTST_DataValidationErrorStyle: array[0..2] of AxUCString = ('stop','warning','information');
type TST_TimePeriod =  (sttpToday,sttpYesterday,sttpTomorrow,sttpLast7Days,sttpThisMonth,sttpLastMonth,sttpNextMonth,sttpThisWeek,sttpLastWeek,sttpNextWeek);
const StrTST_TimePeriod: array[0..9] of AxUCString = ('today','yesterday','tomorrow','last7Days','thisMonth','lastMonth','nextMonth','thisWeek','lastWeek','nextWeek');
type TST_Pane =  (stpBottomRight,stpTopRight,stpBottomLeft,stpTopLeft);
const StrTST_Pane: array[0..3] of AxUCString = ('bottomRight','topRight','bottomLeft','topLeft');
type TST_ConditionalFormattingOperator =  (stcfoLessThan,stcfoLessThanOrEqual,stcfoEqual,stcfoNotEqual,stcfoGreaterThanOrEqual,stcfoGreaterThan,stcfoBetween,stcfoNotBetween,stcfoContainsText,stcfoNotContains,stcfoBeginsWith,stcfoEndsWith);
const StrTST_ConditionalFormattingOperator: array[0..11] of AxUCString = ('lessThan','lessThanOrEqual','equal','notEqual','greaterThanOrEqual','greaterThan','between','notBetween','containsText','notContains','beginsWith','endsWith');
type TST_DataValidationImeMode =  (stdvimNoControl,stdvimOff,stdvimOn,stdvimDisabled,stdvimHiragana,stdvimFullKatakana,stdvimHalfKatakana,stdvimFullAlpha,stdvimHalfAlpha,stdvimFullHangul,stdvimHalfHangul);
const StrTST_DataValidationImeMode: array[0..10] of AxUCString = ('noControl','off','on','disabled','hiragana','fullKatakana','halfKatakana','fullAlpha','halfAlpha','fullHangul','halfHangul');
type TST_SheetViewType =  (stsvtNormal,stsvtPageBreakPreview,stsvtPageLayout);
const StrTST_SheetViewType: array[0..2] of AxUCString = ('normal','pageBreakPreview','pageLayout');
type TST_DataConsolidateFunction =  (stdcfAverage,stdcfCount,stdcfCountNums,stdcfMax,stdcfMin,stdcfProduct,stdcfStdDev,stdcfStdDevp,stdcfSum,stdcfVar,stdcfVarp);
const StrTST_DataConsolidateFunction: array[0..10] of AxUCString = ('average','count','countNums','max','min','product','stdDev','stdDevp','sum','var','varp');
type TST_DataValidationType =  (stdvtNone,stdvtWhole,stdvtDecimal,stdvtList,stdvtDate,stdvtTime,stdvtTextLength,stdvtCustom);
const StrTST_DataValidationType: array[0..7] of AxUCString = ('none','whole','decimal','list','date','time','textLength','custom');
type TST_Orientation =  (stoDefault,stoPortrait,stoLandscape);
const StrTST_Orientation: array[0..2] of AxUCString = ('default','portrait','landscape');

type TST_TableType =  (stttWorksheet,stttXml,stttQueryTable);
const StrTST_TableType: array[0..2] of AxUCString = ('worksheet','xml','queryTable');
type TST_TotalsRowFunction =  (sttrfNone,sttrfSum,sttrfMin,sttrfMax,sttrfAverage,sttrfCount,sttrfCountNums,sttrfStdDev,sttrfVar,sttrfCustom);
const StrTST_TotalsRowFunction: array[0..9] of AxUCString = ('none','sum','min','max','average','count','countNums','stdDev','var','custom');
type TST_XmlDataType =  (stxdtString,stxdtNormalizedString,stxdtToken,stxdtByte,stxdtUnsignedByte,stxdtBase64Binary,stxdtHexBinary,stxdtInteger,stxdtPositiveInteger,stxdtNegativeInteger,stxdtNonPositiveInteger,stxdtNonNegativeInteger,stxdtInt,stxdtUnsignedInt,stxdtLong,stxdtUnsignedLong,stxdtShort,stxdtUnsignedShort,stxdtDecimal,stxdtFloat,stxdtDouble,stxdtBoolean,stxdtTime,stxdtDateTime,stxdtDuration,stxdtDate,stxdtGMonth,stxdtGYear,stxdtGYearMonth,stxdtGDay,stxdtGMonthDay,stxdtName,stxdtQName,stxdtNCName,stxdtAnyURI,stxdtLanguage,stxdtID,stxdtIDREF,stxdtIDREFS,stxdtENTITY,stxdtENTITIES,stxdtNOTATION,stxdtNMTOKEN,stxdtNMTOKENS,stxdtAnyType);
const StrTST_XmlDataType: array[0..44] of AxUCString = ('string','normalizedString','token','byte','unsignedByte','base64Binary','hexBinary','integer','positiveInteger','negativeInteger','nonPositiveInteger','nonNegativeInteger','int','unsignedInt','long','unsignedLong','short','unsignedShort','decimal','float','double','boolean','time','dateTime','duration','date','gMonth','gYear','gYearMonth','gDay','gMonthDay','Name','QName','NCName','anyURI','language','ID','IDREF','IDREFS','ENTITY','ENTITIES','NOTATION','NMTOKEN','NMTOKENS','anyType');

type TXPGDocBase = class(TObject)
protected
     FErrors: TXpgPErrors;
     FOnReadWorksheetColsCol: TNotifyEvent;
     FOnWriteWorksheetColsCol: TWriteElementEvent;
public
     property Errors: TXpgPErrors read FErrors;
     property OnReadWorksheetColsCol: TNotifyEvent read FOnReadWorksheetColsCol write FOnReadWorksheetColsCol;
     property OnWriteWorksheetColsCol: TWriteElementEvent read FOnWriteWorksheetColsCol write FOnWriteWorksheetColsCol;
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

     class procedure AddEnums;
     class function  StrToEnum(const AValue: AxUCString): integer;
     class function  StrToEnumDef(const AValue: AxUCString; ADefault: integer): integer;
     class function  TryStrToEnum(const AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
public
     function  Available: boolean;

     property ElementCount: integer read FElementCount write FElementCount;
     property AttributeCount: integer read FAttributeCount write FAttributeCount;
     property Assigneds: TXpgAssigneds read FAssigneds write FAssigneds;
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

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Uri: AxUCString read FUri write FUri;
     end;

     TCT_ExtensionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Extension;
public
     function  Add: TCT_Extension;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Extension read GetItems; default;
     end;

     TCT_Index = class(TXPGBase)
protected
     FV: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property V: integer read FV write FV;
     end;

     TCT_IndexXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Index;
public
     function  Add: TCT_Index;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Index read GetItems; default;
     end;

     TCT_ExtensionList = class(TXPGBase)
protected
     FExtXpgList: TCT_ExtensionXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ExtXpgList: TCT_ExtensionXpgList read FExtXpgList;
     end;

     TCT_FontName = class(TXPGBase)
protected
     FVal: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: AxUCString read FVal write FVal;
     end;

     TCT_FontNameXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FontName;
public
     function  Add: TCT_FontName;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_FontName read GetItems; default;
     end;

     TCT_IntProperty = class(TXPGBase)
protected
     FVal: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: integer read FVal write FVal;
     end;

     TCT_IntPropertyXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_IntProperty;
public
     function  Add: TCT_IntProperty;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_IntProperty read GetItems; default;
     end;

     TCT_BooleanProperty = class(TXPGBase)
protected
     FVal: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: boolean read FVal write FVal;
     end;

     TCT_Color = class(TXPGBase)
private
     procedure SetRgb(const Value: integer);
protected
     FAuto: boolean;
     FIndexed: integer;
     FRgb: integer;
     // FRgb can have the value -1 (FFFFFFFF).
     FRgbAssigned: boolean;
     FTheme: integer;
     FTint: double;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Auto: boolean read FAuto write FAuto;
     property Indexed: integer read FIndexed write FIndexed;
     property Rgb: integer read FRgb write SetRgb;
     property Theme: integer read FTheme write FTheme;
     property Tint: double read FTint write FTint;
     end;

     TCT_ColorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Color;
public
     function  Add: TCT_Color;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Color read GetItems; default;
     end;

     TCT_FontSize = class(TXPGBase)
protected
     FVal: double;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: double read FVal write FVal;
     end;

     TCT_FontSizeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FontSize;
public
     function  Add: TCT_FontSize;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_FontSize read GetItems; default;
     end;

     TCT_UnderlineProperty = class(TXPGBase)
protected
     FVal: TST_UnderlineValues;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     procedure Clear;

     property Val: TST_UnderlineValues read FVal write FVal;
     end;

     TCT_VerticalAlignFontProperty = class(TXPGBase)
protected
     FVal: TST_VerticalAlignRun;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     procedure Clear;

     property Val: TST_VerticalAlignRun read FVal write FVal;
     end;

     TCT_FontScheme = class(TXPGBase)
protected
     FVal: TST_FontScheme;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: TST_FontScheme read FVal write FVal;
     end;

     TCT_FontSchemeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FontScheme;
public
     function  Add: TCT_FontScheme;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_FontScheme read GetItems; default;
     end;

     TCT_PivotAreaReference = class(TXPGBase)
protected
     FField: integer;
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

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Field: integer read FField write FField;
     property Count_: integer read FCount write FCount;
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
     property Items[Index: integer]: TCT_PivotAreaReference read GetItems; default;
     end;

     TCT_Filter = class(TXPGBase)
protected
     FVal: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Val: AxUCString read FVal write FVal;
     end;

     TCT_FilterXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Filter;
public
     function  Add: TCT_Filter;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Filter read GetItems; default;
     end;

     TCT_DateGroupItem = class(TXPGBase)
protected
     FYear: integer;
     FMonth: integer;
     FDay: integer;
     FHour: integer;
     FMinute: integer;
     FSecond: integer;
     FDateTimeGrouping: TST_DateTimeGrouping;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Year: integer read FYear write FYear;
     property Month: integer read FMonth write FMonth;
     property Day: integer read FDay write FDay;
     property Hour: integer read FHour write FHour;
     property Minute: integer read FMinute write FMinute;
     property Second: integer read FSecond write FSecond;
     property DateTimeGrouping: TST_DateTimeGrouping read FDateTimeGrouping write FDateTimeGrouping;
     end;

     TCT_DateGroupItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DateGroupItem;
public
     function  Add: TCT_DateGroupItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_DateGroupItem read GetItems; default;
     end;

     TCT_CustomFilter = class(TXPGBase)
protected
     FOperator: TST_FilterOperator;
     FVal: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Operator_: TST_FilterOperator read FOperator write FOperator;
     property Val: AxUCString read FVal write FVal;
     end;

     TCT_CustomFilterXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CustomFilter;
public
     function  Add: TCT_CustomFilter;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CustomFilter read GetItems; default;
     end;

     TCT_RPrElt = class(TXPGBase)
protected
     FRFont: TCT_FontName;
     FCharset: TCT_IntProperty;
     FFamily: TCT_IntProperty;
     FB: TCT_BooleanProperty;
     FI: TCT_BooleanProperty;
     FStrike: TCT_BooleanProperty;
     FOutline: TCT_BooleanProperty;
     FShadow: TCT_BooleanProperty;
     FCondense: TCT_BooleanProperty;
     FExtend: TCT_BooleanProperty;
     FColor: TCT_Color;
     FSz: TCT_FontSize;
     FU: TCT_UnderlineProperty;
     FVertAlign: TCT_VerticalAlignFontProperty;
     FScheme: TCT_FontScheme;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RFont: TCT_FontName read FRFont;
     property Charset: TCT_IntProperty read FCharset;
     property Family: TCT_IntProperty read FFamily;
     property B: TCT_BooleanProperty read FB;
     property I: TCT_BooleanProperty read FI;
     property Strike: TCT_BooleanProperty read FStrike;
     property Outline: TCT_BooleanProperty read FOutline;
     property Shadow: TCT_BooleanProperty read FShadow;
     property Condense: TCT_BooleanProperty read FCondense;
     property Extend: TCT_BooleanProperty read FExtend;
     property Color: TCT_Color read FColor;
     property Sz: TCT_FontSize read FSz;
     property U: TCT_UnderlineProperty read FU;
     property VertAlign: TCT_VerticalAlignFontProperty read FVertAlign;
     property Scheme: TCT_FontScheme read FScheme;
     end;

     TCT_GradientStop = class(TXPGBase)
protected
     FPosition: double;
     FColor: TCT_Color;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Position: double read FPosition write FPosition;
     property Color: TCT_Color read FColor;
     end;

     TCT_GradientStopXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_GradientStop;
public
     function  Add: TCT_GradientStop;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_GradientStop read GetItems; default;
     end;

     TCT_PivotAreaReferences = class(TXPGBase)
protected
     FCount: integer;
     FReferenceXpgList: TCT_PivotAreaReferenceXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property ReferenceXpgList: TCT_PivotAreaReferenceXpgList read FReferenceXpgList;
     end;

     TCT_Filters = class(TXPGBase)
protected
     FBlank: boolean;
     FCalendarType: TST_CalendarType;
     FFilterXpgList: TCT_FilterXpgList;
     FDateGroupItemXpgList: TCT_DateGroupItemXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Blank: boolean read FBlank write FBlank;
     property CalendarType: TST_CalendarType read FCalendarType write FCalendarType;
     property FilterXpgList: TCT_FilterXpgList read FFilterXpgList;
     property DateGroupItemXpgList: TCT_DateGroupItemXpgList read FDateGroupItemXpgList;
     end;

     TCT_Top10 = class(TXPGBase)
protected
     FTop: boolean;
     FPercent: boolean;
     FVal: double;
     FFilterVal: double;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Top: boolean read FTop write FTop;
     property Percent: boolean read FPercent write FPercent;
     property Val: double read FVal write FVal;
     property FilterVal: double read FFilterVal write FFilterVal;
     end;

     TCT_CustomFilters = class(TXPGBase)
protected
     FAnd: boolean;
     FCustomFilterXpgList: TCT_CustomFilterXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property And_: boolean read FAnd write FAnd;
     property CustomFilterXpgList: TCT_CustomFilterXpgList read FCustomFilterXpgList;
     end;

     TCT_DynamicFilter = class(TXPGBase)
protected
     FType: TST_DynamicFilterType;
     FVal: double;
     FMaxVal: double;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: TST_DynamicFilterType read FType write FType;
     property Val: double read FVal write FVal;
     property MaxVal: double read FMaxVal write FMaxVal;
     end;

     TCT_ColorFilter = class(TXPGBase)
protected
     FDxfId: integer;
     FCellColor: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DxfId: integer read FDxfId write FDxfId;
     property CellColor: boolean read FCellColor write FCellColor;
     end;

     TCT_IconFilter = class(TXPGBase)
protected
     FIconSet: TST_IconSetType;
     FIconId: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property IconSet: TST_IconSetType read FIconSet write FIconSet;
     property IconId: integer read FIconId write FIconId;
     end;

     TCT_SortCondition = class(TXPGBase)
protected
     FDescending: boolean;
     FSortBy: TST_SortBy;
     FRef: AxUCString;
     FCustomList: AxUCString;
     FDxfId: integer;
     FIconSet: TST_IconSetType;
     FIconId: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Descending: boolean read FDescending write FDescending;
     property SortBy: TST_SortBy read FSortBy write FSortBy;
     property Ref: AxUCString read FRef write FRef;
     property CustomList: AxUCString read FCustomList write FCustomList;
     property DxfId: integer read FDxfId write FDxfId;
     property IconSet: TST_IconSetType read FIconSet write FIconSet;
     property IconId: integer read FIconId write FIconId;
     end;

     TCT_SortConditionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SortCondition;
public
     function  Add: TCT_SortCondition;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_SortCondition read GetItems; default;
     end;

     TCT_RElt = class(TXPGBase)
protected
     FRPr: TCT_RPrElt;
     FT: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RPr: TCT_RPrElt read FRPr;
     property T: AxUCString read FT write FT;
     end;

     TCT_REltXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RElt;
public
     function  Add: TCT_RElt;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_RElt read GetItems; default;
     end;

     TCT_PhoneticRun = class(TXPGBase)
protected
     FSb: integer;
     FEb: integer;
     FT: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Sb: integer read FSb write FSb;
     property Eb: integer read FEb write FEb;
     property T: AxUCString read FT write FT;
     end;

     TCT_PhoneticRunXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PhoneticRun;
public
     function  Add: TCT_PhoneticRun;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_PhoneticRun read GetItems; default;
     end;

     TCT_PhoneticPr = class(TXPGBase)
protected
     FFontId: integer;
     FType: TST_PhoneticType;
     FAlignment: TST_PhoneticAlignment;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property FontId: integer read FFontId write FFontId;
     property Type_: TST_PhoneticType read FType write FType;
     property Alignment: TST_PhoneticAlignment read FAlignment write FAlignment;
     end;

     TCT_PatternFill = class(TXPGBase)
protected
     FPatternType: TST_PatternType;
     FFgColor: TCT_Color;
     FBgColor: TCT_Color;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property PatternType: TST_PatternType read FPatternType write FPatternType;
     property FgColor: TCT_Color read FFgColor;
     property BgColor: TCT_Color read FBgColor;
     end;

     TCT_GradientFill = class(TXPGBase)
protected
     FType: TST_GradientType;
     FDegree: double;
     FLeft: double;
     FRight: double;
     FTop: double;
     FBottom: double;
     FStopXpgList: TCT_GradientStopXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: TST_GradientType read FType write FType;
     property Degree: double read FDegree write FDegree;
     property Left: double read FLeft write FLeft;
     property Right: double read FRight write FRight;
     property Top: double read FTop write FTop;
     property Bottom: double read FBottom write FBottom;
     property StopXpgList: TCT_GradientStopXpgList read FStopXpgList;
     end;

     TCT_BorderPr = class(TXPGBase)
protected
     FStyle: TST_BorderStyle;
     FColor: TCT_Color;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Style: TST_BorderStyle read FStyle write FStyle;
     property Color: TCT_Color read FColor;
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

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

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

     TCT_Break = class(TXPGBase)
protected
     FId: integer;
     FMin: integer;
     FMax: integer;
     FMan: boolean;
     FPt: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Id: integer read FId write FId;
     property Min: integer read FMin write FMin;
     property Max: integer read FMax write FMax;
     property Man: boolean read FMan write FMan;
     property Pt: boolean read FPt write FPt;
     end;

     TCT_BreakXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Break;
public
     function  Add: TCT_Break;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Break read GetItems; default;
     end;

     TCT_FilterColumn = class(TXPGBase)
protected
     FColId: integer;
     FHiddenButton: boolean;
     FShowButton: boolean;
     FFilters: TCT_Filters;
     FTop10: TCT_Top10;
     FCustomFilters: TCT_CustomFilters;
     FDynamicFilter: TCT_DynamicFilter;
     FColorFilter: TCT_ColorFilter;
     FIconFilter: TCT_IconFilter;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ColId: integer read FColId write FColId;
     property HiddenButton: boolean read FHiddenButton write FHiddenButton;
     property ShowButton: boolean read FShowButton write FShowButton;
     property Filters: TCT_Filters read FFilters;
     property Top10: TCT_Top10 read FTop10;
     property CustomFilters: TCT_CustomFilters read FCustomFilters;
     property DynamicFilter: TCT_DynamicFilter read FDynamicFilter;
     property ColorFilter: TCT_ColorFilter read FColorFilter;
     property IconFilter: TCT_IconFilter read FIconFilter;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_FilterColumnXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FilterColumn;
public
     function  Add: TCT_FilterColumn;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_FilterColumn read GetItems; default;
     end;

     TCT_SortState = class(TXPGBase)
protected
     FColumnSort: boolean;
     FCaseSensitive: boolean;
     FSortMethod: TST_SortMethod;
     FRef: AxUCString;
     FSortConditionXpgList: TCT_SortConditionXpgList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ColumnSort: boolean read FColumnSort write FColumnSort;
     property CaseSensitive: boolean read FCaseSensitive write FCaseSensitive;
     property SortMethod: TST_SortMethod read FSortMethod write FSortMethod;
     property Ref: AxUCString read FRef write FRef;
     property SortConditionXpgList: TCT_SortConditionXpgList read FSortConditionXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CellFormula = class(TXPGBase)
protected
     FT: TST_CellFormulaType;
     FAca: boolean;
     FRef: AxUCString;
     FDt2D: boolean;
     FDtr: boolean;
     FDel1: boolean;
     FDel2: boolean;
     FR1: AxUCString;
     FR2: AxUCString;
     FCa: boolean;
     FSi: integer;
     FBx: boolean;
     FContent: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property T: TST_CellFormulaType read FT write FT;
     property Aca: boolean read FAca write FAca;
     property Ref: AxUCString read FRef write FRef;
     property Dt2D: boolean read FDt2D write FDt2D;
     property Dtr: boolean read FDtr write FDtr;
     property Del1: boolean read FDel1 write FDel1;
     property Del2: boolean read FDel2 write FDel2;
     property R1: AxUCString read FR1 write FR1;
     property R2: AxUCString read FR2 write FR2;
     property Ca: boolean read FCa write FCa;
     property Si: integer read FSi write FSi;
     property Bx: boolean read FBx write FBx;
     property Content: AxUCString read FContent write FContent;
     end;

     TCT_Rst = class(TXPGBase)
protected
     _FT: AxUCString;
     FRXpgList: TCT_REltXpgList;
     FRPhXpgList: TCT_PhoneticRunXpgList;
     FPhoneticPr: TCT_PhoneticPr;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property T: AxUCString read _FT write _FT;
     property RXpgList: TCT_REltXpgList read FRXpgList;
     property RPhXpgList: TCT_PhoneticRunXpgList read FRPhXpgList;
     property PhoneticPr: TCT_PhoneticPr read FPhoneticPr;
     end;

     TCT_RstXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Rst;
public
     function  Add: TCT_Rst;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Rst read GetItems; default;
     end;

     TCT_Cfvo = class(TXPGBase)
protected
     FType: TST_CfvoType;
     FVal: AxUCString;
     FGte: boolean;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: TST_CfvoType read FType write FType;
     property Val: AxUCString read FVal write FVal;
     property Gte: boolean read FGte write FGte;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CfvoXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Cfvo;
public
     function  Add: TCT_Cfvo;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Cfvo read GetItems; default;
     end;

     TCT_CellSmartTagPr = class(TXPGBase)
protected
     FKey: AxUCString;
     FVal: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Key: AxUCString read FKey write FKey;
     property Val: AxUCString read FVal write FVal;
     end;

     TCT_CellSmartTagPrXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CellSmartTagPr;
public
     function  Add: TCT_CellSmartTagPr;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CellSmartTagPr read GetItems; default;
     end;

     TCT_CellAlignment = class(TXPGBase)
protected
     FHorizontal: TST_HorizontalAlignment;
     FVertical: TST_VerticalAlignment;
     FTextRotation: integer;
     FWrapText: boolean;
     FIndent: integer;
     FRelativeIndent: integer;
     FJustifyLastLine: boolean;
     FShrinkToFit: boolean;
     FReadingOrder: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Horizontal: TST_HorizontalAlignment read FHorizontal write FHorizontal;
     property Vertical: TST_VerticalAlignment read FVertical write FVertical;
     property TextRotation: integer read FTextRotation write FTextRotation;
     property WrapText: boolean read FWrapText write FWrapText;
     property Indent: integer read FIndent write FIndent;
     property RelativeIndent: integer read FRelativeIndent write FRelativeIndent;
     property JustifyLastLine: boolean read FJustifyLastLine write FJustifyLastLine;
     property ShrinkToFit: boolean read FShrinkToFit write FShrinkToFit;
     property ReadingOrder: integer read FReadingOrder write FReadingOrder;
     end;

     TCT_CellProtection = class(TXPGBase)
protected
     FLocked: boolean;
     FHidden: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Locked: boolean read FLocked write FLocked;
     property Hidden: boolean read FHidden write FHidden;
     end;

     TCT_Font = class(TXPGBase)
protected
     FName: TCT_FontName;
     FCharset: TCT_IntProperty;
     FFamily: TCT_IntProperty;
     FB: TCT_BooleanProperty;
     FI: TCT_BooleanProperty;
     FStrike: TCT_BooleanProperty;
     FOutline: TCT_BooleanProperty;
     FShadow: TCT_BooleanProperty;
     FCondense: TCT_BooleanProperty;
     FExtend: TCT_BooleanProperty;
     FColor: TCT_Color;
     FSz: TCT_FontSize;
     FU: TCT_UnderlineProperty;
     FVertAlign: TCT_VerticalAlignFontProperty;
     FScheme: TCT_FontScheme;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: TCT_FontName read FName;
     property Charset: TCT_IntProperty read FCharset;
     property Family: TCT_IntProperty read FFamily;
     property B: TCT_BooleanProperty read FB;
     property I: TCT_BooleanProperty read FI;
     property Strike: TCT_BooleanProperty read FStrike;
     property Outline: TCT_BooleanProperty read FOutline;
     property Shadow: TCT_BooleanProperty read FShadow;
     property Condense: TCT_BooleanProperty read FCondense;
     property Extend: TCT_BooleanProperty read FExtend;
     property Color: TCT_Color read FColor;
     property Sz: TCT_FontSize read FSz;
     property U: TCT_UnderlineProperty read FU;
     property VertAlign: TCT_VerticalAlignFontProperty read FVertAlign;
     property Scheme: TCT_FontScheme read FScheme;
     end;

     TCT_FontXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Font;
public
     function  Add: TCT_Font;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Font read GetItems; default;
     end;

     TCT_NumFmt = class(TXPGBase)
protected
     FNumFmtId: integer;
     FFormatCode: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property NumFmtId: integer read FNumFmtId write FNumFmtId;
     property FormatCode: AxUCString read FFormatCode write FFormatCode;
     end;

     TCT_NumFmtXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_NumFmt;
public
     function  Add: TCT_NumFmt;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_NumFmt read GetItems; default;
     end;

     TCT_Fill = class(TXPGBase)
protected
     FPatternFill: TCT_PatternFill;
     FGradientFill: TCT_GradientFill;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property PatternFill: TCT_PatternFill read FPatternFill;
     property GradientFill: TCT_GradientFill read FGradientFill;
     end;

     TCT_FillXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Fill;
public
     function  Add: TCT_Fill;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Fill read GetItems; default;
     end;

     TCT_Border = class(TXPGBase)
protected
     FDiagonalUp: boolean;
     FDiagonalDown: boolean;
     FOutline: boolean;
     FLeft: TCT_BorderPr;
     FRight: TCT_BorderPr;
     FTop: TCT_BorderPr;
     FBottom: TCT_BorderPr;
     FDiagonal: TCT_BorderPr;
     FVertical: TCT_BorderPr;
     FHorizontal: TCT_BorderPr;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DiagonalUp: boolean read FDiagonalUp write FDiagonalUp;
     property DiagonalDown: boolean read FDiagonalDown write FDiagonalDown;
     property Outline: boolean read FOutline write FOutline;
     property Left: TCT_BorderPr read FLeft;
     property Right: TCT_BorderPr read FRight;
     property Top: TCT_BorderPr read FTop;
     property Bottom: TCT_BorderPr read FBottom;
     property Diagonal: TCT_BorderPr read FDiagonal;
     property Vertical: TCT_BorderPr read FVertical;
     property Horizontal: TCT_BorderPr read FHorizontal;
     end;

     TCT_BorderXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Border;
public
     function  Add: TCT_Border;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Border read GetItems; default;
     end;

     TCT_TableStyleElement = class(TXPGBase)
protected
     FType: TST_TableStyleType;
     FSize: integer;
     FDxfId: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: TST_TableStyleType read FType write FType;
     property Size: integer read FSize write FSize;
     property DxfId: integer read FDxfId write FDxfId;
     end;

     TCT_TableStyleElementXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TableStyleElement;
public
     function  Add: TCT_TableStyleElement;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_TableStyleElement read GetItems; default;
     end;

     TCT_RgbColor = class(TXPGBase)
protected
     FRgb: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Rgb: integer read FRgb write FRgb;
     end;

     TCT_RgbColorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_RgbColor;
public
     function  Add: TCT_RgbColor;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_RgbColor read GetItems; default;
     end;

     TCT_PageMargins = class(TXPGBase)
protected
     FLeft: double;
     FRight: double;
     FTop: double;
     FBottom: double;
     FHeader: double;
     FFooter: double;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Left: double read FLeft write FLeft;
     property Right: double read FRight write FRight;
     property Top: double read FTop write FTop;
     property Bottom: double read FBottom write FBottom;
     property Header: double read FHeader write FHeader;
     property Footer: double read FFooter write FFooter;
     end;

     TCT_CsPageSetup = class(TXPGBase)
protected
     FPaperSize: integer;
     FFirstPageNumber: integer;
     FOrientation: TST_Orientation;
     FUsePrinterDefaults: boolean;
     FBlackAndWhite: boolean;
     FDraft: boolean;
     FUseFirstPageNumber: boolean;
     FHorizontalDpi: integer;
     FVerticalDpi: integer;
     FCopies: integer;
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property PaperSize: integer read FPaperSize write FPaperSize;
     property FirstPageNumber: integer read FFirstPageNumber write FFirstPageNumber;
     property Orientation: TST_Orientation read FOrientation write FOrientation;
     property UsePrinterDefaults: boolean read FUsePrinterDefaults write FUsePrinterDefaults;
     property BlackAndWhite: boolean read FBlackAndWhite write FBlackAndWhite;
     property Draft: boolean read FDraft write FDraft;
     property UseFirstPageNumber: boolean read FUseFirstPageNumber write FUseFirstPageNumber;
     property HorizontalDpi: integer read FHorizontalDpi write FHorizontalDpi;
     property VerticalDpi: integer read FVerticalDpi write FVerticalDpi;
     property Copies: integer read FCopies write FCopies;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_HeaderFooter = class(TXPGBase)
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

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
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

     TCT_Pane = class(TXPGBase)
protected
     FXSplit: double;
     FYSplit: double;
     FTopLeftCell: AxUCString;
     FActivePane: TST_Pane;
     FState: TST_PaneState;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property XSplit: double read FXSplit write FXSplit;
     property YSplit: double read FYSplit write FYSplit;
     property TopLeftCell: AxUCString read FTopLeftCell write FTopLeftCell;
     property ActivePane: TST_Pane read FActivePane write FActivePane;
     property State: TST_PaneState read FState write FState;
     end;

     TCT_Selection = class(TXPGBase)
protected
     FPane: TST_Pane;
     FActiveCell: AxUCString;
     FActiveCellId: integer;
     FSqrefXpgList: TStringXpgList;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Pane: TST_Pane read FPane write FPane;
     property ActiveCell: AxUCString read FActiveCell write FActiveCell;
     property ActiveCellId: integer read FActiveCellId write FActiveCellId;
     property SqrefXpgList: TStringXpgList read FSqrefXpgList;
     end;

     TCT_SelectionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Selection;
public
     function  Add: TCT_Selection;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Selection read GetItems; default;
     end;

     TCT_PivotSelection = class(TXPGBase)
protected
     FPane: TST_Pane;
     FShowHeader: boolean;
     FLabel: boolean;
     FData: boolean;
     FExtendable: boolean;
     FCount: integer;
     FAxis: TST_Axis;
     FDimension: integer;
     FStart: integer;
     FMin: integer;
     FMax: integer;
     FActiveRow: integer;
     FActiveCol: integer;
     FPreviousRow: integer;
     FPreviousCol: integer;
     FClick: integer;
     FR_Id: AxUCString;
     FPivotArea: TCT_PivotArea;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Pane: TST_Pane read FPane write FPane;
     property ShowHeader: boolean read FShowHeader write FShowHeader;
     property Label_: boolean read FLabel write FLabel;
     property Data: boolean read FData write FData;
     property Extendable: boolean read FExtendable write FExtendable;
     property Count_: integer read FCount write FCount;
     property Axis: TST_Axis read FAxis write FAxis;
     property Dimension: integer read FDimension write FDimension;
     property Start: integer read FStart write FStart;
     property Min: integer read FMin write FMin;
     property Max: integer read FMax write FMax;
     property ActiveRow: integer read FActiveRow write FActiveRow;
     property ActiveCol: integer read FActiveCol write FActiveCol;
     property PreviousRow: integer read FPreviousRow write FPreviousRow;
     property PreviousCol: integer read FPreviousCol write FPreviousCol;
     property Click: integer read FClick write FClick;
     property R_Id: AxUCString read FR_Id write FR_Id;
     property PivotArea: TCT_PivotArea read FPivotArea;
     end;

     TCT_PivotSelectionXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotSelection;
public
     function  Add: TCT_PivotSelection;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_PivotSelection read GetItems; default;
     end;

     TCT_PageBreak = class(TXPGBase)
protected
     FManualBreakCount: integer;
     FBrkXpgList: TCT_BreakXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ManualBreakCount: integer read FManualBreakCount write FManualBreakCount;
     property BrkXpgList: TCT_BreakXpgList read FBrkXpgList;
     end;

     TCT_PrintOptions = class(TXPGBase)
protected
     FHorizontalCentered: boolean;
     FVerticalCentered: boolean;
     FHeadings: boolean;
     FGridLines: boolean;
     FGridLinesSet: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property HorizontalCentered: boolean read FHorizontalCentered write FHorizontalCentered;
     property VerticalCentered: boolean read FVerticalCentered write FVerticalCentered;
     property Headings: boolean read FHeadings write FHeadings;
     property GridLines: boolean read FGridLines write FGridLines;
     property GridLinesSet: boolean read FGridLinesSet write FGridLinesSet;
     end;

     TCT_PageSetup = class(TXPGBase)
protected
     FPaperSize: integer;
     FScale: integer;
     FFirstPageNumber: integer;
     FFitToWidth: integer;
     FFitToHeight: integer;
     FPageOrder: TST_PageOrder;
     FOrientation: TST_Orientation;
     FUsePrinterDefaults: boolean;
     FBlackAndWhite: boolean;
     FDraft: boolean;
     FCellComments: TST_CellComments;
     FUseFirstPageNumber: boolean;
     FErrors: TST_PrintError;
     FHorizontalDpi: integer;
     FVerticalDpi: integer;
     FCopies: integer;
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property PaperSize: integer read FPaperSize write FPaperSize;
     property Scale: integer read FScale write FScale;
     property FirstPageNumber: integer read FFirstPageNumber write FFirstPageNumber;
     property FitToWidth: integer read FFitToWidth write FFitToWidth;
     property FitToHeight: integer read FFitToHeight write FFitToHeight;
     property PageOrder: TST_PageOrder read FPageOrder write FPageOrder;
     property Orientation: TST_Orientation read FOrientation write FOrientation;
     property UsePrinterDefaults: boolean read FUsePrinterDefaults write FUsePrinterDefaults;
     property BlackAndWhite: boolean read FBlackAndWhite write FBlackAndWhite;
     property Draft: boolean read FDraft write FDraft;
     property CellComments: TST_CellComments read FCellComments write FCellComments;
     property UseFirstPageNumber: boolean read FUseFirstPageNumber write FUseFirstPageNumber;
     property Errors: TST_PrintError read FErrors write FErrors;
     property HorizontalDpi: integer read FHorizontalDpi write FHorizontalDpi;
     property VerticalDpi: integer read FVerticalDpi write FVerticalDpi;
     property Copies: integer read FCopies write FCopies;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_AutoFilter = class(TXPGBase)
protected
     FRef: AxUCString;
     FFilterColumnXpgList: TCT_FilterColumnXpgList;
     FSortState: TCT_SortState;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     property FilterColumnXpgList: TCT_FilterColumnXpgList read FFilterColumnXpgList;
     property SortState: TCT_SortState read FSortState;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Cell = class(TXPGBase)
protected
     FR: AxUCString;
     FS: integer;
     FT: TST_CellType;
     FCm: integer;
     FVm: integer;
     FPh: boolean;
     FF: TCT_CellFormula;
     FV: AxUCString;
     FIs: TCT_Rst;
//     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R: AxUCString read FR write FR;
     property S: integer read FS write FS;
     property T: TST_CellType read FT write FT;
     property Cm: integer read FCm write FCm;
     property Vm: integer read FVm write FVm;
     property Ph: boolean read FPh write FPh;
     property F: TCT_CellFormula read FF;
     property V: AxUCString read FV write FV;
     property Is_: TCT_Rst read FIs;
//     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CellXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Cell;
public
     function  Add: TCT_Cell;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Cell read GetItems; default;
     end;

     TCT_InputCells = class(TXPGBase)
protected
     FR: AxUCString;
     FDeleted: boolean;
     FUndone: boolean;
     FVal: AxUCString;
     FNumFmtId: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R: AxUCString read FR write FR;
     property Deleted: boolean read FDeleted write FDeleted;
     property Undone: boolean read FUndone write FUndone;
     property Val: AxUCString read FVal write FVal;
     property NumFmtId: integer read FNumFmtId write FNumFmtId;
     end;

     TCT_InputCellsXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_InputCells;
public
     function  Add: TCT_InputCells;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_InputCells read GetItems; default;
     end;

     TCT_DataRef = class(TXPGBase)
protected
     FRef: AxUCString;
     FName: AxUCString;
     FSheet: AxUCString;
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     property Name: AxUCString read FName write FName;
     property Sheet: AxUCString read FSheet write FSheet;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_DataRefXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DataRef;
public
     function  Add: TCT_DataRef;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_DataRef read GetItems; default;
     end;

     TCT_ColorScale = class(TXPGBase)
protected
     FCfvoXpgList: TCT_CfvoXpgList;
     FColorXpgList: TCT_ColorXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CfvoXpgList: TCT_CfvoXpgList read FCfvoXpgList;
     property ColorXpgList: TCT_ColorXpgList read FColorXpgList;
     end;

     TCT_DataBar = class(TXPGBase)
protected
     FMinLength: integer;
     FMaxLength: integer;
     FShowValue: boolean;
     FCfvoXpgList: TCT_CfvoXpgList;
     FColor: TCT_Color;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property MinLength: integer read FMinLength write FMinLength;
     property MaxLength: integer read FMaxLength write FMaxLength;
     property ShowValue: boolean read FShowValue write FShowValue;
     property CfvoXpgList: TCT_CfvoXpgList read FCfvoXpgList;
     property Color: TCT_Color read FColor;
     end;

     TCT_IconSet = class(TXPGBase)
protected
     FIconSet: TST_IconSetType;
     FShowValue: boolean;
     FPercent: boolean;
     FReverse: boolean;
     FCfvoXpgList: TCT_CfvoXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property IconSet: TST_IconSetType read FIconSet write FIconSet;
     property ShowValue: boolean read FShowValue write FShowValue;
     property Percent: boolean read FPercent write FPercent;
     property Reverse: boolean read FReverse write FReverse;
     property CfvoXpgList: TCT_CfvoXpgList read FCfvoXpgList;
     end;

     TCT_CellSmartTag = class(TXPGBase)
protected
     FType: integer;
     FDeleted: boolean;
     FXmlBased: boolean;
     FCellSmartTagPrXpgList: TCT_CellSmartTagPrXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: integer read FType write FType;
     property Deleted: boolean read FDeleted write FDeleted;
     property XmlBased: boolean read FXmlBased write FXmlBased;
     property CellSmartTagPrXpgList: TCT_CellSmartTagPrXpgList read FCellSmartTagPrXpgList;
     end;

     TCT_CellSmartTagXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CellSmartTag;
public
     function  Add: TCT_CellSmartTag;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CellSmartTag read GetItems; default;
     end;

     TCT_Xf = class(TXPGBase)
protected
     FNumFmtId: integer;
     FFontId: integer;
     FFillId: integer;
     FBorderId: integer;
     FXfId: integer;
     FQuotePrefix: boolean;
     FPivotButton: boolean;
     FApplyNumberFormat: boolean;
     FApplyFont: boolean;
     FApplyFill: boolean;
     FApplyBorder: boolean;
     FApplyAlignment: boolean;
     FApplyProtection: boolean;
     FAlignment: TCT_CellAlignment;
     FProtection: TCT_CellProtection;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property NumFmtId: integer read FNumFmtId write FNumFmtId;
     property FontId: integer read FFontId write FFontId;
     property FillId: integer read FFillId write FFillId;
     property BorderId: integer read FBorderId write FBorderId;
     property XfId: integer read FXfId write FXfId;
     property QuotePrefix: boolean read FQuotePrefix write FQuotePrefix;
     property PivotButton: boolean read FPivotButton write FPivotButton;
     property ApplyNumberFormat: boolean read FApplyNumberFormat write FApplyNumberFormat;
     property ApplyFont: boolean read FApplyFont write FApplyFont;
     property ApplyFill: boolean read FApplyFill write FApplyFill;
     property ApplyBorder: boolean read FApplyBorder write FApplyBorder;
     property ApplyAlignment: boolean read FApplyAlignment write FApplyAlignment;
     property ApplyProtection: boolean read FApplyProtection write FApplyProtection;
     property Alignment: TCT_CellAlignment read FAlignment;
     property Protection: TCT_CellProtection read FProtection;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_XfXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Xf;
public
     function  Add: TCT_Xf;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Xf read GetItems; default;
     end;

     TCT_CellStyle = class(TXPGBase)
protected
     FName: AxUCString;
     FXfId: integer;
     FBuiltinId: integer;
     FILevel: integer;
     FHidden: boolean;
     FCustomBuiltin: boolean;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property XfId: integer read FXfId write FXfId;
     property BuiltinId: integer read FBuiltinId write FBuiltinId;
     property ILevel: integer read FILevel write FILevel;
     property Hidden: boolean read FHidden write FHidden;
     property CustomBuiltin: boolean read FCustomBuiltin write FCustomBuiltin;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CellStyleXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CellStyle;
public
     function  Add: TCT_CellStyle;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CellStyle read GetItems; default;
     end;

     TCT_Dxf = class(TXPGBase)
protected
     FFont: TCT_Font;
     FNumFmt: TCT_NumFmt;
     FFill: TCT_Fill;
     FAlignment: TCT_CellAlignment;
     FBorder: TCT_Border;
     FProtection: TCT_CellProtection;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Font: TCT_Font read FFont;
     property NumFmt: TCT_NumFmt read FNumFmt;
     property Fill: TCT_Fill read FFill;
     property Alignment: TCT_CellAlignment read FAlignment;
     property Border: TCT_Border read FBorder;
     property Protection: TCT_CellProtection read FProtection;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_DxfXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Dxf;
public
     function  Add: TCT_Dxf;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Dxf read GetItems; default;
     end;

     TCT_TableStyle = class(TXPGBase)
protected
     FName: AxUCString;
     FPivot: boolean;
     FTable: boolean;
     FCount: integer;
     FTableStyleElementXpgList: TCT_TableStyleElementXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property Pivot: boolean read FPivot write FPivot;
     property Table: boolean read FTable write FTable;
     property Count_: integer read FCount write FCount;
     property TableStyleElementXpgList: TCT_TableStyleElementXpgList read FTableStyleElementXpgList;
     end;

     TCT_TableStyleXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TableStyle;
public
     function  Add: TCT_TableStyle;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_TableStyle read GetItems; default;
     end;

     TCT_IndexedColors = class(TXPGBase)
protected
     FRgbColorXpgList: TCT_RgbColorXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RgbColorXpgList: TCT_RgbColorXpgList read FRgbColorXpgList;
     end;

     TCT_MRUColors = class(TXPGBase)
protected
     FColorXpgList: TCT_ColorXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ColorXpgList: TCT_ColorXpgList read FColorXpgList;
     end;

     TCT_BookView = class(TXPGBase)
protected
     FVisibility: TST_Visibility;
     FMinimized: boolean;
     FShowHorizontalScroll: boolean;
     FShowVerticalScroll: boolean;
     FShowSheetTabs: boolean;
     FXWindow: integer;
     FYWindow: integer;
     FWindowWidth: integer;
     FWindowHeight: integer;
     FTabRatio: integer;
     FFirstSheet: integer;
     FActiveTab: integer;
     FAutoFilterDateGrouping: boolean;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Visibility: TST_Visibility read FVisibility write FVisibility;
     property Minimized: boolean read FMinimized write FMinimized;
     property ShowHorizontalScroll: boolean read FShowHorizontalScroll write FShowHorizontalScroll;
     property ShowVerticalScroll: boolean read FShowVerticalScroll write FShowVerticalScroll;
     property ShowSheetTabs: boolean read FShowSheetTabs write FShowSheetTabs;
     property XWindow: integer read FXWindow write FXWindow;
     property YWindow: integer read FYWindow write FYWindow;
     property WindowWidth: integer read FWindowWidth write FWindowWidth;
     property WindowHeight: integer read FWindowHeight write FWindowHeight;
     property TabRatio: integer read FTabRatio write FTabRatio;
     property FirstSheet: integer read FFirstSheet write FFirstSheet;
     property ActiveTab: integer read FActiveTab write FActiveTab;
     property AutoFilterDateGrouping: boolean read FAutoFilterDateGrouping write FAutoFilterDateGrouping;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_BookViewXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_BookView;
public
     function  Add: TCT_BookView;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_BookView read GetItems; default;
     end;

     TCT_Sheet = class(TXPGBase)
protected
     FName: AxUCString;
     FSheetId: integer;
     FState: TST_SheetState;
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property SheetId: integer read FSheetId write FSheetId;
     property State: TST_SheetState read FState write FState;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_SheetXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Sheet;
public
     function  Add: TCT_Sheet;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Sheet read GetItems; default;
     end;

     TCT_FunctionGroup = class(TXPGBase)
protected
     FName: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     end;

     TCT_FunctionGroupXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FunctionGroup;
public
     function  Add: TCT_FunctionGroup;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_FunctionGroup read GetItems; default;
     end;

     TCT_ExternalReference = class(TXPGBase)
protected
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_ExternalReferenceXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExternalReference;
public
     function  Add: TCT_ExternalReference;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ExternalReference read GetItems; default;
     end;

     TCT_DefinedName = class(TXPGBase)
protected
     FName: AxUCString;
     FComment: AxUCString;
     FCustomMenu: AxUCString;
     FDescription: AxUCString;
     FHelp: AxUCString;
     FStatusBar: AxUCString;
     FLocalSheetId: integer;
     FHidden: boolean;
     FFunction: boolean;
     FVbProcedure: boolean;
     FXlm: boolean;
     FFunctionGroupId: integer;
     FShortcutKey: AxUCString;
     FPublishToServer: boolean;
     FWorkbookParameter: boolean;
     FContent: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property Comment: AxUCString read FComment write FComment;
     property CustomMenu: AxUCString read FCustomMenu write FCustomMenu;
     property Description: AxUCString read FDescription write FDescription;
     property Help: AxUCString read FHelp write FHelp;
     property StatusBar: AxUCString read FStatusBar write FStatusBar;
     property LocalSheetId: integer read FLocalSheetId write FLocalSheetId;
     property Hidden: boolean read FHidden write FHidden;
     property Function_: boolean read FFunction write FFunction;
     property VbProcedure: boolean read FVbProcedure write FVbProcedure;
     property Xlm: boolean read FXlm write FXlm;
     property FunctionGroupId: integer read FFunctionGroupId write FFunctionGroupId;
     property ShortcutKey: AxUCString read FShortcutKey write FShortcutKey;
     property PublishToServer: boolean read FPublishToServer write FPublishToServer;
     property WorkbookParameter: boolean read FWorkbookParameter write FWorkbookParameter;
     property Content: AxUCString read FContent write FContent;
     end;

     TCT_DefinedNameXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DefinedName;
public
     function  Add: TCT_DefinedName;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_DefinedName read GetItems; default;
     end;

     TCT_CustomWorkbookView = class(TXPGBase)
protected
     FName: AxUCString;
     FGuid: AxUCString;
     FAutoUpdate: boolean;
     FMergeInterval: integer;
     FChangesSavedWin: boolean;
     FOnlySync: boolean;
     FPersonalView: boolean;
     FIncludePrintSettings: boolean;
     FIncludeHiddenRowCol: boolean;
     FMaximized: boolean;
     FMinimized: boolean;
     FShowHorizontalScroll: boolean;
     FShowVerticalScroll: boolean;
     FShowSheetTabs: boolean;
     FXWindow: integer;
     FYWindow: integer;
     FWindowWidth: integer;
     FWindowHeight: integer;
     FTabRatio: integer;
     FActiveSheetId: integer;
     FShowFormulaBar: boolean;
     FShowStatusbar: boolean;
     FShowComments: TST_Comments;
     FShowObjects: TST_Objects;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property Guid: AxUCString read FGuid write FGuid;
     property AutoUpdate: boolean read FAutoUpdate write FAutoUpdate;
     property MergeInterval: integer read FMergeInterval write FMergeInterval;
     property ChangesSavedWin: boolean read FChangesSavedWin write FChangesSavedWin;
     property OnlySync: boolean read FOnlySync write FOnlySync;
     property PersonalView: boolean read FPersonalView write FPersonalView;
     property IncludePrintSettings: boolean read FIncludePrintSettings write FIncludePrintSettings;
     property IncludeHiddenRowCol: boolean read FIncludeHiddenRowCol write FIncludeHiddenRowCol;
     property Maximized: boolean read FMaximized write FMaximized;
     property Minimized: boolean read FMinimized write FMinimized;
     property ShowHorizontalScroll: boolean read FShowHorizontalScroll write FShowHorizontalScroll;
     property ShowVerticalScroll: boolean read FShowVerticalScroll write FShowVerticalScroll;
     property ShowSheetTabs: boolean read FShowSheetTabs write FShowSheetTabs;
     property XWindow: integer read FXWindow write FXWindow;
     property YWindow: integer read FYWindow write FYWindow;
     property WindowWidth: integer read FWindowWidth write FWindowWidth;
     property WindowHeight: integer read FWindowHeight write FWindowHeight;
     property TabRatio: integer read FTabRatio write FTabRatio;
     property ActiveSheetId: integer read FActiveSheetId write FActiveSheetId;
     property ShowFormulaBar: boolean read FShowFormulaBar write FShowFormulaBar;
     property ShowStatusbar: boolean read FShowStatusbar write FShowStatusbar;
     property ShowComments: TST_Comments read FShowComments write FShowComments;
     property ShowObjects: TST_Objects read FShowObjects write FShowObjects;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CustomWorkbookViewXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CustomWorkbookView;
public
     function  Add: TCT_CustomWorkbookView;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CustomWorkbookView read GetItems; default;
     end;

     TCT_PivotCache = class(TXPGBase)
protected
     FCacheId: integer;
     FR_Id   : AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CacheId: integer read FCacheId write FCacheId;
     property R_Id   : AxUCString read FR_Id write FR_Id;
     end;

     TCT_PivotCacheXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_PivotCache;
public
     function  Add: TCT_PivotCache;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);

     property Items[Index: integer]: TCT_PivotCache read GetItems; default;
     end;

     TCT_SmartTagType = class(TXPGBase)
protected
     FNamespaceUri: AxUCString;
     FName: AxUCString;
     FUrl: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property NamespaceUri: AxUCString read FNamespaceUri write FNamespaceUri;
     property Name: AxUCString read FName write FName;
     property Url: AxUCString read FUrl write FUrl;
     end;

     TCT_SmartTagTypeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SmartTagType;
public
     function  Add: TCT_SmartTagType;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_SmartTagType read GetItems; default;
     end;

     TCT_WebPublishObject = class(TXPGBase)
protected
     FId: integer;
     FDivId: AxUCString;
     FSourceObject: AxUCString;
     FDestinationFile: AxUCString;
     FTitle: AxUCString;
     FAutoRepublish: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Id: integer read FId write FId;
     property DivId: AxUCString read FDivId write FDivId;
     property SourceObject: AxUCString read FSourceObject write FSourceObject;
     property DestinationFile: AxUCString read FDestinationFile write FDestinationFile;
     property Title: AxUCString read FTitle write FTitle;
     property AutoRepublish: boolean read FAutoRepublish write FAutoRepublish;
     end;

     TCT_WebPublishObjectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_WebPublishObject;
public
     function  Add: TCT_WebPublishObject;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_WebPublishObject read GetItems; default;
     end;

     TCT_Comment = class(TXPGBase)
protected
     FRef: AxUCString;
     FAuthorId: integer;
     FGuid: AxUCString;
     FText: TCT_Rst;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     property AuthorId: integer read FAuthorId write FAuthorId;
     property Guid: AxUCString read FGuid write FGuid;
     property Text: TCT_Rst read FText;
     end;

     TCT_CommentXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Comment;
public
     function  Add: TCT_Comment;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Comment read GetItems; default;
     end;

     TCT_ChartsheetView = class(TXPGBase)
protected
     FTabSelected: boolean;
     FZoomScale: integer;
     FWorkbookViewId: integer;
     FZoomToFit: boolean;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property TabSelected: boolean read FTabSelected write FTabSelected;
     property ZoomScale: integer read FZoomScale write FZoomScale;
     property WorkbookViewId: integer read FWorkbookViewId write FWorkbookViewId;
     property ZoomToFit: boolean read FZoomToFit write FZoomToFit;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ChartsheetViewXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ChartsheetView;
public
     function  Add: TCT_ChartsheetView;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ChartsheetView read GetItems; default;
     end;

     TCT_CustomChartsheetView = class(TXPGBase)
protected
     FGuid: AxUCString;
     FScale: integer;
     FState: TST_SheetState;
     FZoomToFit: boolean;
     FPageMargins: TCT_PageMargins;
     FPageSetup: TCT_CsPageSetup;
     FHeaderFooter: TCT_HeaderFooter;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Guid: AxUCString read FGuid write FGuid;
     property Scale: integer read FScale write FScale;
     property State: TST_SheetState read FState write FState;
     property ZoomToFit: boolean read FZoomToFit write FZoomToFit;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PageSetup: TCT_CsPageSetup read FPageSetup;
     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     end;

     TCT_CustomChartsheetViewXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CustomChartsheetView;
public
     function  Add: TCT_CustomChartsheetView;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CustomChartsheetView read GetItems; default;
     end;

     TCT_WebPublishItem = class(TXPGBase)
protected
     FId: integer;
     FDivId: AxUCString;
     FSourceType: TST_WebSourceType;
     FSourceRef: AxUCString;
     FSourceObject: AxUCString;
     FDestinationFile: AxUCString;
     FTitle: AxUCString;
     FAutoRepublish: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Id: integer read FId write FId;
     property DivId: AxUCString read FDivId write FDivId;
     property SourceType: TST_WebSourceType read FSourceType write FSourceType;
     property SourceRef: AxUCString read FSourceRef write FSourceRef;
     property SourceObject: AxUCString read FSourceObject write FSourceObject;
     property DestinationFile: AxUCString read FDestinationFile write FDestinationFile;
     property Title: AxUCString read FTitle write FTitle;
     property AutoRepublish: boolean read FAutoRepublish write FAutoRepublish;
     end;

     TCT_WebPublishItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_WebPublishItem;
public
     function  Add: TCT_WebPublishItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_WebPublishItem read GetItems; default;
     end;

     TCT_OutlinePr = class(TXPGBase)
protected
     FApplyStyles: boolean;
     FSummaryBelow: boolean;
     FSummaryRight: boolean;
     FShowOutlineSymbols: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ApplyStyles: boolean read FApplyStyles write FApplyStyles;
     property SummaryBelow: boolean read FSummaryBelow write FSummaryBelow;
     property SummaryRight: boolean read FSummaryRight write FSummaryRight;
     property ShowOutlineSymbols: boolean read FShowOutlineSymbols write FShowOutlineSymbols;
     end;

     TCT_PageSetUpPr = class(TXPGBase)
protected
     FAutoPageBreaks: boolean;
     FFitToPage: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property AutoPageBreaks: boolean read FAutoPageBreaks write FAutoPageBreaks;
     property FitToPage: boolean read FFitToPage write FFitToPage;
     end;

     TCT_SheetView = class(TXPGBase)
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
     FView: TST_SheetViewType;
     FTopLeftCell: AxUCString;
     FColorId: integer;
     FZoomScale: integer;
     FZoomScaleNormal: integer;
     FZoomScaleSheetLayoutView: integer;
     FZoomScalePageLayoutView: integer;
     FWorkbookViewId: integer;
     FPane: TCT_Pane;
     FSelectionXpgList: TCT_SelectionXpgList;
     FPivotSelectionXpgList: TCT_PivotSelectionXpgList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
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
     property View: TST_SheetViewType read FView write FView;
     property TopLeftCell: AxUCString read FTopLeftCell write FTopLeftCell;
     property ColorId: integer read FColorId write FColorId;
     property ZoomScale: integer read FZoomScale write FZoomScale;
     property ZoomScaleNormal: integer read FZoomScaleNormal write FZoomScaleNormal;
     property ZoomScaleSheetLayoutView: integer read FZoomScaleSheetLayoutView write FZoomScaleSheetLayoutView;
     property ZoomScalePageLayoutView: integer read FZoomScalePageLayoutView write FZoomScalePageLayoutView;
     property WorkbookViewId: integer read FWorkbookViewId write FWorkbookViewId;
     property Pane: TCT_Pane read FPane;
     property SelectionXpgList: TCT_SelectionXpgList read FSelectionXpgList;
     property PivotSelectionXpgList: TCT_PivotSelectionXpgList read FPivotSelectionXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SheetViewXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_SheetView;
public
     function  Add: TCT_SheetView;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_SheetView read GetItems; default;
     end;

     TCT_CustomSheetView = class(TXPGBase)
protected
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
     FState: TST_SheetState;
     FFilterUnique: boolean;
     FView: TST_SheetViewType;
     FShowRuler: boolean;
     FTopLeftCell: AxUCString;
     FPane: TCT_Pane;
     FSelection: TCT_Selection;
     FRowBreaks: TCT_PageBreak;
     FColBreaks: TCT_PageBreak;
     FPageMargins: TCT_PageMargins;
     FPrintOptions: TCT_PrintOptions;
     FPageSetup: TCT_PageSetup;
     FHeaderFooter: TCT_HeaderFooter;
     FAutoFilter: TCT_AutoFilter;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Guid: AxUCString read FGuid write FGuid;
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
     property State: TST_SheetState read FState write FState;
     property FilterUnique: boolean read FFilterUnique write FFilterUnique;
     property View: TST_SheetViewType read FView write FView;
     property ShowRuler: boolean read FShowRuler write FShowRuler;
     property TopLeftCell: AxUCString read FTopLeftCell write FTopLeftCell;
     property Pane: TCT_Pane read FPane;
     property Selection: TCT_Selection read FSelection;
     property RowBreaks: TCT_PageBreak read FRowBreaks;
     property ColBreaks: TCT_PageBreak read FColBreaks;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PrintOptions: TCT_PrintOptions read FPrintOptions;
     property PageSetup: TCT_PageSetup read FPageSetup;
     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     property AutoFilter: TCT_AutoFilter read FAutoFilter;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CustomSheetViewXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CustomSheetView;
public
     function  Add: TCT_CustomSheetView;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CustomSheetView read GetItems; default;
     end;

     TCT_OleObject = class(TXPGBase)
protected
     FProgId: AxUCString;
     FDvAspect: TST_DvAspect;
     FLink: AxUCString;
     FOleUpdate: TST_OleUpdate;
     FAutoLoad: boolean;
     FShapeId: integer;
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ProgId: AxUCString read FProgId write FProgId;
     property DvAspect: TST_DvAspect read FDvAspect write FDvAspect;
     property Link: AxUCString read FLink write FLink;
     property OleUpdate: TST_OleUpdate read FOleUpdate write FOleUpdate;
     property AutoLoad: boolean read FAutoLoad write FAutoLoad;
     property ShapeId: integer read FShapeId write FShapeId;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_OleObjectXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_OleObject;
public
     function  Add: TCT_OleObject;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_OleObject read GetItems; default;
     end;

     TCT_Col = class(TXPGBase)
protected
     FMin: integer;
     FMax: integer;
     FWidth: double;
     FStyle: integer;
     FHidden: boolean;
     FBestFit: boolean;
     FCustomWidth: boolean;
     FPhonetic: boolean;
     FOutlineLevel: integer;
     FCollapsed: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Min: integer read FMin write FMin;
     property Max: integer read FMax write FMax;
     property Width: double read FWidth write FWidth;
     property Style: integer read FStyle write FStyle;
     property Hidden: boolean read FHidden write FHidden;
     property BestFit: boolean read FBestFit write FBestFit;
     property CustomWidth: boolean read FCustomWidth write FCustomWidth;
     property Phonetic: boolean read FPhonetic write FPhonetic;
     property OutlineLevel: integer read FOutlineLevel write FOutlineLevel;
     property Collapsed: boolean read FCollapsed write FCollapsed;
     end;

     TCT_ColXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Col;
public
     function  Add: TCT_Col;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Col read GetItems; default;
     end;

     TCT_Row = class(TXPGBase)
protected
     FR: integer;
     FSpansXpgList: TStringXpgList;
     FS: integer;
     FCustomFormat: boolean;
     FHt: double;
     FHidden: boolean;
     FCustomHeight: boolean;
     FOutlineLevel: integer;
     FCollapsed: boolean;
     FThickTop: boolean;
     FThickBot: boolean;
     FPh: boolean;
     FC: TCT_Cell;
     FOnReadC: TNotifyEvent;
     FOnWriteC: TWriteElementEvent;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R: integer read FR write FR;
     property SpansXpgList: TStringXpgList read FSpansXpgList;
     property S: integer read FS write FS;
     property CustomFormat: boolean read FCustomFormat write FCustomFormat;
     property Ht: double read FHt write FHt;
     property Hidden: boolean read FHidden write FHidden;
     property CustomHeight: boolean read FCustomHeight write FCustomHeight;
     property OutlineLevel: integer read FOutlineLevel write FOutlineLevel;
     property Collapsed: boolean read FCollapsed write FCollapsed;
     property ThickTop: boolean read FThickTop write FThickTop;
     property ThickBot: boolean read FThickBot write FThickBot;
     property Ph: boolean read FPh write FPh;
     property _C: TCT_Cell read FC;
     property OnReadC: TNotifyEvent read FOnReadC write FOnReadC;
     property OnWriteC: TWriteElementEvent read FOnWriteC write FOnWriteC;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_RowXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Row;
public
     function  Add: TCT_Row;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Row read GetItems; default;
     end;

     TCT_ProtectedRange = class(TXPGBase)
protected
     FPassword: integer;
     FSqrefXpgList: TStringXpgList;
     FName: AxUCString;
     FSecurityDescriptor: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Password: integer read FPassword write FPassword;
     property SqrefXpgList: TStringXpgList read FSqrefXpgList;
     property Name: AxUCString read FName write FName;
     property SecurityDescriptor: AxUCString read FSecurityDescriptor write FSecurityDescriptor;
     end;

     TCT_ProtectedRangeXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ProtectedRange;
public
     function  Add: TCT_ProtectedRange;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ProtectedRange read GetItems; default;
     end;

     TCT_Scenario = class(TXPGBase)
protected
     FName: AxUCString;
     FLocked: boolean;
     FHidden: boolean;
     FCount: integer;
     FUser: AxUCString;
     FComment: AxUCString;
     FInputCellsXpgList: TCT_InputCellsXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property Locked: boolean read FLocked write FLocked;
     property Hidden: boolean read FHidden write FHidden;
     property Count_: integer read FCount write FCount;
     property User: AxUCString read FUser write FUser;
     property Comment: AxUCString read FComment write FComment;
     property InputCellsXpgList: TCT_InputCellsXpgList read FInputCellsXpgList;
     end;

     TCT_ScenarioXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Scenario;
public
     function  Add: TCT_Scenario;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Scenario read GetItems; default;
     end;

     TCT_DataRefs = class(TXPGBase)
protected
     FCount: integer;
     FDataRefXpgList: TCT_DataRefXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property DataRefXpgList: TCT_DataRefXpgList read FDataRefXpgList;
     end;

     TCT_MergeCell = class(TXPGBase)
protected
     FRef: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     end;

     TCT_MergeCellXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_MergeCell;
public
     function  Add: TCT_MergeCell;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_MergeCell read GetItems; default;
     end;

     TCT_CfRule = class(TXPGBase)
protected
     FType: TST_CfType;
     FDxfId: integer;
     FPriority: integer;
     FStopIfTrue: boolean;
     FAboveAverage: boolean;
     FPercent: boolean;
     FBottom: boolean;
     FOperator: TST_ConditionalFormattingOperator;
     FText: AxUCString;
     FTimePeriod: TST_TimePeriod;
     FRank: integer;
     FStdDev: integer;
     FEqualAverage: boolean;
     FFormulaXpgList: TStringXpgList;
     FColorScale: TCT_ColorScale;
     FDataBar: TCT_DataBar;
     FIconSet: TCT_IconSet;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: TST_CfType read FType write FType;
     property DxfId: integer read FDxfId write FDxfId;
     property Priority: integer read FPriority write FPriority;
     property StopIfTrue: boolean read FStopIfTrue write FStopIfTrue;
     property AboveAverage: boolean read FAboveAverage write FAboveAverage;
     property Percent: boolean read FPercent write FPercent;
     property Bottom: boolean read FBottom write FBottom;
     property Operator_: TST_ConditionalFormattingOperator read FOperator write FOperator;
     property Text: AxUCString read FText write FText;
     property TimePeriod: TST_TimePeriod read FTimePeriod write FTimePeriod;
     property Rank: integer read FRank write FRank;
     property StdDev: integer read FStdDev write FStdDev;
     property EqualAverage: boolean read FEqualAverage write FEqualAverage;
     property FormulaXpgList: TStringXpgList read FFormulaXpgList;
     property ColorScale: TCT_ColorScale read FColorScale;
     property DataBar: TCT_DataBar read FDataBar;
     property IconSet: TCT_IconSet read FIconSet;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_CfRuleXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CfRule;
public
     function  Add: TCT_CfRule;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CfRule read GetItems; default;
     end;

     TCT_DataValidation = class(TXPGBase)
protected
     FType: TST_DataValidationType;
     FErrorStyle: TST_DataValidationErrorStyle;
     FImeMode: TST_DataValidationImeMode;
     FOperator: TST_DataValidationOperator;
     FAllowBlank: boolean;
     FShowDropDown: boolean;
     FShowInputMessage: boolean;
     FShowErrorMessage: boolean;
     FErrorTitle: AxUCString;
     FError: AxUCString;
     FPromptTitle: AxUCString;
     FPrompt: AxUCString;
     FSqrefXpgList: TStringXpgList;
     FFormula1: AxUCString;
     FFormula2: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Type_: TST_DataValidationType read FType write FType;
     property ErrorStyle: TST_DataValidationErrorStyle read FErrorStyle write FErrorStyle;
     property ImeMode: TST_DataValidationImeMode read FImeMode write FImeMode;
     property Operator_: TST_DataValidationOperator read FOperator write FOperator;
     property AllowBlank: boolean read FAllowBlank write FAllowBlank;
     property ShowDropDown: boolean read FShowDropDown write FShowDropDown;
     property ShowInputMessage: boolean read FShowInputMessage write FShowInputMessage;
     property ShowErrorMessage: boolean read FShowErrorMessage write FShowErrorMessage;
     property ErrorTitle: AxUCString read FErrorTitle write FErrorTitle;
     property Error: AxUCString read FError write FError;
     property PromptTitle: AxUCString read FPromptTitle write FPromptTitle;
     property Prompt: AxUCString read FPrompt write FPrompt;
     property SqrefXpgList: TStringXpgList read FSqrefXpgList;
     property Formula1: AxUCString read FFormula1 write FFormula1;
     property Formula2: AxUCString read FFormula2 write FFormula2;
     end;

     TCT_DataValidationXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DataValidation;
public
     function  Add: TCT_DataValidation;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_DataValidation read GetItems; default;
     end;

     TCT_Hyperlink = class(TXPGBase)
protected
     FRef: AxUCString;
     FR_Id: AxUCString;
     FLocation: AxUCString;
     FTooltip: AxUCString;
     FDisplay: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     property R_Id: AxUCString read FR_Id write FR_Id;
     property Location: AxUCString read FLocation write FLocation;
     property Tooltip: AxUCString read FTooltip write FTooltip;
     property Display: AxUCString read FDisplay write FDisplay;
     end;

     TCT_HyperlinkXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Hyperlink;
public
     function  Add: TCT_Hyperlink;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Hyperlink read GetItems; default;
     end;

     TCT_CustomProperty = class(TXPGBase)
protected
     FName: AxUCString;
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_CustomPropertyXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CustomProperty;
public
     function  Add: TCT_CustomProperty;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CustomProperty read GetItems; default;
     end;

     TCT_CellWatch = class(TXPGBase)
protected
     FR: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R: AxUCString read FR write FR;
     end;

     TCT_CellWatchXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CellWatch;
public
     function  Add: TCT_CellWatch;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CellWatch read GetItems; default;
     end;

     TCT_IgnoredError = class(TXPGBase)
protected
     FSqrefXpgList: TStringXpgList;
     FEvalError: boolean;
     FTwoDigitTextYear: boolean;
     FNumberStoredAsText: boolean;
     FFormula: boolean;
     FFormulaRange: boolean;
     FUnlockedFormula: boolean;
     FEmptyCellReference: boolean;
     FListDataValidation: boolean;
     FCalculatedColumn: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SqrefXpgList: TStringXpgList read FSqrefXpgList;
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

     TCT_IgnoredErrorXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_IgnoredError;
public
     function  Add: TCT_IgnoredError;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_IgnoredError read GetItems; default;
     end;

     TCT_CellSmartTags = class(TXPGBase)
protected
     FR: AxUCString;
     FCellSmartTagXpgList: TCT_CellSmartTagXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R: AxUCString read FR write FR;
     property CellSmartTagXpgList: TCT_CellSmartTagXpgList read FCellSmartTagXpgList;
     end;

     TCT_CellSmartTagsXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_CellSmartTags;
public
     function  Add: TCT_CellSmartTags;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_CellSmartTags read GetItems; default;
     end;

     TCT_Control = class(TXPGBase)
protected
     FShapeId: integer;
     FR_Id: AxUCString;
     FName: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ShapeId: integer read FShapeId write FShapeId;
     property R_Id: AxUCString read FR_Id write FR_Id;
     property Name: AxUCString read FName write FName;
     end;

     TCT_ControlXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Control;
public
     function  Add: TCT_Control;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Control read GetItems; default;
     end;

     TCT_TablePart = class(TXPGBase)
protected
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_TablePartXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TablePart;
public
     function  Add: TCT_TablePart;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_TablePart read GetItems; default;
     end;

     TCT_NumFmts = class(TXPGBase)
protected
     FCount: integer;
     FNumFmtXpgList: TCT_NumFmtXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property NumFmtXpgList: TCT_NumFmtXpgList read FNumFmtXpgList;
     end;

     TCT_Fonts = class(TXPGBase)
protected
     FCount: integer;
     FFontXpgList: TCT_FontXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property FontXpgList: TCT_FontXpgList read FFontXpgList;
     end;

     TCT_Fills = class(TXPGBase)
protected
     FCount: integer;
     FFillXpgList: TCT_FillXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property FillXpgList: TCT_FillXpgList read FFillXpgList;
     end;

     TCT_Borders = class(TXPGBase)
protected
     FCount: integer;
     FBorderXpgList: TCT_BorderXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property BorderXpgList: TCT_BorderXpgList read FBorderXpgList;
     end;

     TCT_CellStyleXfs = class(TXPGBase)
protected
     FCount: integer;
     FXfXpgList: TCT_XfXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property XfXpgList: TCT_XfXpgList read FXfXpgList;
     end;

     TCT_CellXfs = class(TXPGBase)
protected
     FCount: integer;
     FXfXpgList: TCT_XfXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property XfXpgList: TCT_XfXpgList read FXfXpgList;
     end;

     TCT_CellStyles = class(TXPGBase)
protected
     FCount: integer;
     FCellStyleXpgList: TCT_CellStyleXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property CellStyleXpgList: TCT_CellStyleXpgList read FCellStyleXpgList;
     end;

     TCT_Dxfs = class(TXPGBase)
protected
     FCount: integer;
     FDxfXpgList: TCT_DxfXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property DxfXpgList: TCT_DxfXpgList read FDxfXpgList;
     end;

     TCT_TableStyles = class(TXPGBase)
protected
     FCount: integer;
     FDefaultTableStyle: AxUCString;
     FDefaultPivotStyle: AxUCString;
     FTableStyleXpgList: TCT_TableStyleXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property DefaultTableStyle: AxUCString read FDefaultTableStyle write FDefaultTableStyle;
     property DefaultPivotStyle: AxUCString read FDefaultPivotStyle write FDefaultPivotStyle;
     property TableStyleXpgList: TCT_TableStyleXpgList read FTableStyleXpgList;
     end;

     TCT_Colors = class(TXPGBase)
protected
     FIndexedColors: TCT_IndexedColors;
     FMruColors: TCT_MRUColors;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property IndexedColors: TCT_IndexedColors read FIndexedColors;
     property MruColors: TCT_MRUColors read FMruColors;
     end;

     TCT_FileVersion = class(TXPGBase)
protected
     FAppName: AxUCString;
     FLastEdited: AxUCString;
     FLowestEdited: AxUCString;
     FRupBuild: AxUCString;
     FCodeName: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property AppName: AxUCString read FAppName write FAppName;
     property LastEdited: AxUCString read FLastEdited write FLastEdited;
     property LowestEdited: AxUCString read FLowestEdited write FLowestEdited;
     property RupBuild: AxUCString read FRupBuild write FRupBuild;
     property CodeName: AxUCString read FCodeName write FCodeName;
     end;

     TCT_FileSharing = class(TXPGBase)
protected
     FReadOnlyRecommended: boolean;
     FUserName: AxUCString;
     FReservationPassword: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ReadOnlyRecommended: boolean read FReadOnlyRecommended write FReadOnlyRecommended;
     property UserName: AxUCString read FUserName write FUserName;
     property ReservationPassword: integer read FReservationPassword write FReservationPassword;
     end;

     TCT_WorkbookPr = class(TXPGBase)
protected
     FDate1904: boolean;
     FShowObjects: TST_Objects;
     FShowBorderUnselectedTables: boolean;
     FFilterPrivacy: boolean;
     FPromptedSolutions: boolean;
     FShowInkAnnotation: boolean;
     FBackupFile: boolean;
     FSaveExternalLinkValues: boolean;
     FUpdateLinks: TST_UpdateLinks;
     FCodeName: AxUCString;
     FHidePivotFieldList: boolean;
     FShowPivotChartFilter: boolean;
     FAllowRefreshQuery: boolean;
     FPublishItems: boolean;
     FCheckCompatibility: boolean;
     FAutoCompressPictures: boolean;
     FRefreshAllConnections: boolean;
     FDefaultThemeVersion: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Date1904: boolean read FDate1904 write FDate1904;
     property ShowObjects: TST_Objects read FShowObjects write FShowObjects;
     property ShowBorderUnselectedTables: boolean read FShowBorderUnselectedTables write FShowBorderUnselectedTables;
     property FilterPrivacy: boolean read FFilterPrivacy write FFilterPrivacy;
     property PromptedSolutions: boolean read FPromptedSolutions write FPromptedSolutions;
     property ShowInkAnnotation: boolean read FShowInkAnnotation write FShowInkAnnotation;
     property BackupFile: boolean read FBackupFile write FBackupFile;
     property SaveExternalLinkValues: boolean read FSaveExternalLinkValues write FSaveExternalLinkValues;
     property UpdateLinks: TST_UpdateLinks read FUpdateLinks write FUpdateLinks;
     property CodeName: AxUCString read FCodeName write FCodeName;
     property HidePivotFieldList: boolean read FHidePivotFieldList write FHidePivotFieldList;
     property ShowPivotChartFilter: boolean read FShowPivotChartFilter write FShowPivotChartFilter;
     property AllowRefreshQuery: boolean read FAllowRefreshQuery write FAllowRefreshQuery;
     property PublishItems: boolean read FPublishItems write FPublishItems;
     property CheckCompatibility: boolean read FCheckCompatibility write FCheckCompatibility;
     property AutoCompressPictures: boolean read FAutoCompressPictures write FAutoCompressPictures;
     property RefreshAllConnections: boolean read FRefreshAllConnections write FRefreshAllConnections;
     property DefaultThemeVersion: integer read FDefaultThemeVersion write FDefaultThemeVersion;
     end;

     TCT_WorkbookProtection = class(TXPGBase)
protected
     FWorkbookPassword: integer;
     FRevisionsPassword: integer;
     FLockStructure: boolean;
     FLockWindows: boolean;
     FLockRevision: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property WorkbookPassword: integer read FWorkbookPassword write FWorkbookPassword;
     property RevisionsPassword: integer read FRevisionsPassword write FRevisionsPassword;
     property LockStructure: boolean read FLockStructure write FLockStructure;
     property LockWindows: boolean read FLockWindows write FLockWindows;
     property LockRevision: boolean read FLockRevision write FLockRevision;
     end;

     TCT_BookViews = class(TXPGBase)
protected
     FWorkbookViewXpgList: TCT_BookViewXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property WorkbookViewXpgList: TCT_BookViewXpgList read FWorkbookViewXpgList;
     end;

     TCT_Sheets = class(TXPGBase)
protected
     FSheetXpgList: TCT_SheetXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetXpgList: TCT_SheetXpgList read FSheetXpgList;
     end;

     TCT_FunctionGroups = class(TXPGBase)
protected
     FBuiltInGroupCount: integer;
     FFunctionGroupXpgList: TCT_FunctionGroupXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property BuiltInGroupCount: integer read FBuiltInGroupCount write FBuiltInGroupCount;
     property FunctionGroupXpgList: TCT_FunctionGroupXpgList read FFunctionGroupXpgList;
     end;

     TCT_ExternalReferences = class(TXPGBase)
protected
     FExternalReferenceXpgList: TCT_ExternalReferenceXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ExternalReferenceXpgList: TCT_ExternalReferenceXpgList read FExternalReferenceXpgList;
     end;

     TCT_DefinedNames = class(TXPGBase)
protected
     FDefinedName: TCT_DefinedName;
     FOnReadDefinedName: TNotifyEvent;
     FOnWriteDefinedName: TWriteElementEvent;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DefinedName: TCT_DefinedName read FDefinedName;
     property OnReadDefinedName: TNotifyEvent read FOnReadDefinedName write FOnReadDefinedName;
     property OnWriteDefinedName: TWriteElementEvent read FOnWriteDefinedName write FOnWriteDefinedName;
     end;

     TCT_CalcPr = class(TXPGBase)
protected
     FCalcId: integer;
     FCalcMode: TST_CalcMode;
     FFullCalcOnLoad: boolean;
     FRefMode: TST_RefMode;
     FIterate: boolean;
     FIterateCount: integer;
     FIterateDelta: double;
     FFullPrecision: boolean;
     FCalcCompleted: boolean;
     FCalcOnSave: boolean;
     FConcurrentCalc: boolean;
     FConcurrentManualCount: integer;
     FForceFullCalc: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CalcId: integer read FCalcId write FCalcId;
     property CalcMode: TST_CalcMode read FCalcMode write FCalcMode;
     property FullCalcOnLoad: boolean read FFullCalcOnLoad write FFullCalcOnLoad;
     property RefMode: TST_RefMode read FRefMode write FRefMode;
     property Iterate: boolean read FIterate write FIterate;
     property IterateCount: integer read FIterateCount write FIterateCount;
     property IterateDelta: double read FIterateDelta write FIterateDelta;
     property FullPrecision: boolean read FFullPrecision write FFullPrecision;
     property CalcCompleted: boolean read FCalcCompleted write FCalcCompleted;
     property CalcOnSave: boolean read FCalcOnSave write FCalcOnSave;
     property ConcurrentCalc: boolean read FConcurrentCalc write FConcurrentCalc;
     property ConcurrentManualCount: integer read FConcurrentManualCount write FConcurrentManualCount;
     property ForceFullCalc: boolean read FForceFullCalc write FForceFullCalc;
     end;

     TCT_OleSize = class(TXPGBase)
protected
     FRef: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     end;

     TCT_CustomWorkbookViews = class(TXPGBase)
protected
     FCustomWorkbookViewXpgList: TCT_CustomWorkbookViewXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CustomWorkbookViewXpgList: TCT_CustomWorkbookViewXpgList read FCustomWorkbookViewXpgList;
     end;

     TCT_PivotCaches = class(TXPGBase)
protected
     FPivotCacheXpgList: TCT_PivotCacheXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property PivotCacheXpgList: TCT_PivotCacheXpgList read FPivotCacheXpgList;
     end;

     TCT_SmartTagPr = class(TXPGBase)
protected
     FEmbed: boolean;
     FShow: TST_SmartTagShow;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Embed: boolean read FEmbed write FEmbed;
     property Show: TST_SmartTagShow read FShow write FShow;
     end;

     TCT_SmartTagTypes = class(TXPGBase)
protected
     FSmartTagTypeXpgList: TCT_SmartTagTypeXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SmartTagTypeXpgList: TCT_SmartTagTypeXpgList read FSmartTagTypeXpgList;
     end;

     TCT_WebPublishing = class(TXPGBase)
protected
     FCss: boolean;
     FThicket: boolean;
     FLongFileNames: boolean;
     FVml: boolean;
     FAllowPng: boolean;
     FTargetScreenSize: TST_TargetScreenSize;
     FDpi: integer;
     FCodePage: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Css: boolean read FCss write FCss;
     property Thicket: boolean read FThicket write FThicket;
     property LongFileNames: boolean read FLongFileNames write FLongFileNames;
     property Vml: boolean read FVml write FVml;
     property AllowPng: boolean read FAllowPng write FAllowPng;
     property TargetScreenSize: TST_TargetScreenSize read FTargetScreenSize write FTargetScreenSize;
     property Dpi: integer read FDpi write FDpi;
     property CodePage: integer read FCodePage write FCodePage;
     end;

     TCT_FileRecoveryPr = class(TXPGBase)
protected
     FAutoRecover: boolean;
     FCrashSave: boolean;
     FDataExtractLoad: boolean;
     FRepairLoad: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property AutoRecover: boolean read FAutoRecover write FAutoRecover;
     property CrashSave: boolean read FCrashSave write FCrashSave;
     property DataExtractLoad: boolean read FDataExtractLoad write FDataExtractLoad;
     property RepairLoad: boolean read FRepairLoad write FRepairLoad;
     end;

     TCT_FileRecoveryPrXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_FileRecoveryPr;
public
     function  Add: TCT_FileRecoveryPr;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_FileRecoveryPr read GetItems; default;
     end;

     TCT_WebPublishObjects = class(TXPGBase)
protected
     FCount: integer;
     FWebPublishObjectXpgList: TCT_WebPublishObjectXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property WebPublishObjectXpgList: TCT_WebPublishObjectXpgList read FWebPublishObjectXpgList;
     end;

     TCT_Authors = class(TXPGBase)
protected
     FAuthorXpgList: TStringXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property AuthorXpgList: TStringXpgList read FAuthorXpgList;
     end;

     TCT_CommentList = class(TXPGBase)
protected
     FComment: TCT_Comment;
     FOnReadComment: TNotifyEvent;
     FOnWriteComment: TWriteElementEvent;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Comment: TCT_Comment read FComment;
     property OnReadComment: TNotifyEvent read FOnReadComment write FOnReadComment;
     property OnWriteComment: TWriteElementEvent read FOnWriteComment write FOnWriteComment;
     end;

     TCT_ChartsheetPr = class(TXPGBase)
protected
     FPublished: boolean;
     FCodeName: AxUCString;
     FTabColor: TCT_Color;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Published_: boolean read FPublished write FPublished;
     property CodeName: AxUCString read FCodeName write FCodeName;
     property TabColor: TCT_Color read FTabColor;
     end;

     TCT_ChartsheetViews = class(TXPGBase)
protected
     FSheetViewXpgList: TCT_ChartsheetViewXpgList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetViewXpgList: TCT_ChartsheetViewXpgList read FSheetViewXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ChartsheetProtection = class(TXPGBase)
protected
     FPassword: integer;
     FContent: boolean;
     FObjects: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Password: integer read FPassword write FPassword;
     property Content: boolean read FContent write FContent;
     property Objects: boolean read FObjects write FObjects;
     end;

     TCT_CustomChartsheetViews = class(TXPGBase)
protected
     FCustomSheetViewXpgList: TCT_CustomChartsheetViewXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CustomSheetViewXpgList: TCT_CustomChartsheetViewXpgList read FCustomSheetViewXpgList;
     end;

     TCT_Drawing = class(TXPGBase)
protected
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_LegacyDrawing = class(TXPGBase)
protected
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_SheetBackgroundPicture = class(TXPGBase)
protected
     FR_Id: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     end;

     TCT_WebPublishItems = class(TXPGBase)
protected
     FCount: integer;
     FWebPublishItemXpgList: TCT_WebPublishItemXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property WebPublishItemXpgList: TCT_WebPublishItemXpgList read FWebPublishItemXpgList;
     end;

     TCT_SheetPr = class(TXPGBase)
protected
     FSyncHorizontal: boolean;
     FSyncVertical: boolean;
     FSyncRef: AxUCString;
     FTransitionEvaluation: boolean;
     FTransitionEntry: boolean;
     FPublished: boolean;
     FCodeName: AxUCString;
     FFilterMode: boolean;
     FEnableFormatConditionsCalculation: boolean;
     FTabColor: TCT_Color;
     FOutlinePr: TCT_OutlinePr;
     FPageSetUpPr: TCT_PageSetUpPr;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SyncHorizontal: boolean read FSyncHorizontal write FSyncHorizontal;
     property SyncVertical: boolean read FSyncVertical write FSyncVertical;
     property SyncRef: AxUCString read FSyncRef write FSyncRef;
     property TransitionEvaluation: boolean read FTransitionEvaluation write FTransitionEvaluation;
     property TransitionEntry: boolean read FTransitionEntry write FTransitionEntry;
     property Published_: boolean read FPublished write FPublished;
     property CodeName: AxUCString read FCodeName write FCodeName;
     property FilterMode: boolean read FFilterMode write FFilterMode;
     property EnableFormatConditionsCalculation: boolean read FEnableFormatConditionsCalculation write FEnableFormatConditionsCalculation;
     property TabColor: TCT_Color read FTabColor;
     property OutlinePr: TCT_OutlinePr read FOutlinePr;
     property PageSetUpPr: TCT_PageSetUpPr read FPageSetUpPr;
     end;

     TCT_SheetViews = class(TXPGBase)
protected
     FSheetViewXpgList: TCT_SheetViewXpgList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetViewXpgList: TCT_SheetViewXpgList read FSheetViewXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SheetFormatPr = class(TXPGBase)
protected
     FBaseColWidth: integer;
     FDefaultColWidth: double;
     FDefaultRowHeight: double;
     FCustomHeight: boolean;
     FZeroHeight: boolean;
     FThickTop: boolean;
     FThickBottom: boolean;
     FOutlineLevelRow: integer;
     FOutlineLevelCol: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

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

     TCT_SheetProtection = class(TXPGBase)
protected
     FPassword: integer;
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

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Password: integer read FPassword write FPassword;
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

     TCT_CustomSheetViews = class(TXPGBase)
protected
     FCustomSheetViewXpgList: TCT_CustomSheetViewXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CustomSheetViewXpgList: TCT_CustomSheetViewXpgList read FCustomSheetViewXpgList;
     end;

     TCT_OleObjects = class(TXPGBase)
protected
     FOleObjectXpgList: TCT_OleObjectXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property OleObjectXpgList: TCT_OleObjectXpgList read FOleObjectXpgList;
     end;

     TCT_SheetDimension = class(TXPGBase)
protected
     FRef: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Ref: AxUCString read FRef write FRef;
     end;

     TCT_Cols = class(TXPGBase)
protected
     FCol: TCT_Col;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Col: TCT_Col read FCol;
     end;

     TCT_ColsXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_Cols;
public
     function  Add: TCT_Cols;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_Cols read GetItems; default;
     end;

     TCT_SheetData = class(TXPGBase)
protected
     FRow: TCT_Row;
     FOnReadRow: TNotifyEvent;
     FOnWriteRow: TWriteElementEvent;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Row: TCT_Row read FRow;
     property OnReadRow: TNotifyEvent read FOnReadRow write FOnReadRow;
     property OnWriteRow: TWriteElementEvent read FOnWriteRow write FOnWriteRow;
     end;

     TCT_SheetCalcPr = class(TXPGBase)
protected
     FFullCalcOnLoad: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property FullCalcOnLoad: boolean read FFullCalcOnLoad write FFullCalcOnLoad;
     end;

     TCT_ProtectedRanges = class(TXPGBase)
protected
     FProtectedRangeXpgList: TCT_ProtectedRangeXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ProtectedRangeXpgList: TCT_ProtectedRangeXpgList read FProtectedRangeXpgList;
     end;

     TCT_Scenarios = class(TXPGBase)
protected
     FCurrent: integer;
     FShow: integer;
     FSqrefXpgList: TStringXpgList;
     FScenarioXpgList: TCT_ScenarioXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Current: integer read FCurrent write FCurrent;
     property Show: integer read FShow write FShow;
     property SqrefXpgList: TStringXpgList read FSqrefXpgList;
     property ScenarioXpgList: TCT_ScenarioXpgList read FScenarioXpgList;
     end;

     TCT_DataConsolidate = class(TXPGBase)
protected
     FFunction: TST_DataConsolidateFunction;
     FLeftLabels: boolean;
     FTopLabels: boolean;
     FLink: boolean;
     FDataRefs: TCT_DataRefs;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Function_: TST_DataConsolidateFunction read FFunction write FFunction;
     property LeftLabels: boolean read FLeftLabels write FLeftLabels;
     property TopLabels: boolean read FTopLabels write FTopLabels;
     property Link: boolean read FLink write FLink;
     property DataRefs: TCT_DataRefs read FDataRefs;
     end;

     TCT_MergeCells = class(TXPGBase)
protected
     FCount: integer;
     FMergeCell: TCT_MergeCell;
     FOnReadMergeCell: TNotifyEvent;
     FOnWriteMergeCell: TWriteElementEvent;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property MergeCell: TCT_MergeCell read FMergeCell;
     property OnReadMergeCell: TNotifyEvent read FOnReadMergeCell write FOnReadMergeCell;
     property OnWriteMergeCell: TWriteElementEvent read FOnWriteMergeCell write FOnWriteMergeCell;
     end;

     TCT_ConditionalFormatting = class(TXPGBase)
protected
     FPivot: boolean;
     FSqrefXpgList: TStringXpgList;
     FCfRuleXpgList: TCT_CfRuleXpgList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Pivot: boolean read FPivot write FPivot;
     property SqrefXpgList: TStringXpgList read FSqrefXpgList;
     property CfRuleXpgList: TCT_CfRuleXpgList read FCfRuleXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_ConditionalFormattingXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ConditionalFormatting;
public
     function  Add: TCT_ConditionalFormatting;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ConditionalFormatting read GetItems; default;
     end;

     TCT_DataValidations = class(TXPGBase)
protected
     FDisablePrompts: boolean;
     FXWindow: integer;
     FYWindow: integer;
     FCount: integer;
     FDataValidationXpgList: TCT_DataValidationXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DisablePrompts: boolean read FDisablePrompts write FDisablePrompts;
     property XWindow: integer read FXWindow write FXWindow;
     property YWindow: integer read FYWindow write FYWindow;
     property Count_: integer read FCount write FCount;
     property DataValidationXpgList: TCT_DataValidationXpgList read FDataValidationXpgList;
     end;

     TCT_Hyperlinks = class(TXPGBase)
protected
     FHyperlink: TCT_Hyperlink;
     FOnReadHyperlink: TNotifyEvent;
     FOnWriteHyperlink: TWriteElementEvent;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Hyperlink: TCT_Hyperlink read FHyperlink;
     property OnReadHyperlink: TNotifyEvent read FOnReadHyperlink write FOnReadHyperlink;
     property OnWriteHyperlink: TWriteElementEvent read FOnWriteHyperlink write FOnWriteHyperlink;
     end;

     TCT_CustomProperties = class(TXPGBase)
protected
     FCustomPrXpgList: TCT_CustomPropertyXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CustomPrXpgList: TCT_CustomPropertyXpgList read FCustomPrXpgList;
     end;

     TCT_CellWatches = class(TXPGBase)
protected
     FCellWatchXpgList: TCT_CellWatchXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CellWatchXpgList: TCT_CellWatchXpgList read FCellWatchXpgList;
     end;

     TCT_IgnoredErrors = class(TXPGBase)
protected
     FIgnoredErrorXpgList: TCT_IgnoredErrorXpgList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property IgnoredErrorXpgList: TCT_IgnoredErrorXpgList read FIgnoredErrorXpgList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_SmartTags = class(TXPGBase)
protected
     FCellSmartTagsXpgList: TCT_CellSmartTagsXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property CellSmartTagsXpgList: TCT_CellSmartTagsXpgList read FCellSmartTagsXpgList;
     end;

     TCT_Controls = class(TXPGBase)
protected
     FControlXpgList: TCT_ControlXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ControlXpgList: TCT_ControlXpgList read FControlXpgList;
     end;

     TCT_TableParts = class(TXPGBase)
protected
     FCount: integer;
     FTablePartXpgList: TCT_TablePartXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property TablePartXpgList: TCT_TablePartXpgList read FTablePartXpgList;
     end;

     TCT_Stylesheet = class(TXPGBase)
protected
     FNumFmts: TCT_NumFmts;
     FFonts: TCT_Fonts;
     FFills: TCT_Fills;
     FBorders: TCT_Borders;
     FCellStyleXfs: TCT_CellStyleXfs;
     FCellXfs: TCT_CellXfs;
     FCellStyles: TCT_CellStyles;
     FDxfs: TCT_Dxfs;
     FTableStyles: TCT_TableStyles;
     FColors: TCT_Colors;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property NumFmts: TCT_NumFmts read FNumFmts;
     property Fonts: TCT_Fonts read FFonts;
     property Fills: TCT_Fills read FFills;
     property Borders: TCT_Borders read FBorders;
     property CellStyleXfs: TCT_CellStyleXfs read FCellStyleXfs;
     property CellXfs: TCT_CellXfs read FCellXfs;
     property CellStyles: TCT_CellStyles read FCellStyles;
     property Dxfs: TCT_Dxfs read FDxfs;
     property TableStyles: TCT_TableStyles read FTableStyles;
     property Colors: TCT_Colors read FColors;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Sst = class(TXPGBase)
protected
     FCount: integer;
     FUniqueCount: integer;
     FSi: TCT_Rst;
     FOnReadSi: TNotifyEvent;
     FOnWriteSi: TWriteElementEvent;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count_: integer read FCount write FCount;
     property UniqueCount: integer read FUniqueCount write FUniqueCount;
     property Si: TCT_Rst read FSi;
     property OnReadSi: TNotifyEvent read FOnReadSi write FOnReadSi;
     property OnWriteSi: TWriteElementEvent read FOnWriteSi write FOnWriteSi;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Workbook = class(TXPGBase)
protected
     FFileVersion: TCT_FileVersion;
     FFileSharing: TCT_FileSharing;
     FWorkbookPr: TCT_WorkbookPr;
     FWorkbookProtection: TCT_WorkbookProtection;
     FBookViews: TCT_BookViews;
     FSheets: TCT_Sheets;
     FFunctionGroups: TCT_FunctionGroups;
     FExternalReferences: TCT_ExternalReferences;
     FDefinedNames: TCT_DefinedNames;
     FOnReadDefinedNames: TNotifyEvent;
     FOnWriteDefinedNames: TWriteElementEvent;
     FCalcPr: TCT_CalcPr;
     FOleSize: TCT_OleSize;
     FCustomWorkbookViews: TCT_CustomWorkbookViews;
     FPivotCaches: TCT_PivotCaches;
     FSmartTagPr: TCT_SmartTagPr;
     FSmartTagTypes: TCT_SmartTagTypes;
     FWebPublishing: TCT_WebPublishing;
     FFileRecoveryPrXpgList: TCT_FileRecoveryPrXpgList;
     FWebPublishObjects: TCT_WebPublishObjects;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure AfterTag; override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property FileVersion: TCT_FileVersion read FFileVersion;
     property FileSharing: TCT_FileSharing read FFileSharing;
     property WorkbookPr: TCT_WorkbookPr read FWorkbookPr;
     property WorkbookProtection: TCT_WorkbookProtection read FWorkbookProtection;
     property BookViews: TCT_BookViews read FBookViews;
     property Sheets: TCT_Sheets read FSheets;
     property FunctionGroups: TCT_FunctionGroups read FFunctionGroups;
     property ExternalReferences: TCT_ExternalReferences read FExternalReferences;
     property DefinedNames: TCT_DefinedNames read FDefinedNames;
     property OnReadDefinedNames: TNotifyEvent read FOnReadDefinedNames write FOnReadDefinedNames;
     property OnWriteDefinedNames: TWriteElementEvent read FOnWriteDefinedNames write FOnWriteDefinedNames;
     property CalcPr: TCT_CalcPr read FCalcPr;
     property OleSize: TCT_OleSize read FOleSize;
     property CustomWorkbookViews: TCT_CustomWorkbookViews read FCustomWorkbookViews;
     property PivotCaches: TCT_PivotCaches read FPivotCaches;
     property SmartTagPr: TCT_SmartTagPr read FSmartTagPr;
     property SmartTagTypes: TCT_SmartTagTypes read FSmartTagTypes;
     property WebPublishing: TCT_WebPublishing read FWebPublishing;
     property FileRecoveryPrXpgList: TCT_FileRecoveryPrXpgList read FFileRecoveryPrXpgList;
     property WebPublishObjects: TCT_WebPublishObjects read FWebPublishObjects;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Comments = class(TXPGBase)
protected
     FAuthors: TCT_Authors;
     FCommentList: TCT_CommentList;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Authors: TCT_Authors read FAuthors;
     property CommentList: TCT_CommentList read FCommentList;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Chartsheet = class(TXPGBase)
protected
     FSheetPr: TCT_ChartsheetPr;
     FSheetViews: TCT_ChartsheetViews;
     FSheetProtection: TCT_ChartsheetProtection;
     FCustomSheetViews: TCT_CustomChartsheetViews;
     FPageMargins: TCT_PageMargins;
     FPageSetup: TCT_CsPageSetup;
     FHeaderFooter: TCT_HeaderFooter;
     FDrawing: TCT_Drawing;
     FLegacyDrawing: TCT_LegacyDrawing;
     FLegacyDrawingHF: TCT_LegacyDrawing;
     FPicture: TCT_SheetBackgroundPicture;
     FWebPublishItems: TCT_WebPublishItems;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetPr: TCT_ChartsheetPr read FSheetPr;
     property SheetViews: TCT_ChartsheetViews read FSheetViews;
     property SheetProtection: TCT_ChartsheetProtection read FSheetProtection;
     property CustomSheetViews: TCT_CustomChartsheetViews read FCustomSheetViews;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PageSetup: TCT_CsPageSetup read FPageSetup;
     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     property Drawing: TCT_Drawing read FDrawing;
     property LegacyDrawing: TCT_LegacyDrawing read FLegacyDrawing;
     property LegacyDrawingHF: TCT_LegacyDrawing read FLegacyDrawingHF;
     property Picture: TCT_SheetBackgroundPicture read FPicture;
     property WebPublishItems: TCT_WebPublishItems read FWebPublishItems;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Dialogsheet = class(TXPGBase)
protected
     FSheetPr: TCT_SheetPr;
     FSheetViews: TCT_SheetViews;
     FSheetFormatPr: TCT_SheetFormatPr;
     FSheetProtection: TCT_SheetProtection;
     FCustomSheetViews: TCT_CustomSheetViews;
     FPrintOptions: TCT_PrintOptions;
     FPageMargins: TCT_PageMargins;
     FPageSetup: TCT_PageSetup;
     FHeaderFooter: TCT_HeaderFooter;
     FDrawing: TCT_Drawing;
     FLegacyDrawing: TCT_LegacyDrawing;
     FLegacyDrawingHF: TCT_LegacyDrawing;
     FOleObjects: TCT_OleObjects;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetPr: TCT_SheetPr read FSheetPr;
     property SheetViews: TCT_SheetViews read FSheetViews;
     property SheetFormatPr: TCT_SheetFormatPr read FSheetFormatPr;
     property SheetProtection: TCT_SheetProtection read FSheetProtection;
     property CustomSheetViews: TCT_CustomSheetViews read FCustomSheetViews;
     property PrintOptions: TCT_PrintOptions read FPrintOptions;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PageSetup: TCT_PageSetup read FPageSetup;
     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     property Drawing: TCT_Drawing read FDrawing;
     property LegacyDrawing: TCT_LegacyDrawing read FLegacyDrawing;
     property LegacyDrawingHF: TCT_LegacyDrawing read FLegacyDrawingHF;
     property OleObjects: TCT_OleObjects read FOleObjects;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Worksheet = class(TXPGBase)
protected
     FSheetPr: TCT_SheetPr;
     FDimension: TCT_SheetDimension;
     FSheetViews: TCT_SheetViews;
     FSheetFormatPr: TCT_SheetFormatPr;
     FColsXpgList: TCT_ColsXpgList;
     FSheetData: TCT_SheetData;
     FSheetCalcPr: TCT_SheetCalcPr;
     FSheetProtection: TCT_SheetProtection;
     FProtectedRanges: TCT_ProtectedRanges;
     FScenarios: TCT_Scenarios;
     FAutoFilter: TCT_AutoFilter;
     FSortState: TCT_SortState;
     FDataConsolidate: TCT_DataConsolidate;
     FCustomSheetViews: TCT_CustomSheetViews;
     FMergeCells: TCT_MergeCells;
     FPhoneticPr: TCT_PhoneticPr;
     FConditionalFormattingXpgList: TCT_ConditionalFormattingXpgList;
     FDataValidations: TCT_DataValidations;
     FHyperlinks: TCT_Hyperlinks;
     FPrintOptions: TCT_PrintOptions;
     FPageMargins: TCT_PageMargins;
     FPageSetup: TCT_PageSetup;
     FHeaderFooter: TCT_HeaderFooter;
     FRowBreaks: TCT_PageBreak;
     FColBreaks: TCT_PageBreak;
     FCustomProperties: TCT_CustomProperties;
     FCellWatches: TCT_CellWatches;
     FIgnoredErrors: TCT_IgnoredErrors;
     FSmartTags: TCT_SmartTags;
     FDrawing: TCT_Drawing;
     FLegacyDrawing: TCT_LegacyDrawing;
     FLegacyDrawingHF: TCT_LegacyDrawing;
     FPicture: TCT_SheetBackgroundPicture;
     FOleObjects: TCT_OleObjects;
     FControls: TCT_Controls;
     FWebPublishItems: TCT_WebPublishItems;
     FTableParts: TCT_TableParts;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetPr: TCT_SheetPr read FSheetPr;
     property Dimension: TCT_SheetDimension read FDimension;
     property SheetViews: TCT_SheetViews read FSheetViews;
     property SheetFormatPr: TCT_SheetFormatPr read FSheetFormatPr;
     property ColsXpgList: TCT_ColsXpgList read FColsXpgList;
     property SheetData: TCT_SheetData read FSheetData;
     property SheetCalcPr: TCT_SheetCalcPr read FSheetCalcPr;
     property SheetProtection: TCT_SheetProtection read FSheetProtection;
     property ProtectedRanges: TCT_ProtectedRanges read FProtectedRanges;
     property Scenarios: TCT_Scenarios read FScenarios;
     property AutoFilter: TCT_AutoFilter read FAutoFilter;
     property SortState: TCT_SortState read FSortState;
     property DataConsolidate: TCT_DataConsolidate read FDataConsolidate;
     property CustomSheetViews: TCT_CustomSheetViews read FCustomSheetViews;
     property MergeCells: TCT_MergeCells read FMergeCells;
     property PhoneticPr: TCT_PhoneticPr read FPhoneticPr;
     property ConditionalFormattingXpgList: TCT_ConditionalFormattingXpgList read FConditionalFormattingXpgList;
     property DataValidations: TCT_DataValidations read FDataValidations;
     property Hyperlinks: TCT_Hyperlinks read FHyperlinks;
     property PrintOptions: TCT_PrintOptions read FPrintOptions;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PageSetup: TCT_PageSetup read FPageSetup;
     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     property RowBreaks: TCT_PageBreak read FRowBreaks;
     property ColBreaks: TCT_PageBreak read FColBreaks;
     property CustomProperties: TCT_CustomProperties read FCustomProperties;
     property CellWatches: TCT_CellWatches read FCellWatches;
     property IgnoredErrors: TCT_IgnoredErrors read FIgnoredErrors;
     property SmartTags: TCT_SmartTags read FSmartTags;
     property Drawing: TCT_Drawing read FDrawing;
     property LegacyDrawing: TCT_LegacyDrawing read FLegacyDrawing;
     property LegacyDrawingHF: TCT_LegacyDrawing read FLegacyDrawingHF;
     property Picture: TCT_SheetBackgroundPicture read FPicture;
     property OleObjects: TCT_OleObjects read FOleObjects;
     property Controls: TCT_Controls read FControls;
     property WebPublishItems: TCT_WebPublishItems read FWebPublishItems;
     property TableParts: TCT_TableParts read FTableParts;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_Macrosheet = class(TXPGBase)
protected
     FSheetPr: TCT_SheetPr;
     FDimension: TCT_SheetDimension;
     FSheetViews: TCT_SheetViews;
     FSheetFormatPr: TCT_SheetFormatPr;
     FColsXpgList: TCT_ColsXpgList;
     FSheetData: TCT_SheetData;
     FSheetProtection: TCT_SheetProtection;
     FAutoFilter: TCT_AutoFilter;
     FSortState: TCT_SortState;
     FDataConsolidate: TCT_DataConsolidate;
     FCustomSheetViews: TCT_CustomSheetViews;
     FPhoneticPr: TCT_PhoneticPr;
     FConditionalFormattingXpgList: TCT_ConditionalFormattingXpgList;
     FPrintOptions: TCT_PrintOptions;
     FPageMargins: TCT_PageMargins;
     FPageSetup: TCT_PageSetup;
     FHeaderFooter: TCT_HeaderFooter;
     FRowBreaks: TCT_PageBreak;
     FColBreaks: TCT_PageBreak;
     FCustomProperties: TCT_CustomProperties;
     FDrawing: TCT_Drawing;
     FLegacyDrawing: TCT_LegacyDrawing;
     FLegacyDrawingHF: TCT_LegacyDrawing;
     FPicture: TCT_SheetBackgroundPicture;
     FOleObjects: TCT_OleObjects;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetPr: TCT_SheetPr read FSheetPr;
     property Dimension: TCT_SheetDimension read FDimension;
     property SheetViews: TCT_SheetViews read FSheetViews;
     property SheetFormatPr: TCT_SheetFormatPr read FSheetFormatPr;
     property ColsXpgList: TCT_ColsXpgList read FColsXpgList;
     property SheetData: TCT_SheetData read FSheetData;
     property SheetProtection: TCT_SheetProtection read FSheetProtection;
     property AutoFilter: TCT_AutoFilter read FAutoFilter;
     property SortState: TCT_SortState read FSortState;
     property DataConsolidate: TCT_DataConsolidate read FDataConsolidate;
     property CustomSheetViews: TCT_CustomSheetViews read FCustomSheetViews;
     property PhoneticPr: TCT_PhoneticPr read FPhoneticPr;
     property ConditionalFormattingXpgList: TCT_ConditionalFormattingXpgList read FConditionalFormattingXpgList;
     property PrintOptions: TCT_PrintOptions read FPrintOptions;
     property PageMargins: TCT_PageMargins read FPageMargins;
     property PageSetup: TCT_PageSetup read FPageSetup;
     property HeaderFooter: TCT_HeaderFooter read FHeaderFooter;
     property RowBreaks: TCT_PageBreak read FRowBreaks;
     property ColBreaks: TCT_PageBreak read FColBreaks;
     property CustomProperties: TCT_CustomProperties read FCustomProperties;
     property Drawing: TCT_Drawing read FDrawing;
     property LegacyDrawing: TCT_LegacyDrawing read FLegacyDrawing;
     property LegacyDrawingHF: TCT_LegacyDrawing read FLegacyDrawingHF;
     property Picture: TCT_SheetBackgroundPicture read FPicture;
     property OleObjects: TCT_OleObjects read FOleObjects;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_TableStyleInfo = class(TXPGBase)
protected
     FName: AxUCString;
     FShowFirstColumn: boolean;
     FShowLastColumn: boolean;
     FShowRowStripes: boolean;
     FShowColumnStripes: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property ShowFirstColumn: boolean read FShowFirstColumn write FShowFirstColumn;
     property ShowLastColumn: boolean read FShowLastColumn write FShowLastColumn;
     property ShowRowStripes: boolean read FShowRowStripes write FShowRowStripes;
     property ShowColumnStripes: boolean read FShowColumnStripes write FShowColumnStripes;
     end;

     TCT_XmlColumnPr = class(TXPGBase)
protected
     FMapId: integer;
     FXpath: AxUCString;
     FDenormalized: boolean;
     FXmlDataType: TST_XmlDataType;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property MapId: integer read FMapId write FMapId;
     property Xpath: AxUCString read FXpath write FXpath;
     property Denormalized: boolean read FDenormalized write FDenormalized;
     property XmlDataType: TST_XmlDataType read FXmlDataType write FXmlDataType;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_TableFormula = class(TXPGBase)
protected
     FArray: boolean;
     FContent: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Array_: boolean read FArray write FArray;
     property Content: AxUCString read FContent write FContent;
     end;

     TCT_TableColumn = class(TXPGBase)
protected
     FId: integer;
     FUniqueName: AxUCString;
     FName: AxUCString;
     FTotalsRowFunction: TST_TotalsRowFunction;
     FTotalsRowLabel: AxUCString;
     FQueryTableFieldId: integer;
     FHeaderRowDxfId: AxUCString;
     FDataDxfId: AxUCString;
     FTotalsRowDxfId: AxUCString;
     FHeaderRowCellStyle: AxUCString;
     FDataCellStyle: AxUCString;
     FTotalsRowCellStyle: AxUCString;
     FCalculatedColumnFormula: TCT_TableFormula;
     FTotalsRowFormula: TCT_TableFormula;
     FXmlColumnPr: TCT_XmlColumnPr;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Id: integer read FId write FId;
     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Name: AxUCString read FName write FName;
     property TotalsRowFunction: TST_TotalsRowFunction read FTotalsRowFunction write FTotalsRowFunction;
     property TotalsRowLabel: AxUCString read FTotalsRowLabel write FTotalsRowLabel;
     property QueryTableFieldId: integer read FQueryTableFieldId write FQueryTableFieldId;
     property HeaderRowDxfId: AxUCString read FHeaderRowDxfId write FHeaderRowDxfId;
     property DataDxfId: AxUCString read FDataDxfId write FDataDxfId;
     property TotalsRowDxfId: AxUCString read FTotalsRowDxfId write FTotalsRowDxfId;
     property HeaderRowCellStyle: AxUCString read FHeaderRowCellStyle write FHeaderRowCellStyle;
     property DataCellStyle: AxUCString read FDataCellStyle write FDataCellStyle;
     property TotalsRowCellStyle: AxUCString read FTotalsRowCellStyle write FTotalsRowCellStyle;
     property CalculatedColumnFormula: TCT_TableFormula read FCalculatedColumnFormula;
     property TotalsRowFormula: TCT_TableFormula read FTotalsRowFormula;
     property XmlColumnPr: TCT_XmlColumnPr read FXmlColumnPr;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     TCT_TableColumnXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_TableColumn;
public
     function  Add: TCT_TableColumn;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_TableColumn read GetItems; default;
     end;

     TCT_TableColumns = class(TXPGBase)
protected
     FCount: integer;
     FTableColumnXpgList: TCT_TableColumnXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Count: integer read FCount write FCount;
     property TableColumnXpgList: TCT_TableColumnXpgList read FTableColumnXpgList;
     end;

     TCT_Table = class(TXPGBase)
protected
     FId: integer;
     FName: AxUCString;
     FDisplayName: AxUCString;
     FComment: AxUCString;
     FRef: AxUCString;
     FTableType: TST_TableType;
     FHeaderRowCount: integer;
     FInsertRow: boolean;
     FInsertRowShift: boolean;
     FTotalsRowCount: integer;
     FTotalsRowShown: boolean;
     FPublished: boolean;
     FHeaderRowDxfId: AxUCString;
     FDataDxfId: AxUCString;
     FTotalsRowDxfId: AxUCString;
     FHeaderRowBorderDxfId: AxUCString;
     FTableBorderDxfId: AxUCString;
     FTotalsRowBorderDxfId: AxUCString;
     FHeaderRowCellStyle: AxUCString;
     FDataCellStyle: AxUCString;
     FTotalsRowCellStyle: AxUCString;
     FConnectionId: integer;
     FAutoFilter: TCT_AutoFilter;
     FSortState: TCT_SortState;
     FTableColumns: TCT_TableColumns;
     FTableStyleInfo: TCT_TableStyleInfo;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Id: integer read FId write FId;
     property Name: AxUCString read FName write FName;
     property DisplayName: AxUCString read FDisplayName write FDisplayName;
     property Comment: AxUCString read FComment write FComment;
     property Ref: AxUCString read FRef write FRef;
     property TableType: TST_TableType read FTableType write FTableType;
     property HeaderRowCount: integer read FHeaderRowCount write FHeaderRowCount;
     property InsertRow: boolean read FInsertRow write FInsertRow;
     property InsertRowShift: boolean read FInsertRowShift write FInsertRowShift;
     property TotalsRowCount: integer read FTotalsRowCount write FTotalsRowCount;
     property TotalsRowShown: boolean read FTotalsRowShown write FTotalsRowShown;
     property Published_: boolean read FPublished write FPublished;
     property HeaderRowDxfId: AxUCString read FHeaderRowDxfId write FHeaderRowDxfId;
     property DataDxfId: AxUCString read FDataDxfId write FDataDxfId;
     property TotalsRowDxfId: AxUCString read FTotalsRowDxfId write FTotalsRowDxfId;
     property HeaderRowBorderDxfId: AxUCString read FHeaderRowBorderDxfId write FHeaderRowBorderDxfId;
     property TableBorderDxfId: AxUCString read FTableBorderDxfId write FTableBorderDxfId;
     property TotalsRowBorderDxfId: AxUCString read FTotalsRowBorderDxfId write FTotalsRowBorderDxfId;
     property HeaderRowCellStyle: AxUCString read FHeaderRowCellStyle write FHeaderRowCellStyle;
     property DataCellStyle: AxUCString read FDataCellStyle write FDataCellStyle;
     property TotalsRowCellStyle: AxUCString read FTotalsRowCellStyle write FTotalsRowCellStyle;
     property ConnectionId: integer read FConnectionId write FConnectionId;
     property AutoFilter: TCT_AutoFilter read FAutoFilter;
     property SortState: TCT_SortState read FSortState;
     property TableColumns: TCT_TableColumns read FTableColumns;
     property TableStyleInfo: TCT_TableStyleInfo read FTableStyleInfo;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FStyleSheet: TCT_Stylesheet;
     FSst: TCT_Sst;
     FWorkbook: TCT_Workbook;
     FComments: TCT_Comments;
     FTable: TCT_Table;
     FChartsheet: TCT_Chartsheet;
     FDialogsheet: TCT_Dialogsheet;
     FWorksheet: TCT_Worksheet;
     FCurrWriteClass: TClass;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML; AClassToWrite: TClass);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property StyleSheet: TCT_Stylesheet read FStyleSheet;
     property Sst: TCT_Sst read FSst;
     property Workbook: TCT_Workbook read FWorkbook;
     property Comments: TCT_Comments read FComments;
     property Table: TCT_Table read FTable;
     property Chartsheet: TCT_Chartsheet read FChartsheet;
     property Dialogsheet: TCT_Dialogsheet read FDialogsheet;
     property Worksheet: TCT_Worksheet read FWorksheet;
     end;

     TCT_XStringElement = class(TXPGBase)
protected
     FV: AxUCString;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property V: AxUCString read FV write FV;
     end;

     TXPGDocXLSX = class(TXPGDocBase)
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;

     function  GetStyleSheet: TCT_Stylesheet;
     function  GetSst: TCT_Sst;
     function  GetWorkbook: TCT_Workbook;
     function  GetComments: TCT_Comments;
     function  GetTable: TCT_Table;
     function  GetChartsheet: TCT_Chartsheet;
     function  GetDialogsheet: TCT_Dialogsheet;
     function  GetWorksheet: TCT_Worksheet;
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromFile(const AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(const AFilename: AxUCString);

     procedure SaveToStream(AStream: TStream; AClassToWrite: TClass);

     property Root: T__ROOT__ read FRoot;
     property StyleSheet: TCT_Stylesheet read GetStyleSheet;
     property Sst: TCT_Sst read GetSst;
     property Workbook: TCT_Workbook read GetWorkbook;
     property Comments: TCT_Comments read GetComments;
     property Table: TCT_Table read GetTable;
     property Chartsheet: TCT_Chartsheet read GetChartsheet;
     property Dialogsheet: TCT_Dialogsheet read GetDialogsheet;
     property Worksheet: TCT_Worksheet read GetWorksheet;
     end;


implementation

{$ifdef DELPHI_5}
var FEnums: TStringList;
{$else}
var FEnums: THashedStringList;
{$endif}


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

class procedure TXPGBase.AddEnums;
begin
  FEnums.AddObject('stuvSingle',TObject(0));
  FEnums.AddObject('stuvDouble',TObject(1));
  FEnums.AddObject('stuvSingleAccounting',TObject(2));
  FEnums.AddObject('stuvDoubleAccounting',TObject(3));
  FEnums.AddObject('stuvNone',TObject(4));
  FEnums.AddObject('stvarBaseline',TObject(0));
  FEnums.AddObject('stvarSuperscript',TObject(1));
  FEnums.AddObject('stvarSubscript',TObject(2));
  FEnums.AddObject('stvaTop',TObject(0));
  FEnums.AddObject('stvaCenter',TObject(1));
  FEnums.AddObject('stvaBottom',TObject(2));
  FEnums.AddObject('stvaJustify',TObject(3));
  FEnums.AddObject('stvaDistributed',TObject(4));
  FEnums.AddObject('stbsNone',TObject(0));
  FEnums.AddObject('stbsThin',TObject(1));
  FEnums.AddObject('stbsMedium',TObject(2));
  FEnums.AddObject('stbsDashed',TObject(3));
  FEnums.AddObject('stbsDotted',TObject(4));
  FEnums.AddObject('stbsThick',TObject(5));
  FEnums.AddObject('stbsDouble',TObject(6));
  FEnums.AddObject('stbsHair',TObject(7));
  FEnums.AddObject('stbsMediumDashed',TObject(8));
  FEnums.AddObject('stbsDashDot',TObject(9));
  FEnums.AddObject('stbsMediumDashDot',TObject(10));
  FEnums.AddObject('stbsDashDotDot',TObject(11));
  FEnums.AddObject('stbsMediumDashDotDot',TObject(12));
  FEnums.AddObject('stbsSlantDashDot',TObject(13));
  FEnums.AddObject('stptNone',TObject(0));
  FEnums.AddObject('stptSolid',TObject(1));
  FEnums.AddObject('stptMediumGray',TObject(2));
  FEnums.AddObject('stptDarkGray',TObject(3));
  FEnums.AddObject('stptLightGray',TObject(4));
  FEnums.AddObject('stptDarkHorizontal',TObject(5));
  FEnums.AddObject('stptDarkVertical',TObject(6));
  FEnums.AddObject('stptDarkDown',TObject(7));
  FEnums.AddObject('stptDarkUp',TObject(8));
  FEnums.AddObject('stptDarkGrid',TObject(9));
  FEnums.AddObject('stptDarkTrellis',TObject(10));
  FEnums.AddObject('stptLightHorizontal',TObject(11));
  FEnums.AddObject('stptLightVertical',TObject(12));
  FEnums.AddObject('stptLightDown',TObject(13));
  FEnums.AddObject('stptLightUp',TObject(14));
  FEnums.AddObject('stptLightGrid',TObject(15));
  FEnums.AddObject('stptLightTrellis',TObject(16));
  FEnums.AddObject('stptGray125',TObject(17));
  FEnums.AddObject('stptGray0625',TObject(18));
  FEnums.AddObject('stfsNone',TObject(0));
  FEnums.AddObject('stfsMajor',TObject(1));
  FEnums.AddObject('stfsMinor',TObject(2));
  FEnums.AddObject('stgtLinear',TObject(0));
  FEnums.AddObject('stgtPath',TObject(1));
  FEnums.AddObject('sttstWholeTable',TObject(0));
  FEnums.AddObject('sttstHeaderRow',TObject(1));
  FEnums.AddObject('sttstTotalRow',TObject(2));
  FEnums.AddObject('sttstFirstColumn',TObject(3));
  FEnums.AddObject('sttstLastColumn',TObject(4));
  FEnums.AddObject('sttstFirstRowStripe',TObject(5));
  FEnums.AddObject('sttstSecondRowStripe',TObject(6));
  FEnums.AddObject('sttstFirstColumnStripe',TObject(7));
  FEnums.AddObject('sttstSecondColumnStripe',TObject(8));
  FEnums.AddObject('sttstFirstHeaderCell',TObject(9));
  FEnums.AddObject('sttstLastHeaderCell',TObject(10));
  FEnums.AddObject('sttstFirstTotalCell',TObject(11));
  FEnums.AddObject('sttstLastTotalCell',TObject(12));
  FEnums.AddObject('sttstFirstSubtotalColumn',TObject(13));
  FEnums.AddObject('sttstSecondSubtotalColumn',TObject(14));
  FEnums.AddObject('sttstThirdSubtotalColumn',TObject(15));
  FEnums.AddObject('sttstFirstSubtotalRow',TObject(16));
  FEnums.AddObject('sttstSecondSubtotalRow',TObject(17));
  FEnums.AddObject('sttstThirdSubtotalRow',TObject(18));
  FEnums.AddObject('sttstBlankRow',TObject(19));
  FEnums.AddObject('sttstFirstColumnSubheading',TObject(20));
  FEnums.AddObject('sttstSecondColumnSubheading',TObject(21));
  FEnums.AddObject('sttstThirdColumnSubheading',TObject(22));
  FEnums.AddObject('sttstFirstRowSubheading',TObject(23));
  FEnums.AddObject('sttstSecondRowSubheading',TObject(24));
  FEnums.AddObject('sttstThirdRowSubheading',TObject(25));
  FEnums.AddObject('sttstPageFieldLabels',TObject(26));
  FEnums.AddObject('sttstPageFieldValues',TObject(27));
  FEnums.AddObject('sthaGeneral',TObject(0));
  FEnums.AddObject('sthaLeft',TObject(1));
  FEnums.AddObject('sthaCenter',TObject(2));
  FEnums.AddObject('sthaRight',TObject(3));
  FEnums.AddObject('sthaFill',TObject(4));
  FEnums.AddObject('sthaJustify',TObject(5));
  FEnums.AddObject('sthaCenterContinuous',TObject(6));
  FEnums.AddObject('sthaDistributed',TObject(7));
  FEnums.AddObject('stpaNoControl',TObject(0));
  FEnums.AddObject('stpaLeft',TObject(1));
  FEnums.AddObject('stpaCenter',TObject(2));
  FEnums.AddObject('stpaDistributed',TObject(3));
  FEnums.AddObject('stptHalfwidthKatakana',TObject(0));
  FEnums.AddObject('stptFullwidthKatakana',TObject(1));
  FEnums.AddObject('stptHiragana',TObject(2));
  FEnums.AddObject('stptNoConversion',TObject(3));
  FEnums.AddObject('stdtgYear',TObject(0));
  FEnums.AddObject('stdtgMonth',TObject(1));
  FEnums.AddObject('stdtgDay',TObject(2));
  FEnums.AddObject('stdtgHour',TObject(3));
  FEnums.AddObject('stdtgMinute',TObject(4));
  FEnums.AddObject('stdtgSecond',TObject(5));
  FEnums.AddObject('stsmStroke',TObject(0));
  FEnums.AddObject('stsmPinYin',TObject(1));
  FEnums.AddObject('stsmNone',TObject(2));
  FEnums.AddObject('stctNone',TObject(0));
  FEnums.AddObject('stctGregorian',TObject(1));
  FEnums.AddObject('stctGregorianUs',TObject(2));
  FEnums.AddObject('stctJapan',TObject(3));
  FEnums.AddObject('stctTaiwan',TObject(4));
  FEnums.AddObject('stctKorea',TObject(5));
  FEnums.AddObject('stctHijri',TObject(6));
  FEnums.AddObject('stctThai',TObject(7));
  FEnums.AddObject('stctHebrew',TObject(8));
  FEnums.AddObject('stctGregorianMeFrench',TObject(9));
  FEnums.AddObject('stctGregorianArabic',TObject(10));
  FEnums.AddObject('stctGregorianXlitEnglish',TObject(11));
  FEnums.AddObject('stctGregorianXlitFrench',TObject(12));
  FEnums.AddObject('stsbValue',TObject(0));
  FEnums.AddObject('stsbCellColor',TObject(1));
  FEnums.AddObject('stsbFontColor',TObject(2));
  FEnums.AddObject('stsbIcon',TObject(3));
  FEnums.AddObject('stist3Arrows',TObject(0));
  FEnums.AddObject('stist3ArrowsGray',TObject(1));
  FEnums.AddObject('stist3Flags',TObject(2));
  FEnums.AddObject('stist3TrafficLights1',TObject(3));
  FEnums.AddObject('stist3TrafficLights2',TObject(4));
  FEnums.AddObject('stist3Signs',TObject(5));
  FEnums.AddObject('stist3Symbols',TObject(6));
  FEnums.AddObject('stist3Symbols2',TObject(7));
  FEnums.AddObject('stist4Arrows',TObject(8));
  FEnums.AddObject('stist4ArrowsGray',TObject(9));
  FEnums.AddObject('stist4RedToBlack',TObject(10));
  FEnums.AddObject('stist4Rating',TObject(11));
  FEnums.AddObject('stist4TrafficLights',TObject(12));
  FEnums.AddObject('stist5Arrows',TObject(13));
  FEnums.AddObject('stist5ArrowsGray',TObject(14));
  FEnums.AddObject('stist5Rating',TObject(15));
  FEnums.AddObject('stist5Quarters',TObject(16));
  FEnums.AddObject('stdftNull',TObject(0));
  FEnums.AddObject('stdftAboveAverage',TObject(1));
  FEnums.AddObject('stdftBelowAverage',TObject(2));
  FEnums.AddObject('stdftTomorrow',TObject(3));
  FEnums.AddObject('stdftToday',TObject(4));
  FEnums.AddObject('stdftYesterday',TObject(5));
  FEnums.AddObject('stdftNextWeek',TObject(6));
  FEnums.AddObject('stdftThisWeek',TObject(7));
  FEnums.AddObject('stdftLastWeek',TObject(8));
  FEnums.AddObject('stdftNextMonth',TObject(9));
  FEnums.AddObject('stdftThisMonth',TObject(10));
  FEnums.AddObject('stdftLastMonth',TObject(11));
  FEnums.AddObject('stdftNextQuarter',TObject(12));
  FEnums.AddObject('stdftThisQuarter',TObject(13));
  FEnums.AddObject('stdftLastQuarter',TObject(14));
  FEnums.AddObject('stdftNextYear',TObject(15));
  FEnums.AddObject('stdftThisYear',TObject(16));
  FEnums.AddObject('stdftLastYear',TObject(17));
  FEnums.AddObject('stdftYearToDate',TObject(18));
  FEnums.AddObject('stdftQ1',TObject(19));
  FEnums.AddObject('stdftQ2',TObject(20));
  FEnums.AddObject('stdftQ3',TObject(21));
  FEnums.AddObject('stdftQ4',TObject(22));
  FEnums.AddObject('stdftM1',TObject(23));
  FEnums.AddObject('stdftM2',TObject(24));
  FEnums.AddObject('stdftM3',TObject(25));
  FEnums.AddObject('stdftM4',TObject(26));
  FEnums.AddObject('stdftM5',TObject(27));
  FEnums.AddObject('stdftM6',TObject(28));
  FEnums.AddObject('stdftM7',TObject(29));
  FEnums.AddObject('stdftM8',TObject(30));
  FEnums.AddObject('stdftM9',TObject(31));
  FEnums.AddObject('stdftM10',TObject(32));
  FEnums.AddObject('stdftM11',TObject(33));
  FEnums.AddObject('stdftM12',TObject(34));
  FEnums.AddObject('stfoEqual',TObject(0));
  FEnums.AddObject('stfoLessThan',TObject(1));
  FEnums.AddObject('stfoLessThanOrEqual',TObject(2));
  FEnums.AddObject('stfoNotEqual',TObject(3));
  FEnums.AddObject('stfoGreaterThanOrEqual',TObject(4));
  FEnums.AddObject('stfoGreaterThan',TObject(5));
  FEnums.AddObject('staAxisRow',TObject(0));
  FEnums.AddObject('staAxisCol',TObject(1));
  FEnums.AddObject('staAxisPage',TObject(2));
  FEnums.AddObject('staAxisValues',TObject(3));
  FEnums.AddObject('stpatNone',TObject(0));
  FEnums.AddObject('stpatNormal',TObject(1));
  FEnums.AddObject('stpatData',TObject(2));
  FEnums.AddObject('stpatAll',TObject(3));
  FEnums.AddObject('stpatOrigin',TObject(4));
  FEnums.AddObject('stpatButton',TObject(5));
  FEnums.AddObject('stpatTopRight',TObject(6));
  FEnums.AddObject('stvVisible',TObject(0));
  FEnums.AddObject('stvHidden',TObject(1));
  FEnums.AddObject('stvVeryHidden',TObject(2));
  FEnums.AddObject('stcmManual',TObject(0));
  FEnums.AddObject('stcmAuto',TObject(1));
  FEnums.AddObject('stcmAutoNoTable',TObject(2));
  FEnums.AddObject('stulUserSet',TObject(0));
  FEnums.AddObject('stulNever',TObject(1));
  FEnums.AddObject('stulAlways',TObject(2));
  FEnums.AddObject('sttss544x376',TObject(0));
  FEnums.AddObject('sttss640x480',TObject(1));
  FEnums.AddObject('sttss720x512',TObject(2));
  FEnums.AddObject('sttss800x600',TObject(3));
  FEnums.AddObject('sttss1024x768',TObject(4));
  FEnums.AddObject('sttss1152x882',TObject(5));
  FEnums.AddObject('sttss1152x900',TObject(6));
  FEnums.AddObject('sttss1280x1024',TObject(7));
  FEnums.AddObject('sttss1600x1200',TObject(8));
  FEnums.AddObject('sttss1800x1440',TObject(9));
  FEnums.AddObject('sttss1920x1200',TObject(10));
  FEnums.AddObject('stcCommNone',TObject(0));
  FEnums.AddObject('stcCommIndicator',TObject(1));
  FEnums.AddObject('stcCommIndAndComment',TObject(2));
  FEnums.AddObject('stoAll',TObject(0));
  FEnums.AddObject('stoPlaceholders',TObject(1));
  FEnums.AddObject('stoNone',TObject(2));
  FEnums.AddObject('ststsAll',TObject(0));
  FEnums.AddObject('ststsNone',TObject(1));
  FEnums.AddObject('ststsNoIndicator',TObject(2));
  FEnums.AddObject('strmA1',TObject(0));
  FEnums.AddObject('strmR1C1',TObject(1));
  FEnums.AddObject('stssVisible',TObject(0));
  FEnums.AddObject('stssHidden',TObject(1));
  FEnums.AddObject('stssVeryHidden',TObject(2));
  FEnums.AddObject('stdaDVASPECT_CONTENT',TObject(0));
  FEnums.AddObject('stdaDVASPECT_ICON',TObject(1));
  FEnums.AddObject('stctNum',TObject(0));
  FEnums.AddObject('stctPercent',TObject(1));
  FEnums.AddObject('stctMax',TObject(2));
  FEnums.AddObject('stctMin',TObject(3));
  FEnums.AddObject('stctFormula',TObject(4));
  FEnums.AddObject('stctPercentile',TObject(5));
  FEnums.AddObject('stpoDownThenOver',TObject(0));
  FEnums.AddObject('stpoOverThenDown',TObject(1));
  FEnums.AddObject('stctB',TObject(0));
  FEnums.AddObject('stctN',TObject(1));
  FEnums.AddObject('stctD',TObject(2));
  FEnums.AddObject('stctE',TObject(3));
  FEnums.AddObject('stctS',TObject(4));
  FEnums.AddObject('stctStr',TObject(5));
  FEnums.AddObject('stctInlineStr',TObject(6));
  FEnums.AddObject('stpsSplit',TObject(0));
  FEnums.AddObject('stpsFrozen',TObject(1));
  FEnums.AddObject('stpsFrozenSplit',TObject(2));
  FEnums.AddObject('stcftNormal',TObject(0));
  FEnums.AddObject('stcftArray',TObject(1));
  FEnums.AddObject('stcftDataTable',TObject(2));
  FEnums.AddObject('stcftShared',TObject(3));
  FEnums.AddObject('stccNone',TObject(0));
  FEnums.AddObject('stccAsDisplayed',TObject(1));
  FEnums.AddObject('stccAtEnd',TObject(2));
  FEnums.AddObject('stdvoBetween',TObject(0));
  FEnums.AddObject('stdvoNotBetween',TObject(1));
  FEnums.AddObject('stdvoEqual',TObject(2));
  FEnums.AddObject('stdvoNotEqual',TObject(3));
  FEnums.AddObject('stdvoLessThan',TObject(4));
  FEnums.AddObject('stdvoLessThanOrEqual',TObject(5));
  FEnums.AddObject('stdvoGreaterThan',TObject(6));
  FEnums.AddObject('stdvoGreaterThanOrEqual',TObject(7));
  FEnums.AddObject('stctExpression',TObject(0));
  FEnums.AddObject('stctCellIs',TObject(1));
  FEnums.AddObject('stctColorScale',TObject(2));
  FEnums.AddObject('stctDataBar',TObject(3));
  FEnums.AddObject('stctIconSet',TObject(4));
  FEnums.AddObject('stctTop10',TObject(5));
  FEnums.AddObject('stctUniqueValues',TObject(6));
  FEnums.AddObject('stctDuplicateValues',TObject(7));
  FEnums.AddObject('stctContainsText',TObject(8));
  FEnums.AddObject('stctNotContainsText',TObject(9));
  FEnums.AddObject('stctBeginsWith',TObject(10));
  FEnums.AddObject('stctEndsWith',TObject(11));
  FEnums.AddObject('stctContainsBlanks',TObject(12));
  FEnums.AddObject('stctNotContainsBlanks',TObject(13));
  FEnums.AddObject('stctContainsErrors',TObject(14));
  FEnums.AddObject('stctNotContainsErrors',TObject(15));
  FEnums.AddObject('stctTimePeriod',TObject(16));
  FEnums.AddObject('stctAboveAverage',TObject(17));
  FEnums.AddObject('stpeDisplayed',TObject(0));
  FEnums.AddObject('stpeBlank',TObject(1));
  FEnums.AddObject('stpeDash',TObject(2));
  FEnums.AddObject('stpeNA',TObject(3));
  FEnums.AddObject('stouOLEUPDATE_ALWAYS',TObject(0));
  FEnums.AddObject('stouOLEUPDATE_ONCALL',TObject(1));
  FEnums.AddObject('stwstSheet',TObject(0));
  FEnums.AddObject('stwstPrintArea',TObject(1));
  FEnums.AddObject('stwstAutoFilter',TObject(2));
  FEnums.AddObject('stwstRange',TObject(3));
  FEnums.AddObject('stwstChart',TObject(4));
  FEnums.AddObject('stwstPivotTable',TObject(5));
  FEnums.AddObject('stwstQuery',TObject(6));
  FEnums.AddObject('stwstLabel',TObject(7));
  FEnums.AddObject('stdvesStop',TObject(0));
  FEnums.AddObject('stdvesWarning',TObject(1));
  FEnums.AddObject('stdvesInformation',TObject(2));
  FEnums.AddObject('sttpToday',TObject(0));
  FEnums.AddObject('sttpYesterday',TObject(1));
  FEnums.AddObject('sttpTomorrow',TObject(2));
  FEnums.AddObject('sttpLast7Days',TObject(3));
  FEnums.AddObject('sttpThisMonth',TObject(4));
  FEnums.AddObject('sttpLastMonth',TObject(5));
  FEnums.AddObject('sttpNextMonth',TObject(6));
  FEnums.AddObject('sttpThisWeek',TObject(7));
  FEnums.AddObject('sttpLastWeek',TObject(8));
  FEnums.AddObject('sttpNextWeek',TObject(9));
  FEnums.AddObject('stpBottomRight',TObject(0));
  FEnums.AddObject('stpTopRight',TObject(1));
  FEnums.AddObject('stpBottomLeft',TObject(2));
  FEnums.AddObject('stpTopLeft',TObject(3));
  FEnums.AddObject('stcfoLessThan',TObject(0));
  FEnums.AddObject('stcfoLessThanOrEqual',TObject(1));
  FEnums.AddObject('stcfoEqual',TObject(2));
  FEnums.AddObject('stcfoNotEqual',TObject(3));
  FEnums.AddObject('stcfoGreaterThanOrEqual',TObject(4));
  FEnums.AddObject('stcfoGreaterThan',TObject(5));
  FEnums.AddObject('stcfoBetween',TObject(6));
  FEnums.AddObject('stcfoNotBetween',TObject(7));
  FEnums.AddObject('stcfoContainsText',TObject(8));
  FEnums.AddObject('stcfoNotContains',TObject(9));
  FEnums.AddObject('stcfoBeginsWith',TObject(10));
  FEnums.AddObject('stcfoEndsWith',TObject(11));
  FEnums.AddObject('stdvimNoControl',TObject(0));
  FEnums.AddObject('stdvimOff',TObject(1));
  FEnums.AddObject('stdvimOn',TObject(2));
  FEnums.AddObject('stdvimDisabled',TObject(3));
  FEnums.AddObject('stdvimHiragana',TObject(4));
  FEnums.AddObject('stdvimFullKatakana',TObject(5));
  FEnums.AddObject('stdvimHalfKatakana',TObject(6));
  FEnums.AddObject('stdvimFullAlpha',TObject(7));
  FEnums.AddObject('stdvimHalfAlpha',TObject(8));
  FEnums.AddObject('stdvimFullHangul',TObject(9));
  FEnums.AddObject('stdvimHalfHangul',TObject(10));
  FEnums.AddObject('stsvtNormal',TObject(0));
  FEnums.AddObject('stsvtPageBreakPreview',TObject(1));
  FEnums.AddObject('stsvtPageLayout',TObject(2));
  FEnums.AddObject('stdcfAverage',TObject(0));
  FEnums.AddObject('stdcfCount',TObject(1));
  FEnums.AddObject('stdcfCountNums',TObject(2));
  FEnums.AddObject('stdcfMax',TObject(3));
  FEnums.AddObject('stdcfMin',TObject(4));
  FEnums.AddObject('stdcfProduct',TObject(5));
  FEnums.AddObject('stdcfStdDev',TObject(6));
  FEnums.AddObject('stdcfStdDevp',TObject(7));
  FEnums.AddObject('stdcfSum',TObject(8));
  FEnums.AddObject('stdcfVar',TObject(9));
  FEnums.AddObject('stdcfVarp',TObject(10));
  FEnums.AddObject('stdvtNone',TObject(0));
  FEnums.AddObject('stdvtWhole',TObject(1));
  FEnums.AddObject('stdvtDecimal',TObject(2));
  FEnums.AddObject('stdvtList',TObject(3));
  FEnums.AddObject('stdvtDate',TObject(4));
  FEnums.AddObject('stdvtTime',TObject(5));
  FEnums.AddObject('stdvtTextLength',TObject(6));
  FEnums.AddObject('stdvtCustom',TObject(7));
  FEnums.AddObject('stoDefault',TObject(0));
  FEnums.AddObject('stoPortrait',TObject(1));
  FEnums.AddObject('stoLandscape',TObject(2));

  FEnums.AddObject('stttworksheet',TObject(0));
  FEnums.AddObject('stttxml',TObject(1));
  FEnums.AddObject('stttqueryTable',TObject(2));

  FEnums.AddObject('sttrfnone',TObject(0));
  FEnums.AddObject('sttrfsum',TObject(1));
  FEnums.AddObject('sttrfmin',TObject(2));
  FEnums.AddObject('sttrfmax',TObject(3));
  FEnums.AddObject('sttrfaverage',TObject(4));
  FEnums.AddObject('sttrfcount',TObject(5));
  FEnums.AddObject('sttrfcountNums',TObject(6));
  FEnums.AddObject('sttrfstdDev',TObject(7));
  FEnums.AddObject('sttrfvar',TObject(8));
  FEnums.AddObject('sttrfcustom',TObject(9));

  FEnums.AddObject('stxdtstring',TObject(0));
  FEnums.AddObject('stxdtnormalizedString',TObject(1));
  FEnums.AddObject('stxdttoken',TObject(2));
  FEnums.AddObject('stxdtbyte',TObject(3));
  FEnums.AddObject('stxdtunsignedByte',TObject(4));
  FEnums.AddObject('stxdtbase64Binary',TObject(5));
  FEnums.AddObject('stxdthexBinary',TObject(6));
  FEnums.AddObject('stxdtinteger',TObject(7));
  FEnums.AddObject('stxdtpositiveInteger',TObject(8));
  FEnums.AddObject('stxdtnegativeInteger',TObject(9));
  FEnums.AddObject('stxdtnonPositiveInteger',TObject(10));
  FEnums.AddObject('stxdtnonNegativeInteger',TObject(11));
  FEnums.AddObject('stxdtint',TObject(12));
  FEnums.AddObject('stxdtunsignedInt',TObject(13));
  FEnums.AddObject('stxdtlong',TObject(14));
  FEnums.AddObject('stxdtunsignedLong',TObject(15));
  FEnums.AddObject('stxdtshort',TObject(16));
  FEnums.AddObject('stxdtunsignedShort',TObject(17));
  FEnums.AddObject('stxdtdecimal',TObject(18));
  FEnums.AddObject('stxdtfloat',TObject(19));
  FEnums.AddObject('stxdtdouble',TObject(20));
  FEnums.AddObject('stxdtboolean',TObject(21));
  FEnums.AddObject('stxdttime',TObject(22));
  FEnums.AddObject('stxdtdateTime',TObject(23));
  FEnums.AddObject('stxdtduration',TObject(24));
  FEnums.AddObject('stxdtdate',TObject(25));
  FEnums.AddObject('stxdtgMonth',TObject(26));
  FEnums.AddObject('stxdtgYear',TObject(27));
  FEnums.AddObject('stxdtgYearMonth',TObject(28));
  FEnums.AddObject('stxdtgDay',TObject(29));
  FEnums.AddObject('stxdtgMonthDay',TObject(30));
  FEnums.AddObject('stxdtName',TObject(31));
  FEnums.AddObject('stxdtQName',TObject(32));
  FEnums.AddObject('stxdtNCName',TObject(33));
  FEnums.AddObject('stxdtanyURI',TObject(34));
  FEnums.AddObject('stxdtlanguage',TObject(35));
  FEnums.AddObject('stxdtID',TObject(36));
  FEnums.AddObject('stxdtIDREF',TObject(37));
  FEnums.AddObject('stxdtIDREFS',TObject(38));
  FEnums.AddObject('stxdtENTITY',TObject(39));
  FEnums.AddObject('stxdtENTITIES',TObject(40));
  FEnums.AddObject('stxdtNOTATION',TObject(41));
  FEnums.AddObject('stxdtNMTOKEN',TObject(42));
  FEnums.AddObject('stxdtNMTOKENS',TObject(43));
  FEnums.AddObject('stxdtanyType',TObject(44));
end;

class function  TXPGBase.StrToEnum(const AValue: AxUCString): integer;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(FEnums.Objects[i])
  else
    Result := 0;
end;

class function  TXPGBase.StrToEnumDef(const AValue: AxUCString; ADefault: integer): integer;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then
    Result := Integer(FEnums.Objects[i])
  else
    Result := ADefault;
end;

class function  TXPGBase.TryStrToEnum(const AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then 
  begin
    i := Integer(FEnums.Objects[i]);
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
  if Attributes.Count > 0 then begin
    FCurrent.AssignAttributes(Attributes);
//    Attributes.Clear;
  end;
end;

procedure TXPGReader.EndTag;
begin
  FCurrent := TXPGBase(FStack.Pop);
  FCurrent.AfterTag;
end;

{ TCT_Extension }

function  TCT_Extension.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FUri <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Extension.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Extension.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FUri <> '' then 
    AWriter.AddAttribute('uri',FUri);
end;

procedure TCT_Extension.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'uri' then 
    FUri := AAttributes.Values[0];
//  else  // TODO Any element
//    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Extension.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_Extension.Destroy;
begin
end;

procedure TCT_Extension.Clear;
begin
  FAssigneds := [];
  FUri := '';
end;

{ TCT_ExtensionXpgList }

function  TCT_ExtensionXpgList.GetItems(Index: integer): TCT_Extension;
begin
  Result := TCT_Extension(inherited Items[Index]);
end;

function  TCT_ExtensionXpgList.Add: TCT_Extension;
begin
  Result := TCT_Extension.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExtensionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExtensionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Index }

function  TCT_Index.CheckAssigned: integer;
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
end;

destructor TCT_Index.Destroy;
begin
end;

procedure TCT_Index.Clear;
begin
  FAssigneds := [];
  FV := 0;
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

{ TCT_ExtensionList }

function  TCT_ExtensionList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FExtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExtensionList.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'ext' then 
    Result := FExtXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExtensionList.Write(AWriter: TXpgWriteXML);
begin
  FExtXpgList.Write(AWriter,'ext');
end;

constructor TCT_ExtensionList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FExtXpgList := TCT_ExtensionXpgList.Create(FOwner);
end;

destructor TCT_ExtensionList.Destroy;
begin
  FExtXpgList.Free;
end;

procedure TCT_ExtensionList.Clear;
begin
  FAssigneds := [];
  FExtXpgList.Clear;
end;

{ TCT_FontName }

function  TCT_FontName.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FontName.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FontName.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',FVal);
end;

procedure TCT_FontName.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FontName.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_FontName.Destroy;
begin
end;

procedure TCT_FontName.Clear;
begin
  FAssigneds := [];
  FVal := '';
end;

{ TCT_FontNameXpgList }

function  TCT_FontNameXpgList.GetItems(Index: integer): TCT_FontName;
begin
  Result := TCT_FontName(inherited Items[Index]);
end;

function  TCT_FontNameXpgList.Add: TCT_FontName;
begin
  Result := TCT_FontName.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FontNameXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FontNameXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_IntProperty }

function  TCT_IntProperty.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_IntProperty.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_IntProperty.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlIntToStr(FVal));
end;

procedure TCT_IntProperty.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_IntProperty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_IntProperty.Destroy;
begin
end;

procedure TCT_IntProperty.Clear;
begin
  FAssigneds := [];
  FVal := 0;
end;

{ TCT_IntPropertyXpgList }

function  TCT_IntPropertyXpgList.GetItems(Index: integer): TCT_IntProperty;
begin
  Result := TCT_IntProperty(inherited Items[Index]);
end;

function  TCT_IntPropertyXpgList.Add: TCT_IntProperty;
begin
  Result := TCT_IntProperty.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_IntPropertyXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_IntPropertyXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_BooleanProperty }

function  TCT_BooleanProperty.CheckAssigned: integer;
//var
//  AttrsAssigned: integer;
begin
//  AttrsAssigned := 0;
//  FAssigneds := [];
//  if FDoWrite then
//    Inc(AttrsAssigned);
//  Result := 0;
//  Inc(Result,AttrsAssigned);
//  if AttrsAssigned > 0 then
//    FAssigneds := [xaAttributes];

  FAssigneds := [];
  if FVal = True then begin
    Result := 1;
    FAssigneds := [xaAttributes];
  end
  else
    Result := 0;
end;

procedure TCT_BooleanProperty.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_BooleanProperty.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> True then 
    AWriter.AddAttribute('val',XmlBoolToStr(FVal));
end;

procedure TCT_BooleanProperty.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := XmlStrToBoolDef(AAttributes.Values[0],True)
  else if AAttributes.Count > 0 then  // TODO Only error if there is an attribute.
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BooleanProperty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := False;
end;

destructor TCT_BooleanProperty.Destroy;
begin
end;

procedure TCT_BooleanProperty.Clear;
begin
  FAssigneds := [];
  FVal := False;
end;

{ TCT_Color }

function  TCT_Color.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAuto <> False then
    Inc(AttrsAssigned);
  if FIndexed <> -1 then
    Inc(AttrsAssigned);
  if FRgbAssigned then
    Inc(AttrsAssigned);
  if FTheme <> -1 then
    Inc(AttrsAssigned);
  if FTint <> 0 then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_Color.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Color.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAuto <> False then
    AWriter.AddAttribute('auto',XmlBoolToStr(FAuto));
  if FIndexed > 0 then
    AWriter.AddAttribute('indexed',XmlIntToStr(FIndexed));
  if FRgbAssigned then
    AWriter.AddAttribute('rgb',XmlIntToHexStr(FRgb));
  if FTheme >= 0 then
    AWriter.AddAttribute('theme',XmlIntToStr(FTheme));
  if FTint <> 0 then 
    AWriter.AddAttribute('tint',XmlFloatToStr(FTint));
end;

procedure TCT_Color.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001B9: FAuto := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002E1: FIndexed := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013B: begin
        FRgb := XmlStrToIntDef('$' + AAttributes.Values[i],0);
        FRgbAssigned := True;
      end;
      $00000213: FTheme := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001BF: FTint := XmlStrToFloatDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Color.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
  FIndexed := -1;
  FRgb := -1;
  FTheme := -1;
  FTint := 0;
end;

destructor TCT_Color.Destroy;
begin
end;

procedure TCT_Color.SetRgb(const Value: integer);
begin
  FRgb := Value;
  FRgbAssigned := True;
end;

procedure TCT_Color.Clear;
begin
  FAssigneds := [];
  FAuto := False;
  FIndexed := -1;
  FRgb := -1;
  FRgbAssigned := False;
  FTheme := -1;
  FTint := 0;
end;

{ TCT_ColorXpgList }

function  TCT_ColorXpgList.GetItems(Index: integer): TCT_Color;
begin
  Result := TCT_Color(inherited Items[Index]);
end;

function  TCT_ColorXpgList.Add: TCT_Color;
begin
  Result := TCT_Color.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ColorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ColorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_FontSize }

function  TCT_FontSize.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FontSize.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FontSize.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('val',XmlFloatToStr(FVal));
end;

procedure TCT_FontSize.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := XmlStrToFloatDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FontSize.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_FontSize.Destroy;
begin
end;

procedure TCT_FontSize.Clear;
begin
  FAssigneds := [];
  FVal := 0;
end;

{ TCT_FontSizeXpgList }

function  TCT_FontSizeXpgList.GetItems(Index: integer): TCT_FontSize;
begin
  Result := TCT_FontSize(inherited Items[Index]);
end;

function  TCT_FontSizeXpgList.Add: TCT_FontSize;
begin
  Result := TCT_FontSize.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FontSizeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FontSizeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_UnderlineProperty }

function  TCT_UnderlineProperty.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if not (FVal in [stuvSingle,stuvNone]) then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_UnderlineProperty.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_UnderlineProperty.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> stuvSingle then 
    AWriter.AddAttribute('val',StrTST_UnderlineValues[Integer(FVal)]);
end;

procedure TCT_UnderlineProperty.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_UnderlineValues(StrToEnum('stuv' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_UnderlineProperty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := stuvNone;
end;

procedure TCT_UnderlineProperty.Clear;
begin
  FAssigneds := [];
  FVal := stuvNone;
end;

{ TCT_VerticalAlignFontProperty }

function  TCT_VerticalAlignFontProperty.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_VerticalAlignFontProperty.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_VerticalAlignFontProperty.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> TST_VerticalAlignRun(XPG_UNKNOWN_ENUM) then
    AWriter.AddAttribute('val',StrTST_VerticalAlignRun[Integer(FVal)]);
end;

procedure TCT_VerticalAlignFontProperty.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then
    FVal := TST_VerticalAlignRun(StrToEnum('stvar' + AAttributes.Values[0]))
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_VerticalAlignFontProperty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_VerticalAlignRun(XPG_UNKNOWN_ENUM);
end;

procedure TCT_VerticalAlignFontProperty.Clear;
begin
  FAssigneds := [];
  FVal := TST_VerticalAlignRun(XPG_UNKNOWN_ENUM);
end;

{ TCT_FontScheme }

function  TCT_FontScheme.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_FontScheme.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FontScheme.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FVal) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('val',StrTST_FontScheme[Integer(FVal)]);
end;

procedure TCT_FontScheme.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := TST_FontScheme(StrToEnum('stfs' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FontScheme.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FVal := TST_FontScheme(XPG_UNKNOWN_ENUM);
end;

destructor TCT_FontScheme.Destroy;
begin
end;

procedure TCT_FontScheme.Clear;
begin
  FAssigneds := [];
  FVal := TST_FontScheme(XPG_UNKNOWN_ENUM);
end;

{ TCT_FontSchemeXpgList }

function  TCT_FontSchemeXpgList.GetItems(Index: integer): TCT_FontScheme;
begin
  Result := TCT_FontScheme(inherited Items[Index]);
end;

function  TCT_FontSchemeXpgList.Add: TCT_FontScheme;
begin
  Result := TCT_FontScheme.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FontSchemeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FontSchemeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PivotAreaReference }

function  TCT_PivotAreaReference.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FField <> 0 then 
    Inc(AttrsAssigned);
  if FCount <> 0 then 
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
  Inc(ElemsAssigned,FXXpgList.CheckAssigned);
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
    $00000078: Result := FXXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PivotAreaReference.Write(AWriter: TXpgWriteXML);
begin
  FXXpgList.Write(AWriter,'x');
  if FExtLst.Assigned then 
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
  if FField <> 0 then 
    AWriter.AddAttribute('field',XmlIntToStr(FField));
  if FCount <> 0 then 
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
  begin
    case AAttributes.HashB[i] of
      $58197860: FField := XmlStrToIntDef(AAttributes.Values[i],0);
      $7C8E2A59: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $CF21447D: FSelected := XmlStrToBoolDef(AAttributes.Values[i],False);
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
  end
end;

constructor TCT_PivotAreaReference.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 17;
  FXXpgList := TCT_IndexXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
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
  FXXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_PivotAreaReference.Clear;
begin
  FAssigneds := [];
  FXXpgList.Clear;
  FExtLst.Clear;
  FField := 0;
  FCount := 0;
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

{ TCT_Filter }

function  TCT_Filter.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVal <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Filter.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Filter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> '' then 
    AWriter.AddAttribute('val',FVal);
end;

procedure TCT_Filter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Filter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_Filter.Destroy;
begin
end;

procedure TCT_Filter.Clear;
begin
  FAssigneds := [];
  FVal := '';
end;

{ TCT_FilterXpgList }

function  TCT_FilterXpgList.GetItems(Index: integer): TCT_Filter;
begin
  Result := TCT_Filter(inherited Items[Index]);
end;

function  TCT_FilterXpgList.Add: TCT_Filter;
begin
  Result := TCT_Filter.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FilterXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FilterXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DateGroupItem }

function  TCT_DateGroupItem.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FYear <> 0 then 
    Inc(AttrsAssigned);
  if FMonth <> 0 then 
    Inc(AttrsAssigned);
  if FDay <> 0 then 
    Inc(AttrsAssigned);
  if FHour <> 0 then 
    Inc(AttrsAssigned);
  if FMinute <> 0 then 
    Inc(AttrsAssigned);
  if FSecond <> 0 then 
    Inc(AttrsAssigned);
  if Integer(FDateTimeGrouping) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_DateGroupItem.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DateGroupItem.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('year',XmlIntToStr(FYear));
  if FMonth <> 0 then
    AWriter.AddAttribute('month',XmlIntToStr(FMonth));
  if FDay <> 0 then
    AWriter.AddAttribute('day',XmlIntToStr(FDay));
  if FHour <> 0 then
    AWriter.AddAttribute('hour',XmlIntToStr(FHour));
  if FMinute <> 0 then
    AWriter.AddAttribute('minute',XmlIntToStr(FMinute));
  if FSecond <> 0 then
    AWriter.AddAttribute('second',XmlIntToStr(FSecond));
  if Integer(FDateTimeGrouping) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('dateTimeGrouping',StrTST_DateTimeGrouping[Integer(FDateTimeGrouping)]);
end;

procedure TCT_DateGroupItem.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001B1: FYear := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000226: FMonth := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013E: FDay := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001BE: FHour := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000292: FMinute := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000027C: FSecond := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000678: FDateTimeGrouping := TST_DateTimeGrouping(StrToEnum('stdtg' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DateGroupItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 7;
  FDateTimeGrouping := TST_DateTimeGrouping(XPG_UNKNOWN_ENUM);
end;

destructor TCT_DateGroupItem.Destroy;
begin
end;

procedure TCT_DateGroupItem.Clear;
begin
  FAssigneds := [];
  FYear := 0;
  FMonth := 0;
  FDay := 0;
  FHour := 0;
  FMinute := 0;
  FSecond := 0;
  FDateTimeGrouping := TST_DateTimeGrouping(XPG_UNKNOWN_ENUM);
end;

{ TCT_DateGroupItemXpgList }

function  TCT_DateGroupItemXpgList.GetItems(Index: integer): TCT_DateGroupItem;
begin
  Result := TCT_DateGroupItem(inherited Items[Index]);
end;

function  TCT_DateGroupItemXpgList.Add: TCT_DateGroupItem;
begin
  Result := TCT_DateGroupItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DateGroupItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DateGroupItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CustomFilter }

function  TCT_CustomFilter.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FOperator <> stfoEqual then 
    Inc(AttrsAssigned);
  if FVal <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CustomFilter.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CustomFilter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FOperator <> stfoEqual then 
    AWriter.AddAttribute('operator',StrTST_FilterOperator[Integer(FOperator)]);
  if FVal <> '' then 
    AWriter.AddAttribute('val',FVal);
end;

procedure TCT_CustomFilter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000036C: FOperator := TST_FilterOperator(StrToEnum('stfo' + AAttributes.Values[i]));
      $00000143: FVal := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CustomFilter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FOperator := stfoEqual;
end;

destructor TCT_CustomFilter.Destroy;
begin
end;

procedure TCT_CustomFilter.Clear;
begin
  FAssigneds := [];
  FOperator := stfoEqual;
  FVal := '';
end;

{ TCT_CustomFilterXpgList }

function  TCT_CustomFilterXpgList.GetItems(Index: integer): TCT_CustomFilter;
begin
  Result := TCT_CustomFilter(inherited Items[Index]);
end;

function  TCT_CustomFilterXpgList.Add: TCT_CustomFilter;
begin
  Result := TCT_CustomFilter.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CustomFilterXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CustomFilterXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_RPrElt }

function  TCT_RPrElt.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FRFont.CheckAssigned);
  Inc(ElemsAssigned,FCharset.CheckAssigned);
  Inc(ElemsAssigned,FFamily.CheckAssigned);
  Inc(ElemsAssigned,FB.CheckAssigned);
  Inc(ElemsAssigned,FI.CheckAssigned);
  Inc(ElemsAssigned,FStrike.CheckAssigned);
  Inc(ElemsAssigned,FOutline.CheckAssigned);
  Inc(ElemsAssigned,FShadow.CheckAssigned);
  Inc(ElemsAssigned,FCondense.CheckAssigned);
  Inc(ElemsAssigned,FExtend.CheckAssigned);
  Inc(ElemsAssigned,FColor.CheckAssigned);
  Inc(ElemsAssigned,FSz.CheckAssigned);
  Inc(ElemsAssigned,FU.CheckAssigned);
  Inc(ElemsAssigned,FVertAlign.CheckAssigned);
  Inc(ElemsAssigned,FScheme.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_RPrElt.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000209: begin
      Result := FRFont;
      TCT_FontName(Result).Clear;
    end;
    $000002EA: begin
      Result := FCharset;
      TCT_IntProperty(Result).Clear;
    end;
    $00000282: begin
      Result := FFamily;
      TCT_IntProperty(Result).Clear;
    end;
    $00000062: begin
      Result := FB;
      TCT_BooleanProperty(Result).Clear;
    end;
    $00000069: begin
      Result := FI;
      TCT_BooleanProperty(Result).Clear;
    end;
    $00000292: begin
      Result := FStrike;
      TCT_BooleanProperty(Result).Clear;
    end;
    $00000300: begin
      Result := FOutline;
      TCT_BooleanProperty(Result).Clear;
    end;
    $00000286: begin
      Result := FShadow;
      TCT_BooleanProperty(Result).Clear;
    end;
    $0000034F: begin
      Result := FCondense;
      TCT_BooleanProperty(Result).Clear;
    end;
    $00000288: begin
      Result := FExtend;
      TCT_BooleanProperty(Result).Clear;
    end;
    $0000021F: begin
      Result := FColor;
      TCT_Color(Result).Clear;
    end;
    $000000ED: begin
      Result := FSz;
      TCT_FontSize(Result).Clear;
    end;
    $00000075: begin
      Result := FU;
      TCT_UnderlineProperty(Result).Clear;
      FU.Val := stuvSingle;
    end;
    $000003AC: begin
      Result := FVertAlign;
      TCT_VerticalAlignFontProperty(Result).Clear;
    end;
    $00000275: begin
      Result := FScheme;
      TCT_FontScheme(Result).Clear;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_RPrElt.Write(AWriter: TXpgWriteXML);
begin
  if FRFont.Assigned then begin
    FRFont.WriteAttributes(AWriter);
    AWriter.SimpleTag('rFont');
  end;
  if FCharset.Assigned then begin
    FCharset.WriteAttributes(AWriter);
    AWriter.SimpleTag('charset');
  end;
  if FFamily.Assigned then begin
    FFamily.WriteAttributes(AWriter);
    AWriter.SimpleTag('family');
  end;
  if FB.Assigned then begin
    FB.WriteAttributes(AWriter);
    AWriter.SimpleTag('b');
  end;
  if FI.Assigned then begin
    FI.WriteAttributes(AWriter);
    AWriter.SimpleTag('i');
  end;
  if FStrike.Assigned then begin
    FStrike.WriteAttributes(AWriter);
    AWriter.SimpleTag('strike');
  end;
  if FOutline.Assigned then begin
    FOutline.WriteAttributes(AWriter);
    AWriter.SimpleTag('outline');
  end;
  if FShadow.Assigned then begin
    FShadow.WriteAttributes(AWriter);
    AWriter.SimpleTag('shadow');
  end;
  if FCondense.Assigned then begin
    FCondense.WriteAttributes(AWriter);
    AWriter.SimpleTag('condense');
  end;
  if FExtend.Assigned then begin
    FExtend.WriteAttributes(AWriter);
    AWriter.SimpleTag('extend');
  end;
  if FColor.Assigned then begin
    FColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('color');
  end;
  if FSz.Assigned then begin
    FSz.WriteAttributes(AWriter);
    AWriter.SimpleTag('sz');
  end;

  if FU.Val = stuvSingle then
    AWriter.SimpleTag('u')
  else if FU.Assigned then begin
    FU.WriteAttributes(AWriter);
    AWriter.SimpleTag('u');
  end;

  if FVertAlign.Assigned then begin
    FVertAlign.WriteAttributes(AWriter);
    AWriter.SimpleTag('vertAlign');
  end;
  if FScheme.Assigned then begin
    FScheme.WriteAttributes(AWriter);
    AWriter.SimpleTag('scheme');
  end;
end;

constructor TCT_RPrElt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 15;
  FAttributeCount := 0;
  FRFont := TCT_FontName.Create(FOwner);
  FCharset := TCT_IntProperty.Create(FOwner);
  FFamily := TCT_IntProperty.Create(FOwner);
  FB := TCT_BooleanProperty.Create(FOwner);
  FI := TCT_BooleanProperty.Create(FOwner);
  FStrike := TCT_BooleanProperty.Create(FOwner);
  FOutline := TCT_BooleanProperty.Create(FOwner);
  FShadow := TCT_BooleanProperty.Create(FOwner);
  FCondense := TCT_BooleanProperty.Create(FOwner);
  FExtend := TCT_BooleanProperty.Create(FOwner);
  FColor := TCT_Color.Create(FOwner);
  FSz := TCT_FontSize.Create(FOwner);
  FU := TCT_UnderlineProperty.Create(FOwner);
  FVertAlign := TCT_VerticalAlignFontProperty.Create(FOwner);
  FScheme := TCT_FontScheme.Create(FOwner);
end;

destructor TCT_RPrElt.Destroy;
begin
  FRFont.Free;
  FCharset.Free;
  FFamily.Free;
  FB.Free;
  FI.Free;
  FStrike.Free;
  FOutline.Free;
  FShadow.Free;
  FCondense.Free;
  FExtend.Free;
  FColor.Free;
  FSz.Free;
  FU.Free;
  FVertAlign.Free;
  FScheme.Free;
end;

procedure TCT_RPrElt.Clear;
begin
  FAssigneds := [];
  FRFont.Clear;
  FCharset.Clear;
  FFamily.Clear;
  FB.Clear;
  FI.Clear;
  FStrike.Clear;
  FOutline.Clear;
  FShadow.Clear;
  FCondense.Clear;
  FExtend.Clear;
  FColor.Clear;
  FSz.Clear;
  FU.Clear;
  FVertAlign.Clear;
  FScheme.Clear;
end;

{ TCT_GradientStop }

function  TCT_GradientStop.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 1;
  FAssigneds := [];
  Inc(ElemsAssigned,FColor.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GradientStop.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'color' then 
    Result := FColor
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_GradientStop.Write(AWriter: TXpgWriteXML);
begin
  if FColor.Assigned then 
  begin
    FColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('color');
  end;
end;

procedure TCT_GradientStop.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('position',XmlFloatToStr(FPosition));
end;

procedure TCT_GradientStop.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'position' then 
    FPosition := XmlStrToFloatDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_GradientStop.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FColor := TCT_Color.Create(FOwner);
end;

destructor TCT_GradientStop.Destroy;
begin
  FColor.Free;
end;

procedure TCT_GradientStop.Clear;
begin
  FAssigneds := [];
  FColor.Clear;
  FPosition := 0;
end;

{ TCT_GradientStopXpgList }

function  TCT_GradientStopXpgList.GetItems(Index: integer): TCT_GradientStop;
begin
  Result := TCT_GradientStop(inherited Items[Index]);
end;

function  TCT_GradientStopXpgList.Add: TCT_GradientStop;
begin
  Result := TCT_GradientStop.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_GradientStopXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_GradientStopXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PivotAreaReferences }

function  TCT_PivotAreaReferences.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FReferenceXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotAreaReferences.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'reference' then 
    Result := FReferenceXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PivotAreaReferences.Write(AWriter: TXpgWriteXML);
begin
  FReferenceXpgList.Write(AWriter,'reference');
end;

procedure TCT_PivotAreaReferences.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
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
  FReferenceXpgList := TCT_PivotAreaReferenceXpgList.Create(FOwner);
end;

destructor TCT_PivotAreaReferences.Destroy;
begin
  FReferenceXpgList.Free;
end;

procedure TCT_PivotAreaReferences.Clear;
begin
  FAssigneds := [];
  FReferenceXpgList.Clear;
  FCount := 0;
end;

{ TCT_Filters }

function  TCT_Filters.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBlank <> False then 
    Inc(AttrsAssigned);
  if FCalendarType <> stctNone then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFilterXpgList.CheckAssigned);
  Inc(ElemsAssigned,FDateGroupItemXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Filters.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000286: Result := FFilterXpgList.Add;
    $0000053A: Result := FDateGroupItemXpgList.Add;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Filters.Write(AWriter: TXpgWriteXML);
begin
  FFilterXpgList.Write(AWriter,'filter');
  FDateGroupItemXpgList.Write(AWriter,'dateGroupItem');
end;

procedure TCT_Filters.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBlank <> False then 
    AWriter.AddAttribute('blank',XmlBoolToStr(FBlank));
  if FCalendarType <> stctNone then 
    AWriter.AddAttribute('calendarType',StrTST_CalendarType[Integer(FCalendarType)]);
end;

procedure TCT_Filters.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000208: FBlank := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004DC: FCalendarType := TST_CalendarType(StrToEnum('stct' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Filters.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 2;
  FFilterXpgList := TCT_FilterXpgList.Create(FOwner);
  FDateGroupItemXpgList := TCT_DateGroupItemXpgList.Create(FOwner);
  FBlank := False;
  FCalendarType := stctNone;
end;

destructor TCT_Filters.Destroy;
begin
  FFilterXpgList.Free;
  FDateGroupItemXpgList.Free;
end;

procedure TCT_Filters.Clear;
begin
  FAssigneds := [];
  FFilterXpgList.Clear;
  FDateGroupItemXpgList.Clear;
  FBlank := False;
  FCalendarType := stctNone;
end;

{ TCT_Top10 }

function  TCT_Top10.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FTop <> True then 
    Inc(AttrsAssigned);
  if FPercent <> False then 
    Inc(AttrsAssigned);
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  if FFilterVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Top10.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Top10.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FTop <> True then 
    AWriter.AddAttribute('top',XmlBoolToStr(FTop));
  if FPercent <> False then 
    AWriter.AddAttribute('percent',XmlBoolToStr(FPercent));
  AWriter.AddAttribute('val',XmlFloatToStr(FVal));
  if FFilterVal <> 0 then 
    AWriter.AddAttribute('filterVal',XmlFloatToStr(FFilterVal));
end;

procedure TCT_Top10.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000153: FTop := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F1: FPercent := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000143: FVal := XmlStrToFloatDef(AAttributes.Values[i],0);
      $000003A9: FFilterVal := XmlStrToFloatDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Top10.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FTop := True;
  FPercent := False;
end;

destructor TCT_Top10.Destroy;
begin
end;

procedure TCT_Top10.Clear;
begin
  FAssigneds := [];
  FTop := True;
  FPercent := False;
  FVal := 0;
  FFilterVal := 0;
end;

{ TCT_CustomFilters }

function  TCT_CustomFilters.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAnd <> True then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCustomFilterXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CustomFilters.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'customFilter' then 
    Result := FCustomFilterXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomFilters.Write(AWriter: TXpgWriteXML);
begin
  FCustomFilterXpgList.Write(AWriter,'customFilter');
end;

procedure TCT_CustomFilters.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAnd <> True then
    AWriter.AddAttribute('and',XmlBoolToStr(FAnd));
end;

procedure TCT_CustomFilters.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'and' then
    FAnd := XmlStrToBoolDef(AAttributes.Values[0],True)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CustomFilters.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCustomFilterXpgList := TCT_CustomFilterXpgList.Create(FOwner);
  FAnd := True;
end;

destructor TCT_CustomFilters.Destroy;
begin
  FCustomFilterXpgList.Free;
end;

procedure TCT_CustomFilters.Clear;
begin
  FAssigneds := [];
  FCustomFilterXpgList.Clear;
  FAnd := True;
end;

{ TCT_DynamicFilter }

function  TCT_DynamicFilter.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FType) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FVal <> 0 then 
    Inc(AttrsAssigned);
  if FMaxVal <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_DynamicFilter.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DynamicFilter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('type',StrTST_DynamicFilterType[Integer(FType)]);
  if FVal <> 0 then 
    AWriter.AddAttribute('val',XmlFloatToStr(FVal));
  if FMaxVal <> 0 then 
    AWriter.AddAttribute('maxVal',XmlFloatToStr(FMaxVal));
end;

procedure TCT_DynamicFilter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_DynamicFilterType(StrToEnum('stdft' + AAttributes.Values[i]));
      $00000143: FVal := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000269: FMaxVal := XmlStrToFloatDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DynamicFilter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FType := TST_DynamicFilterType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_DynamicFilter.Destroy;
begin
end;

procedure TCT_DynamicFilter.Clear;
begin
  FAssigneds := [];
  FType := TST_DynamicFilterType(XPG_UNKNOWN_ENUM);
  FVal := 0;
  FMaxVal := 0;
end;

{ TCT_ColorFilter }

function  TCT_ColorFilter.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDxfId <> -1 then  // TODO Zero is significant
    Inc(AttrsAssigned);
  if FCellColor <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ColorFilter.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ColorFilter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDxfId <> -1 then  // TODO Zero is significant
    AWriter.AddAttribute('dxfId',XmlIntToStr(FDxfId));
  if FCellColor <> True then 
    AWriter.AddAttribute('cellColor',XmlBoolToStr(FCellColor));
end;

procedure TCT_ColorFilter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001EF: FDxfId := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000039F: FCellColor := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_ColorFilter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FCellColor := True;
  FDxfId := -1; // TODO Zero is significant
end;

destructor TCT_ColorFilter.Destroy;
begin
end;

procedure TCT_ColorFilter.Clear;
begin
  FAssigneds := [];
  FDxfId := -1; // TODO Zero is significant
  FCellColor := True;
end;

{ TCT_IconFilter }

function  TCT_IconFilter.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FIconSet) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FIconId <> 0 then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_IconFilter.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_IconFilter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FIconSet) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('iconSet',StrTST_IconSetType[Integer(FIconSet)]);
  if FIconId <> 0 then
    AWriter.AddAttribute('iconId',XmlIntToStr(FIconId));
end;

procedure TCT_IconFilter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000002D5: FIconSet := TST_IconSetType(StrToEnum('stist' + AAttributes.Values[i]));
      $00000256: FIconId := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_IconFilter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FIconSet := TST_IconSetType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_IconFilter.Destroy;
begin
end;

procedure TCT_IconFilter.Clear;
begin
  FAssigneds := [];
  FIconSet := TST_IconSetType(XPG_UNKNOWN_ENUM);
  FIconId := 0;
end;

{ TCT_SortCondition }

function  TCT_SortCondition.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDescending <> False then
    Inc(AttrsAssigned);
  if FSortBy <> stsbValue then
    Inc(AttrsAssigned);
  if FRef <> '' then
    Inc(AttrsAssigned);
  if FCustomList <> '' then
    Inc(AttrsAssigned);
  if FDxfId <> -1 then  // TODO Zero is significant
    Inc(AttrsAssigned);
  if FIconSet <> stist3Arrows then 
    Inc(AttrsAssigned);
  if FIconId <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SortCondition.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SortCondition.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDescending <> False then
    AWriter.AddAttribute('descending',XmlBoolToStr(FDescending));
  if FSortBy <> stsbValue then
    AWriter.AddAttribute('sortBy',StrTST_SortBy[Integer(FSortBy)]);
  AWriter.AddAttribute('ref',FRef);
  if FCustomList <> '' then
    AWriter.AddAttribute('customList',FCustomList);
  if FDxfId <> -1 then  // TODO Zero is significant
    AWriter.AddAttribute('dxfId',XmlIntToStr(FDxfId));
  if FIconSet <> stist3Arrows then
    AWriter.AddAttribute('iconSet',StrTST_IconSetType[Integer(FIconSet)]);
  if FIconId <> 0 then
    AWriter.AddAttribute('iconId',XmlIntToStr(FIconId));
end;

procedure TCT_SortCondition.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000414: FDescending := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000283: FSortBy := TST_SortBy(StrToEnum('stsb' + AAttributes.Values[i]));
      $0000013D: FRef := AAttributes.Values[i];
      $00000437: FCustomList := AAttributes.Values[i];
      $000001EF: FDxfId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002D5: FIconSet := TST_IconSetType(StrToEnum('stist' + AAttributes.Values[i]));
      $00000256: FIconId := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SortCondition.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 7;
  FDescending := False;
  FSortBy := stsbValue;
  FIconSet := stist3Arrows;
  FDxfId := -1;  // TODO Zero is significant
end;

destructor TCT_SortCondition.Destroy;
begin
end;

procedure TCT_SortCondition.Clear;
begin
  FAssigneds := [];
  FDescending := False;
  FSortBy := stsbValue;
  FRef := '';
  FCustomList := '';
  FDxfId := -1;  // TODO Zero is significant
  FIconSet := stist3Arrows;
  FIconId := 0;
end;

{ TCT_SortConditionXpgList }

function  TCT_SortConditionXpgList.GetItems(Index: integer): TCT_SortCondition;
begin
  Result := TCT_SortCondition(inherited Items[Index]);
end;

function  TCT_SortConditionXpgList.Add: TCT_SortCondition;
begin
  Result := TCT_SortCondition.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SortConditionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SortConditionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_RElt }

function  TCT_RElt.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FRPr.CheckAssigned);
  if FT <> '' then
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_RElt.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000134: Result := FRPr;
    $00000074: FT := AReader.Text;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_RElt.Write(AWriter: TXpgWriteXML);
begin
  if FRPr.Assigned then
    if xaElements in FRPr.FAssigneds then
    begin
      AWriter.BeginTag('rPr');
      FRPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('rPr');
  if FT <> '' then begin
    if (AnsiChar(FT[1]) in [#10,#13,' ']) or (AnsiChar(FT[Length(FT)]) in [#10,#13,' ']) then
      AWriter.AddAttribute('xml:space','preserve');  // TODO
    AWriter.SimpleTextTag('t',FT);
  end
  else
    AWriter.EmptyTag('t');  // TODO Empty tags must be written
end;

constructor TCT_RElt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FRPr := TCT_RPrElt.Create(FOwner);
end;

destructor TCT_RElt.Destroy;
begin
  FRPr.Free;
end;

procedure TCT_RElt.Clear;
begin
  FAssigneds := [];
  FRPr.Clear;
  FT := '';
end;

{ TCT_REltXpgList }

function  TCT_REltXpgList.GetItems(Index: integer): TCT_RElt;
begin
  Result := TCT_RElt(inherited Items[Index]);
end;

function  TCT_REltXpgList.Add: TCT_RElt;
begin
  Result := TCT_RElt.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_REltXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_REltXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PhoneticRun }

function  TCT_PhoneticRun.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FSb <> 0 then 
    Inc(AttrsAssigned);
  if FEb <> 0 then 
    Inc(AttrsAssigned);
  if FT <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PhoneticRun.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 't' then 
    FT := AReader.Text
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PhoneticRun.Write(AWriter: TXpgWriteXML);
begin
  if FT <> '' then 
    AWriter.SimpleTextTag('t',FT);
end;

procedure TCT_PhoneticRun.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('sb',XmlIntToStr(FSb));
  AWriter.AddAttribute('eb',XmlIntToStr(FEb));
end;

procedure TCT_PhoneticRun.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000000D5: FSb := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000C7: FEb := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PhoneticRun.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
end;

destructor TCT_PhoneticRun.Destroy;
begin
end;

procedure TCT_PhoneticRun.Clear;
begin
  FAssigneds := [];
  FT := '';
  FSb := 0;
  FEb := 0;
end;

{ TCT_PhoneticRunXpgList }

function  TCT_PhoneticRunXpgList.GetItems(Index: integer): TCT_PhoneticRun;
begin
  Result := TCT_PhoneticRun(inherited Items[Index]);
end;

function  TCT_PhoneticRunXpgList.Add: TCT_PhoneticRun;
begin
  Result := TCT_PhoneticRun.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PhoneticRunXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PhoneticRunXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PhoneticPr }

function  TCT_PhoneticPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FFontId <> -1 then
    Inc(AttrsAssigned);
  if FType <> stptFullwidthKatakana then 
    Inc(AttrsAssigned);
  if FAlignment <> stpaLeft then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PhoneticPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PhoneticPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('fontId',XmlIntToStr(FFontId));
  if FType <> stptFullwidthKatakana then 
    AWriter.AddAttribute('type',StrTST_PhoneticType[Integer(FType)]);
  if FAlignment <> stpaLeft then
    AWriter.AddAttribute('alignment',StrTST_PhoneticAlignment[Integer(FAlignment)]);
end;

procedure TCT_PhoneticPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000264: FFontId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C2: FType := TST_PhoneticType(StrToEnum('stpt' + AAttributes.Values[i]));
      $000003BF: FAlignment := TST_PhoneticAlignment(StrToEnum('stpa' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PhoneticPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FType := stptFullwidthKatakana;
  FAlignment := stpaLeft;
  FFontId := -1;
end;

destructor TCT_PhoneticPr.Destroy;
begin
end;

procedure TCT_PhoneticPr.Clear;
begin
  FAssigneds := [];
  FFontId := -1;
  FType := stptFullwidthKatakana;
  FAlignment := stpaLeft;
end;

{ TCT_PatternFill }

function  TCT_PatternFill.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FPatternType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFgColor.CheckAssigned);
  Inc(ElemsAssigned,FBgColor.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PatternFill.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002CC: Result := FFgColor;
    $000002C8: Result := FBgColor;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PatternFill.Write(AWriter: TXpgWriteXML);
begin
  if FFgColor.Assigned then 
  begin
    FFgColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('fgColor');
  end;
  if FBgColor.Assigned then 
  begin
    FBgColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('bgColor');
  end;
end;

procedure TCT_PatternFill.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FPatternType) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('patternType',StrTST_PatternType[Integer(FPatternType)]);
end;

procedure TCT_PatternFill.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'patternType' then 
    FPatternType := TST_PatternType(StrToEnum('stpt' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_PatternFill.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 1;
  FFgColor := TCT_Color.Create(FOwner);
  FBgColor := TCT_Color.Create(FOwner);
  FPatternType := TST_PatternType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PatternFill.Destroy;
begin
  FFgColor.Free;
  FBgColor.Free;
end;

procedure TCT_PatternFill.Clear;
begin
  FAssigneds := [];
  FFgColor.Clear;
  FBgColor.Clear;
  FPatternType := TST_PatternType(XPG_UNKNOWN_ENUM);
end;

{ TCT_GradientFill }

function  TCT_GradientFill.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FType <> stgtLinear then 
    Inc(AttrsAssigned);
  if FDegree <> 0 then 
    Inc(AttrsAssigned);
  if FLeft <> 0 then 
    Inc(AttrsAssigned);
  if FRight <> 0 then 
    Inc(AttrsAssigned);
  if FTop <> 0 then 
    Inc(AttrsAssigned);
  if FBottom <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FStopXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_GradientFill.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'stop' then 
    Result := FStopXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_GradientFill.Write(AWriter: TXpgWriteXML);
begin
  FStopXpgList.Write(AWriter,'stop');
end;

procedure TCT_GradientFill.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FType <> stgtLinear then 
    AWriter.AddAttribute('type',StrTST_GradientType[Integer(FType)]);
  if FDegree <> 0 then 
    AWriter.AddAttribute('degree',XmlFloatToStr(FDegree));
  if FLeft <> 0 then 
    AWriter.AddAttribute('left',XmlFloatToStr(FLeft));
  if FRight <> 0 then 
    AWriter.AddAttribute('right',XmlFloatToStr(FRight));
  if FTop <> 0 then 
    AWriter.AddAttribute('top',XmlFloatToStr(FTop));
  if FBottom <> 0 then 
    AWriter.AddAttribute('bottom',XmlFloatToStr(FBottom));
end;

procedure TCT_GradientFill.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_GradientType(StrToEnum('stgt' + AAttributes.Values[i]));
      $0000026C: FDegree := XmlStrToFloatDef(AAttributes.Values[i],0);
      $000001AB: FLeft := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000021E: FRight := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000153: FTop := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000295: FBottom := XmlStrToFloatDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_GradientFill.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 6;
  FStopXpgList := TCT_GradientStopXpgList.Create(FOwner);
  FType := stgtLinear;
  FDegree := 0;
  FLeft := 0;
  FRight := 0;
  FTop := 0;
  FBottom := 0;
end;

destructor TCT_GradientFill.Destroy;
begin
  FStopXpgList.Free;
end;

procedure TCT_GradientFill.Clear;
begin
  FAssigneds := [];
  FStopXpgList.Clear;
  FType := stgtLinear;
  FDegree := 0;
  FLeft := 0;
  FRight := 0;
  FTop := 0;
  FBottom := 0;
end;

{ TCT_BorderPr }

function  TCT_BorderPr.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FStyle <> stbsNone then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FColor.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_BorderPr.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'color' then 
    Result := FColor
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_BorderPr.Write(AWriter: TXpgWriteXML);
begin
  if FColor.Assigned then 
  begin
    FColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('color');
  end;
end;

procedure TCT_BorderPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FStyle <> stbsNone then 
    AWriter.AddAttribute('style',StrTST_BorderStyle[Integer(FStyle)]);
end;

procedure TCT_BorderPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'style' then 
    FStyle := TST_BorderStyle(StrToEnum('stbs' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_BorderPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FColor := TCT_Color.Create(FOwner);
  FStyle := stbsNone;
end;

destructor TCT_BorderPr.Destroy;
begin
  FColor.Free;
end;

procedure TCT_BorderPr.Clear;
begin
  FAssigneds := [];
  FColor.Clear;
  FStyle := stbsNone;
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
  if FField <> 0 then 
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
  if FFieldPosition <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FReferences.CheckAssigned);
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
    $00000422: Result := FReferences;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PivotArea.Write(AWriter: TXpgWriteXML);
begin
  if FReferences.Assigned then 
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
  if FExtLst.Assigned then 
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
  if FField <> 0 then 
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
  if FFieldPosition <> 0 then 
    AWriter.AddAttribute('fieldPosition',XmlIntToStr(FFieldPosition));
end;

procedure TCT_PivotArea.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000204: FField := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C2: FType := TST_PivotAreaType(StrToEnum('stpat' + AAttributes.Values[i]));
      $0000033C: FDataOnly := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003A2: FLabelOnly := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000344: FGrandRow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000032A: FGrandCol := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003EC: FCacheIndex := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000300: FOutline := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000287: FOffset := AAttributes.Values[i];
      $00000AFB: FCollapsedLevelsAreSubtotals := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001B5: FAxis := TST_Axis(StrToEnum('sta' + AAttributes.Values[i]));
      $00000559: FFieldPosition := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PivotArea.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 12;
  FReferences := TCT_PivotAreaReferences.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FType := stpatNormal;
  FDataOnly := True;
  FLabelOnly := False;
  FGrandRow := False;
  FGrandCol := False;
  FCacheIndex := False;
  FOutline := True;
  FCollapsedLevelsAreSubtotals := False;
  FAxis := TST_Axis(XPG_UNKNOWN_ENUM);
end;

destructor TCT_PivotArea.Destroy;
begin
  FReferences.Free;
  FExtLst.Free;
end;

procedure TCT_PivotArea.Clear;
begin
  FAssigneds := [];
  FReferences.Clear;
  FExtLst.Clear;
  FField := 0;
  FType := stpatNormal;
  FDataOnly := True;
  FLabelOnly := False;
  FGrandRow := False;
  FGrandCol := False;
  FCacheIndex := False;
  FOutline := True;
  FOffset := '';
  FCollapsedLevelsAreSubtotals := False;
  FAxis := TST_Axis(XPG_UNKNOWN_ENUM);
  FFieldPosition := 0;
end;

{ TCT_Break }

function  TCT_Break.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FId <> 0 then 
    Inc(AttrsAssigned);
  if FMin <> 0 then 
    Inc(AttrsAssigned);
  if FMax <> 0 then 
    Inc(AttrsAssigned);
  if FMan <> False then 
    Inc(AttrsAssigned);
  if FPt <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Break.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Break.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FId <> 0 then 
    AWriter.AddAttribute('id',XmlIntToStr(FId));
  if FMin <> 0 then 
    AWriter.AddAttribute('min',XmlIntToStr(FMin));
  if FMax <> 0 then 
    AWriter.AddAttribute('max',XmlIntToStr(FMax));
  if FMan <> False then 
    AWriter.AddAttribute('man',XmlBoolToStr(FMan));
  if FPt <> False then 
    AWriter.AddAttribute('pt',XmlBoolToStr(FPt));
end;

procedure TCT_Break.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000144: FMin := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000146: FMax := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000013C: FMan := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000E4: FPt := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Break.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
  FId := 0;
  FMin := 0;
  FMax := 0;
  FMan := False;
  FPt := False;
end;

destructor TCT_Break.Destroy;
begin
end;

procedure TCT_Break.Clear;
begin
  FAssigneds := [];
  FId := 0;
  FMin := 0;
  FMax := 0;
  FMan := False;
  FPt := False;
end;

{ TCT_BreakXpgList }

function  TCT_BreakXpgList.GetItems(Index: integer): TCT_Break;
begin
  Result := TCT_Break(inherited Items[Index]);
end;

function  TCT_BreakXpgList.Add: TCT_Break;
begin
  Result := TCT_Break.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BreakXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BreakXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_FilterColumn }

function  TCT_FilterColumn.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FColId <> 0 then 
    Inc(AttrsAssigned);
  if FHiddenButton <> False then 
    Inc(AttrsAssigned);
  if FShowButton <> True then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFilters.CheckAssigned);
  Inc(ElemsAssigned,FTop10.CheckAssigned);
  Inc(ElemsAssigned,FCustomFilters.CheckAssigned);
  Inc(ElemsAssigned,FDynamicFilter.CheckAssigned);
  Inc(ElemsAssigned,FColorFilter.CheckAssigned);
  Inc(ElemsAssigned,FIconFilter.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_FilterColumn.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002F9: Result := FFilters;
    $000001B4: Result := FTop10;
    $00000574: Result := FCustomFilters;
    $0000054B: Result := FDynamicFilter;
    $00000485: Result := FColorFilter;
    $0000040F: Result := FIconFilter;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_FilterColumn.Write(AWriter: TXpgWriteXML);
begin
  if FFilters.Assigned then 
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
  if FTop10.Assigned then 
  begin
    FTop10.WriteAttributes(AWriter);
    AWriter.SimpleTag('top10');
  end;
  if FCustomFilters.Assigned then 
  begin
    FCustomFilters.WriteAttributes(AWriter);
    if xaElements in FCustomFilters.FAssigneds then 
    begin
      AWriter.BeginTag('customFilters');
      FCustomFilters.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customFilters');
  end;
  if FDynamicFilter.Assigned then 
  begin
    FDynamicFilter.WriteAttributes(AWriter);
    AWriter.SimpleTag('dynamicFilter');
  end;
  if FColorFilter.Assigned then 
  begin
    FColorFilter.WriteAttributes(AWriter);
    AWriter.SimpleTag('colorFilter');
  end;
  if FIconFilter.Assigned then 
  begin
    FIconFilter.WriteAttributes(AWriter);
    AWriter.SimpleTag('iconFilter');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_FilterColumn.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('colId',XmlIntToStr(FColId));
  if FHiddenButton <> False then 
    AWriter.AddAttribute('hiddenButton',XmlBoolToStr(FHiddenButton));
  if FShowButton <> True then 
    AWriter.AddAttribute('showButton',XmlBoolToStr(FShowButton));
end;

procedure TCT_FilterColumn.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001EB: FColId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004E8: FHiddenButton := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000043D: FShowButton := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_FilterColumn.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 3;
  FFilters := TCT_Filters.Create(FOwner);
  FTop10 := TCT_Top10.Create(FOwner);
  FCustomFilters := TCT_CustomFilters.Create(FOwner);
  FDynamicFilter := TCT_DynamicFilter.Create(FOwner);
  FColorFilter := TCT_ColorFilter.Create(FOwner);
  FIconFilter := TCT_IconFilter.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FHiddenButton := False;
  FShowButton := True;
end;

destructor TCT_FilterColumn.Destroy;
begin
  FFilters.Free;
  FTop10.Free;
  FCustomFilters.Free;
  FDynamicFilter.Free;
  FColorFilter.Free;
  FIconFilter.Free;
  FExtLst.Free;
end;

procedure TCT_FilterColumn.Clear;
begin
  FAssigneds := [];
  FFilters.Clear;
  FTop10.Clear;
  FCustomFilters.Clear;
  FDynamicFilter.Clear;
  FColorFilter.Clear;
  FIconFilter.Clear;
  FExtLst.Clear;
  FColId := 0;
  FHiddenButton := False;
  FShowButton := True;
end;

{ TCT_FilterColumnXpgList }

function  TCT_FilterColumnXpgList.GetItems(Index: integer): TCT_FilterColumn;
begin
  Result := TCT_FilterColumn(inherited Items[Index]);
end;

function  TCT_FilterColumnXpgList.Add: TCT_FilterColumn;
begin
  Result := TCT_FilterColumn.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FilterColumnXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FilterColumnXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
//    if xaAttributes in Items[i].FAssigneds then
      GetItems(i).WriteAttributes(AWriter);      // TODO Element has required attribute that must be written
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

{ TCT_SortState }

function  TCT_SortState.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FColumnSort <> False then 
    Inc(AttrsAssigned);
  if FCaseSensitive <> False then 
    Inc(AttrsAssigned);
  if FSortMethod <> stsmNone then 
    Inc(AttrsAssigned);
  if FRef <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FSortConditionXpgList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_SortState.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000056F: Result := FSortConditionXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SortState.Write(AWriter: TXpgWriteXML);
begin
  FSortConditionXpgList.Write(AWriter,'sortCondition');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_SortState.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FColumnSort <> False then
    AWriter.AddAttribute('columnSort',XmlBoolToStr(FColumnSort));
  if FCaseSensitive <> False then
    AWriter.AddAttribute('caseSensitive',XmlBoolToStr(FCaseSensitive));
  if FSortMethod <> stsmNone then
    AWriter.AddAttribute('sortMethod',StrTST_SortMethod[Integer(FSortMethod)]);
  AWriter.AddAttribute('ref',FRef);
end;

procedure TCT_SortState.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000436: FColumnSort := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000556: FCaseSensitive := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000429: FSortMethod := TST_SortMethod(StrToEnum('stsm' + AAttributes.Values[i]));
      $0000013D: FRef := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SortState.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 4;
  FSortConditionXpgList := TCT_SortConditionXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FColumnSort := False;
  FCaseSensitive := False;
  FSortMethod := stsmNone;
end;

destructor TCT_SortState.Destroy;
begin
  FSortConditionXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_SortState.Clear;
begin
  FAssigneds := [];
  FSortConditionXpgList.Clear;
  FExtLst.Clear;
  FColumnSort := False;
  FCaseSensitive := False;
  FSortMethod := stsmNone;
  FRef := '';
end;

{ TCT_CellFormula }

function  TCT_CellFormula.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FT <> stcftNormal then
    Inc(AttrsAssigned);
  if FAca <> False then
    Inc(AttrsAssigned);
  if FRef <> '' then
    Inc(AttrsAssigned);
  if FDt2D <> False then
    Inc(AttrsAssigned);
  if FDtr <> False then
    Inc(AttrsAssigned);
  if FDel1 <> False then
    Inc(AttrsAssigned);
  if FDel2 <> False then
    Inc(AttrsAssigned);
  if FR1 <> '' then
    Inc(AttrsAssigned);
  if FR2 <> '' then
    Inc(AttrsAssigned);
  if FCa <> False then
    Inc(AttrsAssigned);
  if FSi <> $FFFF then  // TODO
    Inc(AttrsAssigned);
  if FBx <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
  if FContent <> '' then 
  begin
    FAssigneds := FAssigneds + [xaContent];
    Inc(Result);
  end;
end;

function  TCT_CellFormula.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellFormula.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CellFormula.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FT <> stcftNormal then
    AWriter.AddAttribute('t',StrTST_CellFormulaType[Integer(FT)]);
  if FAca <> False then
    AWriter.AddAttribute('aca',XmlBoolToStr(FAca));
  if FRef <> '' then
    AWriter.AddAttribute('ref',FRef);
  if FDt2D <> False then
    AWriter.AddAttribute('dt2D',XmlBoolToStr(FDt2D));
  if FDtr <> False then
    AWriter.AddAttribute('dtr',XmlBoolToStr(FDtr));
  if FDel1 <> False then
    AWriter.AddAttribute('del1',XmlBoolToStr(FDel1));
  if FDel2 <> False then
    AWriter.AddAttribute('del2',XmlBoolToStr(FDel2));
  if FR1 <> '' then
    AWriter.AddAttribute('r1',FR1);
  if FR2 <> '' then
    AWriter.AddAttribute('r2',FR2);
  if FCa <> False then
    AWriter.AddAttribute('ca',XmlBoolToStr(FCa));
  if FSi <> $FFFF then // TODO
    AWriter.AddAttribute('si',XmlIntToStr(FSi));
  if FBx <> False then 
    AWriter.AddAttribute('bx',XmlBoolToStr(FBx));
end;

procedure TCT_CellFormula.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000074: FT := TST_CellFormulaType(StrToEnum('stcft' + AAttributes.Values[i]));
      $00000125: FAca := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000013D: FRef := AAttributes.Values[i];
      $0000014E: FDt2D := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000014A: FDtr := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000166: FDel1 := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000167: FDel2 := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000A3: FR1 := AAttributes.Values[i];
      $000000A4: FR2 := AAttributes.Values[i];
      $000000C4: FCa := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000DC: FSi := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000DA: FBx := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CellFormula.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 12;
  FT := stcftNormal;
  FAca := False;
  FDt2D := False;
  FDtr := False;
  FDel1 := False;
  FDel2 := False;
  FCa := False;
  FSi := $FFFF; //TODO
  FBx := False;
end;

destructor TCT_CellFormula.Destroy;
begin
end;

procedure TCT_CellFormula.Clear;
begin
  FAssigneds := [];
  FT := stcftNormal;
  FAca := False;
  FRef := '';
  FDt2D := False;
  FDtr := False;
  FDel1 := False;
  FDel2 := False;
  FR1 := '';
  FR2 := '';
  FCa := False;
  FSi := $FFFF; //TODO
  FBx := False;
end;

{ TCT_Rst }

function  TCT_Rst.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if _FT <> '' then
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FRXpgList.CheckAssigned);
  Inc(ElemsAssigned,FRPhXpgList.CheckAssigned);
  Inc(ElemsAssigned,FPhoneticPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Rst.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000074: begin
      _FT := AReader.Text;

      if not ((AReader.Attributes.Count > 0) and (AReader.Attributes.Find('xml:space') >= 0)) then
        _FT := Trim(_FT);
    end;
    $00000072: Result := FRXpgList.Add;
    $0000012A: Result := FRPhXpgList.Add;
    $0000041C: Result := FPhoneticPr;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Rst.Write(AWriter: TXpgWriteXML);
begin
  if _FT <> '' then begin
    if (AnsiChar(_FT[1]) in [#10,#13,' ']) or (AnsiChar(_FT[Length(_FT)]) in [#10,#13,' ']) then
      AWriter.AddAttribute('xml:space','preserve'); // TODO. Check also more possible whitespace chars.
    AWriter.SimpleTextTag('t',_FT);
  end
  else
    AWriter.EmptyTag('t');
  FRXpgList.Write(AWriter,'r');
  FRPhXpgList.Write(AWriter,'rPh');
  if FPhoneticPr.Assigned then
  begin
    FPhoneticPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('phoneticPr');
  end;
end;

constructor TCT_Rst.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FRXpgList := TCT_REltXpgList.Create(FOwner);
  FRPhXpgList := TCT_PhoneticRunXpgList.Create(FOwner);
  FPhoneticPr := TCT_PhoneticPr.Create(FOwner);
end;

destructor TCT_Rst.Destroy;
begin
  FRXpgList.Free;
  FRPhXpgList.Free;
  FPhoneticPr.Free;
end;

procedure TCT_Rst.Clear;
begin
  FAssigneds := [];
  _FT := '';
  FRXpgList.Clear;
  FRPhXpgList.Clear;
  FPhoneticPr.Clear;
end;

{ TCT_RstXpgList }

function  TCT_RstXpgList.GetItems(Index: integer): TCT_Rst;
begin
  Result := TCT_Rst(inherited Items[Index]);
end;

function  TCT_RstXpgList.Add: TCT_Rst;
begin
  Result := TCT_Rst.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RstXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RstXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Cfvo }

function  TCT_Cfvo.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FVal <> '' then
    Inc(AttrsAssigned);
  if FGte <> True then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Cfvo.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'extLst' then
    Result := FExtLst
  else
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Cfvo.Write(AWriter: TXpgWriteXML);
begin
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Cfvo.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('type',StrTST_CfvoType[Integer(FType)]);
  if FVal <> '' then 
    AWriter.AddAttribute('val',FVal);
  if FGte <> True then 
    AWriter.AddAttribute('gte',XmlBoolToStr(FGte));
end;

procedure TCT_Cfvo.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_CfvoType(StrToEnum('stct' + AAttributes.Values[i]));
      $00000143: FVal := AAttributes.Values[i];
      $00000140: FGte := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Cfvo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FType := TST_CfvoType(XPG_UNKNOWN_ENUM);
  FGte := True;
end;

destructor TCT_Cfvo.Destroy;
begin
  FExtLst.Free;
end;

procedure TCT_Cfvo.Clear;
begin
  FAssigneds := [];
  FExtLst.Clear;
  FType := TST_CfvoType(XPG_UNKNOWN_ENUM);
  FVal := '';
  FGte := True;
end;

{ TCT_CfvoXpgList }

function  TCT_CfvoXpgList.GetItems(Index: integer): TCT_Cfvo;
begin
  Result := TCT_Cfvo(inherited Items[Index]);
end;

function  TCT_CfvoXpgList.Add: TCT_Cfvo;
begin
  Result := TCT_Cfvo.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CfvoXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CfvoXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CellSmartTagPr }

function  TCT_CellSmartTagPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FKey <> '' then 
    Inc(AttrsAssigned);
  if FVal <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CellSmartTagPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CellSmartTagPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('key',FKey);
  AWriter.AddAttribute('val',FVal);
end;

procedure TCT_CellSmartTagPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000149: FKey := AAttributes.Values[i];
      $00000143: FVal := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CellSmartTagPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_CellSmartTagPr.Destroy;
begin
end;

procedure TCT_CellSmartTagPr.Clear;
begin
  FAssigneds := [];
  FKey := '';
  FVal := '';
end;

{ TCT_CellSmartTagPrXpgList }

function  TCT_CellSmartTagPrXpgList.GetItems(Index: integer): TCT_CellSmartTagPr;
begin
  Result := TCT_CellSmartTagPr(inherited Items[Index]);
end;

function  TCT_CellSmartTagPrXpgList.Add: TCT_CellSmartTagPr;
begin
  Result := TCT_CellSmartTagPr.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CellSmartTagPrXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CellSmartTagPrXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CellAlignment }

function  TCT_CellAlignment.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FHorizontal) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if Integer(FVertical) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FTextRotation <> 0 then 
    Inc(AttrsAssigned);
  if FWrapText <> False then 
    Inc(AttrsAssigned);
  if FIndent <> 0 then 
    Inc(AttrsAssigned);
  if FRelativeIndent <> 0 then 
    Inc(AttrsAssigned);
  if FJustifyLastLine <> False then 
    Inc(AttrsAssigned);
  if FShrinkToFit <> False then 
    Inc(AttrsAssigned);
  if FReadingOrder <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CellAlignment.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CellAlignment.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FHorizontal) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('horizontal',StrTST_HorizontalAlignment[Integer(FHorizontal)]);
  if Integer(FVertical) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('vertical',StrTST_VerticalAlignment[Integer(FVertical)]);
  if FTextRotation <> 0 then 
    AWriter.AddAttribute('textRotation',XmlIntToStr(FTextRotation));
  if FWrapText <> False then 
    AWriter.AddAttribute('wrapText',XmlBoolToStr(FWrapText));
  if FIndent <> 0 then 
    AWriter.AddAttribute('indent',XmlIntToStr(FIndent));
  if FRelativeIndent <> 0 then 
    AWriter.AddAttribute('relativeIndent',XmlIntToStr(FRelativeIndent));
  if FJustifyLastLine <> False then 
    AWriter.AddAttribute('justifyLastLine',XmlBoolToStr(FJustifyLastLine));
  if FShrinkToFit <> False then 
    AWriter.AddAttribute('shrinkToFit',XmlBoolToStr(FShrinkToFit));
  if FReadingOrder <> 0 then 
    AWriter.AddAttribute('readingOrder',XmlIntToStr(FReadingOrder));
end;

procedure TCT_CellAlignment.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000044A: FHorizontal := TST_HorizontalAlignment(StrToEnum('stha' + AAttributes.Values[i]));
      $0000035A: FVertical := TST_VerticalAlignment(StrToEnum('stva' + AAttributes.Values[i]));
      $00000515: FTextRotation := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000035F: FWrapText := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000282: FIndent := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005BE: FRelativeIndent := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000062A: FJustifyLastLine := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000475: FShrinkToFit := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004D6: FReadingOrder := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CellAlignment.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 9;
  FHorizontal := TST_HorizontalAlignment(XPG_UNKNOWN_ENUM);
  FVertical := TST_VerticalAlignment(XPG_UNKNOWN_ENUM);
end;

destructor TCT_CellAlignment.Destroy;
begin
end;

procedure TCT_CellAlignment.Clear;
begin
  FAssigneds := [];
  FHorizontal := TST_HorizontalAlignment(XPG_UNKNOWN_ENUM);
  FVertical := TST_VerticalAlignment(XPG_UNKNOWN_ENUM);
  FTextRotation := 0;
  FWrapText := False;
  FIndent := 0;
  FRelativeIndent := 0;
  FJustifyLastLine := False;
  FShrinkToFit := False;
  FReadingOrder := 0;
end;

{ TCT_CellProtection }

function  TCT_CellProtection.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FLocked <> True then
    Inc(AttrsAssigned);
  if FHidden <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CellProtection.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CellProtection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FLocked <> True then
    AWriter.AddAttribute('locked',XmlBoolToStr(FLocked));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
end;

procedure TCT_CellProtection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000272: FLocked := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CellProtection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_CellProtection.Destroy;
begin
end;

procedure TCT_CellProtection.Clear;
begin
  FAssigneds := [];
  FLocked := True;
  FHidden := False;
end;

{ TCT_Font }

function  TCT_Font.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FName.CheckAssigned);
  Inc(ElemsAssigned,FCharset.CheckAssigned);
  Inc(ElemsAssigned,FFamily.CheckAssigned);
  Inc(ElemsAssigned,FB.CheckAssigned);
  Inc(ElemsAssigned,FI.CheckAssigned);
  Inc(ElemsAssigned,FStrike.CheckAssigned);
  Inc(ElemsAssigned,FOutline.CheckAssigned);
  Inc(ElemsAssigned,FShadow.CheckAssigned);
  Inc(ElemsAssigned,FCondense.CheckAssigned);
  Inc(ElemsAssigned,FExtend.CheckAssigned);
  Inc(ElemsAssigned,FColor.CheckAssigned);
  Inc(ElemsAssigned,FSz.CheckAssigned);
  Inc(ElemsAssigned,FU.CheckAssigned);
  Inc(ElemsAssigned,FVertAlign.CheckAssigned);
  Inc(ElemsAssigned,FScheme.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Font.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001A1: begin
      Result := FName;
      TCT_FontName(Result).Clear;
    end;
    $000002EA: begin
      Result := FCharset;
      TCT_IntProperty(Result).Clear;
    end;
    $00000282: begin
      Result := FFamily;
      TCT_IntProperty(Result).Clear;
    end;
    $00000062: begin
      Result := FB;
      TCT_BooleanProperty(Result).Clear;
      FB.Val := True;
    end;
    $00000069: begin
      Result := FI;
      TCT_BooleanProperty(Result).Clear;
      FI.Val := True;
    end;
    $00000292: begin
      Result := FStrike;
      TCT_BooleanProperty(Result).Clear;
      FStrike.Val := True;
    end;
    $00000300: begin
      Result := FOutline;
      TCT_BooleanProperty(Result).Clear;
      FOutline.Val := True;
    end;
    $00000286: begin
      Result := FShadow;
      TCT_BooleanProperty(Result).Clear;
      FShadow.Val := True;
    end;
    $0000034F: begin
      Result := FCondense;
      TCT_BooleanProperty(Result).Clear;
      FCondense.Val := True;
    end;
    $00000288: begin
      Result := FExtend;
      TCT_BooleanProperty(Result).Clear;
      FExtend.Val := True;
    end;
    $0000021F: begin
      Result := FColor;
      TCT_Color(Result).Clear;
    end;
    $000000ED: begin
      Result := FSz;
      TCT_FontSize(Result).Clear;
    end;
    $00000075: begin
      Result := FU;
      TCT_UnderlineProperty(Result).Clear;
      FU.Val := stuvSingle;
    end;
    $000003AC: begin
      Result := FVertAlign;
      TCT_VerticalAlignFontProperty(Result).Clear;
    end;
    $00000275: begin
      Result := FScheme;
      TCT_FontScheme(Result).Clear;
    end;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Font.Write(AWriter: TXpgWriteXML);
begin
  if FName.Assigned then begin
    FName.WriteAttributes(AWriter);
    AWriter.SimpleTag('name');
  end;
  if FCharset.Assigned then begin
    FCharset.WriteAttributes(AWriter);
    AWriter.SimpleTag('charset');
  end;
  if FFamily.Assigned then begin
    FFamily.WriteAttributes(AWriter);
    AWriter.SimpleTag('family');
  end;
  if FB.Assigned then begin
    FB.WriteAttributes(AWriter);
    AWriter.SimpleTag('b');
  end;
  if FI.Assigned then begin
    FI.WriteAttributes(AWriter);
    AWriter.SimpleTag('i');
  end;
  if FStrike.Assigned then begin
    FStrike.WriteAttributes(AWriter);
    AWriter.SimpleTag('strike');
  end;
  if FOutline.Assigned then begin
    FOutline.WriteAttributes(AWriter);
    AWriter.SimpleTag('outline');
  end;
  if FShadow.Assigned then begin
    FShadow.WriteAttributes(AWriter);
    AWriter.SimpleTag('shadow');
  end;
  if FCondense.Assigned then begin
    FCondense.WriteAttributes(AWriter);
    AWriter.SimpleTag('condense');
  end;
  if FExtend.Assigned then begin
    FExtend.WriteAttributes(AWriter);
    AWriter.SimpleTag('extend');
  end;
  if FColor.Assigned then begin
    FColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('color');
  end;
  if FSz.Assigned then begin
    FSz.WriteAttributes(AWriter);
    AWriter.SimpleTag('sz');
  end;

  if FU.Val = stuvSingle then
    AWriter.SimpleTag('u')
  else if FU.Assigned then begin
    FU.WriteAttributes(AWriter);
    AWriter.SimpleTag('u');
  end;

  if FVertAlign.Assigned then begin
    FVertAlign.WriteAttributes(AWriter);
    AWriter.SimpleTag('vertAlign');
  end;
  if FScheme.Assigned then begin
    FScheme.WriteAttributes(AWriter);
    AWriter.SimpleTag('scheme');
  end;
end;

constructor TCT_Font.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 15;
  FAttributeCount := 0;
  FName := TCT_FontName.Create(FOwner);
  FCharset := TCT_IntProperty.Create(FOwner);
  FFamily := TCT_IntProperty.Create(FOwner);
  FB := TCT_BooleanProperty.Create(FOwner);
  FI := TCT_BooleanProperty.Create(FOwner);
  FStrike := TCT_BooleanProperty.Create(FOwner);
  FOutline := TCT_BooleanProperty.Create(FOwner);
  FShadow := TCT_BooleanProperty.Create(FOwner);
  FCondense := TCT_BooleanProperty.Create(FOwner);
  FExtend := TCT_BooleanProperty.Create(FOwner);
  FColor := TCT_Color.Create(FOwner);
  FSz := TCT_FontSize.Create(FOwner);
  FU := TCT_UnderlineProperty.Create(FOwner);
  FVertAlign := TCT_VerticalAlignFontProperty.Create(FOwner);
  FScheme := TCT_FontScheme.Create(FOwner);
end;

destructor TCT_Font.Destroy;
begin
  FName.Free;
  FCharset.Free;
  FFamily.Free;
  FB.Free;
  FI.Free;
  FStrike.Free;
  FOutline.Free;
  FShadow.Free;
  FCondense.Free;
  FExtend.Free;
  FColor.Free;
  FSz.Free;
  FU.Free;
  FVertAlign.Free;
  FScheme.Free;
end;

procedure TCT_Font.Clear;
begin
  FAssigneds := [];
  FName.Clear;
  FCharset.Clear;
  FFamily.Clear;
  FB.Clear;
  FI.Clear;
  FStrike.Clear;
  FOutline.Clear;
  FShadow.Clear;
  FCondense.Clear;
  FExtend.Clear;
  FColor.Clear;
  FSz.Clear;
  FU.Clear;
  FVertAlign.Clear;
  FScheme.Clear;
end;

{ TCT_FontXpgList }

function  TCT_FontXpgList.GetItems(Index: integer): TCT_Font;
begin
  Result := TCT_Font(inherited Items[Index]);
end;

function  TCT_FontXpgList.Add: TCT_Font;
begin
  Result := TCT_Font.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FontXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FontXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_NumFmt }

function  TCT_NumFmt.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNumFmtId <> 0 then 
    Inc(AttrsAssigned);
  if FFormatCode <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_NumFmt.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_NumFmt.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('numFmtId',XmlIntToStr(FNumFmtId));
  AWriter.AddAttribute('formatCode',FFormatCode);
end;

procedure TCT_NumFmt.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000324: FNumFmtId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000404: FFormatCode := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_NumFmt.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_NumFmt.Destroy;
begin
end;

procedure TCT_NumFmt.Clear;
begin
  FAssigneds := [];
  FNumFmtId := 0;
  FFormatCode := '';
end;

{ TCT_NumFmtXpgList }

function  TCT_NumFmtXpgList.GetItems(Index: integer): TCT_NumFmt;
begin
  Result := TCT_NumFmt(inherited Items[Index]);
end;

function  TCT_NumFmtXpgList.Add: TCT_NumFmt;
begin
  Result := TCT_NumFmt.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_NumFmtXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_NumFmtXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Fill }

function  TCT_Fill.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FPatternFill.CheckAssigned);
  Inc(ElemsAssigned,FGradientFill.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Fill.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000485: Result := FPatternFill;
    $000004D5: Result := FGradientFill;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Fill.Write(AWriter: TXpgWriteXML);
begin
  if FPatternFill.Assigned then 
  begin
    FPatternFill.WriteAttributes(AWriter);
    if xaElements in FPatternFill.FAssigneds then 
    begin
      AWriter.BeginTag('patternFill');
      FPatternFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('patternFill');
  end;
  if FGradientFill.Assigned then 
  begin
    FGradientFill.WriteAttributes(AWriter);
    if xaElements in FGradientFill.FAssigneds then 
    begin
      AWriter.BeginTag('gradientFill');
      FGradientFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('gradientFill');
  end;
end;

constructor TCT_Fill.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FPatternFill := TCT_PatternFill.Create(FOwner);
  FGradientFill := TCT_GradientFill.Create(FOwner);
end;

destructor TCT_Fill.Destroy;
begin
  FPatternFill.Free;
  FGradientFill.Free;
end;

procedure TCT_Fill.Clear;
begin
  FAssigneds := [];
  FPatternFill.Clear;
  FGradientFill.Clear;
end;

{ TCT_FillXpgList }

function  TCT_FillXpgList.GetItems(Index: integer): TCT_Fill;
begin
  Result := TCT_Fill(inherited Items[Index]);
end;

function  TCT_FillXpgList.Add: TCT_Fill;
begin
  Result := TCT_Fill.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FillXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FillXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Border }

function  TCT_Border.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDiagonalUp <> False then
    Inc(AttrsAssigned);
  if FDiagonalDown <> False then 
    Inc(AttrsAssigned);
  if FOutline <> True then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FLeft.CheckAssigned);
  Inc(ElemsAssigned,FRight.CheckAssigned);
  Inc(ElemsAssigned,FTop.CheckAssigned);
  Inc(ElemsAssigned,FBottom.CheckAssigned);
  Inc(ElemsAssigned,FDiagonal.CheckAssigned);
  Inc(ElemsAssigned,FVertical.CheckAssigned);
  Inc(ElemsAssigned,FHorizontal.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Border.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001AB: Result := FLeft;
    $0000021E: Result := FRight;
    $00000153: Result := FTop;
    $00000295: Result := FBottom;
    $0000033F: Result := FDiagonal;
    $0000035A: Result := FVertical;
    $0000044A: Result := FHorizontal;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Border.Write(AWriter: TXpgWriteXML);
begin
  if FLeft.Assigned then 
  begin
    FLeft.WriteAttributes(AWriter);
    if xaElements in FLeft.FAssigneds then 
    begin
      AWriter.BeginTag('left');
      FLeft.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('left');
  end;
  if FRight.Assigned then 
  begin
    FRight.WriteAttributes(AWriter);
    if xaElements in FRight.FAssigneds then 
    begin
      AWriter.BeginTag('right');
      FRight.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('right');
  end;
  if FTop.Assigned then 
  begin
    FTop.WriteAttributes(AWriter);
    if xaElements in FTop.FAssigneds then 
    begin
      AWriter.BeginTag('top');
      FTop.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('top');
  end;
  if FBottom.Assigned then 
  begin
    FBottom.WriteAttributes(AWriter);
    if xaElements in FBottom.FAssigneds then 
    begin
      AWriter.BeginTag('bottom');
      FBottom.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('bottom');
  end;
  if FDiagonal.Assigned then 
  begin
    FDiagonal.WriteAttributes(AWriter);
    if xaElements in FDiagonal.FAssigneds then 
    begin
      AWriter.BeginTag('diagonal');
      FDiagonal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('diagonal');
  end;
  if FVertical.Assigned then 
  begin
    FVertical.WriteAttributes(AWriter);
    if xaElements in FVertical.FAssigneds then 
    begin
      AWriter.BeginTag('vertical');
      FVertical.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('vertical');
  end;
  if FHorizontal.Assigned then 
  begin
    FHorizontal.WriteAttributes(AWriter);
    if xaElements in FHorizontal.FAssigneds then 
    begin
      AWriter.BeginTag('horizontal');
      FHorizontal.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('horizontal');
  end;
end;

procedure TCT_Border.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDiagonalUp <> False then 
    AWriter.AddAttribute('diagonalUp',XmlBoolToStr(FDiagonalUp));
  if FDiagonalDown <> False then 
    AWriter.AddAttribute('diagonalDown',XmlBoolToStr(FDiagonalDown));
  if FOutline <> True then 
    AWriter.AddAttribute('outline',XmlBoolToStr(FOutline));
end;

procedure TCT_Border.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000404: FDiagonalUp := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004D7: FDiagonalDown := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000300: FOutline := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Border.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 3;
  FLeft := TCT_BorderPr.Create(FOwner);
  FRight := TCT_BorderPr.Create(FOwner);
  FTop := TCT_BorderPr.Create(FOwner);
  FBottom := TCT_BorderPr.Create(FOwner);
  FDiagonal := TCT_BorderPr.Create(FOwner);
  FVertical := TCT_BorderPr.Create(FOwner);
  FHorizontal := TCT_BorderPr.Create(FOwner);
  FOutline := True;
end;

destructor TCT_Border.Destroy;
begin
  FLeft.Free;
  FRight.Free;
  FTop.Free;
  FBottom.Free;
  FDiagonal.Free;
  FVertical.Free;
  FHorizontal.Free;
end;

procedure TCT_Border.Clear;
begin
  FAssigneds := [];
  FLeft.Clear;
  FRight.Clear;
  FTop.Clear;
  FBottom.Clear;
  FDiagonal.Clear;
  FVertical.Clear;
  FHorizontal.Clear;
  FDiagonalUp := False;
  FDiagonalDown := False;
  FOutline := True;
end;

{ TCT_BorderXpgList }

function  TCT_BorderXpgList.GetItems(Index: integer): TCT_Border;
begin
  Result := TCT_Border(inherited Items[Index]);
end;

function  TCT_BorderXpgList.Add: TCT_Border;
begin
  Result := TCT_Border.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BorderXpgList.CheckAssigned: integer;
var
  i: integer;
begin
//  Result := 0;    // TODO At least one has to be written.
  Result := Count;  // Consider just checking Count, and if it's one or more, write the items.

  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BorderXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_TableStyleElement }

function  TCT_TableStyleElement.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FSize <> 1 then
    Inc(AttrsAssigned);
  if FDxfId <> 0 then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_TableStyleElement.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TableStyleElement.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('type',StrTST_TableStyleType[Integer(FType)]);
  if FSize <> 1 then 
    AWriter.AddAttribute('size',XmlIntToStr(FSize));
  if FDxfId <> 0 then 
    AWriter.AddAttribute('dxfId',XmlIntToStr(FDxfId));
end;

procedure TCT_TableStyleElement.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_TableStyleType(StrToEnum('sttst' + AAttributes.Values[i]));
      $000001BB: FSize := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001EF: FDxfId := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_TableStyleElement.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FType := TST_TableStyleType(XPG_UNKNOWN_ENUM);
  FSize := 1;
end;

destructor TCT_TableStyleElement.Destroy;
begin
end;

procedure TCT_TableStyleElement.Clear;
begin
  FAssigneds := [];
  FType := TST_TableStyleType(XPG_UNKNOWN_ENUM);
  FSize := 1;
  FDxfId := 0;
end;

{ TCT_TableStyleElementXpgList }

function  TCT_TableStyleElementXpgList.GetItems(Index: integer): TCT_TableStyleElement;
begin
  Result := TCT_TableStyleElement(inherited Items[Index]);
end;

function  TCT_TableStyleElementXpgList.Add: TCT_TableStyleElement;
begin
  Result := TCT_TableStyleElement.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TableStyleElementXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TableStyleElementXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_RgbColor }

function  TCT_RgbColor.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRgb <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_RgbColor.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_RgbColor.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('rgb',XmlIntToHexStr(FRgb));
end;

procedure TCT_RgbColor.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'rgb' then 
    FRgb := XmlStrToIntDef('$' + AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_RgbColor.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_RgbColor.Destroy;
begin
end;

procedure TCT_RgbColor.Clear;
begin
  FAssigneds := [];
  FRgb := 0;
end;

{ TCT_RgbColorXpgList }

function  TCT_RgbColorXpgList.GetItems(Index: integer): TCT_RgbColor;
begin
  Result := TCT_RgbColor(inherited Items[Index]);
end;

function  TCT_RgbColorXpgList.Add: TCT_RgbColor;
begin
  Result := TCT_RgbColor.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RgbColorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RgbColorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
//    if xaAttributes in Items[i].FAssigneds then
    GetItems(i).WriteAttributes(AWriter);
    if xaElements in Items[i].FAssigneds then begin
      AWriter.BeginTag(AName);
      GetItems(i).Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag(AName);
  end
end;

{ TCT_PageMargins }

function  TCT_PageMargins.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FLeft <> 0 then 
    Inc(AttrsAssigned);
  if FRight <> 0 then 
    Inc(AttrsAssigned);
  if FTop <> 0 then 
    Inc(AttrsAssigned);
  if FBottom <> 0 then 
    Inc(AttrsAssigned);
  if FHeader <> 0 then 
    Inc(AttrsAssigned);
  if FFooter <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PageMargins.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PageMargins.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('left',XmlFloatToStr(FLeft));
  AWriter.AddAttribute('right',XmlFloatToStr(FRight));
  AWriter.AddAttribute('top',XmlFloatToStr(FTop));
  AWriter.AddAttribute('bottom',XmlFloatToStr(FBottom));
  AWriter.AddAttribute('header',XmlFloatToStr(FHeader));
  AWriter.AddAttribute('footer',XmlFloatToStr(FFooter));
end;

procedure TCT_PageMargins.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001AB: FLeft := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000021E: FRight := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000153: FTop := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000295: FBottom := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000269: FHeader := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000028F: FFooter := XmlStrToFloatDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PageMargins.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
end;

destructor TCT_PageMargins.Destroy;
begin
end;

procedure TCT_PageMargins.Clear;
begin
  FAssigneds := [];
  FLeft := 0;
  FRight := 0;
  FTop := 0;
  FBottom := 0;
  FHeader := 0;
  FFooter := 0;
end;

{ TCT_CsPageSetup }

function  TCT_CsPageSetup.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPaperSize <> 1 then 
    Inc(AttrsAssigned);
  if FFirstPageNumber <> 1 then 
    Inc(AttrsAssigned);
  if FOrientation <> stoDefault then 
    Inc(AttrsAssigned);
  if FUsePrinterDefaults <> True then 
    Inc(AttrsAssigned);
  if FBlackAndWhite <> False then 
    Inc(AttrsAssigned);
  if FDraft <> False then 
    Inc(AttrsAssigned);
  if FUseFirstPageNumber <> False then 
    Inc(AttrsAssigned);
  if FHorizontalDpi <> 600 then 
    Inc(AttrsAssigned);
  if FVerticalDpi <> 600 then 
    Inc(AttrsAssigned);
  if FCopies <> 1 then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CsPageSetup.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CsPageSetup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPaperSize <> 1 then 
    AWriter.AddAttribute('paperSize',XmlIntToStr(FPaperSize));
  if FFirstPageNumber <> 1 then 
    AWriter.AddAttribute('firstPageNumber',XmlIntToStr(FFirstPageNumber));
  if FOrientation <> stoDefault then 
    AWriter.AddAttribute('orientation',StrTST_Orientation[Integer(FOrientation)]);
  if FUsePrinterDefaults <> True then 
    AWriter.AddAttribute('usePrinterDefaults',XmlBoolToStr(FUsePrinterDefaults));
  if FBlackAndWhite <> False then 
    AWriter.AddAttribute('blackAndWhite',XmlBoolToStr(FBlackAndWhite));
  if FDraft <> False then 
    AWriter.AddAttribute('draft',XmlBoolToStr(FDraft));
  if FUseFirstPageNumber <> False then 
    AWriter.AddAttribute('useFirstPageNumber',XmlBoolToStr(FUseFirstPageNumber));
  if FHorizontalDpi <> 600 then 
    AWriter.AddAttribute('horizontalDpi',XmlIntToStr(FHorizontalDpi));
  if FVerticalDpi <> 600 then 
    AWriter.AddAttribute('verticalDpi',XmlIntToStr(FVerticalDpi));
  if FCopies <> 1 then 
    AWriter.AddAttribute('copies',XmlIntToStr(FCopies));
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_CsPageSetup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000003B3: FPaperSize := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000060E: FFirstPageNumber := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004AC: FOrientation := TST_Orientation(StrToEnum('sto' + AAttributes.Values[i]));
      $00000769: FUsePrinterDefaults := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000511: FBlackAndWhite := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000211: FDraft := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000073B: FUseFirstPageNumber := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000567: FHorizontalDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000477: FVerticalDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000283: FCopies := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CsPageSetup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 11;
  FPaperSize := 1;
  FFirstPageNumber := 1;
  FOrientation := stoDefault;
  FUsePrinterDefaults := True;
  FBlackAndWhite := False;
  FDraft := False;
  FUseFirstPageNumber := False;
  FHorizontalDpi := 600;
  FVerticalDpi := 600;
  FCopies := 1;
end;

destructor TCT_CsPageSetup.Destroy;
begin
end;

procedure TCT_CsPageSetup.Clear;
begin
  FAssigneds := [];
  FPaperSize := 1;
  FFirstPageNumber := 1;
  FOrientation := stoDefault;
  FUsePrinterDefaults := True;
  FBlackAndWhite := False;
  FDraft := False;
  FUseFirstPageNumber := False;
  FHorizontalDpi := 600;
  FVerticalDpi := 600;
  FCopies := 1;
  FR_Id := '';
end;

{ TCT_HeaderFooter }

function  TCT_HeaderFooter.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDifferentOddEven <> False then 
    Inc(AttrsAssigned);
  if FDifferentFirst <> False then 
    Inc(AttrsAssigned);
  if FScaleWithDoc <> True then 
    Inc(AttrsAssigned);
  if FAlignWithMargins <> True then 
    Inc(AttrsAssigned);
  if FOddHeader <> '' then 
    Inc(ElemsAssigned);
  if FOddFooter <> '' then 
    Inc(ElemsAssigned);
  if FEvenHeader <> '' then 
    Inc(ElemsAssigned);
  if FEvenFooter <> '' then 
    Inc(ElemsAssigned);
  if FFirstHeader <> '' then 
    Inc(ElemsAssigned);
  if FFirstFooter <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_HeaderFooter.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000380: FOddHeader := AReader.Text;
    $000003A6: FOddFooter := AReader.Text;
    $000003F7: FEvenHeader := AReader.Text;
    $0000041D: FEvenFooter := AReader.Text;
    $00000471: FFirstHeader := AReader.Text;
    $00000497: FFirstFooter := AReader.Text;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_HeaderFooter.Write(AWriter: TXpgWriteXML);
begin
  if FOddHeader <> '' then 
    AWriter.SimpleTextTag('oddHeader',FOddHeader);
  if FOddFooter <> '' then 
    AWriter.SimpleTextTag('oddFooter',FOddFooter);
  if FEvenHeader <> '' then 
    AWriter.SimpleTextTag('evenHeader',FEvenHeader);
  if FEvenFooter <> '' then 
    AWriter.SimpleTextTag('evenFooter',FEvenFooter);
  if FFirstHeader <> '' then 
    AWriter.SimpleTextTag('firstHeader',FFirstHeader);
  if FFirstFooter <> '' then 
    AWriter.SimpleTextTag('firstFooter',FFirstFooter);
end;

procedure TCT_HeaderFooter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDifferentOddEven <> False then 
    AWriter.AddAttribute('differentOddEven',XmlBoolToStr(FDifferentOddEven));
  if FDifferentFirst <> False then 
    AWriter.AddAttribute('differentFirst',XmlBoolToStr(FDifferentFirst));
  if FScaleWithDoc <> True then 
    AWriter.AddAttribute('scaleWithDoc',XmlBoolToStr(FScaleWithDoc));
  if FAlignWithMargins <> True then 
    AWriter.AddAttribute('alignWithMargins',XmlBoolToStr(FAlignWithMargins));
end;

procedure TCT_HeaderFooter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000065C: FDifferentOddEven := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005BF: FDifferentFirst := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004BA: FScaleWithDoc := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000678: FAlignWithMargins := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_HeaderFooter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 6;
  FAttributeCount := 4;
  FDifferentOddEven := False;
  FDifferentFirst := False;
  FScaleWithDoc := True;
  FAlignWithMargins := True;
end;

destructor TCT_HeaderFooter.Destroy;
begin
end;

procedure TCT_HeaderFooter.Clear;
begin
  FAssigneds := [];
  FOddHeader := '';
  FOddFooter := '';
  FEvenHeader := '';
  FEvenFooter := '';
  FFirstHeader := '';
  FFirstFooter := '';
  FDifferentOddEven := False;
  FDifferentFirst := False;
  FScaleWithDoc := True;
  FAlignWithMargins := True;
end;

{ TCT_Pane }

function  TCT_Pane.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FXSplit <> 0 then
    Inc(AttrsAssigned);
  if FYSplit <> 0 then
    Inc(AttrsAssigned);
  if FTopLeftCell <> '' then
    Inc(AttrsAssigned);
  if FActivePane <> stpTopLeft then
    Inc(AttrsAssigned);
  if FState <> stpsSplit then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Pane.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Pane.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FXSplit <> 0 then 
    AWriter.AddAttribute('xSplit',XmlFloatToStr(FXSplit));
  if FYSplit <> 0 then 
    AWriter.AddAttribute('ySplit',XmlFloatToStr(FYSplit));
  if FTopLeftCell <> '' then 
    AWriter.AddAttribute('topLeftCell',FTopLeftCell);
  if FActivePane <> stpTopLeft then
    AWriter.AddAttribute('activePane',StrTST_Pane[Integer(FActivePane)]);
  if FState <> stpsSplit then
    AWriter.AddAttribute('state',StrTST_PaneState[Integer(FState)]);
end;

procedure TCT_Pane.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000284: FXSplit := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000285: FYSplit := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000045E: FTopLeftCell := AAttributes.Values[i];
      $00000400: FActivePane := TST_Pane(StrToEnum('stp' + AAttributes.Values[i]));
      $00000221: FState := TST_PaneState(StrToEnum('stps' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Pane.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
  FXSplit := 0;
  FYSplit := 0;
  FActivePane := stpTopLeft;
  FState := stpsSplit;
end;

destructor TCT_Pane.Destroy;
begin
end;

procedure TCT_Pane.Clear;
begin
  FAssigneds := [];
  FXSplit := 0;
  FYSplit := 0;
  FTopLeftCell := '';
  FActivePane := stpTopLeft;
  FState := stpsSplit;
end;

{ TCT_Selection }

function  TCT_Selection.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPane <> stpTopLeft then 
    Inc(AttrsAssigned);
  if FActiveCell <> '' then 
    Inc(AttrsAssigned);
  if FActiveCellId <> 0 then 
    Inc(AttrsAssigned);
  if FSqrefXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Selection.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Selection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPane <> stpTopLeft then 
    AWriter.AddAttribute('pane',StrTST_Pane[Integer(FPane)]);
  if FActiveCell <> '' then 
    AWriter.AddAttribute('activeCell',FActiveCell);
  if FActiveCellId <> 0 then 
    AWriter.AddAttribute('activeCellId',XmlIntToStr(FActiveCellId));
  if FSqrefXpgList.Count > 0 then 
    AWriter.AddAttribute('sqref',FSqrefXpgList.DelimitedText);
end;

procedure TCT_Selection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A4: FPane := TST_Pane(StrToEnum('stp' + AAttributes.Values[i]));
      $000003FC: FActiveCell := AAttributes.Values[i];
      $000004A9: FActiveCellId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000221: FSqrefXpgList.DelimitedText := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Selection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FPane := stpTopLeft;
  FActiveCellId := 0;
  FSqrefXpgList := TStringXpgList.Create;
  FSqrefXpgList.DelimitedText := 'A1';
end;

destructor TCT_Selection.Destroy;
begin
  FSqrefXpgList.Free;
end;

procedure TCT_Selection.Clear;
begin
  FAssigneds := [];
  FPane := stpTopLeft;
  FActiveCell := '';
  FActiveCellId := 0;
  FSqrefXpgList.DelimitedText := 'A1';
end;

{ TCT_SelectionXpgList }

function  TCT_SelectionXpgList.GetItems(Index: integer): TCT_Selection;
begin
  Result := TCT_Selection(inherited Items[Index]);
end;

function  TCT_SelectionXpgList.Add: TCT_Selection;
begin
  Result := TCT_Selection.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SelectionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SelectionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PivotSelection }

function  TCT_PivotSelection.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPane <> stpTopLeft then 
    Inc(AttrsAssigned);
  if FShowHeader <> False then 
    Inc(AttrsAssigned);
  if FLabel <> False then 
    Inc(AttrsAssigned);
  if FData <> False then 
    Inc(AttrsAssigned);
  if FExtendable <> False then 
    Inc(AttrsAssigned);
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if Integer(FAxis) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FDimension <> 0 then 
    Inc(AttrsAssigned);
  if FStart <> 0 then 
    Inc(AttrsAssigned);
  if FMin <> 0 then 
    Inc(AttrsAssigned);
  if FMax <> 0 then 
    Inc(AttrsAssigned);
  if FActiveRow <> 0 then 
    Inc(AttrsAssigned);
  if FActiveCol <> 0 then 
    Inc(AttrsAssigned);
  if FPreviousRow <> 0 then 
    Inc(AttrsAssigned);
  if FPreviousCol <> 0 then 
    Inc(AttrsAssigned);
  if FClick <> 0 then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FPivotArea.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PivotSelection.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'pivotArea' then 
    Result := FPivotArea
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PivotSelection.Write(AWriter: TXpgWriteXML);
begin
  if FPivotArea.Assigned then 
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

procedure TCT_PivotSelection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPane <> stpTopLeft then 
    AWriter.AddAttribute('pane',StrTST_Pane[Integer(FPane)]);
  if FShowHeader <> False then 
    AWriter.AddAttribute('showHeader',XmlBoolToStr(FShowHeader));
  if FLabel <> False then 
    AWriter.AddAttribute('label',XmlBoolToStr(FLabel));
  if FData <> False then 
    AWriter.AddAttribute('data',XmlBoolToStr(FData));
  if FExtendable <> False then 
    AWriter.AddAttribute('extendable',XmlBoolToStr(FExtendable));
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if Integer(FAxis) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('axis',StrTST_Axis[Integer(FAxis)]);
  if FDimension <> 0 then 
    AWriter.AddAttribute('dimension',XmlIntToStr(FDimension));
  if FStart <> 0 then 
    AWriter.AddAttribute('start',XmlIntToStr(FStart));
  if FMin <> 0 then 
    AWriter.AddAttribute('min',XmlIntToStr(FMin));
  if FMax <> 0 then 
    AWriter.AddAttribute('max',XmlIntToStr(FMax));
  if FActiveRow <> 0 then 
    AWriter.AddAttribute('activeRow',XmlIntToStr(FActiveRow));
  if FActiveCol <> 0 then 
    AWriter.AddAttribute('activeCol',XmlIntToStr(FActiveCol));
  if FPreviousRow <> 0 then 
    AWriter.AddAttribute('previousRow',XmlIntToStr(FPreviousRow));
  if FPreviousCol <> 0 then 
    AWriter.AddAttribute('previousCol',XmlIntToStr(FPreviousCol));
  if FClick <> 0 then 
    AWriter.AddAttribute('click',XmlIntToStr(FClick));
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_PivotSelection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A4: FPane := TST_Pane(StrToEnum('stp' + AAttributes.Values[i]));
      $0000040A: FShowHeader := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000200: FLabel := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000019A: FData := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000041C: FExtendable := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001B5: FAxis := TST_Axis(StrToEnum('sta' + AAttributes.Values[i]));
      $000003C6: FDimension := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000022E: FStart := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000144: FMin := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000146: FMax := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003B4: FActiveRow := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000039A: FActiveCol := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004B5: FPreviousRow := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000049B: FPreviousCol := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000206: FClick := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PivotSelection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 17;
  FPivotArea := TCT_PivotArea.Create(FOwner);
  FPane := TST_Pane(XPG_UNKNOWN_ENUM);
  FPane := stpTopLeft;
  FShowHeader := False;
  FLabel := False;
  FData := False;
  FExtendable := False;
  FCount := 0;
  FAxis := TST_Axis(XPG_UNKNOWN_ENUM);
  FDimension := 0;
  FStart := 0;
  FMin := 0;
  FMax := 0;
  FActiveRow := 0;
  FActiveCol := 0;
  FPreviousRow := 0;
  FPreviousCol := 0;
  FClick := 0;
end;

destructor TCT_PivotSelection.Destroy;
begin
  FPivotArea.Free;
end;

procedure TCT_PivotSelection.Clear;
begin
  FAssigneds := [];
  FPivotArea.Clear;
  FPane := stpTopLeft;
  FShowHeader := False;
  FLabel := False;
  FData := False;
  FExtendable := False;
  FCount := 0;
  FAxis := TST_Axis(XPG_UNKNOWN_ENUM);
  FDimension := 0;
  FStart := 0;
  FMin := 0;
  FMax := 0;
  FActiveRow := 0;
  FActiveCol := 0;
  FPreviousRow := 0;
  FPreviousCol := 0;
  FClick := 0;
  FR_Id := '';
end;

{ TCT_PivotSelectionXpgList }

function  TCT_PivotSelectionXpgList.GetItems(Index: integer): TCT_PivotSelection;
begin
  Result := TCT_PivotSelection(inherited Items[Index]);
end;

function  TCT_PivotSelectionXpgList.Add: TCT_PivotSelection;
begin
  Result := TCT_PivotSelection.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotSelectionXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotSelectionXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PageBreak }

function  TCT_PageBreak.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBrkXpgList.Count <> 0 then
    Inc(AttrsAssigned);
  if FManualBreakCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FBrkXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_PageBreak.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'brk' then 
    Result := FBrkXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PageBreak.Write(AWriter: TXpgWriteXML);
begin
  FBrkXpgList.Write(AWriter,'brk');
end;

procedure TCT_PageBreak.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('count',XmlIntToStr(FBrkXpgList.Count));
  if FManualBreakCount <> 0 then 
    AWriter.AddAttribute('manualBreakCount',XmlIntToStr(FManualBreakCount));
end;

procedure TCT_PageBreak.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000229: ; // Count
      $0000066C: FManualBreakCount := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PageBreak.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FBrkXpgList := TCT_BreakXpgList.Create(FOwner);
  FManualBreakCount := 0;
end;

destructor TCT_PageBreak.Destroy;
begin
  FBrkXpgList.Free;
end;

procedure TCT_PageBreak.Clear;
begin
  FAssigneds := [];
  FBrkXpgList.Clear;
  FManualBreakCount := 0;
end;

{ TCT_PrintOptions }

function  TCT_PrintOptions.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FHorizontalCentered <> False then 
    Inc(AttrsAssigned);
  if FVerticalCentered <> False then 
    Inc(AttrsAssigned);
  if FHeadings <> False then 
    Inc(AttrsAssigned);
  if FGridLines <> False then 
    Inc(AttrsAssigned);
  if FGridLinesSet <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PrintOptions.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PrintOptions.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FHorizontalCentered <> False then 
    AWriter.AddAttribute('horizontalCentered',XmlBoolToStr(FHorizontalCentered));
  if FVerticalCentered <> False then 
    AWriter.AddAttribute('verticalCentered',XmlBoolToStr(FVerticalCentered));
  if FHeadings <> False then 
    AWriter.AddAttribute('headings',XmlBoolToStr(FHeadings));
  if FGridLines <> False then 
    AWriter.AddAttribute('gridLines',XmlBoolToStr(FGridLines));
  if FGridLinesSet <> True then 
    AWriter.AddAttribute('gridLinesSet',XmlBoolToStr(FGridLinesSet));
end;

procedure TCT_PrintOptions.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000774: FHorizontalCentered := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000684: FVerticalCentered := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000343: FHeadings := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003A1: FGridLines := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004CD: FGridLinesSet := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PrintOptions.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
  FHorizontalCentered := False;
  FVerticalCentered := False;
  FHeadings := False;
  FGridLines := False;
  FGridLinesSet := True;
end;

destructor TCT_PrintOptions.Destroy;
begin
end;

procedure TCT_PrintOptions.Clear;
begin
  FAssigneds := [];
  FHorizontalCentered := False;
  FVerticalCentered := False;
  FHeadings := False;
  FGridLines := False;
  FGridLinesSet := True;
end;

{ TCT_PageSetup }

function  TCT_PageSetup.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPaperSize <> 1 then 
    Inc(AttrsAssigned);
  if FScale <> 100 then 
    Inc(AttrsAssigned);
  if FFirstPageNumber <> 1 then 
    Inc(AttrsAssigned);
  if FFitToWidth <> 1 then 
    Inc(AttrsAssigned);
  if FFitToHeight <> 1 then 
    Inc(AttrsAssigned);
  if FPageOrder <> stpoDownThenOver then 
    Inc(AttrsAssigned);
  if FOrientation <> stoDefault then 
    Inc(AttrsAssigned);
  if FUsePrinterDefaults <> True then 
    Inc(AttrsAssigned);
  if FBlackAndWhite <> False then 
    Inc(AttrsAssigned);
  if FDraft <> False then 
    Inc(AttrsAssigned);
  if FCellComments <> stccNone then 
    Inc(AttrsAssigned);
  if FUseFirstPageNumber <> False then 
    Inc(AttrsAssigned);
  if FErrors <> stpeDisplayed then 
    Inc(AttrsAssigned);
  if FHorizontalDpi <> 600 then 
    Inc(AttrsAssigned);
  if FVerticalDpi <> 600 then 
    Inc(AttrsAssigned);
  if FCopies <> 1 then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PageSetup.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PageSetup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPaperSize <> 1 then 
    AWriter.AddAttribute('paperSize',XmlIntToStr(FPaperSize));
  if FScale <> 100 then 
    AWriter.AddAttribute('scale',XmlIntToStr(FScale));
  if FFirstPageNumber <> 1 then 
    AWriter.AddAttribute('firstPageNumber',XmlIntToStr(FFirstPageNumber));
  if FFitToWidth <> 1 then 
    AWriter.AddAttribute('fitToWidth',XmlIntToStr(FFitToWidth));
  if FFitToHeight <> 1 then 
    AWriter.AddAttribute('fitToHeight',XmlIntToStr(FFitToHeight));
  if FPageOrder <> stpoDownThenOver then 
    AWriter.AddAttribute('pageOrder',StrTST_PageOrder[Integer(FPageOrder)]);
  if FOrientation <> stoDefault then 
    AWriter.AddAttribute('orientation',StrTST_Orientation[Integer(FOrientation)]);
  if FUsePrinterDefaults <> True then 
    AWriter.AddAttribute('usePrinterDefaults',XmlBoolToStr(FUsePrinterDefaults));
  if FBlackAndWhite <> False then 
    AWriter.AddAttribute('blackAndWhite',XmlBoolToStr(FBlackAndWhite));
  if FDraft <> False then 
    AWriter.AddAttribute('draft',XmlBoolToStr(FDraft));
  if FCellComments <> stccNone then 
    AWriter.AddAttribute('cellComments',StrTST_CellComments[Integer(FCellComments)]);
  if FUseFirstPageNumber <> False then 
    AWriter.AddAttribute('useFirstPageNumber',XmlBoolToStr(FUseFirstPageNumber));
  if FErrors <> stpeDisplayed then 
    AWriter.AddAttribute('errors',StrTST_PrintError[Integer(FErrors)]);
  if FHorizontalDpi <> 600 then 
    AWriter.AddAttribute('horizontalDpi',XmlIntToStr(FHorizontalDpi));
  if FVerticalDpi <> 600 then 
    AWriter.AddAttribute('verticalDpi',XmlIntToStr(FVerticalDpi));
  if FCopies <> 1 then 
    AWriter.AddAttribute('copies',XmlIntToStr(FCopies));
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_PageSetup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000003B3: FPaperSize := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000208: FScale := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000060E: FFirstPageNumber := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000406: FFitToWidth := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000045F: FFitToHeight := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000399: FPageOrder := TST_PageOrder(StrToEnum('stpo' + AAttributes.Values[i]));
      $000004AC: FOrientation := TST_Orientation(StrToEnum('sto' + AAttributes.Values[i]));
      $00000769: FUsePrinterDefaults := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000511: FBlackAndWhite := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000211: FDraft := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004E6: FCellComments := TST_CellComments(StrToEnum('stcc' + AAttributes.Values[i]));
      $0000073B: FUseFirstPageNumber := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000029D: FErrors := TST_PrintError(StrToEnum('stpe' + AAttributes.Values[i]));
      $00000567: FHorizontalDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000477: FVerticalDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000283: FCopies := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PageSetup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 17;
  FPaperSize := 1;
  FScale := 100;
  FFirstPageNumber := 1;
  FFitToWidth := 1;
  FFitToHeight := 1;
  FPageOrder := stpoDownThenOver;
  FOrientation := stoDefault;
  FUsePrinterDefaults := True;
  FBlackAndWhite := False;
  FDraft := False;
  FCellComments := stccNone;
  FUseFirstPageNumber := False;
  FErrors := stpeDisplayed;
  FHorizontalDpi := 600;
  FVerticalDpi := 600;
  FCopies := 1;
end;

destructor TCT_PageSetup.Destroy;
begin
end;

procedure TCT_PageSetup.Clear;
begin
  FAssigneds := [];
  FPaperSize := 1;
  FScale := 100;
  FFirstPageNumber := 1;
  FFitToWidth := 1;
  FFitToHeight := 1;
  FPageOrder := stpoDownThenOver;
  FOrientation := stoDefault;
  FUsePrinterDefaults := True;
  FBlackAndWhite := False;
  FDraft := False;
  FCellComments := stccNone;
  FUseFirstPageNumber := False;
  FErrors := stpeDisplayed;
  FHorizontalDpi := 600;
  FVerticalDpi := 600;
  FCopies := 1;
  FR_Id := '';
end;

{ TCT_AutoFilter }

function  TCT_AutoFilter.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFilterColumnXpgList.CheckAssigned);
  Inc(ElemsAssigned,FSortState.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_AutoFilter.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004F4: Result := FFilterColumnXpgList.Add;
    $000003C9: Result := FSortState;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_AutoFilter.Write(AWriter: TXpgWriteXML);
begin
  FFilterColumnXpgList.Write(AWriter,'filterColumn');
  if FSortState.Assigned then 
  begin
    FSortState.WriteAttributes(AWriter);
    if xaElements in FSortState.FAssigneds then 
    begin
      AWriter.BeginTag('sortState');
      FSortState.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sortState');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_AutoFilter.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRef <> '' then 
    AWriter.AddAttribute('ref',FRef);
end;

procedure TCT_AutoFilter.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'ref' then 
    FRef := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_AutoFilter.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 1;
  FFilterColumnXpgList := TCT_FilterColumnXpgList.Create(FOwner);
  FSortState := TCT_SortState.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_AutoFilter.Destroy;
begin
  FFilterColumnXpgList.Free;
  FSortState.Free;
  FExtLst.Free;
end;

procedure TCT_AutoFilter.Clear;
begin
  FAssigneds := [];
  FFilterColumnXpgList.Clear;
  FSortState.Clear;
  FExtLst.Clear;
  FRef := '';
end;

{ TCT_Cell }

function  TCT_Cell.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> '' then 
    Inc(AttrsAssigned);
  if FS <> 0 then 
    Inc(AttrsAssigned);
  if FT <> stctN then 
    Inc(AttrsAssigned);
  if FCm <> 0 then 
    Inc(AttrsAssigned);
  if FVm <> 0 then 
    Inc(AttrsAssigned);
  if FPh <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FF.CheckAssigned);
  if FV <> '' then 
    Inc(ElemsAssigned);
  Inc(ElemsAssigned,FIs.CheckAssigned);
//  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Cell.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QNameIsV then
    FV := AReader.Text
  else begin
    case AReader.QNameHashA of
      $00000066: begin
        Result := FF;
        if AReader.HasText then
          TCT_CellFormula(Result).Content := AReader.Text;
      end;
      $00000076: FV := AReader.Text;
      $000000DC: Result := FIs;
  //    $00000284: Result := FExtLst;
      else
        FOwner.Errors.Error(xemUnknownElement,AReader.QName);
    end;
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Cell.Write(AWriter: TXpgWriteXML);
begin
  if FF.Assigned then 
  begin
    AWriter.Text := FF.Content;
    FF.WriteAttributes(AWriter);
    AWriter.SimpleTag('f');
  end;
  if FV <> '' then 
    AWriter.SimpleTextTag('v',FV);
  if FIs.Assigned then 
    if xaElements in FIs.FAssigneds then 
    begin
      AWriter.BeginTag('is');
      FIs.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('is');
//  if FExtLst.Assigned then
//    if xaElements in FExtLst.FAssigneds then
//    begin
//      AWriter.BeginTag('extLst');
//      FExtLst.Write(AWriter);
//      AWriter.EndTag;
//    end
//    else
//      AWriter.SimpleTag('extLst');
end;

procedure TCT_Cell.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FR <> '' then 
    AWriter.AddAttribute('r',FR);
  if FS <> 0 then 
    AWriter.AddAttribute('s',XmlIntToStr(FS));
  if FT <> stctN then 
    AWriter.AddAttribute('t',StrTST_CellType[Integer(FT)]);
  if FCm <> 0 then 
    AWriter.AddAttribute('cm',XmlIntToStr(FCm));
  if FVm <> 0 then 
    AWriter.AddAttribute('vm',XmlIntToStr(FVm));
  if FPh <> False then 
    AWriter.AddAttribute('ph',XmlBoolToStr(FPh));
end;

procedure TCT_Cell.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000072: FR := AAttributes.Values[i];
      $00000073: FS := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000074: FT := TST_CellType(StrToEnum('stct' + AAttributes.Values[i]));
      $000000D0: FCm := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000E3: FVm := XmlStrToIntDef(AAttributes.Values[i],0);
      $000000D8: FPh := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Cell.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 6;
  FF := TCT_CellFormula.Create(FOwner);
  FIs := TCT_Rst.Create(FOwner);
//  FExtLst := TCT_ExtensionList.Create(FOwner);
  FS := 0;
  FT := stctN;
  FCm := 0;
  FVm := 0;
  FPh := False;
end;

destructor TCT_Cell.Destroy;
begin
  FF.Free;
  FIs.Free;
//  FExtLst.Free;
end;

procedure TCT_Cell.Clear;
begin
  FAssigneds := [];
  FF.Content := '';
  FF.Clear;
  FV := '';
  FIs.Clear;
//  FExtLst.Clear;
  FR := '';
  FS := 0;
  FT := stctN;
  FCm := 0;
  FVm := 0;
  FPh := False;
end;

{ TCT_CellXpgList }

function  TCT_CellXpgList.GetItems(Index: integer): TCT_Cell;
begin
  Result := TCT_Cell(inherited Items[Index]);
end;

function  TCT_CellXpgList.Add: TCT_Cell;
begin
  Result := TCT_Cell.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CellXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CellXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_InputCells }

function  TCT_InputCells.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> '' then 
    Inc(AttrsAssigned);
  if FDeleted <> False then 
    Inc(AttrsAssigned);
  if FUndone <> False then 
    Inc(AttrsAssigned);
  if FVal <> '' then 
    Inc(AttrsAssigned);
  if FNumFmtId <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_InputCells.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_InputCells.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r',FR);
  if FDeleted <> False then 
    AWriter.AddAttribute('deleted',XmlBoolToStr(FDeleted));
  if FUndone <> False then 
    AWriter.AddAttribute('undone',XmlBoolToStr(FUndone));
  AWriter.AddAttribute('val',FVal);
  if FNumFmtId <> 0 then 
    AWriter.AddAttribute('numFmtId',XmlIntToStr(FNumFmtId));
end;

procedure TCT_InputCells.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000072: FR := AAttributes.Values[i];
      $000002D7: FDeleted := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000289: FUndone := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000143: FVal := AAttributes.Values[i];
      $00000324: FNumFmtId := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_InputCells.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
  FDeleted := False;
  FUndone := False;
end;

destructor TCT_InputCells.Destroy;
begin
end;

procedure TCT_InputCells.Clear;
begin
  FAssigneds := [];
  FR := '';
  FDeleted := False;
  FUndone := False;
  FVal := '';
  FNumFmtId := 0;
end;

{ TCT_InputCellsXpgList }

function  TCT_InputCellsXpgList.GetItems(Index: integer): TCT_InputCells;
begin
  Result := TCT_InputCells(inherited Items[Index]);
end;

function  TCT_InputCellsXpgList.Add: TCT_InputCells;
begin
  Result := TCT_InputCells.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_InputCellsXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_InputCellsXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DataRef }

function  TCT_DataRef.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
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

procedure TCT_DataRef.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DataRef.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRef <> '' then 
    AWriter.AddAttribute('ref',FRef);
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FSheet <> '' then 
    AWriter.AddAttribute('sheet',FSheet);
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_DataRef.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000013D: FRef := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $00000219: FSheet := AAttributes.Values[i];
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DataRef.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
end;

destructor TCT_DataRef.Destroy;
begin
end;

procedure TCT_DataRef.Clear;
begin
  FAssigneds := [];
  FRef := '';
  FName := '';
  FSheet := '';
  FR_Id := '';
end;

{ TCT_DataRefXpgList }

function  TCT_DataRefXpgList.GetItems(Index: integer): TCT_DataRef;
begin
  Result := TCT_DataRef(inherited Items[Index]);
end;

function  TCT_DataRefXpgList.Add: TCT_DataRef;
begin
  Result := TCT_DataRef.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DataRefXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DataRefXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ColorScale }

function  TCT_ColorScale.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCfvoXpgList.CheckAssigned);
  Inc(ElemsAssigned,FColorXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ColorScale.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001AE: Result := FCfvoXpgList.Add;
    $0000021F: Result := FColorXpgList.Add;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ColorScale.Write(AWriter: TXpgWriteXML);
begin
  FCfvoXpgList.Write(AWriter,'cfvo');
  FColorXpgList.Write(AWriter,'color');
end;

constructor TCT_ColorScale.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FCfvoXpgList := TCT_CfvoXpgList.Create(FOwner);
  FColorXpgList := TCT_ColorXpgList.Create(FOwner);
end;

destructor TCT_ColorScale.Destroy;
begin
  FCfvoXpgList.Free;
  FColorXpgList.Free;
end;

procedure TCT_ColorScale.Clear;
begin
  FAssigneds := [];
  FCfvoXpgList.Clear;
  FColorXpgList.Clear;
end;

{ TCT_DataBar }

function  TCT_DataBar.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMinLength <> 10 then 
    Inc(AttrsAssigned);
  if FMaxLength <> 90 then 
    Inc(AttrsAssigned);
  if FShowValue <> True then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCfvoXpgList.CheckAssigned);
  Inc(ElemsAssigned,FColor.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DataBar.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001AE: Result := FCfvoXpgList.Add;
    $0000021F: Result := FColor;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DataBar.Write(AWriter: TXpgWriteXML);
begin
  if FCfvoXpgList.Count > 0 then begin  // TODO Sequence. color must be written if cfvo is.
    FCfvoXpgList.Write(AWriter,'cfvo');
    FColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('color');
  end;
end;

procedure TCT_DataBar.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FMinLength <> 10 then 
    AWriter.AddAttribute('minLength',XmlIntToStr(FMinLength));
  if FMaxLength <> 90 then 
    AWriter.AddAttribute('maxLength',XmlIntToStr(FMaxLength));
  if FShowValue <> True then 
    AWriter.AddAttribute('showValue',XmlBoolToStr(FShowValue));
end;

procedure TCT_DataBar.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000003A6: FMinLength := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003A8: FMaxLength := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003BE: FShowValue := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DataBar.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 3;
  FCfvoXpgList := TCT_CfvoXpgList.Create(FOwner);
  FColor := TCT_Color.Create(FOwner);
  FMinLength := 10;
  FMaxLength := 90;
  FShowValue := True;
end;

destructor TCT_DataBar.Destroy;
begin
  FCfvoXpgList.Free;
  FColor.Free;
end;

procedure TCT_DataBar.Clear;
begin
  FAssigneds := [];
  FCfvoXpgList.Clear;
  FColor.Clear;
  FMinLength := 10;
  FMaxLength := 90;
  FShowValue := True;
end;

{ TCT_IconSet }

function  TCT_IconSet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FIconSet <> stist3TrafficLights1 then 
    Inc(AttrsAssigned);
  if FShowValue <> True then 
    Inc(AttrsAssigned);
  if FPercent <> True then 
    Inc(AttrsAssigned);
  if FReverse <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCfvoXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_IconSet.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cfvo' then 
    Result := FCfvoXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_IconSet.Write(AWriter: TXpgWriteXML);
begin
  FCfvoXpgList.Write(AWriter,'cfvo');
end;

procedure TCT_IconSet.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FIconSet <> stist3TrafficLights1 then 
    AWriter.AddAttribute('iconSet',StrTST_IconSetType[Integer(FIconSet)]);
  if FShowValue <> True then 
    AWriter.AddAttribute('showValue',XmlBoolToStr(FShowValue));
  if FPercent <> True then 
    AWriter.AddAttribute('percent',XmlBoolToStr(FPercent));
  if FReverse <> False then 
    AWriter.AddAttribute('reverse',XmlBoolToStr(FReverse));
end;

procedure TCT_IconSet.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000002D5: FIconSet := TST_IconSetType(StrToEnum('stist' + AAttributes.Values[i]));
      $000003BE: FShowValue := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F1: FPercent := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002FC: FReverse := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_IconSet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FCfvoXpgList := TCT_CfvoXpgList.Create(FOwner);
  FIconSet := stist3TrafficLights1;
  FShowValue := True;
  FPercent := True;
  FReverse := False;
end;

destructor TCT_IconSet.Destroy;
begin
  FCfvoXpgList.Free;
end;

procedure TCT_IconSet.Clear;
begin
  FAssigneds := [];
  FCfvoXpgList.Clear;
  FIconSet := stist3TrafficLights1;
  FShowValue := True;
  FPercent := True;
  FReverse := False;
end;

{ TCT_CellSmartTag }

function  TCT_CellSmartTag.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FType <> 0 then 
    Inc(AttrsAssigned);
  if FDeleted <> False then 
    Inc(AttrsAssigned);
  if FXmlBased <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCellSmartTagPrXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CellSmartTag.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cellSmartTagPr' then 
    Result := FCellSmartTagPrXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellSmartTag.Write(AWriter: TXpgWriteXML);
begin
  FCellSmartTagPrXpgList.Write(AWriter,'cellSmartTagPr');
end;

procedure TCT_CellSmartTag.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('type',XmlIntToStr(FType));
  if FDeleted <> False then 
    AWriter.AddAttribute('deleted',XmlBoolToStr(FDeleted));
  if FXmlBased <> False then 
    AWriter.AddAttribute('xmlBased',XmlBoolToStr(FXmlBased));
end;

procedure TCT_CellSmartTag.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002D7: FDeleted := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000330: FXmlBased := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CellSmartTag.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FCellSmartTagPrXpgList := TCT_CellSmartTagPrXpgList.Create(FOwner);
  FDeleted := False;
  FXmlBased := False;
end;

destructor TCT_CellSmartTag.Destroy;
begin
  FCellSmartTagPrXpgList.Free;
end;

procedure TCT_CellSmartTag.Clear;
begin
  FAssigneds := [];
  FCellSmartTagPrXpgList.Clear;
  FType := 0;
  FDeleted := False;
  FXmlBased := False;
end;

{ TCT_CellSmartTagXpgList }

function  TCT_CellSmartTagXpgList.GetItems(Index: integer): TCT_CellSmartTag;
begin
  Result := TCT_CellSmartTag(inherited Items[Index]);
end;

function  TCT_CellSmartTagXpgList.Add: TCT_CellSmartTag;
begin
  Result := TCT_CellSmartTag.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CellSmartTagXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CellSmartTagXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Xf }

function  TCT_Xf.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNumFmtId <> 0 then
    Inc(AttrsAssigned);
  if FFontId <> 0 then
    Inc(AttrsAssigned);
  if FFillId <> 0 then
    Inc(AttrsAssigned);
  if FBorderId <> 0 then
    Inc(AttrsAssigned);
  if FXfId <> 0 then
    Inc(AttrsAssigned);
  if FQuotePrefix <> False then
    Inc(AttrsAssigned);
  if FPivotButton <> False then
    Inc(AttrsAssigned);
  if FApplyNumberFormat <> False then
    Inc(AttrsAssigned);
  if FApplyFont <> False then
    Inc(AttrsAssigned);
  if FApplyFill <> False then
    Inc(AttrsAssigned);
  if FApplyBorder <> False then
    Inc(AttrsAssigned);
  if FApplyAlignment <> False then
    Inc(AttrsAssigned);
  if FApplyProtection <> False then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FAlignment.CheckAssigned);
  Inc(ElemsAssigned,FProtection.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);

  // TODO
  if FAssigneds = [] then begin
    Result := 1;
    FAssigneds := [xaAttributes];
  end;
end;

function  TCT_Xf.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003BF: Result := FAlignment;
    $00000447: Result := FProtection;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Xf.Write(AWriter: TXpgWriteXML);
begin
  if FAlignment.Assigned then
  begin
    FAlignment.WriteAttributes(AWriter);
    AWriter.SimpleTag('alignment');
  end;
  if FProtection.Assigned then
  begin
    FProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('protection');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Xf.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('numFmtId',XmlIntToStr(FNumFmtId));
  AWriter.AddAttribute('fontId',XmlIntToStr(FFontId));
  AWriter.AddAttribute('fillId',XmlIntToStr(FFillId));
  AWriter.AddAttribute('borderId',XmlIntToStr(FBorderId));
  AWriter.AddAttribute('xfId',XmlIntToStr(FXfId));
  if FQuotePrefix <> False then 
    AWriter.AddAttribute('quotePrefix',XmlBoolToStr(FQuotePrefix));
  if FPivotButton <> False then
    AWriter.AddAttribute('pivotButton',XmlBoolToStr(FPivotButton));
  if FApplyNumberFormat <> False then 
    AWriter.AddAttribute('applyNumberFormat',XmlBoolToStr(FApplyNumberFormat));
  if FApplyFont <> False then 
    AWriter.AddAttribute('applyFont',XmlBoolToStr(FApplyFont));
  if FApplyFill <> False then 
    AWriter.AddAttribute('applyFill',XmlBoolToStr(FApplyFill));
  if FApplyBorder <> False then 
    AWriter.AddAttribute('applyBorder',XmlBoolToStr(FApplyBorder));
  if FApplyAlignment <> False then 
    AWriter.AddAttribute('applyAlignment',XmlBoolToStr(FApplyAlignment));
  if FApplyProtection <> False then 
    AWriter.AddAttribute('applyProtection',XmlBoolToStr(FApplyProtection));
end;

procedure TCT_Xf.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000324: FNumFmtId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000264: FFontId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000254: FFillId := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000032B: FBorderId := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000018B: FXfId := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000049C: FQuotePrefix := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004AE: FPivotButton := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006F8: FApplyNumberFormat := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003BD: FApplyFont := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003AD: FApplyFill := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000484: FApplyBorder := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005C5: FApplyAlignment := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000064D: FApplyProtection := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Xf.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 13;
  FAlignment := TCT_CellAlignment.Create(FOwner);
  FProtection := TCT_CellProtection.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FQuotePrefix := False;
  FPivotButton := False;
end;

destructor TCT_Xf.Destroy;
begin
  FAlignment.Free;
  FProtection.Free;
  FExtLst.Free;
end;

procedure TCT_Xf.Clear;
begin
  FAssigneds := [];
  FAlignment.Clear;
  FProtection.Clear;
  FExtLst.Clear;
  FNumFmtId := 0;
  FFontId := 0;
  FFillId := 0;
  FBorderId := 0;
  FXfId := 0;
  FQuotePrefix := False;
  FPivotButton := False;
  FApplyNumberFormat := False;
  FApplyFont := False;
  FApplyFill := False;
  FApplyBorder := False;
  FApplyAlignment := False;
  FApplyProtection := False;
end;

{ TCT_XfXpgList }

function  TCT_XfXpgList.GetItems(Index: integer): TCT_Xf;
begin
  Result := TCT_Xf(inherited Items[Index]);
end;

function  TCT_XfXpgList.Add: TCT_Xf;
begin
  Result := TCT_Xf.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_XfXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_XfXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CellStyle }

function  TCT_CellStyle.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FXfId <> 0 then 
    Inc(AttrsAssigned);
  if FBuiltinId <> 0 then 
    Inc(AttrsAssigned);
  if FILevel <> 0 then 
    Inc(AttrsAssigned);
  if FHidden <> False then 
    Inc(AttrsAssigned);
  if FCustomBuiltin <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CellStyle.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'extLst' then 
    Result := FExtLst
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellStyle.Write(AWriter: TXpgWriteXML);
begin
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CellStyle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('xfId',XmlIntToStr(FXfId));
  if FBuiltinId <> 0 then 
    AWriter.AddAttribute('builtinId',XmlIntToStr(FBuiltinId));
  if FILevel <> 0 then 
    AWriter.AddAttribute('iLevel',XmlIntToStr(FILevel));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
  if FCustomBuiltin <> False then 
    AWriter.AddAttribute('customBuiltin',XmlBoolToStr(FCustomBuiltin));
end;

procedure TCT_CellStyle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $0000018B: FXfId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003A4: FBuiltinId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000261: FILevel := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000572: FCustomBuiltin := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CellStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 6;
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_CellStyle.Destroy;
begin
  FExtLst.Free;
end;

procedure TCT_CellStyle.Clear;
begin
  FAssigneds := [];
  FExtLst.Clear;
  FName := '';
  FXfId := 0;
  FBuiltinId := 0;
  FILevel := 0;
  FHidden := False;
  FCustomBuiltin := False;
end;

{ TCT_CellStyleXpgList }

function  TCT_CellStyleXpgList.GetItems(Index: integer): TCT_CellStyle;
begin
  Result := TCT_CellStyle(inherited Items[Index]);
end;

function  TCT_CellStyleXpgList.Add: TCT_CellStyle;
begin
  Result := TCT_CellStyle.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CellStyleXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CellStyleXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Dxf }

function  TCT_Dxf.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FFont.CheckAssigned);
  Inc(ElemsAssigned,FNumFmt.CheckAssigned);
  Inc(ElemsAssigned,FFill.CheckAssigned);
  Inc(ElemsAssigned,FAlignment.CheckAssigned);
  Inc(ElemsAssigned,FBorder.CheckAssigned);
  Inc(ElemsAssigned,FProtection.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Dxf.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001B7: Result := FFont;
    $00000277: Result := FNumFmt;
    $000001A7: Result := FFill;
    $000003BF: Result := FAlignment;
    $0000027E: Result := FBorder;
    $00000447: Result := FProtection;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Dxf.Write(AWriter: TXpgWriteXML);
begin
  if FFont.Assigned then 
    if xaElements in FFont.FAssigneds then 
    begin
      AWriter.BeginTag('font');
      FFont.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('font');
  if FNumFmt.Assigned then 
  begin
    FNumFmt.WriteAttributes(AWriter);
    AWriter.SimpleTag('numFmt');
  end;
  if FFill.Assigned then 
    if xaElements in FFill.FAssigneds then 
    begin
      AWriter.BeginTag('fill');
      FFill.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('fill');
  if FAlignment.Assigned then 
  begin
    FAlignment.WriteAttributes(AWriter);
    AWriter.SimpleTag('alignment');
  end;
  if FBorder.Assigned then 
  begin
    FBorder.WriteAttributes(AWriter);
    if xaElements in FBorder.FAssigneds then 
    begin
      AWriter.BeginTag('border');
      FBorder.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('border');
  end;
  if FProtection.Assigned then 
  begin
    FProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('protection');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Dxf.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 7;
  FAttributeCount := 0;
  FFont := TCT_Font.Create(FOwner);
  FNumFmt := TCT_NumFmt.Create(FOwner);
  FFill := TCT_Fill.Create(FOwner);
  FAlignment := TCT_CellAlignment.Create(FOwner);
  FBorder := TCT_Border.Create(FOwner);
  FProtection := TCT_CellProtection.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Dxf.Destroy;
begin
  FFont.Free;
  FNumFmt.Free;
  FFill.Free;
  FAlignment.Free;
  FBorder.Free;
  FProtection.Free;
  FExtLst.Free;
end;

procedure TCT_Dxf.Clear;
begin
  FAssigneds := [];
  FFont.Clear;
  FNumFmt.Clear;
  FFill.Clear;
  FAlignment.Clear;
  FBorder.Clear;
  FProtection.Clear;
  FExtLst.Clear;
end;

{ TCT_DxfXpgList }

function  TCT_DxfXpgList.GetItems(Index: integer): TCT_Dxf;
begin
  Result := TCT_Dxf(inherited Items[Index]);
end;

function  TCT_DxfXpgList.Add: TCT_Dxf;
begin
  Result := TCT_Dxf.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DxfXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DxfXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_TableStyle }

function  TCT_TableStyle.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FPivot <> True then 
    Inc(AttrsAssigned);
  if FTable <> True then 
    Inc(AttrsAssigned);
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FTableStyleElementXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TableStyle.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'tableStyleElement' then 
    Result := FTableStyleElementXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_TableStyle.Write(AWriter: TXpgWriteXML);
begin
  FTableStyleElementXpgList.Write(AWriter,'tableStyleElement');
end;

procedure TCT_TableStyle.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  if FPivot <> True then 
    AWriter.AddAttribute('pivot',XmlBoolToStr(FPivot));
  if FTable <> True then 
    AWriter.AddAttribute('table',XmlBoolToStr(FTable));
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_TableStyle.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000232: FPivot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000208: FTable := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_TableStyle.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FTableStyleElementXpgList := TCT_TableStyleElementXpgList.Create(FOwner);
  FPivot := True;
  FTable := True;
end;

destructor TCT_TableStyle.Destroy;
begin
  FTableStyleElementXpgList.Free;
end;

procedure TCT_TableStyle.Clear;
begin
  FAssigneds := [];
  FTableStyleElementXpgList.Clear;
  FName := '';
  FPivot := True;
  FTable := True;
  FCount := 0;
end;

{ TCT_TableStyleXpgList }

function  TCT_TableStyleXpgList.GetItems(Index: integer): TCT_TableStyle;
begin
  Result := TCT_TableStyle(inherited Items[Index]);
end;

function  TCT_TableStyleXpgList.Add: TCT_TableStyle;
begin
  Result := TCT_TableStyle.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TableStyleXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TableStyleXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_IndexedColors }

function  TCT_IndexedColors.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FRgbColorXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_IndexedColors.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'rgbColor' then 
    Result := FRgbColorXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_IndexedColors.Write(AWriter: TXpgWriteXML);
begin
  FRgbColorXpgList.Write(AWriter,'rgbColor');
end;

constructor TCT_IndexedColors.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FRgbColorXpgList := TCT_RgbColorXpgList.Create(FOwner);
end;

destructor TCT_IndexedColors.Destroy;
begin
  FRgbColorXpgList.Free;
end;

procedure TCT_IndexedColors.Clear;
begin
  FAssigneds := [];
  FRgbColorXpgList.Clear;
end;

{ TCT_MRUColors }

function  TCT_MRUColors.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FColorXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_MRUColors.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'color' then 
    Result := FColorXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_MRUColors.Write(AWriter: TXpgWriteXML);
begin
  FColorXpgList.Write(AWriter,'color');
end;

constructor TCT_MRUColors.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FColorXpgList := TCT_ColorXpgList.Create(FOwner);
end;

destructor TCT_MRUColors.Destroy;
begin
  FColorXpgList.Free;
end;

procedure TCT_MRUColors.Clear;
begin
  FAssigneds := [];
  FColorXpgList.Clear;
end;

{ TCT_BookView }

function  TCT_BookView.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FVisibility <> stvVisible then 
    Inc(AttrsAssigned);
  if FMinimized <> False then 
    Inc(AttrsAssigned);
  if FShowHorizontalScroll <> True then 
    Inc(AttrsAssigned);
  if FShowVerticalScroll <> True then 
    Inc(AttrsAssigned);
  if FShowSheetTabs <> True then 
    Inc(AttrsAssigned);
  if FXWindow <> 0 then 
    Inc(AttrsAssigned);
  if FYWindow <> 0 then 
    Inc(AttrsAssigned);
  if FWindowWidth <> 0 then 
    Inc(AttrsAssigned);
  if FWindowHeight <> 0 then 
    Inc(AttrsAssigned);
  if FTabRatio <> 600 then 
    Inc(AttrsAssigned);
  if FFirstSheet <> 0 then 
    Inc(AttrsAssigned);
  if FActiveTab <> 0 then 
    Inc(AttrsAssigned);
  if FAutoFilterDateGrouping <> True then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_BookView.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'extLst' then 
    Result := FExtLst
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_BookView.Write(AWriter: TXpgWriteXML);
begin
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_BookView.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVisibility <> stvVisible then 
    AWriter.AddAttribute('visibility',StrTST_Visibility[Integer(FVisibility)]);
  if FMinimized <> False then 
    AWriter.AddAttribute('minimized',XmlBoolToStr(FMinimized));
  if FShowHorizontalScroll <> True then 
    AWriter.AddAttribute('showHorizontalScroll',XmlBoolToStr(FShowHorizontalScroll));
  if FShowVerticalScroll <> True then 
    AWriter.AddAttribute('showVerticalScroll',XmlBoolToStr(FShowVerticalScroll));
  if FShowSheetTabs <> True then 
    AWriter.AddAttribute('showSheetTabs',XmlBoolToStr(FShowSheetTabs));
  if FXWindow <> 0 then 
    AWriter.AddAttribute('xWindow',XmlIntToStr(FXWindow));
  if FYWindow <> 0 then 
    AWriter.AddAttribute('yWindow',XmlIntToStr(FYWindow));
  if FWindowWidth <> 0 then 
    AWriter.AddAttribute('windowWidth',XmlIntToStr(FWindowWidth));
  if FWindowHeight <> 0 then 
    AWriter.AddAttribute('windowHeight',XmlIntToStr(FWindowHeight));
  if FTabRatio <> 600 then 
    AWriter.AddAttribute('tabRatio',XmlIntToStr(FTabRatio));
  if FFirstSheet <> 0 then 
    AWriter.AddAttribute('firstSheet',XmlIntToStr(FFirstSheet));
  if FActiveTab <> 0 then 
    AWriter.AddAttribute('activeTab',XmlIntToStr(FActiveTab));
  if FAutoFilterDateGrouping <> True then 
    AWriter.AddAttribute('autoFilterDateGrouping',XmlBoolToStr(FAutoFilterDateGrouping));
end;

procedure TCT_BookView.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000448: FVisibility := TST_Visibility(StrToEnum('stv' + AAttributes.Values[i]));
      $000003C6: FMinimized := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000085A: FShowHorizontalScroll := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000076A: FShowVerticalScroll := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000544: FShowSheetTabs := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F0: FXWindow := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002F1: FYWindow := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000498: FWindowWidth := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004F1: FWindowHeight := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000336: FTabRatio := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000421: FFirstSheet := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000393: FActiveTab := XmlStrToIntDef(AAttributes.Values[i],0);
      $000008E8: FAutoFilterDateGrouping := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_BookView.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 13;
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FVisibility := stvVisible;
  FMinimized := False;
  FShowHorizontalScroll := True;
  FShowVerticalScroll := True;
  FShowSheetTabs := True;
  FTabRatio := 600;
  FFirstSheet := 0;
  FActiveTab := 0;
  FAutoFilterDateGrouping := True;
end;

destructor TCT_BookView.Destroy;
begin
  FExtLst.Free;
end;

procedure TCT_BookView.Clear;
begin
  FAssigneds := [];
  FExtLst.Clear;
  FVisibility := stvVisible;
  FMinimized := False;
  FShowHorizontalScroll := True;
  FShowVerticalScroll := True;
  FShowSheetTabs := True;
  FXWindow := 0;
  FYWindow := 0;
  FWindowWidth := 0;
  FWindowHeight := 0;
  FTabRatio := 600;
  FFirstSheet := 0;
  FActiveTab := 0;
  FAutoFilterDateGrouping := True;
end;

{ TCT_BookViewXpgList }

function  TCT_BookViewXpgList.GetItems(Index: integer): TCT_BookView;
begin
  Result := TCT_BookView(inherited Items[Index]);
end;

function  TCT_BookViewXpgList.Add: TCT_BookView;
begin
  Result := TCT_BookView.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_BookViewXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_BookViewXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Sheet }

function  TCT_Sheet.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FSheetId <> 0 then 
    Inc(AttrsAssigned);
  if FState <> stssVisible then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Sheet.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Sheet.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('sheetId',XmlIntToStr(FSheetId));
  if FState <> stssVisible then 
    AWriter.AddAttribute('state',StrTST_SheetState[Integer(FState)]);
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_Sheet.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000002C6: FSheetId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000221: FState := TST_SheetState(StrToEnum('stss' + AAttributes.Values[i]));
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Sheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FState := stssVisible;
end;

destructor TCT_Sheet.Destroy;
begin
end;

procedure TCT_Sheet.Clear;
begin
  FAssigneds := [];
  FName := '';
  FSheetId := 0;
  FState := stssVisible;
  FR_Id := '';
end;

{ TCT_SheetXpgList }

function  TCT_SheetXpgList.GetItems(Index: integer): TCT_Sheet;
begin
  Result := TCT_Sheet(inherited Items[Index]);
end;

function  TCT_SheetXpgList.Add: TCT_Sheet;
begin
  Result := TCT_Sheet.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SheetXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SheetXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_FunctionGroup }

function  TCT_FunctionGroup.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FunctionGroup.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FunctionGroup.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
end;

procedure TCT_FunctionGroup.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'name' then 
    FName := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FunctionGroup.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_FunctionGroup.Destroy;
begin
end;

procedure TCT_FunctionGroup.Clear;
begin
  FAssigneds := [];
  FName := '';
end;

{ TCT_FunctionGroupXpgList }

function  TCT_FunctionGroupXpgList.GetItems(Index: integer): TCT_FunctionGroup;
begin
  Result := TCT_FunctionGroup(inherited Items[Index]);
end;

function  TCT_FunctionGroupXpgList.Add: TCT_FunctionGroup;
begin
  Result := TCT_FunctionGroup.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FunctionGroupXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FunctionGroupXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ExternalReference }

function  TCT_ExternalReference.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ExternalReference.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ExternalReference.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_ExternalReference.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if NameOfQName(AAttributes[0]) = 'id' then
    FR_Id := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ExternalReference.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_ExternalReference.Destroy;
begin
end;

procedure TCT_ExternalReference.Clear;
begin
  FAssigneds := [];
  FR_Id := '';
end;

{ TCT_ExternalReferenceXpgList }

function  TCT_ExternalReferenceXpgList.GetItems(Index: integer): TCT_ExternalReference;
begin
  Result := TCT_ExternalReference(inherited Items[Index]);
end;

function  TCT_ExternalReferenceXpgList.Add: TCT_ExternalReference;
begin
  Result := TCT_ExternalReference.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExternalReferenceXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExternalReferenceXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DefinedName }

function  TCT_DefinedName.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then
    Inc(AttrsAssigned);
  if FComment <> '' then
    Inc(AttrsAssigned);
  if FCustomMenu <> '' then
    Inc(AttrsAssigned);
  if FDescription <> '' then
    Inc(AttrsAssigned);
  if FHelp <> '' then
    Inc(AttrsAssigned);
  if FStatusBar <> '' then
    Inc(AttrsAssigned);
  if FLocalSheetId <> -1 then  // TODO Zero is valid
    Inc(AttrsAssigned);
  if FHidden <> False then
    Inc(AttrsAssigned);
  if FFunction <> False then 
    Inc(AttrsAssigned);
  if FVbProcedure <> False then 
    Inc(AttrsAssigned);
  if FXlm <> False then 
    Inc(AttrsAssigned);
  if FFunctionGroupId <> 0 then 
    Inc(AttrsAssigned);
  if FShortcutKey <> '' then 
    Inc(AttrsAssigned);
  if FPublishToServer <> False then 
    Inc(AttrsAssigned);
  if FWorkbookParameter <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
  if FContent <> '' then 
  begin
    FAssigneds := FAssigneds + [xaContent];
    Inc(Result);
  end;
end;

function  TCT_DefinedName.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DefinedName.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_DefinedName.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  if FComment <> '' then 
    AWriter.AddAttribute('comment',FComment);
  if FCustomMenu <> '' then 
    AWriter.AddAttribute('customMenu',FCustomMenu);
  if FDescription <> '' then 
    AWriter.AddAttribute('description',FDescription);
  if FHelp <> '' then 
    AWriter.AddAttribute('help',FHelp);
  if FStatusBar <> '' then 
    AWriter.AddAttribute('statusBar',FStatusBar);
  if FLocalSheetId <> -1 then      // TODO Zero is valid
    AWriter.AddAttribute('localSheetId',XmlIntToStr(FLocalSheetId));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
  if FFunction <> False then 
    AWriter.AddAttribute('function',XmlBoolToStr(FFunction));
  if FVbProcedure <> False then 
    AWriter.AddAttribute('vbProcedure',XmlBoolToStr(FVbProcedure));
  if FXlm <> False then 
    AWriter.AddAttribute('xlm',XmlBoolToStr(FXlm));
  if FFunctionGroupId <> 0 then 
    AWriter.AddAttribute('functionGroupId',XmlIntToStr(FFunctionGroupId));
  if FShortcutKey <> '' then 
    AWriter.AddAttribute('shortcutKey',FShortcutKey);
  if FPublishToServer <> False then 
    AWriter.AddAttribute('publishToServer',XmlBoolToStr(FPublishToServer));
  if FWorkbookParameter <> False then 
    AWriter.AddAttribute('workbookParameter',XmlBoolToStr(FWorkbookParameter));
end;

procedure TCT_DefinedName.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000002F3: FComment := AAttributes.Values[i];
      $00000430: FCustomMenu := AAttributes.Values[i];
      $000004A4: FDescription := AAttributes.Values[i];
      $000001A9: FHelp := AAttributes.Values[i];
      $000003B9: FStatusBar := AAttributes.Values[i];
      $000004B1: FLocalSheetId := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000366: FFunction := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000481: FVbProcedure := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000151: FXlm := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000620: FFunctionGroupId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004A5: FShortcutKey := AAttributes.Values[i];
      $00000631: FPublishToServer := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000070F: FWorkbookParameter := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DefinedName.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 15;
  FHidden := False;
  FFunction := False;
  FVbProcedure := False;
  FXlm := False;
  FPublishToServer := False;
  FWorkbookParameter := False;
  FLocalSheetId := -1;  // TODO Zero is valid
end;

destructor TCT_DefinedName.Destroy;
begin
end;

procedure TCT_DefinedName.Clear;
begin
  FAssigneds := [];
  FName := '';
  FComment := '';
  FCustomMenu := '';
  FDescription := '';
  FHelp := '';
  FStatusBar := '';
  FLocalSheetId := -1;  // TODO Zero is valid
  FHidden := False;
  FFunction := False;
  FVbProcedure := False;
  FXlm := False;
  FFunctionGroupId := 0;
  FShortcutKey := '';
  FPublishToServer := False;
  FWorkbookParameter := False;
end;

{ TCT_DefinedNameXpgList }

function  TCT_DefinedNameXpgList.GetItems(Index: integer): TCT_DefinedName;
begin
  Result := TCT_DefinedName(inherited Items[Index]);
end;

function  TCT_DefinedNameXpgList.Add: TCT_DefinedName;
begin
  Result := TCT_DefinedName.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DefinedNameXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DefinedNameXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do 
  begin
    if xaAttributes in Items[i].FAssigneds then 
      GetItems(i).WriteAttributes(AWriter);
    if xaContent in Items[i].FAssigneds then 
      AWriter.Text := GetItems(i).Content;
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

{ TCT_CustomWorkbookView }

function  TCT_CustomWorkbookView.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FGuid <> '' then 
    Inc(AttrsAssigned);
  if FAutoUpdate <> False then 
    Inc(AttrsAssigned);
  if FMergeInterval <> 0 then 
    Inc(AttrsAssigned);
  if FChangesSavedWin <> False then 
    Inc(AttrsAssigned);
  if FOnlySync <> False then 
    Inc(AttrsAssigned);
  if FPersonalView <> False then 
    Inc(AttrsAssigned);
  if FIncludePrintSettings <> True then 
    Inc(AttrsAssigned);
  if FIncludeHiddenRowCol <> True then 
    Inc(AttrsAssigned);
  if FMaximized <> False then 
    Inc(AttrsAssigned);
  if FMinimized <> False then 
    Inc(AttrsAssigned);
  if FShowHorizontalScroll <> True then 
    Inc(AttrsAssigned);
  if FShowVerticalScroll <> True then 
    Inc(AttrsAssigned);
  if FShowSheetTabs <> True then 
    Inc(AttrsAssigned);
  if FXWindow <> 0 then 
    Inc(AttrsAssigned);
  if FYWindow <> 0 then 
    Inc(AttrsAssigned);
  if FWindowWidth <> 0 then 
    Inc(AttrsAssigned);
  if FWindowHeight <> 0 then 
    Inc(AttrsAssigned);
  if FTabRatio <> 600 then 
    Inc(AttrsAssigned);
  if FActiveSheetId <> 0 then 
    Inc(AttrsAssigned);
  if FShowFormulaBar <> True then 
    Inc(AttrsAssigned);
  if FShowStatusbar <> True then 
    Inc(AttrsAssigned);
  if FShowComments <> stcCommIndicator then 
    Inc(AttrsAssigned);
  if FShowObjects <> stoAll then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CustomWorkbookView.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'extLst' then 
    Result := FExtLst
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomWorkbookView.Write(AWriter: TXpgWriteXML);
begin
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CustomWorkbookView.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('guid',FGuid);
  if FAutoUpdate <> False then 
    AWriter.AddAttribute('autoUpdate',XmlBoolToStr(FAutoUpdate));
  if FMergeInterval <> 0 then 
    AWriter.AddAttribute('mergeInterval',XmlIntToStr(FMergeInterval));
  if FChangesSavedWin <> False then 
    AWriter.AddAttribute('changesSavedWin',XmlBoolToStr(FChangesSavedWin));
  if FOnlySync <> False then 
    AWriter.AddAttribute('onlySync',XmlBoolToStr(FOnlySync));
  if FPersonalView <> False then 
    AWriter.AddAttribute('personalView',XmlBoolToStr(FPersonalView));
  if FIncludePrintSettings <> True then 
    AWriter.AddAttribute('includePrintSettings',XmlBoolToStr(FIncludePrintSettings));
  if FIncludeHiddenRowCol <> True then 
    AWriter.AddAttribute('includeHiddenRowCol',XmlBoolToStr(FIncludeHiddenRowCol));
  if FMaximized <> False then 
    AWriter.AddAttribute('maximized',XmlBoolToStr(FMaximized));
  if FMinimized <> False then 
    AWriter.AddAttribute('minimized',XmlBoolToStr(FMinimized));
  if FShowHorizontalScroll <> True then 
    AWriter.AddAttribute('showHorizontalScroll',XmlBoolToStr(FShowHorizontalScroll));
  if FShowVerticalScroll <> True then 
    AWriter.AddAttribute('showVerticalScroll',XmlBoolToStr(FShowVerticalScroll));
  if FShowSheetTabs <> True then 
    AWriter.AddAttribute('showSheetTabs',XmlBoolToStr(FShowSheetTabs));
  if FXWindow <> 0 then 
    AWriter.AddAttribute('xWindow',XmlIntToStr(FXWindow));
  if FYWindow <> 0 then 
    AWriter.AddAttribute('yWindow',XmlIntToStr(FYWindow));
  AWriter.AddAttribute('windowWidth',XmlIntToStr(FWindowWidth));
  AWriter.AddAttribute('windowHeight',XmlIntToStr(FWindowHeight));
  if FTabRatio <> 600 then 
    AWriter.AddAttribute('tabRatio',XmlIntToStr(FTabRatio));
  AWriter.AddAttribute('activeSheetId',XmlIntToStr(FActiveSheetId));
  if FShowFormulaBar <> True then 
    AWriter.AddAttribute('showFormulaBar',XmlBoolToStr(FShowFormulaBar));
  if FShowStatusbar <> True then 
    AWriter.AddAttribute('showStatusbar',XmlBoolToStr(FShowStatusbar));
  if FShowComments <> stcCommIndicator then 
    AWriter.AddAttribute('showComments',StrTST_Comments[Integer(FShowComments)]);
  if FShowObjects <> stoAll then 
    AWriter.AddAttribute('showObjects',StrTST_Objects[Integer(FShowObjects)]);
end;

procedure TCT_CustomWorkbookView.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $000001A9: FGuid := AAttributes.Values[i];
      $0000041C: FAutoUpdate := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000555: FMergeInterval := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005FA: FChangesSavedWin := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000035F: FOnlySync := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004FF: FPersonalView := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000842: FIncludePrintSettings := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000786: FIncludeHiddenRowCol := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003C8: FMaximized := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003C6: FMinimized := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000085A: FShowHorizontalScroll := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000076A: FShowVerticalScroll := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000544: FShowSheetTabs := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F0: FXWindow := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002F1: FYWindow := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000498: FWindowWidth := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004F1: FWindowHeight := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000336: FTabRatio := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000522: FActiveSheetId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005AC: FShowFormulaBar := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000057A: FShowStatusbar := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000507: FShowComments := TST_Comments(StrToEnum('stc' + AAttributes.Values[i]));
      $0000048B: FShowObjects := TST_Objects(StrToEnum('sto' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CustomWorkbookView.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 24;
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FAutoUpdate := False;
  FChangesSavedWin := False;
  FOnlySync := False;
  FPersonalView := False;
  FIncludePrintSettings := True;
  FIncludeHiddenRowCol := True;
  FMaximized := False;
  FMinimized := False;
  FShowHorizontalScroll := True;
  FShowVerticalScroll := True;
  FShowSheetTabs := True;
  FXWindow := 0;
  FYWindow := 0;
  FTabRatio := 600;
  FShowFormulaBar := True;
  FShowStatusbar := True;
  FShowComments := stcCommIndicator;
  FShowObjects := stoAll;
end;

destructor TCT_CustomWorkbookView.Destroy;
begin
  FExtLst.Free;
end;

procedure TCT_CustomWorkbookView.Clear;
begin
  FAssigneds := [];
  FExtLst.Clear;
  FName := '';
  FGuid := '';
  FAutoUpdate := False;
  FMergeInterval := 0;
  FChangesSavedWin := False;
  FOnlySync := False;
  FPersonalView := False;
  FIncludePrintSettings := True;
  FIncludeHiddenRowCol := True;
  FMaximized := False;
  FMinimized := False;
  FShowHorizontalScroll := True;
  FShowVerticalScroll := True;
  FShowSheetTabs := True;
  FXWindow := 0;
  FYWindow := 0;
  FWindowWidth := 0;
  FWindowHeight := 0;
  FTabRatio := 600;
  FActiveSheetId := 0;
  FShowFormulaBar := True;
  FShowStatusbar := True;
  FShowComments := stcCommIndicator;
  FShowObjects := stoAll;
end;

{ TCT_CustomWorkbookViewXpgList }

function  TCT_CustomWorkbookViewXpgList.GetItems(Index: integer): TCT_CustomWorkbookView;
begin
  Result := TCT_CustomWorkbookView(inherited Items[Index]);
end;

function  TCT_CustomWorkbookViewXpgList.Add: TCT_CustomWorkbookView;
begin
  Result := TCT_CustomWorkbookView.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CustomWorkbookViewXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CustomWorkbookViewXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_PivotCache }

function  TCT_PivotCache.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCacheId <> 0 then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PivotCache.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PivotCache.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('cacheId',XmlIntToStr(FCacheId));
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_PivotCache.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000002A1: FCacheId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PivotCache.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_PivotCache.Destroy;
begin
end;

procedure TCT_PivotCache.Clear;
begin
  FAssigneds := [];
  FCacheId := 0;
  FR_Id := '';
end;

{ TCT_PivotCacheXpgList }

function  TCT_PivotCacheXpgList.GetItems(Index: integer): TCT_PivotCache;
begin
  Result := TCT_PivotCache(inherited Items[Index]);
end;

function  TCT_PivotCacheXpgList.Add: TCT_PivotCache;
begin
  Result := TCT_PivotCache.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_PivotCacheXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_PivotCacheXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_SmartTagType }

function  TCT_SmartTagType.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FNamespaceUri <> '' then 
    Inc(AttrsAssigned);
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FUrl <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SmartTagType.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SmartTagType.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FNamespaceUri <> '' then 
    AWriter.AddAttribute('namespaceUri',FNamespaceUri);
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FUrl <> '' then 
    AWriter.AddAttribute('url',FUrl);
end;

procedure TCT_SmartTagType.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000004DD: FNamespaceUri := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $00000153: FUrl := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SmartTagType.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
end;

destructor TCT_SmartTagType.Destroy;
begin
end;

procedure TCT_SmartTagType.Clear;
begin
  FAssigneds := [];
  FNamespaceUri := '';
  FName := '';
  FUrl := '';
end;

{ TCT_SmartTagTypeXpgList }

function  TCT_SmartTagTypeXpgList.GetItems(Index: integer): TCT_SmartTagType;
begin
  Result := TCT_SmartTagType(inherited Items[Index]);
end;

function  TCT_SmartTagTypeXpgList.Add: TCT_SmartTagType;
begin
  Result := TCT_SmartTagType.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SmartTagTypeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SmartTagTypeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_WebPublishObject }

function  TCT_WebPublishObject.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FId <> 0 then 
    Inc(AttrsAssigned);
  if FDivId <> '' then 
    Inc(AttrsAssigned);
  if FSourceObject <> '' then 
    Inc(AttrsAssigned);
  if FDestinationFile <> '' then 
    Inc(AttrsAssigned);
  if FTitle <> '' then 
    Inc(AttrsAssigned);
  if FAutoRepublish <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_WebPublishObject.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_WebPublishObject.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  AWriter.AddAttribute('divId',FDivId);
  if FSourceObject <> '' then 
    AWriter.AddAttribute('sourceObject',FSourceObject);
  AWriter.AddAttribute('destinationFile',FDestinationFile);
  if FTitle <> '' then 
    AWriter.AddAttribute('title',FTitle);
  if FAutoRepublish <> False then 
    AWriter.AddAttribute('autoRepublish',XmlBoolToStr(FAutoRepublish));
end;

procedure TCT_WebPublishObject.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001F0: FDivId := AAttributes.Values[i];
      $000004E8: FSourceObject := AAttributes.Values[i];
      $00000622: FDestinationFile := AAttributes.Values[i];
      $00000222: FTitle := AAttributes.Values[i];
      $00000567: FAutoRepublish := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_WebPublishObject.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 6;
  FAutoRepublish := False;
end;

destructor TCT_WebPublishObject.Destroy;
begin
end;

procedure TCT_WebPublishObject.Clear;
begin
  FAssigneds := [];
  FId := 0;
  FDivId := '';
  FSourceObject := '';
  FDestinationFile := '';
  FTitle := '';
  FAutoRepublish := False;
end;

{ TCT_WebPublishObjectXpgList }

function  TCT_WebPublishObjectXpgList.GetItems(Index: integer): TCT_WebPublishObject;
begin
  Result := TCT_WebPublishObject(inherited Items[Index]);
end;

function  TCT_WebPublishObjectXpgList.Add: TCT_WebPublishObject;
begin
  Result := TCT_WebPublishObject.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_WebPublishObjectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_WebPublishObjectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Comment }

function  TCT_Comment.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  if FAuthorId <> 0 then 
    Inc(AttrsAssigned);
  if FGuid <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FText.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Comment.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'text' then 
    Result := FText
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Comment.Write(AWriter: TXpgWriteXML);
begin
  if FText.Assigned then 
    if xaElements in FText.FAssigneds then 
    begin
      AWriter.BeginTag('text');
      FText.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('text');
end;

procedure TCT_Comment.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ref',FRef);
  AWriter.AddAttribute('authorId',XmlIntToStr(FAuthorId));
  if FGuid <> '' then 
    AWriter.AddAttribute('guid',FGuid);
end;

procedure TCT_Comment.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000013D: FRef := AAttributes.Values[i];
      $00000340: FAuthorId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A9: FGuid := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Comment.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FText := TCT_Rst.Create(FOwner);
end;

destructor TCT_Comment.Destroy;
begin
  FText.Free;
end;

procedure TCT_Comment.Clear;
begin
  FAssigneds := [];
  FText.Clear;
  FRef := '';
  FAuthorId := 0;
  FGuid := '';
end;

{ TCT_CommentXpgList }

function  TCT_CommentXpgList.GetItems(Index: integer): TCT_Comment;
begin
  Result := TCT_Comment(inherited Items[Index]);
end;

function  TCT_CommentXpgList.Add: TCT_Comment;
begin
  Result := TCT_Comment.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CommentXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CommentXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ChartsheetView }

function  TCT_ChartsheetView.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FTabSelected <> False then 
    Inc(AttrsAssigned);
  if FZoomScale <> 100 then 
    Inc(AttrsAssigned);
  if FWorkbookViewId <> 0 then 
    Inc(AttrsAssigned);
  if FZoomToFit <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ChartsheetView.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'extLst' then 
    Result := FExtLst
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ChartsheetView.Write(AWriter: TXpgWriteXML);
begin
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_ChartsheetView.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FTabSelected <> False then 
    AWriter.AddAttribute('tabSelected',XmlBoolToStr(FTabSelected));
  if FZoomScale <> 100 then 
    AWriter.AddAttribute('zoomScale',XmlIntToStr(FZoomScale));
  AWriter.AddAttribute('workbookViewId',XmlIntToStr(FWorkbookViewId));
  if FZoomToFit <> False then 
    AWriter.AddAttribute('zoomToFit',XmlBoolToStr(FZoomToFit));
end;

procedure TCT_ChartsheetView.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000460: FTabSelected := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003AD: FZoomScale := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005B6: FWorkbookViewId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003AB: FZoomToFit := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_ChartsheetView.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FTabSelected := False;
  FZoomScale := 100;
  FZoomToFit := False;
end;

destructor TCT_ChartsheetView.Destroy;
begin
  FExtLst.Free;
end;

procedure TCT_ChartsheetView.Clear;
begin
  FAssigneds := [];
  FExtLst.Clear;
  FTabSelected := False;
  FZoomScale := 100;
  FWorkbookViewId := 0;
  FZoomToFit := False;
end;

{ TCT_ChartsheetViewXpgList }

function  TCT_ChartsheetViewXpgList.GetItems(Index: integer): TCT_ChartsheetView;
begin
  Result := TCT_ChartsheetView(inherited Items[Index]);
end;

function  TCT_ChartsheetViewXpgList.Add: TCT_ChartsheetView;
begin
  Result := TCT_ChartsheetView.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ChartsheetViewXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ChartsheetViewXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CustomChartsheetView }

function  TCT_CustomChartsheetView.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FGuid <> '' then 
    Inc(AttrsAssigned);
  if FScale <> 100 then 
    Inc(AttrsAssigned);
  if FState <> stssVisible then 
    Inc(AttrsAssigned);
  if FZoomToFit <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CustomChartsheetView.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000046E: Result := FPageMargins;
    $000003AE: Result := FPageSetup;
    $000004D8: Result := FHeaderFooter;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomChartsheetView.Write(AWriter: TXpgWriteXML);
begin
  if FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageMargins');
  end;
  if FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetup');
  end;
  if FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.FAssigneds then 
    begin
      AWriter.BeginTag('headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('headerFooter');
  end;
end;

procedure TCT_CustomChartsheetView.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('guid',FGuid);
  if FScale <> 100 then 
    AWriter.AddAttribute('scale',XmlIntToStr(FScale));
  if FState <> stssVisible then 
    AWriter.AddAttribute('state',StrTST_SheetState[Integer(FState)]);
  if FZoomToFit <> False then 
    AWriter.AddAttribute('zoomToFit',XmlBoolToStr(FZoomToFit));
end;

procedure TCT_CustomChartsheetView.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A9: FGuid := AAttributes.Values[i];
      $00000208: FScale := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000221: FState := TST_SheetState(StrToEnum('stss' + AAttributes.Values[i]));
      $000003AB: FZoomToFit := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CustomChartsheetView.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 4;
  FPageMargins := TCT_PageMargins.Create(FOwner);
  FPageSetup := TCT_CsPageSetup.Create(FOwner);
  FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
  FScale := 100;
  FState := stssVisible;
  FZoomToFit := False;
end;

destructor TCT_CustomChartsheetView.Destroy;
begin
  FPageMargins.Free;
  FPageSetup.Free;
  FHeaderFooter.Free;
end;

procedure TCT_CustomChartsheetView.Clear;
begin
  FAssigneds := [];
  FPageMargins.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FGuid := '';
  FScale := 100;
  FState := stssVisible;
  FZoomToFit := False;
end;

{ TCT_CustomChartsheetViewXpgList }

function  TCT_CustomChartsheetViewXpgList.GetItems(Index: integer): TCT_CustomChartsheetView;
begin
  Result := TCT_CustomChartsheetView(inherited Items[Index]);
end;

function  TCT_CustomChartsheetViewXpgList.Add: TCT_CustomChartsheetView;
begin
  Result := TCT_CustomChartsheetView.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CustomChartsheetViewXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CustomChartsheetViewXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_WebPublishItem }

function  TCT_WebPublishItem.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FId <> 0 then 
    Inc(AttrsAssigned);
  if FDivId <> '' then 
    Inc(AttrsAssigned);
  if Integer(FSourceType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FSourceRef <> '' then
    Inc(AttrsAssigned);
  if FSourceObject <> '' then
    Inc(AttrsAssigned);
  if FDestinationFile <> '' then
    Inc(AttrsAssigned);
  if FTitle <> '' then
    Inc(AttrsAssigned);
  if FAutoRepublish <> False then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_WebPublishItem.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_WebPublishItem.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  AWriter.AddAttribute('divId',FDivId);
  if Integer(FSourceType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('sourceType',StrTST_WebSourceType[Integer(FSourceType)]);
  if FSourceRef <> '' then 
    AWriter.AddAttribute('sourceRef',FSourceRef);
  if FSourceObject <> '' then 
    AWriter.AddAttribute('sourceObject',FSourceObject);
  AWriter.AddAttribute('destinationFile',FDestinationFile);
  if FTitle <> '' then 
    AWriter.AddAttribute('title',FTitle);
  if FAutoRepublish <> False then 
    AWriter.AddAttribute('autoRepublish',XmlBoolToStr(FAutoRepublish));
end;

procedure TCT_WebPublishItem.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001F0: FDivId := AAttributes.Values[i];
      $00000433: FSourceType := TST_WebSourceType(StrToEnum('stwst' + AAttributes.Values[i]));
      $000003AE: FSourceRef := AAttributes.Values[i];
      $000004E8: FSourceObject := AAttributes.Values[i];
      $00000622: FDestinationFile := AAttributes.Values[i];
      $00000222: FTitle := AAttributes.Values[i];
      $00000567: FAutoRepublish := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_WebPublishItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 8;
  FSourceType := TST_WebSourceType(XPG_UNKNOWN_ENUM);
  FAutoRepublish := False;
end;

destructor TCT_WebPublishItem.Destroy;
begin
end;

procedure TCT_WebPublishItem.Clear;
begin
  FAssigneds := [];
  FId := 0;
  FDivId := '';
  FSourceType := TST_WebSourceType(XPG_UNKNOWN_ENUM);
  FSourceRef := '';
  FSourceObject := '';
  FDestinationFile := '';
  FTitle := '';
  FAutoRepublish := False;
end;

{ TCT_WebPublishItemXpgList }

function  TCT_WebPublishItemXpgList.GetItems(Index: integer): TCT_WebPublishItem;
begin
  Result := TCT_WebPublishItem(inherited Items[Index]);
end;

function  TCT_WebPublishItemXpgList.Add: TCT_WebPublishItem;
begin
  Result := TCT_WebPublishItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_WebPublishItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_WebPublishItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_OutlinePr }

function  TCT_OutlinePr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FApplyStyles <> False then 
    Inc(AttrsAssigned);
  if FSummaryBelow <> True then 
    Inc(AttrsAssigned);
  if FSummaryRight <> True then 
    Inc(AttrsAssigned);
  if FShowOutlineSymbols <> True then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_OutlinePr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_OutlinePr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FApplyStyles <> False then 
    AWriter.AddAttribute('applyStyles',XmlBoolToStr(FApplyStyles));
  if FSummaryBelow <> True then 
    AWriter.AddAttribute('summaryBelow',XmlBoolToStr(FSummaryBelow));
  if FSummaryRight <> True then 
    AWriter.AddAttribute('summaryRight',XmlBoolToStr(FSummaryRight));
  if FShowOutlineSymbols <> True then 
    AWriter.AddAttribute('showOutlineSymbols',XmlBoolToStr(FShowOutlineSymbols));
end;

procedure TCT_OutlinePr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000004AA: FApplyStyles := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000507: FSummaryBelow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000050C: FSummaryRight := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000078A: FShowOutlineSymbols := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_OutlinePr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FApplyStyles := False;
  FSummaryBelow := True;
  FSummaryRight := True;
  FShowOutlineSymbols := True;
end;

destructor TCT_OutlinePr.Destroy;
begin
end;

procedure TCT_OutlinePr.Clear;
begin
  FAssigneds := [];
  FApplyStyles := False;
  FSummaryBelow := True;
  FSummaryRight := True;
  FShowOutlineSymbols := True;
end;

{ TCT_PageSetUpPr }

function  TCT_PageSetUpPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAutoPageBreaks <> True then 
    Inc(AttrsAssigned);
  if FFitToPage <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_PageSetUpPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_PageSetUpPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAutoPageBreaks <> True then 
    AWriter.AddAttribute('autoPageBreaks',XmlBoolToStr(FAutoPageBreaks));
  if FFitToPage <> False then 
    AWriter.AddAttribute('fitToPage',XmlBoolToStr(FFitToPage));
end;

procedure TCT_PageSetUpPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000058E: FAutoPageBreaks := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000383: FFitToPage := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_PageSetUpPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FAutoPageBreaks := True;
  FFitToPage := False;
end;

destructor TCT_PageSetUpPr.Destroy;
begin
end;

procedure TCT_PageSetUpPr.Clear;
begin
  FAssigneds := [];
  FAutoPageBreaks := True;
  FFitToPage := False;
end;

{ TCT_SheetView }

function  TCT_SheetView.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FWindowProtection <> False then 
    Inc(AttrsAssigned);
  if FShowFormulas <> False then 
    Inc(AttrsAssigned);
  if FShowGridLines <> True then 
    Inc(AttrsAssigned);
  if FShowRowColHeaders <> True then 
    Inc(AttrsAssigned);
  if FShowZeros <> True then 
    Inc(AttrsAssigned);
  if FRightToLeft <> False then 
    Inc(AttrsAssigned);
  if FTabSelected <> False then 
    Inc(AttrsAssigned);
  if FShowRuler <> True then 
    Inc(AttrsAssigned);
  if FShowOutlineSymbols <> True then 
    Inc(AttrsAssigned);
  if FDefaultGridColor <> True then 
    Inc(AttrsAssigned);
  if FShowWhiteSpace <> True then 
    Inc(AttrsAssigned);
  if FView <> stsvtNormal then 
    Inc(AttrsAssigned);
  if FTopLeftCell <> '' then 
    Inc(AttrsAssigned);
  if FColorId <> 64 then 
    Inc(AttrsAssigned);
  if FZoomScale <> 100 then 
    Inc(AttrsAssigned);
  if FZoomScaleNormal <> 0 then 
    Inc(AttrsAssigned);
  if FZoomScaleSheetLayoutView <> 0 then 
    Inc(AttrsAssigned);
  if FZoomScalePageLayoutView <> 0 then 
    Inc(AttrsAssigned);
  if FWorkbookViewId <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FPane.CheckAssigned);
  Inc(ElemsAssigned,FSelectionXpgList.CheckAssigned);
  Inc(ElemsAssigned,FPivotSelectionXpgList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_SheetView.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001A4: Result := FPane;
    $000003C6: Result := FSelectionXpgList.Add;
    $000005D8: Result := FPivotSelectionXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SheetView.Write(AWriter: TXpgWriteXML);
begin
  if FPane.Assigned then
  begin
    FPane.WriteAttributes(AWriter);
    AWriter.SimpleTag('pane');
  end;
  FSelectionXpgList.Write(AWriter,'selection');
  FPivotSelectionXpgList.Write(AWriter,'pivotSelection');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_SheetView.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FWindowProtection <> False then
    AWriter.AddAttribute('windowProtection',XmlBoolToStr(FWindowProtection));
  if FShowFormulas <> False then
    AWriter.AddAttribute('showFormulas',XmlBoolToStr(FShowFormulas));
  if FShowGridLines <> True then
    AWriter.AddAttribute('showGridLines',XmlBoolToStr(FShowGridLines));
  if FShowRowColHeaders <> True then
    AWriter.AddAttribute('showRowColHeaders',XmlBoolToStr(FShowRowColHeaders));
  if FShowZeros <> True then
    AWriter.AddAttribute('showZeros',XmlBoolToStr(FShowZeros));
  if FRightToLeft <> False then
    AWriter.AddAttribute('rightToLeft',XmlBoolToStr(FRightToLeft));
  if FTabSelected <> False then
    AWriter.AddAttribute('tabSelected',XmlBoolToStr(FTabSelected));
  if FShowRuler <> True then
    AWriter.AddAttribute('showRuler',XmlBoolToStr(FShowRuler));
  if FShowOutlineSymbols <> True then
    AWriter.AddAttribute('showOutlineSymbols',XmlBoolToStr(FShowOutlineSymbols));
  if FDefaultGridColor <> True then
    AWriter.AddAttribute('defaultGridColor',XmlBoolToStr(FDefaultGridColor));
  if FShowWhiteSpace <> True then
    AWriter.AddAttribute('showWhiteSpace',XmlBoolToStr(FShowWhiteSpace));
  if FView <> stsvtNormal then
    AWriter.AddAttribute('view',StrTST_SheetViewType[Integer(FView)]);
  if FTopLeftCell <> '' then
    AWriter.AddAttribute('topLeftCell',FTopLeftCell);
  if FColorId <> 64 then
    AWriter.AddAttribute('colorId',XmlIntToStr(FColorId));
  if FZoomScale <> 100 then
    AWriter.AddAttribute('zoomScale',XmlIntToStr(FZoomScale));
  if FZoomScaleNormal <> 0 then
    AWriter.AddAttribute('zoomScaleNormal',XmlIntToStr(FZoomScaleNormal));
  if FZoomScaleSheetLayoutView <> 0 then
    AWriter.AddAttribute('zoomScaleSheetLayoutView',XmlIntToStr(FZoomScaleSheetLayoutView));
  if FZoomScalePageLayoutView <> 0 then
    AWriter.AddAttribute('zoomScalePageLayoutView',XmlIntToStr(FZoomScalePageLayoutView));
  AWriter.AddAttribute('workbookViewId',XmlIntToStr(FWorkbookViewId));
end;

procedure TCT_SheetView.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000006BF: FWindowProtection := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000050A: FShowFormulas := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000542: FShowGridLines := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006D3: FShowRowColHeaders := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003D4: FShowZeros := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000046C: FRightToLeft := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000460: FTabSelected := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003CB: FShowRuler := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000078A: FShowOutlineSymbols := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000066A: FDefaultGridColor := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005AE: FShowWhiteSpace := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001BB: FView := TST_SheetViewType(StrToEnum('stsvt' + AAttributes.Values[i]));
      $0000045E: FTopLeftCell := AAttributes.Values[i];
      $000002CC: FColorId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003AD: FZoomScale := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000616: FZoomScaleNormal := XmlStrToIntDef(AAttributes.Values[i],0);
      $000009BF: FZoomScaleSheetLayoutView := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000943: FZoomScalePageLayoutView := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005B6: FWorkbookViewId := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SheetView.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 19;
  FPane := TCT_Pane.Create(FOwner);
  FSelectionXpgList := TCT_SelectionXpgList.Create(FOwner);
  FPivotSelectionXpgList := TCT_PivotSelectionXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
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
  FView := stsvtNormal;
  FColorId := 64;
  FZoomScale := 100;
  FZoomScaleNormal := 0;
  FZoomScaleSheetLayoutView := 0;
  FZoomScalePageLayoutView := 0;
end;

destructor TCT_SheetView.Destroy;
begin
  FPane.Free;
  FSelectionXpgList.Free;
  FPivotSelectionXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_SheetView.Clear;
begin
  FAssigneds := [];
  FPane.Clear;
  FSelectionXpgList.Clear;
  FPivotSelectionXpgList.Clear;
  FExtLst.Clear;
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
  FView := stsvtNormal;
  FTopLeftCell := '';
  FColorId := 64;
  FZoomScale := 100;
  FZoomScaleNormal := 0;
  FZoomScaleSheetLayoutView := 0;
  FZoomScalePageLayoutView := 0;
  FWorkbookViewId := 0;
end;

{ TCT_SheetViewXpgList }

function  TCT_SheetViewXpgList.GetItems(Index: integer): TCT_SheetView;
begin
  Result := TCT_SheetView(inherited Items[Index]);
end;

function  TCT_SheetViewXpgList.Add: TCT_SheetView;
begin
  Result := TCT_SheetView.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_SheetViewXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_SheetViewXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
  begin
//    if xaAttributes in Items[i].FAssigneds then  // TODO There are required attribut that must be written.
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

{ TCT_CustomSheetView }

function  TCT_CustomSheetView.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FGuid <> '' then 
    Inc(AttrsAssigned);
  if FScale <> 100 then 
    Inc(AttrsAssigned);
  if FColorId <> 64 then 
    Inc(AttrsAssigned);
  if FShowPageBreaks <> False then 
    Inc(AttrsAssigned);
  if FShowFormulas <> False then 
    Inc(AttrsAssigned);
  if FShowGridLines <> True then 
    Inc(AttrsAssigned);
  if FShowRowCol <> True then 
    Inc(AttrsAssigned);
  if FOutlineSymbols <> True then 
    Inc(AttrsAssigned);
  if FZeroValues <> True then 
    Inc(AttrsAssigned);
  if FFitToPage <> False then 
    Inc(AttrsAssigned);
  if FPrintArea <> False then 
    Inc(AttrsAssigned);
  if FFilter <> False then 
    Inc(AttrsAssigned);
  if FShowAutoFilter <> False then 
    Inc(AttrsAssigned);
  if FHiddenRows <> False then 
    Inc(AttrsAssigned);
  if FHiddenColumns <> False then 
    Inc(AttrsAssigned);
  if FState <> stssVisible then 
    Inc(AttrsAssigned);
  if FFilterUnique <> False then 
    Inc(AttrsAssigned);
  if FView <> stsvtNormal then 
    Inc(AttrsAssigned);
  if FShowRuler <> True then 
    Inc(AttrsAssigned);
  if FTopLeftCell <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FPane.CheckAssigned);
  Inc(ElemsAssigned,FSelection.CheckAssigned);
  Inc(ElemsAssigned,FRowBreaks.CheckAssigned);
  Inc(ElemsAssigned,FColBreaks.CheckAssigned);
  Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  Inc(ElemsAssigned,FPrintOptions.CheckAssigned);
  Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  Inc(ElemsAssigned,FAutoFilter.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CustomSheetView.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000001A4: Result := FPane;
    $000003C6: Result := FSelection;
    $000003B0: Result := FRowBreaks;
    $00000396: Result := FColBreaks;
    $0000046E: Result := FPageMargins;
    $00000519: Result := FPrintOptions;
    $000003AE: Result := FPageSetup;
    $000004D8: Result := FHeaderFooter;
    $0000041F: Result := FAutoFilter;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomSheetView.Write(AWriter: TXpgWriteXML);
begin
  if FPane.Assigned then 
  begin
    FPane.WriteAttributes(AWriter);
    AWriter.SimpleTag('pane');
  end;
  if FSelection.Assigned then 
  begin
    FSelection.WriteAttributes(AWriter);
    AWriter.SimpleTag('selection');
  end;
  if FRowBreaks.Assigned then 
  begin
    FRowBreaks.WriteAttributes(AWriter);
    if xaElements in FRowBreaks.FAssigneds then 
    begin
      AWriter.BeginTag('rowBreaks');
      FRowBreaks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rowBreaks');
  end;
  if FColBreaks.Assigned then 
  begin
    FColBreaks.WriteAttributes(AWriter);
    if xaElements in FColBreaks.FAssigneds then 
    begin
      AWriter.BeginTag('colBreaks');
      FColBreaks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colBreaks');
  end;
  if FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageMargins');
  end;
  if FPrintOptions.Assigned then 
  begin
    FPrintOptions.WriteAttributes(AWriter);
    AWriter.SimpleTag('printOptions');
  end;
  if FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetup');
  end;
  if FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.FAssigneds then 
    begin
      AWriter.BeginTag('headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('headerFooter');
  end;
  if FAutoFilter.Assigned then 
  begin
    FAutoFilter.WriteAttributes(AWriter);
    if xaElements in FAutoFilter.FAssigneds then 
    begin
      AWriter.BeginTag('autoFilter');
      FAutoFilter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('autoFilter');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CustomSheetView.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('guid',FGuid);
  if FScale <> 100 then 
    AWriter.AddAttribute('scale',XmlIntToStr(FScale));
  if FColorId <> 64 then 
    AWriter.AddAttribute('colorId',XmlIntToStr(FColorId));
  if FShowPageBreaks <> False then 
    AWriter.AddAttribute('showPageBreaks',XmlBoolToStr(FShowPageBreaks));
  if FShowFormulas <> False then 
    AWriter.AddAttribute('showFormulas',XmlBoolToStr(FShowFormulas));
  if FShowGridLines <> True then 
    AWriter.AddAttribute('showGridLines',XmlBoolToStr(FShowGridLines));
  if FShowRowCol <> True then 
    AWriter.AddAttribute('showRowCol',XmlBoolToStr(FShowRowCol));
  if FOutlineSymbols <> True then 
    AWriter.AddAttribute('outlineSymbols',XmlBoolToStr(FOutlineSymbols));
  if FZeroValues <> True then 
    AWriter.AddAttribute('zeroValues',XmlBoolToStr(FZeroValues));
  if FFitToPage <> False then 
    AWriter.AddAttribute('fitToPage',XmlBoolToStr(FFitToPage));
  if FPrintArea <> False then 
    AWriter.AddAttribute('printArea',XmlBoolToStr(FPrintArea));
  if FFilter <> False then 
    AWriter.AddAttribute('filter',XmlBoolToStr(FFilter));
  if FShowAutoFilter <> False then 
    AWriter.AddAttribute('showAutoFilter',XmlBoolToStr(FShowAutoFilter));
  if FHiddenRows <> False then 
    AWriter.AddAttribute('hiddenRows',XmlBoolToStr(FHiddenRows));
  if FHiddenColumns <> False then 
    AWriter.AddAttribute('hiddenColumns',XmlBoolToStr(FHiddenColumns));
  if FState <> stssVisible then 
    AWriter.AddAttribute('state',StrTST_SheetState[Integer(FState)]);
  if FFilterUnique <> False then 
    AWriter.AddAttribute('filterUnique',XmlBoolToStr(FFilterUnique));
  if FView <> stsvtNormal then 
    AWriter.AddAttribute('view',StrTST_SheetViewType[Integer(FView)]);
  if FShowRuler <> True then 
    AWriter.AddAttribute('showRuler',XmlBoolToStr(FShowRuler));
  if FTopLeftCell <> '' then 
    AWriter.AddAttribute('topLeftCell',FTopLeftCell);
end;

procedure TCT_CustomSheetView.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashB[i] of
      $72D34C6D: FGuid := AAttributes.Values[i];
      $E3CD1FC8: FScale := XmlStrToIntDef(AAttributes.Values[i],0);
      $C08847EC: FColorId := XmlStrToIntDef(AAttributes.Values[i],0);
      $15C3DBDC: FShowPageBreaks := XmlStrToBoolDef(AAttributes.Values[i],False);
      $F5F64EC8: FShowFormulas := XmlStrToBoolDef(AAttributes.Values[i],False);
      $850FA8A0: FShowGridLines := XmlStrToBoolDef(AAttributes.Values[i],False);
      $D2A80CEB: FShowRowCol := XmlStrToBoolDef(AAttributes.Values[i],False);
      $24CA59BB: FOutlineSymbols := XmlStrToBoolDef(AAttributes.Values[i],False);
      $524E199E: FZeroValues := XmlStrToBoolDef(AAttributes.Values[i],False);
      $A5D14157: FFitToPage := XmlStrToBoolDef(AAttributes.Values[i],False);
      $9187ED5A: FPrintArea := XmlStrToBoolDef(AAttributes.Values[i],False);
      $1F1C5E28: FFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $AE978826: FShowAutoFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $52CE0581: FHiddenRows := XmlStrToBoolDef(AAttributes.Values[i],False);
      $BAC60055: FHiddenColumns := XmlStrToBoolDef(AAttributes.Values[i],False);
      $6B23329F: FState := TST_SheetState(StrToEnum('stss' + AAttributes.Values[i]));
      $BF29EE59: FFilterUnique := XmlStrToBoolDef(AAttributes.Values[i],False);
      $E6369937: FView := TST_SheetViewType(StrToEnum('stsvt' + AAttributes.Values[i]));
      $F5DE3B43: FShowRuler := XmlStrToBoolDef(AAttributes.Values[i],False);
      $6435B76C: FTopLeftCell := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CustomSheetView.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 10;
  FAttributeCount := 20;
  FPane := TCT_Pane.Create(FOwner);
  FSelection := TCT_Selection.Create(FOwner);
  FRowBreaks := TCT_PageBreak.Create(FOwner);
  FColBreaks := TCT_PageBreak.Create(FOwner);
  FPageMargins := TCT_PageMargins.Create(FOwner);
  FPrintOptions := TCT_PrintOptions.Create(FOwner);
  FPageSetup := TCT_PageSetup.Create(FOwner);
  FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
  FAutoFilter := TCT_AutoFilter.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
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
  FState := stssVisible;
  FFilterUnique := False;
  FView := stsvtNormal;
  FShowRuler := True;
end;

destructor TCT_CustomSheetView.Destroy;
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
  FExtLst.Free;
end;

procedure TCT_CustomSheetView.Clear;
begin
  FAssigneds := [];
  FPane.Clear;
  FSelection.Clear;
  FRowBreaks.Clear;
  FColBreaks.Clear;
  FPageMargins.Clear;
  FPrintOptions.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FAutoFilter.Clear;
  FExtLst.Clear;
  FGuid := '';
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
  FState := stssVisible;
  FFilterUnique := False;
  FView := stsvtNormal;
  FShowRuler := True;
  FTopLeftCell := '';
end;

{ TCT_CustomSheetViewXpgList }

function  TCT_CustomSheetViewXpgList.GetItems(Index: integer): TCT_CustomSheetView;
begin
  Result := TCT_CustomSheetView(inherited Items[Index]);
end;

function  TCT_CustomSheetViewXpgList.Add: TCT_CustomSheetView;
begin
  Result := TCT_CustomSheetView.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CustomSheetViewXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CustomSheetViewXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_OleObject }

function  TCT_OleObject.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FProgId <> '' then 
    Inc(AttrsAssigned);
  if FDvAspect <> stdaDVASPECT_CONTENT then 
    Inc(AttrsAssigned);
  if FLink <> '' then 
    Inc(AttrsAssigned);
  if Integer(FOleUpdate) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FAutoLoad <> False then 
    Inc(AttrsAssigned);
  if FShapeId <> 0 then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_OleObject.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_OleObject.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FProgId <> '' then 
    AWriter.AddAttribute('progId',FProgId);
  if FDvAspect <> stdaDVASPECT_CONTENT then 
    AWriter.AddAttribute('dvAspect',StrTST_DvAspect[Integer(FDvAspect)]);
  if FLink <> '' then 
    AWriter.AddAttribute('link',FLink);
  if Integer(FOleUpdate) <> XPG_UNKNOWN_ENUM then 
    AWriter.AddAttribute('oleUpdate',StrTST_OleUpdate[Integer(FOleUpdate)]);
  if FAutoLoad <> False then 
    AWriter.AddAttribute('autoLoad',XmlBoolToStr(FAutoLoad));
  AWriter.AddAttribute('shapeId',XmlIntToStr(FShapeId));
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_OleObject.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000265: FProgId := AAttributes.Values[i];
      $0000033A: FDvAspect := TST_DvAspect(StrToEnum('stda' + AAttributes.Values[i]));
      $000001AE: FLink := AAttributes.Values[i];
      $000003A3: FOleUpdate := TST_OleUpdate(StrToEnum('stou' + AAttributes.Values[i]));
      $00000339: FAutoLoad := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002BE: FShapeId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_OleObject.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 7;
  FDvAspect := stdaDVASPECT_CONTENT;
  FOleUpdate := TST_OleUpdate(XPG_UNKNOWN_ENUM);
  FAutoLoad := False;
end;

destructor TCT_OleObject.Destroy;
begin
end;

procedure TCT_OleObject.Clear;
begin
  FAssigneds := [];
  FProgId := '';
  FDvAspect := stdaDVASPECT_CONTENT;
  FLink := '';
  FOleUpdate := TST_OleUpdate(XPG_UNKNOWN_ENUM);
  FAutoLoad := False;
  FShapeId := 0;
  FR_Id := '';
end;

{ TCT_OleObjectXpgList }

function  TCT_OleObjectXpgList.GetItems(Index: integer): TCT_OleObject;
begin
  Result := TCT_OleObject(inherited Items[Index]);
end;

function  TCT_OleObjectXpgList.Add: TCT_OleObject;
begin
  Result := TCT_OleObject.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OleObjectXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OleObjectXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Col }

function  TCT_Col.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMin <> 0 then 
    Inc(AttrsAssigned);
  if FMax <> 0 then 
    Inc(AttrsAssigned);
  if FWidth <> 0 then 
    Inc(AttrsAssigned);
  if FStyle <> 0 then 
    Inc(AttrsAssigned);
  if FHidden <> False then 
    Inc(AttrsAssigned);
  if FBestFit <> False then 
    Inc(AttrsAssigned);
  if FCustomWidth <> False then 
    Inc(AttrsAssigned);
  if FPhonetic <> False then 
    Inc(AttrsAssigned);
  if FOutlineLevel <> 0 then 
    Inc(AttrsAssigned);
  if FCollapsed <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Col.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Col.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('min',XmlIntToStr(FMin));
  AWriter.AddAttribute('max',XmlIntToStr(FMax));
  if FWidth <> 0 then 
    AWriter.AddAttribute('width',XmlFloatToStr(FWidth));
  if FStyle <> 0 then 
    AWriter.AddAttribute('style',XmlIntToStr(FStyle));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
  if FBestFit <> False then 
    AWriter.AddAttribute('bestFit',XmlBoolToStr(FBestFit));
  if FCustomWidth <> False then 
    AWriter.AddAttribute('customWidth',XmlBoolToStr(FCustomWidth));
  if FPhonetic <> False then 
    AWriter.AddAttribute('phonetic',XmlBoolToStr(FPhonetic));
  if FOutlineLevel <> 0 then 
    AWriter.AddAttribute('outlineLevel',XmlIntToStr(FOutlineLevel));
  if FCollapsed <> False then 
    AWriter.AddAttribute('collapsed',XmlBoolToStr(FCollapsed));
end;

procedure TCT_Col.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000144: FMin := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000146: FMax := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000220: FWidth := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000231: FStyle := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002D1: FBestFit := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000049B: FCustomWidth := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000035A: FPhonetic := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F8: FOutlineLevel := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003B7: FCollapsed := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Col.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 10;
  FStyle := 0;
  FHidden := False;
  FBestFit := False;
  FCustomWidth := False;
  FPhonetic := False;
  FOutlineLevel := 0;
  FCollapsed := False;
end;

destructor TCT_Col.Destroy;
begin
end;

procedure TCT_Col.Clear;
begin
  FAssigneds := [];
  FMin := 0;
  FMax := 0;
  FWidth := 0;
  FStyle := 0;
  FHidden := False;
  FBestFit := False;
  FCustomWidth := False;
  FPhonetic := False;
  FOutlineLevel := 0;
  FCollapsed := False;
end;

{ TCT_ColXpgList }

function  TCT_ColXpgList.GetItems(Index: integer): TCT_Col;
begin
  Result := TCT_Col(inherited Items[Index]);
end;

function  TCT_ColXpgList.Add: TCT_Col;
begin
  Result := TCT_Col.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ColXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ColXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Row }

function  TCT_Row.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> 0 then 
    Inc(AttrsAssigned);
  if FSpansXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  if FS <> 0 then 
    Inc(AttrsAssigned);
  if FCustomFormat <> False then 
    Inc(AttrsAssigned);
  if FHt <> 0 then 
    Inc(AttrsAssigned);
  if FHidden <> False then 
    Inc(AttrsAssigned);
  if FCustomHeight <> False then 
    Inc(AttrsAssigned);
  if FOutlineLevel <> 0 then 
    Inc(AttrsAssigned);
  if FCollapsed <> False then 
    Inc(AttrsAssigned);
  if FThickTop <> False then 
    Inc(AttrsAssigned);
  if FThickBot <> False then 
    Inc(AttrsAssigned);
  if FPh <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FC.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Row.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  if AReader.QNameIsC then begin
    Result := FC;
    TCT_Cell(Result).Clear;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,AReader.QName);

//  QName := AReader.QName;
//  case CalcHash_A(QName) of
//    $00000063: begin
//      Result := FC;
//      TCT_Cell(Result).Clear;
//    end;
//    $00000284: Result := FExtLst;
//    else
//      FOwner.Errors.Error(xemUnknownElement,QName);
//  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Row.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteC) then
    while True do 
    begin
      WriteElement := False;
      FOnWriteC(FC,WriteElement);
      if WriteElement = False then 
        Break;
      FC.CheckAssigned;
      if FC.Assigned then 
      begin
        FC.WriteAttributes(AWriter);
        if xaElements in FC.FAssigneds then 
        begin
          AWriter.BeginTag('c');
          FC.Write(AWriter);
          AWriter.EndTag;
        end
        else 
          AWriter.SimpleTag('c');
        FC.Clear;
      end;
    end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Row.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FR <> 0 then 
    AWriter.AddAttribute('r',XmlIntToStr(FR));
  if FSpansXpgList.Count > 0 then 
    AWriter.AddAttribute('spans',FSpansXpgList.DelimitedText);
  if FS <> 0 then 
    AWriter.AddAttribute('s',XmlIntToStr(FS));
  if FCustomFormat <> False then 
    AWriter.AddAttribute('customFormat',XmlBoolToStr(FCustomFormat));
  if FHt <> 0 then 
    AWriter.AddAttribute('ht',XmlFloatToStr(FHt));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
  if FCustomHeight <> False then 
    AWriter.AddAttribute('customHeight',XmlBoolToStr(FCustomHeight));
  if FOutlineLevel <> 0 then 
    AWriter.AddAttribute('outlineLevel',XmlIntToStr(FOutlineLevel));
  if FCollapsed <> False then 
    AWriter.AddAttribute('collapsed',XmlBoolToStr(FCollapsed));
  if FThickTop <> False then 
    AWriter.AddAttribute('thickTop',XmlBoolToStr(FThickTop));
  if FThickBot <> False then 
    AWriter.AddAttribute('thickBot',XmlBoolToStr(FThickBot));
  if FPh <> False then 
    AWriter.AddAttribute('ph',XmlBoolToStr(FPh));
end;

procedure TCT_Row.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000072: FR := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000225: FSpansXpgList.DelimitedText := AAttributes.Values[i];
      $00000073: FS := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000504: FCustomFormat := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000DC: FHt := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F4: FCustomHeight := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F8: FOutlineLevel := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003B7: FCollapsed := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000346: FThickTop := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000338: FThickBot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000000D8: FPh := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

procedure TCT_Row.AfterTag;
begin
  if (FC.FAssigneds <> []) and System.Assigned(FOnReadC) then begin
    FOnReadC(FC);
    FC.Clear;
  end;
  FC.FAssigneds := [];
end;

constructor TCT_Row.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 12;
  FC := TCT_Cell.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FSpansXpgList := TStringXpgList.Create;
  FS := 0;
  FCustomFormat := False;
  FHidden := False;
  FCustomHeight := False;
  FOutlineLevel := 0;
  FCollapsed := False;
  FThickTop := False;
  FThickBot := False;
  FPh := False;
end;

destructor TCT_Row.Destroy;
begin
  FC.Free;
  FExtLst.Free;
  FSpansXpgList.Free;
end;

procedure TCT_Row.Clear;
begin
  FAssigneds := [];
  FC.Clear;
  FExtLst.Clear;
  FR := 0;
  FSpansXpgList.DelimitedText := '';
  FS := 0;
  FCustomFormat := False;
  FHt := 0;
  FHidden := False;
  FCustomHeight := False;
  FOutlineLevel := 0;
  FCollapsed := False;
  FThickTop := False;
  FThickBot := False;
  FPh := False;
end;

{ TCT_RowXpgList }

function  TCT_RowXpgList.GetItems(Index: integer): TCT_Row;
begin
  Result := TCT_Row(inherited Items[Index]);
end;

function  TCT_RowXpgList.Add: TCT_Row;
begin
  Result := TCT_Row.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_RowXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_RowXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ProtectedRange }

function  TCT_ProtectedRange.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPassword <> 0 then 
    Inc(AttrsAssigned);
  if FSqrefXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FSecurityDescriptor <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ProtectedRange.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ProtectedRange.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPassword <> 0 then 
    AWriter.AddAttribute('password',XmlIntToHexStr(FPassword));
  AWriter.AddAttribute('sqref',FSqrefXpgList.DelimitedText);
  AWriter.AddAttribute('name',FName);
  if FSecurityDescriptor <> '' then 
    AWriter.AddAttribute('securityDescriptor',FSecurityDescriptor);
end;

procedure TCT_ProtectedRange.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000373: FPassword := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000221: FSqrefXpgList.DelimitedText := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $00000797: FSecurityDescriptor := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_ProtectedRange.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FSqrefXpgList := TStringXpgList.Create;
end;

destructor TCT_ProtectedRange.Destroy;
begin
  FSqrefXpgList.Free;
end;

procedure TCT_ProtectedRange.Clear;
begin
  FAssigneds := [];
  FPassword := 0;
  FSqrefXpgList.DelimitedText := '';
  FName := '';
  FSecurityDescriptor := '';
end;

{ TCT_ProtectedRangeXpgList }

function  TCT_ProtectedRangeXpgList.GetItems(Index: integer): TCT_ProtectedRange;
begin
  Result := TCT_ProtectedRange(inherited Items[Index]);
end;

function  TCT_ProtectedRangeXpgList.Add: TCT_ProtectedRange;
begin
  Result := TCT_ProtectedRange.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ProtectedRangeXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ProtectedRangeXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Scenario }

function  TCT_Scenario.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FLocked <> False then 
    Inc(AttrsAssigned);
  if FHidden <> False then 
    Inc(AttrsAssigned);
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FUser <> '' then 
    Inc(AttrsAssigned);
  if FComment <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FInputCellsXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Scenario.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'inputCells' then 
    Result := FInputCellsXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Scenario.Write(AWriter: TXpgWriteXML);
begin
  FInputCellsXpgList.Write(AWriter,'inputCells');
end;

procedure TCT_Scenario.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  if FLocked <> False then 
    AWriter.AddAttribute('locked',XmlBoolToStr(FLocked));
  if FHidden <> False then 
    AWriter.AddAttribute('hidden',XmlBoolToStr(FHidden));
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if FUser <> '' then 
    AWriter.AddAttribute('user',FUser);
  if FComment <> '' then 
    AWriter.AddAttribute('comment',FComment);
end;

procedure TCT_Scenario.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000272: FLocked := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000026C: FHidden := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001BF: FUser := AAttributes.Values[i];
      $000002F3: FComment := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Scenario.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 6;
  FInputCellsXpgList := TCT_InputCellsXpgList.Create(FOwner);
  FLocked := False;
  FHidden := False;
end;

destructor TCT_Scenario.Destroy;
begin
  FInputCellsXpgList.Free;
end;

procedure TCT_Scenario.Clear;
begin
  FAssigneds := [];
  FInputCellsXpgList.Clear;
  FName := '';
  FLocked := False;
  FHidden := False;
  FCount := 0;
  FUser := '';
  FComment := '';
end;

{ TCT_ScenarioXpgList }

function  TCT_ScenarioXpgList.GetItems(Index: integer): TCT_Scenario;
begin
  Result := TCT_Scenario(inherited Items[Index]);
end;

function  TCT_ScenarioXpgList.Add: TCT_Scenario;
begin
  Result := TCT_Scenario.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ScenarioXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ScenarioXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DataRefs }

function  TCT_DataRefs.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FDataRefXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DataRefs.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'dataRef' then 
    Result := FDataRefXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DataRefs.Write(AWriter: TXpgWriteXML);
begin
  FDataRefXpgList.Write(AWriter,'dataRef');
end;

procedure TCT_DataRefs.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_DataRefs.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_DataRefs.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FDataRefXpgList := TCT_DataRefXpgList.Create(FOwner);
end;

destructor TCT_DataRefs.Destroy;
begin
  FDataRefXpgList.Free;
end;

procedure TCT_DataRefs.Clear;
begin
  FAssigneds := [];
  FDataRefXpgList.Clear;
  FCount := 0;
end;

{ TCT_MergeCell }

function  TCT_MergeCell.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_MergeCell.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_MergeCell.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ref',FRef);
end;

procedure TCT_MergeCell.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'ref' then 
    FRef := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_MergeCell.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_MergeCell.Destroy;
begin
end;

procedure TCT_MergeCell.Clear;
begin
  FAssigneds := [];
  FRef := '';
end;

{ TCT_MergeCellXpgList }

function  TCT_MergeCellXpgList.GetItems(Index: integer): TCT_MergeCell;
begin
  Result := TCT_MergeCell(inherited Items[Index]);
end;

function  TCT_MergeCellXpgList.Add: TCT_MergeCell;
begin
  Result := TCT_MergeCell.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_MergeCellXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_MergeCellXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CfRule }

function  TCT_CfRule.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  if FDxfId >= 0 then   // TODO zero is valid
    Inc(AttrsAssigned);
  if FPriority <> 0 then
    Inc(AttrsAssigned);
  if FStopIfTrue <> False then 
    Inc(AttrsAssigned);
  if FAboveAverage <> True then 
    Inc(AttrsAssigned);
  if FPercent <> False then 
    Inc(AttrsAssigned);
  if FBottom <> False then 
    Inc(AttrsAssigned);
  if Integer(FOperator) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FText <> '' then 
    Inc(AttrsAssigned);
  if Integer(FTimePeriod) <> XPG_UNKNOWN_ENUM then 
    Inc(AttrsAssigned);
  if FRank <> 0 then 
    Inc(AttrsAssigned);
  if FStdDev <> 0 then 
    Inc(AttrsAssigned);
  if FEqualAverage <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFormulaXpgList.Count);
  Inc(ElemsAssigned,FColorScale.CheckAssigned);
  Inc(ElemsAssigned,FDataBar.CheckAssigned);
  Inc(ElemsAssigned,FIconSet.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CfRule.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002F6: FFormulaXpgList.Add(AReader.Text);
    $00000407: Result := FColorScale;
    $000002AF: Result := FDataBar;
    $000002D5: Result := FIconSet;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CfRule.Write(AWriter: TXpgWriteXML);
begin
  FFormulaXpgList.WriteElements(AWriter,'formula');
  if FColorScale.Assigned then 
    if xaElements in FColorScale.FAssigneds then 
    begin
      AWriter.BeginTag('colorScale');
      FColorScale.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colorScale');
  if FDataBar.Assigned then 
  begin
    FDataBar.WriteAttributes(AWriter);
    if xaElements in FDataBar.FAssigneds then 
    begin
      AWriter.BeginTag('dataBar');
      FDataBar.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dataBar');
  end;
  if FIconSet.Assigned then 
  begin
    FIconSet.WriteAttributes(AWriter);
    if xaElements in FIconSet.FAssigneds then 
    begin
      AWriter.BeginTag('iconSet');
      FIconSet.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('iconSet');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_CfRule.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if Integer(FType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('type',StrTST_CfType[Integer(FType)]);
  if FDxfId >= 0 then        // TODO zero is valid
    AWriter.AddAttribute('dxfId',XmlIntToStr(FDxfId));
  AWriter.AddAttribute('priority',XmlIntToStr(FPriority));
  if FStopIfTrue <> False then
    AWriter.AddAttribute('stopIfTrue',XmlBoolToStr(FStopIfTrue));
  if FAboveAverage <> True then
    AWriter.AddAttribute('aboveAverage',XmlBoolToStr(FAboveAverage));
  if FPercent <> False then
    AWriter.AddAttribute('percent',XmlBoolToStr(FPercent));
  if FBottom <> False then
    AWriter.AddAttribute('bottom',XmlBoolToStr(FBottom));
  if Integer(FOperator) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('operator',StrTST_ConditionalFormattingOperator[Integer(FOperator)]);
  if FText <> '' then
    AWriter.AddAttribute('text',FText);
  if Integer(FTimePeriod) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('timePeriod',StrTST_TimePeriod[Integer(FTimePeriod)]);
  if FRank <> 0 then 
    AWriter.AddAttribute('rank',XmlIntToStr(FRank));
  if FStdDev <> 0 then 
    AWriter.AddAttribute('stdDev',XmlIntToStr(FStdDev));
  if FEqualAverage <> False then 
    AWriter.AddAttribute('equalAverage',XmlBoolToStr(FEqualAverage));
end;

procedure TCT_CfRule.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_CfType(StrToEnum('stct' + AAttributes.Values[i]));
      $000001EF: FDxfId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000382: FPriority := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000415: FStopIfTrue := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004C8: FAboveAverage := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F1: FPercent := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000295: FBottom := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000036C: FOperator := TST_ConditionalFormattingOperator(StrToEnum('stcfo' + AAttributes.Values[i]));
      $000001C5: FText := AAttributes.Values[i];
      $00000412: FTimePeriod := TST_TimePeriod(StrToEnum('sttp' + AAttributes.Values[i]));
      $000001AC: FRank := XmlStrToIntDef(AAttributes.Values[i],0);
      $0000026A: FStdDev := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004D3: FEqualAverage := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CfRule.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 13;
  FFormulaXpgList := TStringXpgList.Create;
  FColorScale := TCT_ColorScale.Create(FOwner);
  FDataBar := TCT_DataBar.Create(FOwner);
  FIconSet := TCT_IconSet.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FType := TST_CfType(XPG_UNKNOWN_ENUM);
  FDxfId := -1;  // TODO Zero is valid
  FStopIfTrue := False;
  FAboveAverage := True;
  FPercent := False;
  FBottom := False;
  FOperator := TST_ConditionalFormattingOperator(XPG_UNKNOWN_ENUM);
  FTimePeriod := TST_TimePeriod(XPG_UNKNOWN_ENUM);
  FEqualAverage := False;
end;

destructor TCT_CfRule.Destroy;
begin
  FFormulaXpgList.Free;
  FColorScale.Free;
  FDataBar.Free;
  FIconSet.Free;
  FExtLst.Free;
end;

procedure TCT_CfRule.Clear;
begin
  FAssigneds := [];
  FFormulaXpgList.DelimitedText := '';
  FColorScale.Clear;
  FDataBar.Clear;
  FIconSet.Clear;
  FExtLst.Clear;
  FType := TST_CfType(XPG_UNKNOWN_ENUM);
  FDxfId := -1; // TODO Must be possible to check if value is assigned, as zero is a valid number.
  FPriority := 0;
  FStopIfTrue := False;
  FAboveAverage := True;
  FPercent := False;
  FBottom := False;
  FOperator := TST_ConditionalFormattingOperator(XPG_UNKNOWN_ENUM);
  FText := '';
  FTimePeriod := TST_TimePeriod(XPG_UNKNOWN_ENUM);
  FRank := 0;
  FStdDev := 0;
  FEqualAverage := False;
end;

{ TCT_CfRuleXpgList }

function  TCT_CfRuleXpgList.GetItems(Index: integer): TCT_CfRule;
begin
  Result := TCT_CfRule(inherited Items[Index]);
end;

function  TCT_CfRuleXpgList.Add: TCT_CfRule;
begin
  Result := TCT_CfRule.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CfRuleXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CfRuleXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DataValidation }

function  TCT_DataValidation.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FType <> stdvtNone then 
    Inc(AttrsAssigned);
  if FErrorStyle <> stdvesStop then 
    Inc(AttrsAssigned);
  if FImeMode <> stdvimNoControl then 
    Inc(AttrsAssigned);
  if FOperator <> stdvoBetween then 
    Inc(AttrsAssigned);
  if FAllowBlank <> False then 
    Inc(AttrsAssigned);
  if FShowDropDown <> False then 
    Inc(AttrsAssigned);
  if FShowInputMessage <> False then 
    Inc(AttrsAssigned);
  if FShowErrorMessage <> False then 
    Inc(AttrsAssigned);
  if FErrorTitle <> '' then 
    Inc(AttrsAssigned);
  if FError <> '' then 
    Inc(AttrsAssigned);
  if FPromptTitle <> '' then 
    Inc(AttrsAssigned);
  if FPrompt <> '' then 
    Inc(AttrsAssigned);
  if FSqrefXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  if FFormula1 <> '' then 
    Inc(ElemsAssigned);
  if FFormula2 <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DataValidation.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000327: FFormula1 := AReader.Text;
    $00000328: FFormula2 := AReader.Text;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DataValidation.Write(AWriter: TXpgWriteXML);
begin
  if FFormula1 <> '' then 
    AWriter.SimpleTextTag('formula1',FFormula1);
  if FFormula2 <> '' then 
    AWriter.SimpleTextTag('formula2',FFormula2);
end;

procedure TCT_DataValidation.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FType <> stdvtNone then 
    AWriter.AddAttribute('type',StrTST_DataValidationType[Integer(FType)]);
  if FErrorStyle <> stdvesStop then 
    AWriter.AddAttribute('errorStyle',StrTST_DataValidationErrorStyle[Integer(FErrorStyle)]);
  if FImeMode <> stdvimNoControl then 
    AWriter.AddAttribute('imeMode',StrTST_DataValidationImeMode[Integer(FImeMode)]);
  if FOperator <> stdvoBetween then 
    AWriter.AddAttribute('operator',StrTST_DataValidationOperator[Integer(FOperator)]);
  if FAllowBlank <> False then 
    AWriter.AddAttribute('allowBlank',XmlBoolToStr(FAllowBlank));
  if FShowDropDown <> False then 
    AWriter.AddAttribute('showDropDown',XmlBoolToStr(FShowDropDown));
  if FShowInputMessage <> False then 
    AWriter.AddAttribute('showInputMessage',XmlBoolToStr(FShowInputMessage));
  if FShowErrorMessage <> False then 
    AWriter.AddAttribute('showErrorMessage',XmlBoolToStr(FShowErrorMessage));
  if FErrorTitle <> '' then 
    AWriter.AddAttribute('errorTitle',FErrorTitle);
  if FError <> '' then 
    AWriter.AddAttribute('error',FError);
  if FPromptTitle <> '' then 
    AWriter.AddAttribute('promptTitle',FPromptTitle);
  if FPrompt <> '' then 
    AWriter.AddAttribute('prompt',FPrompt);
  AWriter.AddAttribute('sqref',FSqrefXpgList.DelimitedText);
end;

procedure TCT_DataValidation.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001C2: FType := TST_DataValidationType(StrToEnum('stdvt' + AAttributes.Values[i]));
      $0000043B: FErrorStyle := TST_DataValidationErrorStyle(StrToEnum('stdves' + AAttributes.Values[i]));
      $000002C0: FImeMode := TST_DataValidationImeMode(StrToEnum('stdvim' + AAttributes.Values[i]));
      $0000036C: FOperator := TST_DataValidationOperator(StrToEnum('stdvo' + AAttributes.Values[i]));
      $00000407: FAllowBlank := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004EE: FShowDropDown := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000696: FShowInputMessage := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000690: FShowErrorMessage := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000042C: FErrorTitle := AAttributes.Values[i];
      $0000022A: FError := AAttributes.Values[i];
      $000004A4: FPromptTitle := AAttributes.Values[i];
      $000002A2: FPrompt := AAttributes.Values[i];
      $00000221: FSqrefXpgList.DelimitedText := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DataValidation.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 13;
  FType := stdvtNone;
  FErrorStyle := stdvesStop;
  FImeMode := stdvimNoControl;
  FOperator := stdvoBetween;
  FAllowBlank := False;
  FShowDropDown := False;
  FShowInputMessage := False;
  FShowErrorMessage := False;
  FSqrefXpgList := TStringXpgList.Create;
end;

destructor TCT_DataValidation.Destroy;
begin
  FSqrefXpgList.Free;
end;

procedure TCT_DataValidation.Clear;
begin
  FAssigneds := [];
  FFormula1 := '';
  FFormula2 := '';
  FType := stdvtNone;
  FErrorStyle := stdvesStop;
  FImeMode := stdvimNoControl;
  FOperator := stdvoBetween;
  FAllowBlank := False;
  FShowDropDown := False;
  FShowInputMessage := False;
  FShowErrorMessage := False;
  FErrorTitle := '';
  FError := '';
  FPromptTitle := '';
  FPrompt := '';
  FSqrefXpgList.DelimitedText := '';
end;

{ TCT_DataValidationXpgList }

function  TCT_DataValidationXpgList.GetItems(Index: integer): TCT_DataValidation;
begin
  Result := TCT_DataValidation(inherited Items[Index]);
end;

function  TCT_DataValidationXpgList.Add: TCT_DataValidation;
begin
  Result := TCT_DataValidation.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DataValidationXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DataValidationXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Hyperlink }

function  TCT_Hyperlink.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  if FLocation <> '' then 
    Inc(AttrsAssigned);
  if FTooltip <> '' then 
    Inc(AttrsAssigned);
  if FDisplay <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Hyperlink.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Hyperlink.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ref',FRef);
  if FR_Id <> '' then 
    AWriter.AddAttribute('r:id',FR_Id);
  if FLocation <> '' then 
    AWriter.AddAttribute('location',FLocation);
  if FTooltip <> '' then 
    AWriter.AddAttribute('tooltip',FTooltip);
  if FDisplay <> '' then 
    AWriter.AddAttribute('display',FDisplay);
end;

procedure TCT_Hyperlink.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000013D: FRef := AAttributes.Values[i];
      $00000179: FR_Id := AAttributes.Values[i];
      $00000359: FLocation := AAttributes.Values[i];
      $0000030B: FTooltip := AAttributes.Values[i];
      $000002F6: FDisplay := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Hyperlink.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
end;

destructor TCT_Hyperlink.Destroy;
begin
end;

procedure TCT_Hyperlink.Clear;
begin
  FAssigneds := [];
  FRef := '';
  FR_Id := '';
  FLocation := '';
  FTooltip := '';
  FDisplay := '';
end;

{ TCT_HyperlinkXpgList }

function  TCT_HyperlinkXpgList.GetItems(Index: integer): TCT_Hyperlink;
begin
  Result := TCT_Hyperlink(inherited Items[Index]);
end;

function  TCT_HyperlinkXpgList.Add: TCT_Hyperlink;
begin
  Result := TCT_Hyperlink.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_HyperlinkXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_HyperlinkXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CustomProperty }

function  TCT_CustomProperty.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CustomProperty.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CustomProperty.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_CustomProperty.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000179: FR_Id := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CustomProperty.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
end;

destructor TCT_CustomProperty.Destroy;
begin
end;

procedure TCT_CustomProperty.Clear;
begin
  FAssigneds := [];
  FName := '';
  FR_Id := '';
end;

{ TCT_CustomPropertyXpgList }

function  TCT_CustomPropertyXpgList.GetItems(Index: integer): TCT_CustomProperty;
begin
  Result := TCT_CustomProperty(inherited Items[Index]);
end;

function  TCT_CustomPropertyXpgList.Add: TCT_CustomProperty;
begin
  Result := TCT_CustomProperty.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CustomPropertyXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CustomPropertyXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CellWatch }

function  TCT_CellWatch.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CellWatch.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CellWatch.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r',FR);
end;

procedure TCT_CellWatch.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'r' then 
    FR := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CellWatch.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_CellWatch.Destroy;
begin
end;

procedure TCT_CellWatch.Clear;
begin
  FAssigneds := [];
  FR := '';
end;

{ TCT_CellWatchXpgList }

function  TCT_CellWatchXpgList.GetItems(Index: integer): TCT_CellWatch;
begin
  Result := TCT_CellWatch(inherited Items[Index]);
end;

function  TCT_CellWatchXpgList.Add: TCT_CellWatch;
begin
  Result := TCT_CellWatch.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CellWatchXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CellWatchXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_IgnoredError }

function  TCT_IgnoredError.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FSqrefXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  if FEvalError <> False then 
    Inc(AttrsAssigned);
  if FTwoDigitTextYear <> False then 
    Inc(AttrsAssigned);
  if FNumberStoredAsText <> False then 
    Inc(AttrsAssigned);
  if FFormula <> False then 
    Inc(AttrsAssigned);
  if FFormulaRange <> False then 
    Inc(AttrsAssigned);
  if FUnlockedFormula <> False then 
    Inc(AttrsAssigned);
  if FEmptyCellReference <> False then 
    Inc(AttrsAssigned);
  if FListDataValidation <> False then 
    Inc(AttrsAssigned);
  if FCalculatedColumn <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_IgnoredError.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_IgnoredError.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('sqref',FSqrefXpgList.DelimitedText);
  if FEvalError <> False then 
    AWriter.AddAttribute('evalError',XmlBoolToStr(FEvalError));
  if FTwoDigitTextYear <> False then 
    AWriter.AddAttribute('twoDigitTextYear',XmlBoolToStr(FTwoDigitTextYear));
  if FNumberStoredAsText <> False then 
    AWriter.AddAttribute('numberStoredAsText',XmlBoolToStr(FNumberStoredAsText));
  if FFormula <> False then 
    AWriter.AddAttribute('formula',XmlBoolToStr(FFormula));
  if FFormulaRange <> False then 
    AWriter.AddAttribute('formulaRange',XmlBoolToStr(FFormulaRange));
  if FUnlockedFormula <> False then 
    AWriter.AddAttribute('unlockedFormula',XmlBoolToStr(FUnlockedFormula));
  if FEmptyCellReference <> False then 
    AWriter.AddAttribute('emptyCellReference',XmlBoolToStr(FEmptyCellReference));
  if FListDataValidation <> False then 
    AWriter.AddAttribute('listDataValidation',XmlBoolToStr(FListDataValidation));
  if FCalculatedColumn <> False then 
    AWriter.AddAttribute('calculatedColumn',XmlBoolToStr(FCalculatedColumn));
end;

procedure TCT_IgnoredError.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000221: FSqrefXpgList.DelimitedText := AAttributes.Values[i];
      $000003B2: FEvalError := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000681: FTwoDigitTextYear := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000753: FNumberStoredAsText := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F6: FFormula := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004E3: FFormulaRange := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000062B: FUnlockedFormula := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000073E: FEmptyCellReference := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000741: FListDataValidation := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000680: FCalculatedColumn := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_IgnoredError.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 10;
  FSqrefXpgList := TStringXpgList.Create;
  FEvalError := False;
  FTwoDigitTextYear := False;
  FNumberStoredAsText := False;
  FFormula := False;
  FFormulaRange := False;
  FUnlockedFormula := False;
  FEmptyCellReference := False;
  FListDataValidation := False;
  FCalculatedColumn := False;
end;

destructor TCT_IgnoredError.Destroy;
begin
  FSqrefXpgList.Free;
end;

procedure TCT_IgnoredError.Clear;
begin
  FAssigneds := [];
  FSqrefXpgList.DelimitedText := '';
  FEvalError := False;
  FTwoDigitTextYear := False;
  FNumberStoredAsText := False;
  FFormula := False;
  FFormulaRange := False;
  FUnlockedFormula := False;
  FEmptyCellReference := False;
  FListDataValidation := False;
  FCalculatedColumn := False;
end;

{ TCT_IgnoredErrorXpgList }

function  TCT_IgnoredErrorXpgList.GetItems(Index: integer): TCT_IgnoredError;
begin
  Result := TCT_IgnoredError(inherited Items[Index]);
end;

function  TCT_IgnoredErrorXpgList.Add: TCT_IgnoredError;
begin
  Result := TCT_IgnoredError.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_IgnoredErrorXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_IgnoredErrorXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_CellSmartTags }

function  TCT_CellSmartTags.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCellSmartTagXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CellSmartTags.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cellSmartTag' then 
    Result := FCellSmartTagXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellSmartTags.Write(AWriter: TXpgWriteXML);
begin
  FCellSmartTagXpgList.Write(AWriter,'cellSmartTag');
end;

procedure TCT_CellSmartTags.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r',FR);
end;

procedure TCT_CellSmartTags.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'r' then 
    FR := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CellSmartTags.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCellSmartTagXpgList := TCT_CellSmartTagXpgList.Create(FOwner);
end;

destructor TCT_CellSmartTags.Destroy;
begin
  FCellSmartTagXpgList.Free;
end;

procedure TCT_CellSmartTags.Clear;
begin
  FAssigneds := [];
  FCellSmartTagXpgList.Clear;
  FR := '';
end;

{ TCT_CellSmartTagsXpgList }

function  TCT_CellSmartTagsXpgList.GetItems(Index: integer): TCT_CellSmartTags;
begin
  Result := TCT_CellSmartTags(inherited Items[Index]);
end;

function  TCT_CellSmartTagsXpgList.Add: TCT_CellSmartTags;
begin
  Result := TCT_CellSmartTags.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_CellSmartTagsXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_CellSmartTagsXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_Control }

function  TCT_Control.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FShapeId <> 0 then 
    Inc(AttrsAssigned);
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  if FName <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Control.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Control.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('shapeId',XmlIntToStr(FShapeId));
  AWriter.AddAttribute('r:id',FR_Id);
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
end;

procedure TCT_Control.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000002BE: FShapeId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000179: FR_Id := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Control.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
end;

destructor TCT_Control.Destroy;
begin
end;

procedure TCT_Control.Clear;
begin
  FAssigneds := [];
  FShapeId := 0;
  FR_Id := '';
  FName := '';
end;

{ TCT_ControlXpgList }

function  TCT_ControlXpgList.GetItems(Index: integer): TCT_Control;
begin
  Result := TCT_Control(inherited Items[Index]);
end;

function  TCT_ControlXpgList.Add: TCT_Control;
begin
  Result := TCT_Control.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ControlXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ControlXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_TablePart }

function  TCT_TablePart.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_TablePart.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TablePart.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_TablePart.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if NameOfQName(AAttributes[0]) = 'id' then
    FR_Id := AAttributes.Values[0]
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TablePart.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_TablePart.Destroy;
begin
end;

procedure TCT_TablePart.Clear;
begin
  FAssigneds := [];
  FR_Id := '';
end;

{ TCT_TablePartXpgList }

function  TCT_TablePartXpgList.GetItems(Index: integer): TCT_TablePart;
begin
  Result := TCT_TablePart(inherited Items[Index]);
end;

function  TCT_TablePartXpgList.Add: TCT_TablePart;
begin
  Result := TCT_TablePart.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TablePartXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TablePartXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_NumFmts }

function  TCT_NumFmts.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FNumFmtXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_NumFmts.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'numFmt' then 
    Result := FNumFmtXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_NumFmts.Write(AWriter: TXpgWriteXML);
begin
  FNumFmtXpgList.Write(AWriter,'numFmt');
end;

procedure TCT_NumFmts.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_NumFmts.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_NumFmts.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FNumFmtXpgList := TCT_NumFmtXpgList.Create(FOwner);
end;

destructor TCT_NumFmts.Destroy;
begin
  FNumFmtXpgList.Free;
end;

procedure TCT_NumFmts.Clear;
begin
  FAssigneds := [];
  FNumFmtXpgList.Clear;
  FCount := 0;
end;

{ TCT_Fonts }

function  TCT_Fonts.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFontXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Fonts.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'font' then 
    Result := FFontXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Fonts.Write(AWriter: TXpgWriteXML);
begin
  FFontXpgList.Write(AWriter,'font');
end;

procedure TCT_Fonts.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Fonts.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Fonts.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FFontXpgList := TCT_FontXpgList.Create(FOwner);
end;

destructor TCT_Fonts.Destroy;
begin
  FFontXpgList.Free;
end;

procedure TCT_Fonts.Clear;
begin
  FAssigneds := [];
  FFontXpgList.Clear;
  FCount := 0;
end;

{ TCT_Fills }

function  TCT_Fills.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFillXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Fills.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'fill' then 
    Result := FFillXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Fills.Write(AWriter: TXpgWriteXML);
begin
  FFillXpgList.Write(AWriter,'fill');
end;

procedure TCT_Fills.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Fills.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Fills.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FFillXpgList := TCT_FillXpgList.Create(FOwner);
end;

destructor TCT_Fills.Destroy;
begin
  FFillXpgList.Free;
end;

procedure TCT_Fills.Clear;
begin
  FAssigneds := [];
  FFillXpgList.Clear;
  FCount := 0;
end;

{ TCT_Borders }

function  TCT_Borders.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FBorderXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Borders.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'border' then 
    Result := FBorderXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Borders.Write(AWriter: TXpgWriteXML);
begin
  FBorderXpgList.Write(AWriter,'border');
end;

procedure TCT_Borders.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Borders.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Borders.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FBorderXpgList := TCT_BorderXpgList.Create(FOwner);
end;

destructor TCT_Borders.Destroy;
begin
  FBorderXpgList.Free;
end;

procedure TCT_Borders.Clear;
begin
  FAssigneds := [];
  FBorderXpgList.Clear;
  FCount := 0;
end;

{ TCT_CellStyleXfs }

function  TCT_CellStyleXfs.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FXfXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CellStyleXfs.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'xf' then 
    Result := FXfXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellStyleXfs.Write(AWriter: TXpgWriteXML);
begin
  FXfXpgList.Write(AWriter,'xf');
end;

procedure TCT_CellStyleXfs.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_CellStyleXfs.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CellStyleXfs.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FXfXpgList := TCT_XfXpgList.Create(FOwner);
end;

destructor TCT_CellStyleXfs.Destroy;
begin
  FXfXpgList.Free;
end;

procedure TCT_CellStyleXfs.Clear;
begin
  FAssigneds := [];
  FXfXpgList.Clear;
  FCount := 0;
end;

{ TCT_CellXfs }

function  TCT_CellXfs.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FXfXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CellXfs.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'xf' then 
    Result := FXfXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellXfs.Write(AWriter: TXpgWriteXML);
begin
  FXfXpgList.Write(AWriter,'xf');
end;

procedure TCT_CellXfs.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_CellXfs.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CellXfs.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FXfXpgList := TCT_XfXpgList.Create(FOwner);
end;

destructor TCT_CellXfs.Destroy;
begin
  FXfXpgList.Free;
end;

procedure TCT_CellXfs.Clear;
begin
  FAssigneds := [];
  FXfXpgList.Clear;
  FCount := 0;
end;

{ TCT_CellStyles }

function  TCT_CellStyles.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCellStyleXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_CellStyles.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cellStyle' then 
    Result := FCellStyleXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellStyles.Write(AWriter: TXpgWriteXML);
begin
  FCellStyleXpgList.Write(AWriter,'cellStyle');
end;

procedure TCT_CellStyles.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_CellStyles.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_CellStyles.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCellStyleXpgList := TCT_CellStyleXpgList.Create(FOwner);
end;

destructor TCT_CellStyles.Destroy;
begin
  FCellStyleXpgList.Free;
end;

procedure TCT_CellStyles.Clear;
begin
  FAssigneds := [];
  FCellStyleXpgList.Clear;
  FCount := 0;
end;

{ TCT_Dxfs }

function  TCT_Dxfs.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FDxfXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Dxfs.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'dxf' then 
    Result := FDxfXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Dxfs.Write(AWriter: TXpgWriteXML);
begin
  FDxfXpgList.Write(AWriter,'dxf');
end;

procedure TCT_Dxfs.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_Dxfs.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Dxfs.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FDxfXpgList := TCT_DxfXpgList.Create(FOwner);
end;

destructor TCT_Dxfs.Destroy;
begin
  FDxfXpgList.Free;
end;

procedure TCT_Dxfs.Clear;
begin
  FAssigneds := [];
  FDxfXpgList.Clear;
  FCount := 0;
end;

{ TCT_TableStyles }

function  TCT_TableStyles.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FDefaultTableStyle <> '' then 
    Inc(AttrsAssigned);
  if FDefaultPivotStyle <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FTableStyleXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TableStyles.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'tableStyle' then 
    Result := FTableStyleXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_TableStyles.Write(AWriter: TXpgWriteXML);
begin
  FTableStyleXpgList.Write(AWriter,'tableStyle');
end;

procedure TCT_TableStyles.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if FDefaultTableStyle <> '' then 
    AWriter.AddAttribute('defaultTableStyle',FDefaultTableStyle);
  if FDefaultPivotStyle <> '' then 
    AWriter.AddAttribute('defaultPivotStyle',FDefaultPivotStyle);
end;

procedure TCT_TableStyles.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000006DE: FDefaultTableStyle := AAttributes.Values[i];
      $00000708: FDefaultPivotStyle := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_TableStyles.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FTableStyleXpgList := TCT_TableStyleXpgList.Create(FOwner);
end;

destructor TCT_TableStyles.Destroy;
begin
  FTableStyleXpgList.Free;
end;

procedure TCT_TableStyles.Clear;
begin
  FAssigneds := [];
  FTableStyleXpgList.Clear;
  FCount := 0;
  FDefaultTableStyle := '';
  FDefaultPivotStyle := '';
end;

{ TCT_Colors }

function  TCT_Colors.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FIndexedColors.CheckAssigned);
  Inc(ElemsAssigned,FMruColors.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Colors.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000553: Result := FIndexedColors;
    $000003C6: Result := FMruColors;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Colors.Write(AWriter: TXpgWriteXML);
begin
  if FIndexedColors.Assigned then 
    if xaElements in FIndexedColors.FAssigneds then 
    begin
      AWriter.BeginTag('indexedColors');
      FIndexedColors.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('indexedColors');
  if FMruColors.Assigned then 
    if xaElements in FMruColors.FAssigneds then 
    begin
      AWriter.BeginTag('mruColors');
      FMruColors.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('mruColors');
end;

constructor TCT_Colors.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FIndexedColors := TCT_IndexedColors.Create(FOwner);
  FMruColors := TCT_MRUColors.Create(FOwner);
end;

destructor TCT_Colors.Destroy;
begin
  FIndexedColors.Free;
  FMruColors.Free;
end;

procedure TCT_Colors.Clear;
begin
  FAssigneds := [];
  FIndexedColors.Clear;
  FMruColors.Clear;
end;

{ TCT_FileVersion }

function  TCT_FileVersion.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAppName <> '' then 
    Inc(AttrsAssigned);
  if FLastEdited <> '' then 
    Inc(AttrsAssigned);
  if FLowestEdited <> '' then 
    Inc(AttrsAssigned);
  if FRupBuild <> '' then 
    Inc(AttrsAssigned);
  if FCodeName <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FileVersion.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FileVersion.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAppName <> '' then 
    AWriter.AddAttribute('appName',FAppName);
  if FLastEdited <> '' then 
    AWriter.AddAttribute('lastEdited',FLastEdited);
  if FLowestEdited <> '' then 
    AWriter.AddAttribute('lowestEdited',FLowestEdited);
  if FRupBuild <> '' then 
    AWriter.AddAttribute('rupBuild',FRupBuild);
  if FCodeName <> '' then 
    AWriter.AddAttribute('codeName',FCodeName);
end;

procedure TCT_FileVersion.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000002C2: FAppName := AAttributes.Values[i];
      $00000403: FLastEdited := AAttributes.Values[i];
      $000004ED: FLowestEdited := AAttributes.Values[i];
      $00000347: FRupBuild := AAttributes.Values[i];
      $0000031C: FCodeName := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_FileVersion.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
end;

destructor TCT_FileVersion.Destroy;
begin
end;

procedure TCT_FileVersion.Clear;
begin
  FAssigneds := [];
  FAppName := '';
  FLastEdited := '';
  FLowestEdited := '';
  FRupBuild := '';
  FCodeName := '';
end;

{ TCT_FileSharing }

function  TCT_FileSharing.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FReadOnlyRecommended <> False then 
    Inc(AttrsAssigned);
  if FUserName <> '' then 
    Inc(AttrsAssigned);
  if FReservationPassword <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FileSharing.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FileSharing.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FReadOnlyRecommended <> False then 
    AWriter.AddAttribute('readOnlyRecommended',XmlBoolToStr(FReadOnlyRecommended));
  if FUserName <> '' then 
    AWriter.AddAttribute('userName',FUserName);
  if FReservationPassword <> 0 then 
    AWriter.AddAttribute('reservationPassword',XmlIntToHexStr(FReservationPassword));
end;

procedure TCT_FileSharing.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000007A1: FReadOnlyRecommended := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000340: FUserName := AAttributes.Values[i];
      $00000805: FReservationPassword := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_FileSharing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FReadOnlyRecommended := False;
end;

destructor TCT_FileSharing.Destroy;
begin
end;

procedure TCT_FileSharing.Clear;
begin
  FAssigneds := [];
  FReadOnlyRecommended := False;
  FUserName := '';
  FReservationPassword := 0;
end;

{ TCT_WorkbookPr }

function  TCT_WorkbookPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDate1904 <> False then 
    Inc(AttrsAssigned);
  if FShowObjects <> stoAll then 
    Inc(AttrsAssigned);
  if FShowBorderUnselectedTables <> True then 
    Inc(AttrsAssigned);
  if FFilterPrivacy <> False then 
    Inc(AttrsAssigned);
  if FPromptedSolutions <> False then 
    Inc(AttrsAssigned);
  if FShowInkAnnotation <> True then 
    Inc(AttrsAssigned);
  if FBackupFile <> False then 
    Inc(AttrsAssigned);
  if FSaveExternalLinkValues <> True then 
    Inc(AttrsAssigned);
  if FUpdateLinks <> stulUserSet then 
    Inc(AttrsAssigned);
  if FCodeName <> '' then 
    Inc(AttrsAssigned);
  if FHidePivotFieldList <> False then 
    Inc(AttrsAssigned);
  if FShowPivotChartFilter <> False then 
    Inc(AttrsAssigned);
  if FAllowRefreshQuery <> False then 
    Inc(AttrsAssigned);
  if FPublishItems <> False then 
    Inc(AttrsAssigned);
  if FCheckCompatibility <> False then 
    Inc(AttrsAssigned);
  if FAutoCompressPictures <> True then 
    Inc(AttrsAssigned);
  if FRefreshAllConnections <> False then 
    Inc(AttrsAssigned);
  if FDefaultThemeVersion <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_WorkbookPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_WorkbookPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDate1904 <> False then 
    AWriter.AddAttribute('date1904',XmlBoolToStr(FDate1904));
  if FShowObjects <> stoAll then 
    AWriter.AddAttribute('showObjects',StrTST_Objects[Integer(FShowObjects)]);
  if FShowBorderUnselectedTables <> True then 
    AWriter.AddAttribute('showBorderUnselectedTables',XmlBoolToStr(FShowBorderUnselectedTables));
  if FFilterPrivacy <> False then 
    AWriter.AddAttribute('filterPrivacy',XmlBoolToStr(FFilterPrivacy));
  if FPromptedSolutions <> False then 
    AWriter.AddAttribute('promptedSolutions',XmlBoolToStr(FPromptedSolutions));
  if FShowInkAnnotation <> True then 
    AWriter.AddAttribute('showInkAnnotation',XmlBoolToStr(FShowInkAnnotation));
  if FBackupFile <> False then 
    AWriter.AddAttribute('backupFile',XmlBoolToStr(FBackupFile));
  if FSaveExternalLinkValues <> True then 
    AWriter.AddAttribute('saveExternalLinkValues',XmlBoolToStr(FSaveExternalLinkValues));
  if FUpdateLinks <> stulUserSet then 
    AWriter.AddAttribute('updateLinks',StrTST_UpdateLinks[Integer(FUpdateLinks)]);
  if FCodeName <> '' then 
    AWriter.AddAttribute('codeName',FCodeName);
  if FHidePivotFieldList <> False then 
    AWriter.AddAttribute('hidePivotFieldList',XmlBoolToStr(FHidePivotFieldList));
  if FShowPivotChartFilter <> False then 
    AWriter.AddAttribute('showPivotChartFilter',XmlBoolToStr(FShowPivotChartFilter));
  if FAllowRefreshQuery <> False then 
    AWriter.AddAttribute('allowRefreshQuery',XmlBoolToStr(FAllowRefreshQuery));
  if FPublishItems <> False then 
    AWriter.AddAttribute('publishItems',XmlBoolToStr(FPublishItems));
  if FCheckCompatibility <> False then 
    AWriter.AddAttribute('checkCompatibility',XmlBoolToStr(FCheckCompatibility));
  if FAutoCompressPictures <> True then 
    AWriter.AddAttribute('autoCompressPictures',XmlBoolToStr(FAutoCompressPictures));
  if FRefreshAllConnections <> False then 
    AWriter.AddAttribute('refreshAllConnections',XmlBoolToStr(FRefreshAllConnections));
  if FDefaultThemeVersion <> 0 then 
    AWriter.AddAttribute('defaultThemeVersion',XmlIntToStr(FDefaultThemeVersion));
end;

procedure TCT_WorkbookPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000026C: FDate1904 := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000048B: FShowObjects := TST_Objects(StrToEnum('sto' + AAttributes.Values[i]));
      $00000A86: FShowBorderUnselectedTables := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000564: FFilterPrivacy := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000073B: FPromptedSolutions := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006FE: FShowInkAnnotation := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003F6: FBackupFile := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000008F0: FSaveExternalLinkValues := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000484: FUpdateLinks := TST_UpdateLinks(StrToEnum('stul' + AAttributes.Values[i]));
      $0000031C: FCodeName := AAttributes.Values[i];
      $0000072C: FHidePivotFieldList := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000082B: FShowPivotChartFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000704: FAllowRefreshQuery := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F9: FPublishItems := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000758: FCheckCompatibility := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000854: FAutoCompressPictures := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000088B: FRefreshAllConnections := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000007BE: FDefaultThemeVersion := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_WorkbookPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 18;
  FDate1904 := False;
  FShowObjects := stoAll;
  FShowBorderUnselectedTables := True;
  FFilterPrivacy := False;
  FPromptedSolutions := False;
  FShowInkAnnotation := True;
  FBackupFile := False;
  FSaveExternalLinkValues := True;
  FUpdateLinks := stulUserSet;
  FHidePivotFieldList := False;
  FShowPivotChartFilter := False;
  FAllowRefreshQuery := False;
  FPublishItems := False;
  FCheckCompatibility := False;
  FAutoCompressPictures := True;
  FRefreshAllConnections := False;
end;

destructor TCT_WorkbookPr.Destroy;
begin
end;

procedure TCT_WorkbookPr.Clear;
begin
  FAssigneds := [];
  FDate1904 := False;
  FShowObjects := stoAll;
  FShowBorderUnselectedTables := True;
  FFilterPrivacy := False;
  FPromptedSolutions := False;
  FShowInkAnnotation := True;
  FBackupFile := False;
  FSaveExternalLinkValues := True;
  FUpdateLinks := stulUserSet;
  FCodeName := '';
  FHidePivotFieldList := False;
  FShowPivotChartFilter := False;
  FAllowRefreshQuery := False;
  FPublishItems := False;
  FCheckCompatibility := False;
  FAutoCompressPictures := True;
  FRefreshAllConnections := False;
  FDefaultThemeVersion := 0;
end;

{ TCT_WorkbookProtection }

function  TCT_WorkbookProtection.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FWorkbookPassword <> 0 then 
    Inc(AttrsAssigned);
  if FRevisionsPassword <> 0 then 
    Inc(AttrsAssigned);
  if FLockStructure <> False then 
    Inc(AttrsAssigned);
  if FLockWindows <> False then 
    Inc(AttrsAssigned);
  if FLockRevision <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_WorkbookProtection.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_WorkbookProtection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FWorkbookPassword <> 0 then 
    AWriter.AddAttribute('workbookPassword',XmlIntToHexStr(FWorkbookPassword));
  if FRevisionsPassword <> 0 then 
    AWriter.AddAttribute('revisionsPassword',XmlIntToHexStr(FRevisionsPassword));
  if FLockStructure <> False then 
    AWriter.AddAttribute('lockStructure',XmlBoolToStr(FLockStructure));
  if FLockWindows <> False then 
    AWriter.AddAttribute('lockWindows',XmlBoolToStr(FLockWindows));
  if FLockRevision <> False then 
    AWriter.AddAttribute('lockRevision',XmlBoolToStr(FLockRevision));
end;

procedure TCT_WorkbookProtection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000006C1: FWorkbookPassword := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000735: FRevisionsPassword := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $0000057A: FLockStructure := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000494: FLockWindows := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F8: FLockRevision := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_WorkbookProtection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
  FLockStructure := False;
  FLockWindows := False;
  FLockRevision := False;
end;

destructor TCT_WorkbookProtection.Destroy;
begin
end;

procedure TCT_WorkbookProtection.Clear;
begin
  FAssigneds := [];
  FWorkbookPassword := 0;
  FRevisionsPassword := 0;
  FLockStructure := False;
  FLockWindows := False;
  FLockRevision := False;
end;

{ TCT_BookViews }

function  TCT_BookViews.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FWorkbookViewXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_BookViews.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'workbookView' then 
    Result := FWorkbookViewXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_BookViews.Write(AWriter: TXpgWriteXML);
begin
  FWorkbookViewXpgList.Write(AWriter,'workbookView');
end;

constructor TCT_BookViews.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FWorkbookViewXpgList := TCT_BookViewXpgList.Create(FOwner);
end;

destructor TCT_BookViews.Destroy;
begin
  FWorkbookViewXpgList.Free;
end;

procedure TCT_BookViews.Clear;
begin
  FAssigneds := [];
  FWorkbookViewXpgList.Clear;
end;

{ TCT_Sheets }

function  TCT_Sheets.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Sheets.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'sheet' then 
    Result := FSheetXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Sheets.Write(AWriter: TXpgWriteXML);
begin
  FSheetXpgList.Write(AWriter,'sheet');
end;

constructor TCT_Sheets.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FSheetXpgList := TCT_SheetXpgList.Create(FOwner);
end;

destructor TCT_Sheets.Destroy;
begin
  FSheetXpgList.Free;
end;

procedure TCT_Sheets.Clear;
begin
  FAssigneds := [];
  FSheetXpgList.Clear;
end;

{ TCT_FunctionGroups }

function  TCT_FunctionGroups.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBuiltInGroupCount <> 16 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FFunctionGroupXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_FunctionGroups.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'functionGroup' then 
    Result := FFunctionGroupXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_FunctionGroups.Write(AWriter: TXpgWriteXML);
begin
  FFunctionGroupXpgList.Write(AWriter,'functionGroup');
end;

procedure TCT_FunctionGroups.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBuiltInGroupCount <> 16 then 
    AWriter.AddAttribute('builtInGroupCount',XmlIntToStr(FBuiltInGroupCount));
end;

procedure TCT_FunctionGroups.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'builtInGroupCount' then 
    FBuiltInGroupCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_FunctionGroups.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FFunctionGroupXpgList := TCT_FunctionGroupXpgList.Create(FOwner);
  FBuiltInGroupCount := 16;
end;

destructor TCT_FunctionGroups.Destroy;
begin
  FFunctionGroupXpgList.Free;
end;

procedure TCT_FunctionGroups.Clear;
begin
  FAssigneds := [];
  FFunctionGroupXpgList.Clear;
  FBuiltInGroupCount := 16;
end;

{ TCT_ExternalReferences }

function  TCT_ExternalReferences.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FExternalReferenceXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExternalReferences.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'externalReference' then 
    Result := FExternalReferenceXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalReferences.Write(AWriter: TXpgWriteXML);
begin
  FExternalReferenceXpgList.Write(AWriter,'externalReference');
end;

constructor TCT_ExternalReferences.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FExternalReferenceXpgList := TCT_ExternalReferenceXpgList.Create(FOwner);
end;

destructor TCT_ExternalReferences.Destroy;
begin
  FExternalReferenceXpgList.Free;
end;

procedure TCT_ExternalReferences.Clear;
begin
  FAssigneds := [];
  FExternalReferenceXpgList.Clear;
end;

{ TCT_DefinedNames }

function  TCT_DefinedNames.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
//  ElemsAssigned := 0;
  FAssigneds := [];
  ElemsAssigned := 1; // TODO Event
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DefinedNames.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'definedName' then 
  begin
    Result := FDefinedName;
    TCT_DefinedName(Result).Clear;
    if AReader.HasText then 
      TCT_DefinedName(Result).Content := AReader.Text;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DefinedNames.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteDefinedName) then 
    while True do 
    begin
      WriteElement := False;
      FOnWriteDefinedName(FDefinedName,WriteElement);
      if WriteElement = False then 
        Break;
      FDefinedName.CheckAssigned;  // TODO
      if FDefinedName.Assigned then
      begin
        AWriter.Text := FDefinedName.Content;
        FDefinedName.WriteAttributes(AWriter);
        AWriter.SimpleTag('definedName');
      end;
    end;
end;

procedure TCT_DefinedNames.AfterTag;
begin
  if (FDefinedName.FAssigneds <> []) and System.Assigned(FOnReadDefinedName) then 
    FOnReadDefinedName(FDefinedName);
  FDefinedName.FAssigneds := [];
end;

constructor TCT_DefinedNames.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FDefinedName := TCT_DefinedName.Create(FOwner);
end;

destructor TCT_DefinedNames.Destroy;
begin
  FDefinedName.Free;
end;

procedure TCT_DefinedNames.Clear;
begin
  FAssigneds := [];
  FDefinedName.Clear;
end;

{ TCT_CalcPr }

function  TCT_CalcPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCalcId <> 0 then 
    Inc(AttrsAssigned);
  if FCalcMode <> stcmAuto then 
    Inc(AttrsAssigned);
  if FFullCalcOnLoad <> False then 
    Inc(AttrsAssigned);
  if FRefMode <> strmA1 then 
    Inc(AttrsAssigned);
  if FIterate <> False then 
    Inc(AttrsAssigned);
  if FIterateCount <> 100 then 
    Inc(AttrsAssigned);
  if FIterateDelta <> 0.000999999999748979 then 
    Inc(AttrsAssigned);
  if FFullPrecision <> True then 
    Inc(AttrsAssigned);
  if FCalcCompleted <> True then 
    Inc(AttrsAssigned);
  if FCalcOnSave <> True then 
    Inc(AttrsAssigned);
  if FConcurrentCalc <> True then 
    Inc(AttrsAssigned);
  if FConcurrentManualCount <> 0 then 
    Inc(AttrsAssigned);
  if FForceFullCalc <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_CalcPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_CalcPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
//  if FCalcId <> 0 then
    AWriter.AddAttribute('calcId',XmlIntToStr(FCalcId));
  if FCalcMode <> stcmAuto then 
    AWriter.AddAttribute('calcMode',StrTST_CalcMode[Integer(FCalcMode)]);
  if FFullCalcOnLoad <> False then 
    AWriter.AddAttribute('fullCalcOnLoad',XmlBoolToStr(FFullCalcOnLoad));
  if FRefMode <> strmA1 then 
    AWriter.AddAttribute('refMode',StrTST_RefMode[Integer(FRefMode)]);
  if FIterate <> False then 
    AWriter.AddAttribute('iterate',XmlBoolToStr(FIterate));
  if FIterateCount <> 100 then 
    AWriter.AddAttribute('iterateCount',XmlIntToStr(FIterateCount));
  if FIterateDelta <> 0.000999999999748979 then 
    AWriter.AddAttribute('iterateDelta',XmlFloatToStr(FIterateDelta));
  if FFullPrecision <> True then 
    AWriter.AddAttribute('fullPrecision',XmlBoolToStr(FFullPrecision));
  if FCalcCompleted <> True then 
    AWriter.AddAttribute('calcCompleted',XmlBoolToStr(FCalcCompleted));
  if FCalcOnSave <> True then 
    AWriter.AddAttribute('calcOnSave',XmlBoolToStr(FCalcOnSave));
  if FConcurrentCalc <> True then 
    AWriter.AddAttribute('concurrentCalc',XmlBoolToStr(FConcurrentCalc));
  if FConcurrentManualCount <> 0 then 
    AWriter.AddAttribute('concurrentManualCount',XmlIntToStr(FConcurrentManualCount));
  if FForceFullCalc <> False then 
    AWriter.AddAttribute('forceFullCalc',XmlBoolToStr(FForceFullCalc));
end;

procedure TCT_CalcPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000240: FCalcId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000318: FCalcMode := TST_CalcMode(StrToEnum('stcm' + AAttributes.Values[i]));
      $00000563: FFullCalcOnLoad := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002C2: FRefMode := TST_RefMode(StrToEnum('strm' + AAttributes.Values[i]));
      $000002EE: FIterate := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F7: FIterateCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004D8: FIterateDelta := XmlStrToFloatDef(AAttributes.Values[i],0);
      $0000055F: FFullPrecision := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000530: FCalcCompleted := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003DF: FCalcOnSave := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005B6: FConcurrentCalc := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000008AA: FConcurrentManualCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000515: FForceFullCalc := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_CalcPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 13;
  FCalcMode := stcmAuto;
  FCalcId := 125725;
  FFullCalcOnLoad := False;
  FRefMode := strmA1;
  FIterate := False;
  FIterateCount := 100;
  FIterateDelta := 0.000999999999748979;
  FFullPrecision := True;
  FCalcCompleted := True;
  FCalcOnSave := True;
  FConcurrentCalc := True;
end;

destructor TCT_CalcPr.Destroy;
begin
end;

procedure TCT_CalcPr.Clear;
begin
  FAssigneds := [];
  FCalcId := 125725;
  FCalcMode := stcmAuto;
  FFullCalcOnLoad := False;
  FRefMode := strmA1;
  FIterate := False;
  FIterateCount := 100;
  FIterateDelta := 0.000999999999748979;
  FFullPrecision := True;
  FCalcCompleted := True;
  FCalcOnSave := True;
  FConcurrentCalc := True;
  FConcurrentManualCount := 0;
  FForceFullCalc := False;
end;

{ TCT_OleSize }

function  TCT_OleSize.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_OleSize.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_OleSize.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ref',FRef);
end;

procedure TCT_OleSize.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'ref' then 
    FRef := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_OleSize.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_OleSize.Destroy;
begin
end;

procedure TCT_OleSize.Clear;
begin
  FAssigneds := [];
  FRef := '';
end;

{ TCT_CustomWorkbookViews }

function  TCT_CustomWorkbookViews.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCustomWorkbookViewXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustomWorkbookViews.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'customWorkbookView' then 
    Result := FCustomWorkbookViewXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomWorkbookViews.Write(AWriter: TXpgWriteXML);
begin
  FCustomWorkbookViewXpgList.Write(AWriter,'customWorkbookView');
end;

constructor TCT_CustomWorkbookViews.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCustomWorkbookViewXpgList := TCT_CustomWorkbookViewXpgList.Create(FOwner);
end;

destructor TCT_CustomWorkbookViews.Destroy;
begin
  FCustomWorkbookViewXpgList.Free;
end;

procedure TCT_CustomWorkbookViews.Clear;
begin
  FAssigneds := [];
  FCustomWorkbookViewXpgList.Clear;
end;

{ TCT_PivotCaches }

function  TCT_PivotCaches.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FPivotCacheXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_PivotCaches.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'pivotCache' then 
    Result := FPivotCacheXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_PivotCaches.Write(AWriter: TXpgWriteXML);
begin
  FPivotCacheXpgList.Write(AWriter,'pivotCache');
end;

constructor TCT_PivotCaches.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FPivotCacheXpgList := TCT_PivotCacheXpgList.Create(FOwner);
end;

destructor TCT_PivotCaches.Destroy;
begin
  FPivotCacheXpgList.Free;
end;

procedure TCT_PivotCaches.Clear;
begin
  FAssigneds := [];
  FPivotCacheXpgList.Clear;
end;

{ TCT_SmartTagPr }

function  TCT_SmartTagPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FEmbed <> False then 
    Inc(AttrsAssigned);
  if FShow <> ststsAll then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SmartTagPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SmartTagPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FEmbed <> False then 
    AWriter.AddAttribute('embed',XmlBoolToStr(FEmbed));
  if FShow <> ststsAll then 
    AWriter.AddAttribute('show',StrTST_SmartTagShow[Integer(FShow)]);
end;

procedure TCT_SmartTagPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001FD: FEmbed := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001C1: FShow := TST_SmartTagShow(StrToEnum('ststs' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SmartTagPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 2;
  FEmbed := False;
  FShow := ststsAll;
end;

destructor TCT_SmartTagPr.Destroy;
begin
end;

procedure TCT_SmartTagPr.Clear;
begin
  FAssigneds := [];
  FEmbed := False;
  FShow := ststsAll;
end;

{ TCT_SmartTagTypes }

function  TCT_SmartTagTypes.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSmartTagTypeXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SmartTagTypes.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'smartTagType' then 
    Result := FSmartTagTypeXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SmartTagTypes.Write(AWriter: TXpgWriteXML);
begin
  FSmartTagTypeXpgList.Write(AWriter,'smartTagType');
end;

constructor TCT_SmartTagTypes.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FSmartTagTypeXpgList := TCT_SmartTagTypeXpgList.Create(FOwner);
end;

destructor TCT_SmartTagTypes.Destroy;
begin
  FSmartTagTypeXpgList.Free;
end;

procedure TCT_SmartTagTypes.Clear;
begin
  FAssigneds := [];
  FSmartTagTypeXpgList.Clear;
end;

{ TCT_WebPublishing }

function  TCT_WebPublishing.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCss <> True then 
    Inc(AttrsAssigned);
  if FThicket <> True then 
    Inc(AttrsAssigned);
  if FLongFileNames <> True then 
    Inc(AttrsAssigned);
  if FVml <> False then 
    Inc(AttrsAssigned);
  if FAllowPng <> False then 
    Inc(AttrsAssigned);
  if FTargetScreenSize <> sttss800x600 then 
    Inc(AttrsAssigned);
  if FDpi <> 96 then 
    Inc(AttrsAssigned);
  if FCodePage <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_WebPublishing.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_WebPublishing.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCss <> True then 
    AWriter.AddAttribute('css',XmlBoolToStr(FCss));
  if FThicket <> True then 
    AWriter.AddAttribute('thicket',XmlBoolToStr(FThicket));
  if FLongFileNames <> True then 
    AWriter.AddAttribute('longFileNames',XmlBoolToStr(FLongFileNames));
  if FVml <> False then 
    AWriter.AddAttribute('vml',XmlBoolToStr(FVml));
  if FAllowPng <> False then 
    AWriter.AddAttribute('allowPng',XmlBoolToStr(FAllowPng));
  if FTargetScreenSize <> sttss800x600 then 
    AWriter.AddAttribute('targetScreenSize',StrTST_TargetScreenSize[Integer(FTargetScreenSize)]);
  if FDpi <> 96 then 
    AWriter.AddAttribute('dpi',XmlIntToStr(FDpi));
  if FCodePage <> 0 then 
    AWriter.AddAttribute('codePage',XmlIntToStr(FCodePage));
end;

procedure TCT_WebPublishing.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000149: FCss := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002EC: FThicket := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000524: FLongFileNames := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000014F: FVml := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000344: FAllowPng := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000682: FTargetScreenSize := TST_TargetScreenSize(StrToEnum('sttss' + AAttributes.Values[i]));
      $0000013D: FDpi := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000318: FCodePage := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_WebPublishing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 8;
  FCss := True;
  FThicket := True;
  FLongFileNames := True;
  FVml := False;
  FAllowPng := False;
  FTargetScreenSize := sttss800x600;
  FDpi := 96;
end;

destructor TCT_WebPublishing.Destroy;
begin
end;

procedure TCT_WebPublishing.Clear;
begin
  FAssigneds := [];
  FCss := True;
  FThicket := True;
  FLongFileNames := True;
  FVml := False;
  FAllowPng := False;
  FTargetScreenSize := sttss800x600;
  FDpi := 96;
  FCodePage := 0;
end;

{ TCT_FileRecoveryPr }

function  TCT_FileRecoveryPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FAutoRecover <> True then 
    Inc(AttrsAssigned);
  if FCrashSave <> False then 
    Inc(AttrsAssigned);
  if FDataExtractLoad <> False then 
    Inc(AttrsAssigned);
  if FRepairLoad <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_FileRecoveryPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_FileRecoveryPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FAutoRecover <> True then 
    AWriter.AddAttribute('autoRecover',XmlBoolToStr(FAutoRecover));
  if FCrashSave <> False then 
    AWriter.AddAttribute('crashSave',XmlBoolToStr(FCrashSave));
  if FDataExtractLoad <> False then 
    AWriter.AddAttribute('dataExtractLoad',XmlBoolToStr(FDataExtractLoad));
  if FRepairLoad <> False then 
    AWriter.AddAttribute('repairLoad',XmlBoolToStr(FRepairLoad));
end;

procedure TCT_FileRecoveryPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $0000048F: FAutoRecover := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003A0: FCrashSave := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005F5: FDataExtractLoad := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000403: FRepairLoad := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_FileRecoveryPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FAutoRecover := True;
  FCrashSave := False;
  FDataExtractLoad := False;
  FRepairLoad := False;
end;

destructor TCT_FileRecoveryPr.Destroy;
begin
end;

procedure TCT_FileRecoveryPr.Clear;
begin
  FAssigneds := [];
  FAutoRecover := True;
  FCrashSave := False;
  FDataExtractLoad := False;
  FRepairLoad := False;
end;

{ TCT_FileRecoveryPrXpgList }

function  TCT_FileRecoveryPrXpgList.GetItems(Index: integer): TCT_FileRecoveryPr;
begin
  Result := TCT_FileRecoveryPr(inherited Items[Index]);
end;

function  TCT_FileRecoveryPrXpgList.Add: TCT_FileRecoveryPr;
begin
  Result := TCT_FileRecoveryPr.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_FileRecoveryPrXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_FileRecoveryPrXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_WebPublishObjects }

function  TCT_WebPublishObjects.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FWebPublishObjectXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_WebPublishObjects.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'webPublishObject' then 
    Result := FWebPublishObjectXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_WebPublishObjects.Write(AWriter: TXpgWriteXML);
begin
  FWebPublishObjectXpgList.Write(AWriter,'webPublishObject');
end;

procedure TCT_WebPublishObjects.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_WebPublishObjects.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_WebPublishObjects.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FWebPublishObjectXpgList := TCT_WebPublishObjectXpgList.Create(FOwner);
end;

destructor TCT_WebPublishObjects.Destroy;
begin
  FWebPublishObjectXpgList.Free;
end;

procedure TCT_WebPublishObjects.Clear;
begin
  FAssigneds := [];
  FWebPublishObjectXpgList.Clear;
  FCount := 0;
end;

{ TCT_Authors }

function  TCT_Authors.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FAuthorXpgList.Count);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Authors.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'author' then 
    FAuthorXpgList.Add(AReader.Text)
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Authors.Write(AWriter: TXpgWriteXML);
begin
  FAuthorXpgList.WriteElements(AWriter,'author');
end;

constructor TCT_Authors.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FAuthorXpgList := TStringXpgList.Create;
end;

destructor TCT_Authors.Destroy;
begin
  FAuthorXpgList.Free;
end;

procedure TCT_Authors.Clear;
begin
  FAssigneds := [];
  FAuthorXpgList.DelimitedText := '';
end;

{ TCT_CommentList }

function  TCT_CommentList.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FComment.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CommentList.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'comment' then 
  begin
    Result := FComment;
    TCT_Comment(Result).Clear;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CommentList.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteComment) then 
    while True do 
    begin
      WriteElement := False;
      FComment.Clear; // TODO
      FOnWriteComment(FComment,WriteElement);
      if WriteElement = False then
        Break;
      FComment.CheckAssigned;  // TODO
      if FComment.Assigned then
      begin
        FComment.WriteAttributes(AWriter);
        if xaElements in FComment.FAssigneds then
        begin
          AWriter.BeginTag('comment');
          FComment.Write(AWriter);
          AWriter.EndTag;
        end
        else
          AWriter.SimpleTag('comment');
      end;
    end;
end;

procedure TCT_CommentList.AfterTag;
begin
  if (FComment.FAssigneds <> []) and System.Assigned(FOnReadComment) then begin
    FOnReadComment(FComment);
    FComment.Clear; // TODO
  end;
  FComment.FAssigneds := [];
end;

constructor TCT_CommentList.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FComment := TCT_Comment.Create(FOwner);
end;

destructor TCT_CommentList.Destroy;
begin
  FComment.Free;
end;

procedure TCT_CommentList.Clear;
begin
  FAssigneds := [];
  FComment.Clear;
end;

{ TCT_ChartsheetPr }

function  TCT_ChartsheetPr.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPublished <> True then 
    Inc(AttrsAssigned);
  if FCodeName <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FTabColor.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ChartsheetPr.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'tabColor' then 
    Result := FTabColor
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ChartsheetPr.Write(AWriter: TXpgWriteXML);
begin
  if FTabColor.Assigned then 
  begin
    FTabColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('tabColor');
  end;
end;

procedure TCT_ChartsheetPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPublished <> True then 
    AWriter.AddAttribute('published',XmlBoolToStr(FPublished));
  if FCodeName <> '' then 
    AWriter.AddAttribute('codeName',FCodeName);
end;

procedure TCT_ChartsheetPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000003C0: FPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000031C: FCodeName := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_ChartsheetPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FTabColor := TCT_Color.Create(FOwner);
  FPublished := True;
end;

destructor TCT_ChartsheetPr.Destroy;
begin
  FTabColor.Free;
end;

procedure TCT_ChartsheetPr.Clear;
begin
  FAssigneds := [];
  FTabColor.Clear;
  FPublished := True;
  FCodeName := '';
end;

{ TCT_ChartsheetViews }

function  TCT_ChartsheetViews.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetViewXpgList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ChartsheetViews.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003B4: Result := FSheetViewXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ChartsheetViews.Write(AWriter: TXpgWriteXML);
begin
  FSheetViewXpgList.Write(AWriter,'sheetView');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_ChartsheetViews.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FSheetViewXpgList := TCT_ChartsheetViewXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_ChartsheetViews.Destroy;
begin
  FSheetViewXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_ChartsheetViews.Clear;
begin
  FAssigneds := [];
  FSheetViewXpgList.Clear;
  FExtLst.Clear;
end;

{ TCT_ChartsheetProtection }

function  TCT_ChartsheetProtection.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPassword <> 0 then 
    Inc(AttrsAssigned);
  if FContent <> False then 
    Inc(AttrsAssigned);
  if FObjects <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ChartsheetProtection.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ChartsheetProtection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPassword <> 0 then 
    AWriter.AddAttribute('password',XmlIntToHexStr(FPassword));
  if FContent <> False then 
    AWriter.AddAttribute('content',XmlBoolToStr(FContent));
  if FObjects <> False then 
    AWriter.AddAttribute('objects',XmlBoolToStr(FObjects));
end;

procedure TCT_ChartsheetProtection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000373: FPassword := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $000002FB: FContent := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002EA: FObjects := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_ChartsheetProtection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
  FContent := False;
  FObjects := False;
end;

destructor TCT_ChartsheetProtection.Destroy;
begin
end;

procedure TCT_ChartsheetProtection.Clear;
begin
  FAssigneds := [];
  FPassword := 0;
  FContent := False;
  FObjects := False;
end;

{ TCT_CustomChartsheetViews }

function  TCT_CustomChartsheetViews.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCustomSheetViewXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustomChartsheetViews.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'customSheetView' then 
    Result := FCustomSheetViewXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomChartsheetViews.Write(AWriter: TXpgWriteXML);
begin
  FCustomSheetViewXpgList.Write(AWriter,'customSheetView');
end;

constructor TCT_CustomChartsheetViews.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCustomSheetViewXpgList := TCT_CustomChartsheetViewXpgList.Create(FOwner);
end;

destructor TCT_CustomChartsheetViews.Destroy;
begin
  FCustomSheetViewXpgList.Free;
end;

procedure TCT_CustomChartsheetViews.Clear;
begin
  FAssigneds := [];
  FCustomSheetViewXpgList.Clear;
end;

{ TCT_Drawing }

function  TCT_Drawing.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_Drawing.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_Drawing.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_Drawing.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if NameOfQName(AAttributes[0]) = 'id' then
    FR_Id := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_Drawing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_Drawing.Destroy;
begin
end;

procedure TCT_Drawing.Clear;
begin
  FAssigneds := [];
  FR_Id := '';
end;

{ TCT_LegacyDrawing }

function  TCT_LegacyDrawing.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_LegacyDrawing.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_LegacyDrawing.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_LegacyDrawing.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if NameOfQName(AAttributes[0]) = 'id' then
    FR_Id := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_LegacyDrawing.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_LegacyDrawing.Destroy;
begin
end;

procedure TCT_LegacyDrawing.Clear;
begin
  FAssigneds := [];
  FR_Id := '';
end;

{ TCT_SheetBackgroundPicture }

function  TCT_SheetBackgroundPicture.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SheetBackgroundPicture.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SheetBackgroundPicture.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_SheetBackgroundPicture.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if NameOfQName(AAttributes[0]) = 'id' then
    FR_Id := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SheetBackgroundPicture.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_SheetBackgroundPicture.Destroy;
begin
end;

procedure TCT_SheetBackgroundPicture.Clear;
begin
  FAssigneds := [];
  FR_Id := '';
end;

{ TCT_WebPublishItems }

function  TCT_WebPublishItems.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FWebPublishItemXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_WebPublishItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'webPublishItem' then 
    Result := FWebPublishItemXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_WebPublishItems.Write(AWriter: TXpgWriteXML);
begin
  FWebPublishItemXpgList.Write(AWriter,'webPublishItem');
end;

procedure TCT_WebPublishItems.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_WebPublishItems.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_WebPublishItems.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FWebPublishItemXpgList := TCT_WebPublishItemXpgList.Create(FOwner);
end;

destructor TCT_WebPublishItems.Destroy;
begin
  FWebPublishItemXpgList.Free;
end;

procedure TCT_WebPublishItems.Clear;
begin
  FAssigneds := [];
  FWebPublishItemXpgList.Clear;
  FCount := 0;
end;

{ TCT_SheetPr }

function  TCT_SheetPr.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FSyncHorizontal <> False then 
    Inc(AttrsAssigned);
  if FSyncVertical <> False then 
    Inc(AttrsAssigned);
  if FSyncRef <> '' then 
    Inc(AttrsAssigned);
  if FTransitionEvaluation <> False then 
    Inc(AttrsAssigned);
  if FTransitionEntry <> False then 
    Inc(AttrsAssigned);
  if FPublished <> True then 
    Inc(AttrsAssigned);
  if FCodeName <> '' then 
    Inc(AttrsAssigned);
  if FFilterMode <> False then 
    Inc(AttrsAssigned);
  if FEnableFormatConditionsCalculation <> True then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FTabColor.CheckAssigned);
  Inc(ElemsAssigned,FOutlinePr.CheckAssigned);
  Inc(ElemsAssigned,FPageSetUpPr.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_SheetPr.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000336: Result := FTabColor;
    $000003C2: Result := FOutlinePr;
    $00000450: Result := FPageSetUpPr;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SheetPr.Write(AWriter: TXpgWriteXML);
begin
  if FTabColor.Assigned then 
  begin
    FTabColor.WriteAttributes(AWriter);
    AWriter.SimpleTag('tabColor');
  end;
  if FOutlinePr.Assigned then 
  begin
    FOutlinePr.WriteAttributes(AWriter);
    AWriter.SimpleTag('outlinePr');
  end;
  if FPageSetUpPr.Assigned then 
  begin
    FPageSetUpPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetUpPr');
  end;
end;

procedure TCT_SheetPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FSyncHorizontal <> False then 
    AWriter.AddAttribute('syncHorizontal',XmlBoolToStr(FSyncHorizontal));
  if FSyncVertical <> False then 
    AWriter.AddAttribute('syncVertical',XmlBoolToStr(FSyncVertical));
  if FSyncRef <> '' then 
    AWriter.AddAttribute('syncRef',FSyncRef);
  if FTransitionEvaluation <> False then 
    AWriter.AddAttribute('transitionEvaluation',XmlBoolToStr(FTransitionEvaluation));
  if FTransitionEntry <> False then 
    AWriter.AddAttribute('transitionEntry',XmlBoolToStr(FTransitionEntry));
  if FPublished <> True then 
    AWriter.AddAttribute('published',XmlBoolToStr(FPublished));
  if FCodeName <> '' then 
    AWriter.AddAttribute('codeName',FCodeName);
  if FFilterMode <> False then 
    AWriter.AddAttribute('filterMode',XmlBoolToStr(FFilterMode));
  if FEnableFormatConditionsCalculation <> True then 
    AWriter.AddAttribute('enableFormatConditionsCalculation',XmlBoolToStr(FEnableFormatConditionsCalculation));
end;

procedure TCT_SheetPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000005E7: FSyncHorizontal := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000004F7: FSyncVertical := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002DA: FSyncRef := AAttributes.Values[i];
      $00000863: FTransitionEvaluation := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000065D: FTransitionEntry := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003C0: FPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000031C: FCodeName := AAttributes.Values[i];
      $0000040B: FFilterMode := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000D59: FEnableFormatConditionsCalculation := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SheetPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 9;
  FTabColor := TCT_Color.Create(FOwner);
  FOutlinePr := TCT_OutlinePr.Create(FOwner);
  FPageSetUpPr := TCT_PageSetUpPr.Create(FOwner);
  FSyncHorizontal := False;
  FSyncVertical := False;
  FTransitionEvaluation := False;
  FTransitionEntry := False;
  FPublished := True;
  FFilterMode := False;
  FEnableFormatConditionsCalculation := True;
end;

destructor TCT_SheetPr.Destroy;
begin
  FTabColor.Free;
  FOutlinePr.Free;
  FPageSetUpPr.Free;
end;

procedure TCT_SheetPr.Clear;
begin
  FAssigneds := [];
  FTabColor.Clear;
  FOutlinePr.Clear;
  FPageSetUpPr.Clear;
  FSyncHorizontal := False;
  FSyncVertical := False;
  FSyncRef := '';
  FTransitionEvaluation := False;
  FTransitionEntry := False;
  FPublished := True;
  FCodeName := '';
  FFilterMode := False;
  FEnableFormatConditionsCalculation := True;
end;

{ TCT_SheetViews }

function  TCT_SheetViews.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetViewXpgList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SheetViews.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000003B4: Result := FSheetViewXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SheetViews.Write(AWriter: TXpgWriteXML);
begin
  FSheetViewXpgList.Write(AWriter,'sheetView');
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_SheetViews.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FSheetViewXpgList := TCT_SheetViewXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_SheetViews.Destroy;
begin
  FSheetViewXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_SheetViews.Clear;
begin
  FAssigneds := [];
  FSheetViewXpgList.Clear;
  FExtLst.Clear;
end;

{ TCT_SheetFormatPr }

function  TCT_SheetFormatPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FBaseColWidth <> 8 then 
    Inc(AttrsAssigned);
  if FDefaultColWidth <> 0 then 
    Inc(AttrsAssigned);
  if FDefaultRowHeight <> 0 then 
    Inc(AttrsAssigned);
  if FCustomHeight <> False then 
    Inc(AttrsAssigned);
  if FZeroHeight <> False then 
    Inc(AttrsAssigned);
  if FThickTop <> False then 
    Inc(AttrsAssigned);
  if FThickBottom <> False then 
    Inc(AttrsAssigned);
  if FOutlineLevelRow <> 0 then 
    Inc(AttrsAssigned);
  if FOutlineLevelCol <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SheetFormatPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SheetFormatPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FBaseColWidth <> 8 then 
    AWriter.AddAttribute('baseColWidth',XmlIntToStr(FBaseColWidth));
  if FDefaultColWidth <> 0 then 
    AWriter.AddAttribute('defaultColWidth',XmlFloatToStr(FDefaultColWidth));
  AWriter.AddAttribute('defaultRowHeight',XmlFloatToStr(FDefaultRowHeight));
  if FCustomHeight <> False then 
    AWriter.AddAttribute('customHeight',XmlBoolToStr(FCustomHeight));
  if FZeroHeight <> False then 
    AWriter.AddAttribute('zeroHeight',XmlBoolToStr(FZeroHeight));
  if FThickTop <> False then 
    AWriter.AddAttribute('thickTop',XmlBoolToStr(FThickTop));
  if FThickBottom <> False then 
    AWriter.AddAttribute('thickBottom',XmlBoolToStr(FThickBottom));
  if FOutlineLevelRow <> 0 then 
    AWriter.AddAttribute('outlineLevelRow',XmlIntToStr(FOutlineLevelRow));
  if FOutlineLevelCol <> 0 then 
    AWriter.AddAttribute('outlineLevelCol',XmlIntToStr(FOutlineLevelCol));
end;

procedure TCT_SheetFormatPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000004B9: FBaseColWidth := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000603: FDefaultColWidth := XmlStrToFloatDef(AAttributes.Values[i],0);
      $00000676: FDefaultRowHeight := XmlStrToFloatDef(AAttributes.Values[i],0);
      $000004F4: FCustomHeight := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000419: FZeroHeight := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000346: FThickTop := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000488: FThickBottom := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000630: FOutlineLevelRow := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000616: FOutlineLevelCol := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SheetFormatPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 9;
  FBaseColWidth := 8;
  FCustomHeight := False;
  FZeroHeight := False;
  FThickTop := False;
  FThickBottom := False;
  FOutlineLevelRow := 0;
  FOutlineLevelCol := 0;
end;

destructor TCT_SheetFormatPr.Destroy;
begin
end;

procedure TCT_SheetFormatPr.Clear;
begin
  FAssigneds := [];
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

{ TCT_SheetProtection }

function  TCT_SheetProtection.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPassword <> 0 then 
    Inc(AttrsAssigned);
  if FSheet <> False then 
    Inc(AttrsAssigned);
  if FObjects <> False then 
    Inc(AttrsAssigned);
  if FScenarios <> False then 
    Inc(AttrsAssigned);
  if FFormatCells <> True then 
    Inc(AttrsAssigned);
  if FFormatColumns <> True then 
    Inc(AttrsAssigned);
  if FFormatRows <> True then 
    Inc(AttrsAssigned);
  if FInsertColumns <> True then 
    Inc(AttrsAssigned);
  if FInsertRows <> True then 
    Inc(AttrsAssigned);
  if FInsertHyperlinks <> True then 
    Inc(AttrsAssigned);
  if FDeleteColumns <> True then 
    Inc(AttrsAssigned);
  if FDeleteRows <> True then 
    Inc(AttrsAssigned);
  if FSelectLockedCells <> False then 
    Inc(AttrsAssigned);
  if FSort <> True then 
    Inc(AttrsAssigned);
  if FAutoFilter <> True then 
    Inc(AttrsAssigned);
  if FPivotTables <> True then 
    Inc(AttrsAssigned);
  if FSelectUnlockedCells <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SheetProtection.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SheetProtection.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPassword <> 0 then
    AWriter.AddAttribute('password',XmlIntToHexStr(FPassword,4));
  if FSheet <> False then 
    AWriter.AddAttribute('sheet',XmlBoolToStr(FSheet));
  if FObjects <> False then 
    AWriter.AddAttribute('objects',XmlBoolToStr(FObjects));
  if FScenarios <> False then 
    AWriter.AddAttribute('scenarios',XmlBoolToStr(FScenarios));
  if FFormatCells <> True then 
    AWriter.AddAttribute('formatCells',XmlBoolToStr(FFormatCells));
  if FFormatColumns <> True then 
    AWriter.AddAttribute('formatColumns',XmlBoolToStr(FFormatColumns));
  if FFormatRows <> True then 
    AWriter.AddAttribute('formatRows',XmlBoolToStr(FFormatRows));
  if FInsertColumns <> True then 
    AWriter.AddAttribute('insertColumns',XmlBoolToStr(FInsertColumns));
  if FInsertRows <> True then 
    AWriter.AddAttribute('insertRows',XmlBoolToStr(FInsertRows));
  if FInsertHyperlinks <> True then 
    AWriter.AddAttribute('insertHyperlinks',XmlBoolToStr(FInsertHyperlinks));
  if FDeleteColumns <> True then 
    AWriter.AddAttribute('deleteColumns',XmlBoolToStr(FDeleteColumns));
  if FDeleteRows <> True then 
    AWriter.AddAttribute('deleteRows',XmlBoolToStr(FDeleteRows));
  if FSelectLockedCells <> False then 
    AWriter.AddAttribute('selectLockedCells',XmlBoolToStr(FSelectLockedCells));
  if FSort <> True then 
    AWriter.AddAttribute('sort',XmlBoolToStr(FSort));
  if FAutoFilter <> True then 
    AWriter.AddAttribute('autoFilter',XmlBoolToStr(FAutoFilter));
  if FPivotTables <> True then 
    AWriter.AddAttribute('pivotTables',XmlBoolToStr(FPivotTables));
  if FSelectUnlockedCells <> False then 
    AWriter.AddAttribute('selectUnlockedCells',XmlBoolToStr(FSelectUnlockedCells));
end;

procedure TCT_SheetProtection.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000373: FPassword := XmlStrToIntDef('$' + AAttributes.Values[i],0);
      $00000219: FSheet := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002EA: FObjects := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003C7: FScenarios := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000047C: FFormatCells := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000056A: FFormatColumns := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000434: FFormatRows := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000576: FInsertColumns := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000440: FInsertRows := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006BE: FInsertHyperlinks := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000554: FDeleteColumns := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000041E: FDeleteRows := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000006C5: FSelectLockedCells := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001C8: FSort := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000041F: FAutoFilter := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000048D: FPivotTables := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000007A8: FSelectUnlockedCells := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_SheetProtection.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 17;
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

destructor TCT_SheetProtection.Destroy;
begin
end;

procedure TCT_SheetProtection.Clear;
begin
  FAssigneds := [];
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

{ TCT_CustomSheetViews }

function  TCT_CustomSheetViews.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCustomSheetViewXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustomSheetViews.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'customSheetView' then 
    Result := FCustomSheetViewXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomSheetViews.Write(AWriter: TXpgWriteXML);
begin
  FCustomSheetViewXpgList.Write(AWriter,'customSheetView');
end;

constructor TCT_CustomSheetViews.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCustomSheetViewXpgList := TCT_CustomSheetViewXpgList.Create(FOwner);
end;

destructor TCT_CustomSheetViews.Destroy;
begin
  FCustomSheetViewXpgList.Free;
end;

procedure TCT_CustomSheetViews.Clear;
begin
  FAssigneds := [];
  FCustomSheetViewXpgList.Clear;
end;

{ TCT_OleObjects }

function  TCT_OleObjects.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FOleObjectXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_OleObjects.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'oleObject' then 
    Result := FOleObjectXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_OleObjects.Write(AWriter: TXpgWriteXML);
begin
  FOleObjectXpgList.Write(AWriter,'oleObject');
end;

constructor TCT_OleObjects.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FOleObjectXpgList := TCT_OleObjectXpgList.Create(FOwner);
end;

destructor TCT_OleObjects.Destroy;
begin
  FOleObjectXpgList.Free;
end;

procedure TCT_OleObjects.Clear;
begin
  FAssigneds := [];
  FOleObjectXpgList.Clear;
end;

{ TCT_SheetDimension }

function  TCT_SheetDimension.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRef <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SheetDimension.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SheetDimension.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ref',FRef);
end;

procedure TCT_SheetDimension.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'ref' then 
    FRef := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SheetDimension.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_SheetDimension.Destroy;
begin
end;

procedure TCT_SheetDimension.Clear;
begin
  FAssigneds := [];
  FRef := '';
end;

{ TCT_Cols }

function  TCT_Cols.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  if System.Assigned(FOwner.OnWriteWorksheetColsCol) then // TODO
    ElemsAssigned := 1;

//  Inc(ElemsAssigned,FCol.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Cols.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'col' then 
  begin
    Result := FCol;
    TCT_Col(Result).Clear;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Cols.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOwner.FOnWriteWorksheetColsCol) then 
    while True do 
    begin
      WriteElement := False;
      FOwner.FOnWriteWorksheetColsCol(FCol,WriteElement);
      if WriteElement = False then 
        Break;
//      if FCol.Assigned then   // TODO
//      begin
        FCol.WriteAttributes(AWriter);
        AWriter.SimpleTag('col');
//      end;
    end;
end;

procedure TCT_Cols.AfterTag;
begin
  if (FCol.FAssigneds <> []) and System.Assigned(FOwner.OnReadWorksheetColsCol) then 
    FOwner.OnReadWorksheetColsCol(FCol);
  FCol.FAssigneds := [];
end;

constructor TCT_Cols.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCol := TCT_Col.Create(FOwner);
end;

destructor TCT_Cols.Destroy;
begin
  FCol.Free;
end;

procedure TCT_Cols.Clear;
begin
  FAssigneds := [];
  FCol.Clear;
end;

{ TCT_ColsXpgList }

function  TCT_ColsXpgList.GetItems(Index: integer): TCT_Cols;
begin
  Result := TCT_Cols(inherited Items[Index]);
end;

function  TCT_ColsXpgList.Add: TCT_Cols;
begin
  Result := TCT_Cols.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ColsXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ColsXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_SheetData }

function  TCT_SheetData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
//  ElemsAssigned := 0;
  FAssigneds := [];
  // TODO This element must be written even if empty.
  // Use WriteCondition = "True"
  if True then
    ElemsAssigned := 1;
//  Inc(ElemsAssigned,FRow.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SheetData.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'row' then 
  begin
    Result := FRow;
    TCT_Row(Result).Clear;
  end
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SheetData.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteRow) then 
    while True do 
    begin
      WriteElement := False;
      FOnWriteRow(FRow,WriteElement);
      if WriteElement = False then 
        Break;
      FRow.CheckAssigned; // TODO CheckAssigned and CheckAssigned must return True if there childs with write event.
      if FRow.Assigned then
      begin
        FRow.WriteAttributes(AWriter);
//        if xaElements in FRow.FAssigneds then
//        begin
          AWriter.BeginTag('row');
          FRow.Write(AWriter);
          AWriter.EndTag;
//        end
//        else
//          AWriter.SimpleTag('row');
      end;
    end;
end;

procedure TCT_SheetData.AfterTag;
begin
  if (FRow.FAssigneds <> []) and System.Assigned(FOnReadRow) then 
    FOnReadRow(FRow);
  FRow.FAssigneds := [];
end;

constructor TCT_SheetData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FRow := TCT_Row.Create(FOwner);
end;

destructor TCT_SheetData.Destroy;
begin
  FRow.Free;
end;

procedure TCT_SheetData.Clear;
begin
  FAssigneds := [];
  FRow.Clear;
end;

{ TCT_SheetCalcPr }

function  TCT_SheetCalcPr.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FFullCalcOnLoad <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_SheetCalcPr.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_SheetCalcPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FFullCalcOnLoad <> False then 
    AWriter.AddAttribute('fullCalcOnLoad',XmlBoolToStr(FFullCalcOnLoad));
end;

procedure TCT_SheetCalcPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'fullCalcOnLoad' then 
    FFullCalcOnLoad := XmlStrToBoolDef(AAttributes.Values[0],False)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_SheetCalcPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FFullCalcOnLoad := False;
end;

destructor TCT_SheetCalcPr.Destroy;
begin
end;

procedure TCT_SheetCalcPr.Clear;
begin
  FAssigneds := [];
  FFullCalcOnLoad := False;
end;

{ TCT_ProtectedRanges }

function  TCT_ProtectedRanges.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FProtectedRangeXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ProtectedRanges.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'protectedRange' then 
    Result := FProtectedRangeXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ProtectedRanges.Write(AWriter: TXpgWriteXML);
begin
  FProtectedRangeXpgList.Write(AWriter,'protectedRange');
end;

constructor TCT_ProtectedRanges.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FProtectedRangeXpgList := TCT_ProtectedRangeXpgList.Create(FOwner);
end;

destructor TCT_ProtectedRanges.Destroy;
begin
  FProtectedRangeXpgList.Free;
end;

procedure TCT_ProtectedRanges.Clear;
begin
  FAssigneds := [];
  FProtectedRangeXpgList.Clear;
end;

{ TCT_Scenarios }

function  TCT_Scenarios.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCurrent <> 0 then 
    Inc(AttrsAssigned);
  if FShow <> 0 then 
    Inc(AttrsAssigned);
  if FSqrefXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FScenarioXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Scenarios.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'scenario' then 
    Result := FScenarioXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Scenarios.Write(AWriter: TXpgWriteXML);
begin
  FScenarioXpgList.Write(AWriter,'scenario');
end;

procedure TCT_Scenarios.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCurrent <> 0 then 
    AWriter.AddAttribute('current',XmlIntToStr(FCurrent));
  if FShow <> 0 then 
    AWriter.AddAttribute('show',XmlIntToStr(FShow));
  if FSqrefXpgList.Count > 0 then 
    AWriter.AddAttribute('sqref',FSqrefXpgList.DelimitedText);
end;

procedure TCT_Scenarios.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000303: FCurrent := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001C1: FShow := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000221: FSqrefXpgList.DelimitedText := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Scenarios.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FScenarioXpgList := TCT_ScenarioXpgList.Create(FOwner);
  FSqrefXpgList := TStringXpgList.Create;
end;

destructor TCT_Scenarios.Destroy;
begin
  FScenarioXpgList.Free;
  FSqrefXpgList.Free; // TODO This line is not created by XPG
end;

procedure TCT_Scenarios.Clear;
begin
  FAssigneds := [];
  FScenarioXpgList.Clear;
  FCurrent := 0;
  FShow := 0;
  FSqrefXpgList.DelimitedText := '';
end;

{ TCT_DataConsolidate }

function  TCT_DataConsolidate.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FFunction <> stdcfSum then 
    Inc(AttrsAssigned);
  if FLeftLabels <> False then 
    Inc(AttrsAssigned);
  if FTopLabels <> False then 
    Inc(AttrsAssigned);
  if FLink <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FDataRefs.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DataConsolidate.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'dataRefs' then 
    Result := FDataRefs
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DataConsolidate.Write(AWriter: TXpgWriteXML);
begin
  if FDataRefs.Assigned then 
  begin
    FDataRefs.WriteAttributes(AWriter);
    if xaElements in FDataRefs.FAssigneds then 
    begin
      AWriter.BeginTag('dataRefs');
      FDataRefs.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dataRefs');
  end;
end;

procedure TCT_DataConsolidate.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FFunction <> stdcfSum then 
    AWriter.AddAttribute('function',StrTST_DataConsolidateFunction[Integer(FFunction)]);
  if FLeftLabels <> False then 
    AWriter.AddAttribute('leftLabels',XmlBoolToStr(FLeftLabels));
  if FTopLabels <> False then 
    AWriter.AddAttribute('topLabels',XmlBoolToStr(FTopLabels));
  if FLink <> False then 
    AWriter.AddAttribute('link',XmlBoolToStr(FLink));
end;

procedure TCT_DataConsolidate.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000366: FFunction := TST_DataConsolidateFunction(StrToEnum('stdcf' + AAttributes.Values[i]));
      $000003FE: FLeftLabels := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003A6: FTopLabels := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000001AE: FLink := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DataConsolidate.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FDataRefs := TCT_DataRefs.Create(FOwner);
  FFunction := stdcfSum;
  FLeftLabels := False;
  FTopLabels := False;
  FLink := False;
end;

destructor TCT_DataConsolidate.Destroy;
begin
  FDataRefs.Free;
end;

procedure TCT_DataConsolidate.Clear;
begin
  FAssigneds := [];
  FDataRefs.Clear;
  FFunction := stdcfSum;
  FLeftLabels := False;
  FTopLabels := False;
  FLink := False;
end;

{ TCT_MergeCells }

function  TCT_MergeCells.CheckAssigned: integer;
var
  B: boolean;
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then
    Inc(AttrsAssigned);
//  Inc(ElemsAssigned,FMergeCell.CheckAssigned);

  // TODO Must know if there is any child elements to write. Writing parent
  // element without childs may result in invalid file.
  if System.Assigned(FOnWriteMergeCell) then begin
    B := False;
    FOnWriteMergeCell(Nil,B);
    if B then
      Inc(ElemsAssigned);
  end;
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_MergeCells.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'mergeCell' then
  begin
    Result := FMergeCell;
    TCT_MergeCell(Result).Clear;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_MergeCells.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteMergeCell) then
    while True do
    begin
      WriteElement := False;
      FOnWriteMergeCell(FMergeCell,WriteElement);
      if WriteElement = True then begin
//      if FMergeCell.Assigned then   // TODO
//      begin
        FMergeCell.WriteAttributes(AWriter);
        AWriter.SimpleTag('mergeCell');
        FMergeCell.Clear;
//      end;
      end
      else
        Break;
    end;
end;

procedure TCT_MergeCells.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_MergeCells.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

procedure TCT_MergeCells.AfterTag;
begin
  if (FMergeCell.FAssigneds <> []) and System.Assigned(FOnReadMergeCell) then 
    FOnReadMergeCell(FMergeCell);
  FMergeCell.FAssigneds := [];
end;

constructor TCT_MergeCells.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FMergeCell := TCT_MergeCell.Create(FOwner);
end;

destructor TCT_MergeCells.Destroy;
begin
  FMergeCell.Free;
end;

procedure TCT_MergeCells.Clear;
begin
  FAssigneds := [];
  FMergeCell.Clear;
  FCount := 0;
end;

{ TCT_ConditionalFormatting }

function  TCT_ConditionalFormatting.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FPivot <> False then 
    Inc(AttrsAssigned);
  if FSqrefXpgList.Count > 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCfRuleXpgList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ConditionalFormatting.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000261: Result := FCfRuleXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ConditionalFormatting.Write(AWriter: TXpgWriteXML);
begin
  FCfRuleXpgList.Write(AWriter,'cfRule');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_ConditionalFormatting.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FPivot <> False then 
    AWriter.AddAttribute('pivot',XmlBoolToStr(FPivot));
  if FSqrefXpgList.Count > 0 then 
    AWriter.AddAttribute('sqref',FSqrefXpgList.DelimitedText);
end;

procedure TCT_ConditionalFormatting.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000232: FPivot := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000221: FSqrefXpgList.DelimitedText := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_ConditionalFormatting.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 2;
  FCfRuleXpgList := TCT_CfRuleXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FPivot := False;
  FSqrefXpgList := TStringXpgList.Create;
end;

destructor TCT_ConditionalFormatting.Destroy;
begin
  FCfRuleXpgList.Free;
  FSqrefXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_ConditionalFormatting.Clear;
begin
  FAssigneds := [];
  FCfRuleXpgList.Clear;
  FExtLst.Clear;
  FPivot := False;
  FSqrefXpgList.DelimitedText := '';
end;

{ TCT_ConditionalFormattingXpgList }

function  TCT_ConditionalFormattingXpgList.GetItems(Index: integer): TCT_ConditionalFormatting;
begin
  Result := TCT_ConditionalFormatting(inherited Items[Index]);
end;

function  TCT_ConditionalFormattingXpgList.Add: TCT_ConditionalFormatting;
begin
  Result := TCT_ConditionalFormatting.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ConditionalFormattingXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ConditionalFormattingXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DataValidations }

function  TCT_DataValidations.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDisablePrompts <> False then 
    Inc(AttrsAssigned);
  if FXWindow <> 0 then 
    Inc(AttrsAssigned);
  if FYWindow <> 0 then 
    Inc(AttrsAssigned);
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FDataValidationXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DataValidations.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'dataValidation' then 
    Result := FDataValidationXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DataValidations.Write(AWriter: TXpgWriteXML);
begin
  FDataValidationXpgList.Write(AWriter,'dataValidation');
end;

procedure TCT_DataValidations.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FDisablePrompts <> False then 
    AWriter.AddAttribute('disablePrompts',XmlBoolToStr(FDisablePrompts));
  if FXWindow <> 0 then 
    AWriter.AddAttribute('xWindow',XmlIntToStr(FXWindow));
  if FYWindow <> 0 then 
    AWriter.AddAttribute('yWindow',XmlIntToStr(FYWindow));
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_DataValidations.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000005C9: FDisablePrompts := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000002F0: FXWindow := XmlStrToIntDef(AAttributes.Values[i],0);
      $000002F1: FYWindow := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_DataValidations.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FDataValidationXpgList := TCT_DataValidationXpgList.Create(FOwner);
  FDisablePrompts := False;
end;

destructor TCT_DataValidations.Destroy;
begin
  FDataValidationXpgList.Free;
end;

procedure TCT_DataValidations.Clear;
begin
  FAssigneds := [];
  FDataValidationXpgList.Clear;
  FDisablePrompts := False;
  FXWindow := 0;
  FYWindow := 0;
  FCount := 0;
end;

{ TCT_Hyperlinks }

function  TCT_Hyperlinks.CheckAssigned: integer;
var
  B: boolean;
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
//  Inc(ElemsAssigned,FHyperlink.CheckAssigned);  // TODO See TCT_MergeCells.CheckAssigned
  if System.Assigned(FOnWriteHyperlink) then begin
    B := False;
    FOnWriteHyperlink(Nil,B);
    if B then
      Inc(ElemsAssigned);
  end;
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Hyperlinks.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'hyperlink' then
  begin
    Result := FHyperlink;
    TCT_Hyperlink(Result).Clear;
  end
  else
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Hyperlinks.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteHyperlink) then
    while True do 
    begin
      WriteElement := False;
      FHyperlink.Clear;
      FOnWriteHyperlink(FHyperlink,WriteElement);
      if WriteElement = False then 
        Break;
//      if FHyperlink.Assigned then   // TODO
//      begin
        FHyperlink.WriteAttributes(AWriter);
        AWriter.SimpleTag('hyperlink');
//      end;
    end;
end;

procedure TCT_Hyperlinks.AfterTag;
begin
  if (FHyperlink.FAssigneds <> []) and System.Assigned(FOnReadHyperlink) then 
    FOnReadHyperlink(FHyperlink);
  FHyperlink.FAssigneds := [];
end;

constructor TCT_Hyperlinks.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FHyperlink := TCT_Hyperlink.Create(FOwner);
end;

destructor TCT_Hyperlinks.Destroy;
begin
  FHyperlink.Free;
end;

procedure TCT_Hyperlinks.Clear;
begin
  FAssigneds := [];
  FHyperlink.Clear;
end;

{ TCT_CustomProperties }

function  TCT_CustomProperties.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCustomPrXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CustomProperties.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'customPr' then 
    Result := FCustomPrXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CustomProperties.Write(AWriter: TXpgWriteXML);
begin
  FCustomPrXpgList.Write(AWriter,'customPr');
end;

constructor TCT_CustomProperties.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCustomPrXpgList := TCT_CustomPropertyXpgList.Create(FOwner);
end;

destructor TCT_CustomProperties.Destroy;
begin
  FCustomPrXpgList.Free;
end;

procedure TCT_CustomProperties.Clear;
begin
  FAssigneds := [];
  FCustomPrXpgList.Clear;
end;

{ TCT_CellWatches }

function  TCT_CellWatches.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCellWatchXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_CellWatches.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cellWatch' then 
    Result := FCellWatchXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_CellWatches.Write(AWriter: TXpgWriteXML);
begin
  FCellWatchXpgList.Write(AWriter,'cellWatch');
end;

constructor TCT_CellWatches.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCellWatchXpgList := TCT_CellWatchXpgList.Create(FOwner);
end;

destructor TCT_CellWatches.Destroy;
begin
  FCellWatchXpgList.Free;
end;

procedure TCT_CellWatches.Clear;
begin
  FAssigneds := [];
  FCellWatchXpgList.Clear;
end;

{ TCT_IgnoredErrors }

function  TCT_IgnoredErrors.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FIgnoredErrorXpgList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_IgnoredErrors.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000004F2: Result := FIgnoredErrorXpgList.Add;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_IgnoredErrors.Write(AWriter: TXpgWriteXML);
begin
  FIgnoredErrorXpgList.Write(AWriter,'ignoredError');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_IgnoredErrors.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 0;
  FIgnoredErrorXpgList := TCT_IgnoredErrorXpgList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_IgnoredErrors.Destroy;
begin
  FIgnoredErrorXpgList.Free;
  FExtLst.Free;
end;

procedure TCT_IgnoredErrors.Clear;
begin
  FAssigneds := [];
  FIgnoredErrorXpgList.Clear;
  FExtLst.Clear;
end;

{ TCT_SmartTags }

function  TCT_SmartTags.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FCellSmartTagsXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_SmartTags.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cellSmartTags' then 
    Result := FCellSmartTagsXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_SmartTags.Write(AWriter: TXpgWriteXML);
begin
  FCellSmartTagsXpgList.Write(AWriter,'cellSmartTags');
end;

constructor TCT_SmartTags.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FCellSmartTagsXpgList := TCT_CellSmartTagsXpgList.Create(FOwner);
end;

destructor TCT_SmartTags.Destroy;
begin
  FCellSmartTagsXpgList.Free;
end;

procedure TCT_SmartTags.Clear;
begin
  FAssigneds := [];
  FCellSmartTagsXpgList.Clear;
end;

{ TCT_Controls }

function  TCT_Controls.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FControlXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Controls.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'control' then 
    Result := FControlXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Controls.Write(AWriter: TXpgWriteXML);
begin
  FControlXpgList.Write(AWriter,'control');
end;

constructor TCT_Controls.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FControlXpgList := TCT_ControlXpgList.Create(FOwner);
end;

destructor TCT_Controls.Destroy;
begin
  FControlXpgList.Free;
end;

procedure TCT_Controls.Clear;
begin
  FAssigneds := [];
  FControlXpgList.Clear;
end;

{ TCT_TableParts }

function  TCT_TableParts.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FTablePartXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TableParts.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'tablePart' then 
    Result := FTablePartXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_TableParts.Write(AWriter: TXpgWriteXML);
begin
  FTablePartXpgList.Write(AWriter,'tablePart');
end;

procedure TCT_TableParts.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_TableParts.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then 
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TableParts.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FTablePartXpgList := TCT_TablePartXpgList.Create(FOwner);
end;

destructor TCT_TableParts.Destroy;
begin
  FTablePartXpgList.Free;
end;

procedure TCT_TableParts.Clear;
begin
  FAssigneds := [];
  FTablePartXpgList.Clear;
  FCount := 0;
end;

{ TCT_Stylesheet }

function  TCT_Stylesheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FNumFmts.CheckAssigned);
  Inc(ElemsAssigned,FFonts.CheckAssigned);
  Inc(ElemsAssigned,FFills.CheckAssigned);
  Inc(ElemsAssigned,FBorders.CheckAssigned);
  Inc(ElemsAssigned,FCellStyleXfs.CheckAssigned);
  Inc(ElemsAssigned,FCellXfs.CheckAssigned);
  Inc(ElemsAssigned,FCellStyles.CheckAssigned);
  Inc(ElemsAssigned,FDxfs.CheckAssigned);
  Inc(ElemsAssigned,FTableStyles.CheckAssigned);
  Inc(ElemsAssigned,FColors.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Stylesheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002EA: Result := FNumFmts;
    $0000022A: Result := FFonts;
    $0000021A: Result := FFills;
    $000002F1: Result := FBorders;
    $000004E2: Result := FCellStyleXfs;
    $000002D1: Result := FCellXfs;
    $00000424: Result := FCellStyles;
    $000001B5: Result := FDxfs;
    $0000048C: Result := FTableStyles;
    $00000292: Result := FColors;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Stylesheet.Write(AWriter: TXpgWriteXML);
begin
  if FNumFmts.Assigned then 
  begin
    FNumFmts.WriteAttributes(AWriter);
    if xaElements in FNumFmts.FAssigneds then 
    begin
      AWriter.BeginTag('numFmts');
      FNumFmts.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('numFmts');
  end;
  if FFonts.Assigned then 
  begin
    FFonts.WriteAttributes(AWriter);
    if xaElements in FFonts.FAssigneds then 
    begin
      AWriter.BeginTag('fonts');
      FFonts.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('fonts');
  end;
  if FFills.Assigned then 
  begin
    FFills.WriteAttributes(AWriter);
    if xaElements in FFills.FAssigneds then 
    begin
      AWriter.BeginTag('fills');
      FFills.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('fills');
  end;
  if FBorders.Assigned then 
  begin
    FBorders.WriteAttributes(AWriter);
    if xaElements in FBorders.FAssigneds then 
    begin
      AWriter.BeginTag('borders');
      FBorders.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('borders');
  end;
  if FCellStyleXfs.Assigned then 
  begin
    FCellStyleXfs.WriteAttributes(AWriter);
    if xaElements in FCellStyleXfs.FAssigneds then 
    begin
      AWriter.BeginTag('cellStyleXfs');
      FCellStyleXfs.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cellStyleXfs');
  end;
  if FCellXfs.Assigned then 
  begin
    FCellXfs.WriteAttributes(AWriter);
    if xaElements in FCellXfs.FAssigneds then 
    begin
      AWriter.BeginTag('cellXfs');
      FCellXfs.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cellXfs');
  end;
  if FCellStyles.Assigned then 
  begin
    FCellStyles.WriteAttributes(AWriter);
    if xaElements in FCellStyles.FAssigneds then 
    begin
      AWriter.BeginTag('cellStyles');
      FCellStyles.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cellStyles');
  end;
  if FDxfs.Assigned then 
  begin
    FDxfs.WriteAttributes(AWriter);
    if xaElements in FDxfs.FAssigneds then 
    begin
      AWriter.BeginTag('dxfs');
      FDxfs.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dxfs');
  end;
  if FTableStyles.Assigned then 
  begin
    FTableStyles.WriteAttributes(AWriter);
    if xaElements in FTableStyles.FAssigneds then 
    begin
      AWriter.BeginTag('tableStyles');
      FTableStyles.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('tableStyles');
  end;
  if FColors.Assigned then 
    if xaElements in FColors.FAssigneds then 
    begin
      AWriter.BeginTag('colors');
      FColors.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colors');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Stylesheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 11;
  FAttributeCount := 0;
  FNumFmts := TCT_NumFmts.Create(FOwner);
  FFonts := TCT_Fonts.Create(FOwner);
  FFills := TCT_Fills.Create(FOwner);
  FBorders := TCT_Borders.Create(FOwner);
  FCellStyleXfs := TCT_CellStyleXfs.Create(FOwner);
  FCellXfs := TCT_CellXfs.Create(FOwner);
  FCellStyles := TCT_CellStyles.Create(FOwner);
  FDxfs := TCT_Dxfs.Create(FOwner);
  FTableStyles := TCT_TableStyles.Create(FOwner);
  FColors := TCT_Colors.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Stylesheet.Destroy;
begin
  FNumFmts.Free;
  FFonts.Free;
  FFills.Free;
  FBorders.Free;
  FCellStyleXfs.Free;
  FCellXfs.Free;
  FCellStyles.Free;
  FDxfs.Free;
  FTableStyles.Free;
  FColors.Free;
  FExtLst.Free;
end;

procedure TCT_Stylesheet.Clear;
begin
  FAssigneds := [];
  FNumFmts.Clear;
  FFonts.Clear;
  FFills.Clear;
  FBorders.Clear;
  FCellStyleXfs.Clear;
  FCellXfs.Clear;
  FCellStyles.Clear;
  FDxfs.Clear;
  FTableStyles.Clear;
  FColors.Clear;
  FExtLst.Clear;
end;

{ TCT_Sst }

function  TCT_Sst.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then 
    Inc(AttrsAssigned);
  if FUniqueCount <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FSi.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Sst.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000000DC: begin
      Result := FSi;
      TCT_Rst(Result).Clear;
    end;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Sst.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if System.Assigned(FOnWriteSi) then
    while True do begin
      WriteElement := False;
      FOnWriteSi(FSi,WriteElement);
      if WriteElement = False then
        Break;
      FSi.CheckAssigned;

// TODO  Must always be written, even if empty.
// Otherwise SST index will go out of sync.
//     if FSi.Assigned then begin
//        if xaElements in FSi.FAssigneds then
//        begin
          AWriter.BeginTag('si');
          FSi.Write(AWriter);
          AWriter.EndTag;
//        end
//        else
//          AWriter.SimpleTag('si');
        FSi.Clear; // TODO
//      end;
    end;
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Sst.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then 
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
  if FUniqueCount <> 0 then 
    AWriter.AddAttribute('uniqueCount',XmlIntToStr(FUniqueCount));
end;

procedure TCT_Sst.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $00000229: FCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004A0: FUniqueCount := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

procedure TCT_Sst.AfterTag;
begin
  if (FSi.FAssigneds <> []) and System.Assigned(FOnReadSi) then 
    FOnReadSi(FSi);
  FSi.FAssigneds := [];
end;

constructor TCT_Sst.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 2;
  FAttributeCount := 2;
  FSi := TCT_Rst.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Sst.Destroy;
begin
  FSi.Free;
  FExtLst.Free;
end;

procedure TCT_Sst.Clear;
begin
  FAssigneds := [];
  FSi.Clear;
  FExtLst.Clear;
  FCount := 0;
  FUniqueCount := 0;
end;

{ TCT_Workbook }

function  TCT_Workbook.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FFileVersion.CheckAssigned);
  Inc(ElemsAssigned,FFileSharing.CheckAssigned);
  Inc(ElemsAssigned,FWorkbookPr.CheckAssigned);
  Inc(ElemsAssigned,FWorkbookProtection.CheckAssigned);
  Inc(ElemsAssigned,FBookViews.CheckAssigned);
  Inc(ElemsAssigned,FSheets.CheckAssigned);
  Inc(ElemsAssigned,FFunctionGroups.CheckAssigned);
  Inc(ElemsAssigned,FExternalReferences.CheckAssigned);
  Inc(ElemsAssigned,FDefinedNames.CheckAssigned);
  Inc(ElemsAssigned,FCalcPr.CheckAssigned);
  Inc(ElemsAssigned,FOleSize.CheckAssigned);
  Inc(ElemsAssigned,FCustomWorkbookViews.CheckAssigned);
  Inc(ElemsAssigned,FPivotCaches.CheckAssigned);
  Inc(ElemsAssigned,FSmartTagPr.CheckAssigned);
  Inc(ElemsAssigned,FSmartTagTypes.CheckAssigned);
  Inc(ElemsAssigned,FWebPublishing.CheckAssigned);
  Inc(ElemsAssigned,FFileRecoveryPrXpgList.CheckAssigned);
  Inc(ElemsAssigned,FWebPublishObjects.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Workbook.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000486: Result := FFileVersion;
    $0000046C: Result := FFileSharing;
    $00000430: Result := FWorkbookPr;
    $00000795: Result := FWorkbookProtection;
    $000003B9: Result := FBookViews;
    $0000028C: Result := FSheets;
    $000005E6: Result := FFunctionGroups;
    $00000765: Result := FExternalReferences;
    $000004C3: begin
      Result := FDefinedNames;
      TCT_DefinedNames(Result).Clear;
    end;
    $00000255: Result := FCalcPr;
    $000002DB: Result := FOleSize;
    $000007F7: Result := FCustomWorkbookViews;
    $00000479: Result := FPivotCaches;
    $00000405: Result := FSmartTagPr;
    $00000558: Result := FSmartTagTypes;
    $00000553: Result := FWebPublishing;
    $000005B1: Result := FFileRecoveryPrXpgList.Add;
    $000006DF: Result := FWebPublishObjects;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Workbook.Write(AWriter: TXpgWriteXML);
var
  WriteElement: boolean;
begin
  if FFileVersion.Assigned then 
  begin
    FFileVersion.WriteAttributes(AWriter);
    AWriter.SimpleTag('fileVersion');
  end;
  if FFileSharing.Assigned then 
  begin
    FFileSharing.WriteAttributes(AWriter);
    AWriter.SimpleTag('fileSharing');
  end;
  if FWorkbookPr.Assigned then 
  begin
    FWorkbookPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('workbookPr');
  end;
  if FWorkbookProtection.Assigned then 
  begin
    FWorkbookProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('workbookProtection');
  end;
  if FBookViews.Assigned then 
    if xaElements in FBookViews.FAssigneds then 
    begin
      AWriter.BeginTag('bookViews');
      FBookViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('bookViews');
  if FSheets.Assigned then 
    if xaElements in FSheets.FAssigneds then 
    begin
      AWriter.BeginTag('sheets');
      FSheets.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheets');
  if FFunctionGroups.Assigned then 
  begin
    FFunctionGroups.WriteAttributes(AWriter);
    if xaElements in FFunctionGroups.FAssigneds then 
    begin
      AWriter.BeginTag('functionGroups');
      FFunctionGroups.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('functionGroups');
  end;
  if FExternalReferences.Assigned then 
    if xaElements in FExternalReferences.FAssigneds then 
    begin
      AWriter.BeginTag('externalReferences');
      FExternalReferences.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('externalReferences');
  if System.Assigned(FOnWriteDefinedNames) then
    while True do 
    begin
      WriteElement := False;
      FOnWriteDefinedNames(FDefinedNames,WriteElement);
      if WriteElement = False then 
        Break;
      if FDefinedNames.Assigned then 
        if xaElements in FDefinedNames.FAssigneds then 
        begin
          AWriter.BeginTag('definedNames');
          FDefinedNames.Write(AWriter);
          AWriter.EndTag;
        end
        else 
          AWriter.SimpleTag('definedNames');
    end;
  if FCalcPr.Assigned then 
  begin
    FCalcPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('calcPr');
  end;
  if FOleSize.Assigned then 
  begin
    FOleSize.WriteAttributes(AWriter);
    AWriter.SimpleTag('oleSize');
  end;
  if FCustomWorkbookViews.Assigned then 
    if xaElements in FCustomWorkbookViews.FAssigneds then 
    begin
      AWriter.BeginTag('customWorkbookViews');
      FCustomWorkbookViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customWorkbookViews');
  if FPivotCaches.Assigned then 
    if xaElements in FPivotCaches.FAssigneds then 
    begin
      AWriter.BeginTag('pivotCaches');
      FPivotCaches.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('pivotCaches');
  if FSmartTagPr.Assigned then 
  begin
    FSmartTagPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('smartTagPr');
  end;
  if FSmartTagTypes.Assigned then 
    if xaElements in FSmartTagTypes.FAssigneds then 
    begin
      AWriter.BeginTag('smartTagTypes');
      FSmartTagTypes.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('smartTagTypes');
  if FWebPublishing.Assigned then 
  begin
    FWebPublishing.WriteAttributes(AWriter);
    AWriter.SimpleTag('webPublishing');
  end;
  FFileRecoveryPrXpgList.Write(AWriter,'fileRecoveryPr');
  if FWebPublishObjects.Assigned then 
  begin
    FWebPublishObjects.WriteAttributes(AWriter);
    if xaElements in FWebPublishObjects.FAssigneds then 
    begin
      AWriter.BeginTag('webPublishObjects');
      FWebPublishObjects.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('webPublishObjects');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Workbook.AfterTag;
begin
  if (FDefinedNames.FAssigneds <> []) and System.Assigned(FOnReadDefinedNames) then 
    FOnReadDefinedNames(FDefinedNames);
  FDefinedNames.FAssigneds := [];
end;

constructor TCT_Workbook.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 19;
  FAttributeCount := 0;
  FFileVersion := TCT_FileVersion.Create(FOwner);
  FFileSharing := TCT_FileSharing.Create(FOwner);
  FWorkbookPr := TCT_WorkbookPr.Create(FOwner);
  FWorkbookProtection := TCT_WorkbookProtection.Create(FOwner);
  FBookViews := TCT_BookViews.Create(FOwner);
  FSheets := TCT_Sheets.Create(FOwner);
  FFunctionGroups := TCT_FunctionGroups.Create(FOwner);
  FExternalReferences := TCT_ExternalReferences.Create(FOwner);
  FDefinedNames := TCT_DefinedNames.Create(FOwner);
  FCalcPr := TCT_CalcPr.Create(FOwner);
  FOleSize := TCT_OleSize.Create(FOwner);
  FCustomWorkbookViews := TCT_CustomWorkbookViews.Create(FOwner);
  FPivotCaches := TCT_PivotCaches.Create(FOwner);
  FSmartTagPr := TCT_SmartTagPr.Create(FOwner);
  FSmartTagTypes := TCT_SmartTagTypes.Create(FOwner);
  FWebPublishing := TCT_WebPublishing.Create(FOwner);
  FFileRecoveryPrXpgList := TCT_FileRecoveryPrXpgList.Create(FOwner);
  FWebPublishObjects := TCT_WebPublishObjects.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Workbook.Destroy;
begin
  FFileVersion.Free;
  FFileSharing.Free;
  FWorkbookPr.Free;
  FWorkbookProtection.Free;
  FBookViews.Free;
  FSheets.Free;
  FFunctionGroups.Free;
  FExternalReferences.Free;
  FDefinedNames.Free;
  FCalcPr.Free;
  FOleSize.Free;
  FCustomWorkbookViews.Free;
  FPivotCaches.Free;
  FSmartTagPr.Free;
  FSmartTagTypes.Free;
  FWebPublishing.Free;
  FFileRecoveryPrXpgList.Free;
  FWebPublishObjects.Free;
  FExtLst.Free;
end;

procedure TCT_Workbook.Clear;
begin
  FAssigneds := [];
  FFileVersion.Clear;
  FFileSharing.Clear;
  FWorkbookPr.Clear;
  FWorkbookProtection.Clear;
  FBookViews.Clear;
  FSheets.Clear;
  FFunctionGroups.Clear;
  FExternalReferences.Clear;
  FDefinedNames.Clear;
  FCalcPr.Clear;
  FOleSize.Clear;
  FCustomWorkbookViews.Clear;
  FPivotCaches.Clear;
  FSmartTagPr.Clear;
  FSmartTagTypes.Clear;
  FWebPublishing.Clear;
  FFileRecoveryPrXpgList.Clear;
  FWebPublishObjects.Clear;
  FExtLst.Clear;
end;

{ TCT_Comments }

function  TCT_Comments.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FAuthors.CheckAssigned);
  Inc(ElemsAssigned,FCommentList.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Comments.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000306: Result := FAuthors;
    $0000048F: Result := FCommentList;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Comments.Write(AWriter: TXpgWriteXML);
begin
  if FAuthors.Assigned then 
    if xaElements in FAuthors.FAssigneds then 
    begin
      AWriter.BeginTag('authors');
      FAuthors.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('authors');
//  if FCommentList.Assigned then  // TODO Has child with write event
//    if xaElements in FCommentList.FAssigneds then
//    begin
      AWriter.BeginTag('commentList');
      FCommentList.Write(AWriter);
      AWriter.EndTag;
//    end
//    else
//      AWriter.SimpleTag('commentList');
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Comments.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 0;
  FAuthors := TCT_Authors.Create(FOwner);
  FCommentList := TCT_CommentList.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Comments.Destroy;
begin
  FAuthors.Free;
  FCommentList.Free;
  FExtLst.Free;
end;

procedure TCT_Comments.Clear;
begin
  FAssigneds := [];
  FAuthors.Clear;
  FCommentList.Clear;
  FExtLst.Clear;
end;

{ TCT_Chartsheet }

function  TCT_Chartsheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetPr.CheckAssigned);
  Inc(ElemsAssigned,FSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FSheetProtection.CheckAssigned);
  Inc(ElemsAssigned,FCustomSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  Inc(ElemsAssigned,FDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawingHF.CheckAssigned);
  Inc(ElemsAssigned,FPicture.CheckAssigned);
  Inc(ElemsAssigned,FWebPublishItems.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Chartsheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002DB: Result := FSheetPr;
    $00000427: Result := FSheetViews;
    $00000640: Result := FSheetProtection;
    $000006A2: Result := FCustomSheetViews;
    $0000046E: Result := FPageMargins;
    $000003AE: Result := FPageSetup;
    $000004D8: Result := FHeaderFooter;
    $000002EC: Result := FDrawing;
    $00000541: Result := FLegacyDrawing;
    $000005CF: Result := FLegacyDrawingHF;
    $000002FC: Result := FPicture;
    $00000617: Result := FWebPublishItems;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Chartsheet.Write(AWriter: TXpgWriteXML);
begin
  if FSheetPr.Assigned then 
  begin
    FSheetPr.WriteAttributes(AWriter);
    if xaElements in FSheetPr.FAssigneds then 
    begin
      AWriter.BeginTag('sheetPr');
      FSheetPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetPr');
  end;
  if FSheetViews.Assigned then 
    if xaElements in FSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('sheetViews');
      FSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetViews');
  if FSheetProtection.Assigned then 
  begin
    FSheetProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetProtection');
  end;
  if FCustomSheetViews.Assigned then 
    if xaElements in FCustomSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('customSheetViews');
      FCustomSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customSheetViews');
  if FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageMargins');
  end;
  if FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetup');
  end;
  if FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.FAssigneds then 
    begin
      AWriter.BeginTag('headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('headerFooter');
  end;
  if FDrawing.Assigned then 
  begin
    FDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('drawing');
  end;
  if FLegacyDrawing.Assigned then 
  begin
    FLegacyDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawing');
  end;
  if FLegacyDrawingHF.Assigned then 
  begin
    FLegacyDrawingHF.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawingHF');
  end;
  if FPicture.Assigned then 
  begin
    FPicture.WriteAttributes(AWriter);
    AWriter.SimpleTag('picture');
  end;
  if FWebPublishItems.Assigned then 
  begin
    FWebPublishItems.WriteAttributes(AWriter);
    if xaElements in FWebPublishItems.FAssigneds then 
    begin
      AWriter.BeginTag('webPublishItems');
      FWebPublishItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('webPublishItems');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Chartsheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 13;
  FAttributeCount := 0;
  FSheetPr := TCT_ChartsheetPr.Create(FOwner);
  FSheetViews := TCT_ChartsheetViews.Create(FOwner);
  FSheetProtection := TCT_ChartsheetProtection.Create(FOwner);
  FCustomSheetViews := TCT_CustomChartsheetViews.Create(FOwner);
  FPageMargins := TCT_PageMargins.Create(FOwner);
  FPageSetup := TCT_CsPageSetup.Create(FOwner);
  FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
  FDrawing := TCT_Drawing.Create(FOwner);
  FLegacyDrawing := TCT_LegacyDrawing.Create(FOwner);
  FLegacyDrawingHF := TCT_LegacyDrawing.Create(FOwner);
  FPicture := TCT_SheetBackgroundPicture.Create(FOwner);
  FWebPublishItems := TCT_WebPublishItems.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Chartsheet.Destroy;
begin
  FSheetPr.Free;
  FSheetViews.Free;
  FSheetProtection.Free;
  FCustomSheetViews.Free;
  FPageMargins.Free;
  FPageSetup.Free;
  FHeaderFooter.Free;
  FDrawing.Free;
  FLegacyDrawing.Free;
  FLegacyDrawingHF.Free;
  FPicture.Free;
  FWebPublishItems.Free;
  FExtLst.Free;
end;

procedure TCT_Chartsheet.Clear;
begin
  FAssigneds := [];
  FSheetPr.Clear;
  FSheetViews.Clear;
  FSheetProtection.Clear;
  FCustomSheetViews.Clear;
  FPageMargins.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FDrawing.Clear;
  FLegacyDrawing.Clear;
  FLegacyDrawingHF.Clear;
  FPicture.Clear;
  FWebPublishItems.Clear;
  FExtLst.Clear;
end;

{ TCT_Dialogsheet }

function  TCT_Dialogsheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetPr.CheckAssigned);
  Inc(ElemsAssigned,FSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FSheetFormatPr.CheckAssigned);
  Inc(ElemsAssigned,FSheetProtection.CheckAssigned);
  Inc(ElemsAssigned,FCustomSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FPrintOptions.CheckAssigned);
  Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  Inc(ElemsAssigned,FDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawingHF.CheckAssigned);
  Inc(ElemsAssigned,FOleObjects.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Dialogsheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002DB: Result := FSheetPr;
    $00000427: Result := FSheetViews;
    $00000544: Result := FSheetFormatPr;
    $00000640: Result := FSheetProtection;
    $000006A2: Result := FCustomSheetViews;
    $00000519: Result := FPrintOptions;
    $0000046E: Result := FPageMargins;
    $000003AE: Result := FPageSetup;
    $000004D8: Result := FHeaderFooter;
    $000002EC: Result := FDrawing;
    $00000541: Result := FLegacyDrawing;
    $000005CF: Result := FLegacyDrawingHF;
    $0000040A: Result := FOleObjects;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Dialogsheet.Write(AWriter: TXpgWriteXML);
begin
  if FSheetPr.Assigned then 
  begin
    FSheetPr.WriteAttributes(AWriter);
    if xaElements in FSheetPr.FAssigneds then 
    begin
      AWriter.BeginTag('sheetPr');
      FSheetPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetPr');
  end;
  if FSheetViews.Assigned then 
    if xaElements in FSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('sheetViews');
      FSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetViews');
  if FSheetFormatPr.Assigned then 
  begin
    FSheetFormatPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetFormatPr');
  end;
  if FSheetProtection.Assigned then 
  begin
    FSheetProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetProtection');
  end;
  if FCustomSheetViews.Assigned then 
    if xaElements in FCustomSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('customSheetViews');
      FCustomSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customSheetViews');
  if FPrintOptions.Assigned then 
  begin
    FPrintOptions.WriteAttributes(AWriter);
    AWriter.SimpleTag('printOptions');
  end;
  if FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageMargins');
  end;
  if FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetup');
  end;
  if FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.FAssigneds then 
    begin
      AWriter.BeginTag('headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('headerFooter');
  end;
  if FDrawing.Assigned then 
  begin
    FDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('drawing');
  end;
  if FLegacyDrawing.Assigned then 
  begin
    FLegacyDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawing');
  end;
  if FLegacyDrawingHF.Assigned then 
  begin
    FLegacyDrawingHF.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawingHF');
  end;
  if FOleObjects.Assigned then 
    if xaElements in FOleObjects.FAssigneds then 
    begin
      AWriter.BeginTag('oleObjects');
      FOleObjects.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('oleObjects');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Dialogsheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 14;
  FAttributeCount := 0;
  FSheetPr := TCT_SheetPr.Create(FOwner);
  FSheetViews := TCT_SheetViews.Create(FOwner);
  FSheetFormatPr := TCT_SheetFormatPr.Create(FOwner);
  FSheetProtection := TCT_SheetProtection.Create(FOwner);
  FCustomSheetViews := TCT_CustomSheetViews.Create(FOwner);
  FPrintOptions := TCT_PrintOptions.Create(FOwner);
  FPageMargins := TCT_PageMargins.Create(FOwner);
  FPageSetup := TCT_PageSetup.Create(FOwner);
  FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
  FDrawing := TCT_Drawing.Create(FOwner);
  FLegacyDrawing := TCT_LegacyDrawing.Create(FOwner);
  FLegacyDrawingHF := TCT_LegacyDrawing.Create(FOwner);
  FOleObjects := TCT_OleObjects.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Dialogsheet.Destroy;
begin
  FSheetPr.Free;
  FSheetViews.Free;
  FSheetFormatPr.Free;
  FSheetProtection.Free;
  FCustomSheetViews.Free;
  FPrintOptions.Free;
  FPageMargins.Free;
  FPageSetup.Free;
  FHeaderFooter.Free;
  FDrawing.Free;
  FLegacyDrawing.Free;
  FLegacyDrawingHF.Free;
  FOleObjects.Free;
  FExtLst.Free;
end;

procedure TCT_Dialogsheet.Clear;
begin
  FAssigneds := [];
  FSheetPr.Clear;
  FSheetViews.Clear;
  FSheetFormatPr.Clear;
  FSheetProtection.Clear;
  FCustomSheetViews.Clear;
  FPrintOptions.Clear;
  FPageMargins.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FDrawing.Clear;
  FLegacyDrawing.Clear;
  FLegacyDrawingHF.Clear;
  FOleObjects.Clear;
  FExtLst.Clear;
end;

{ TCT_Worksheet }

function  TCT_Worksheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetPr.CheckAssigned);
  Inc(ElemsAssigned,FDimension.CheckAssigned);
  Inc(ElemsAssigned,FSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FSheetFormatPr.CheckAssigned);
  Inc(ElemsAssigned,FColsXpgList.CheckAssigned);
  Inc(ElemsAssigned,FSheetData.CheckAssigned);
  Inc(ElemsAssigned,FSheetCalcPr.CheckAssigned);
  Inc(ElemsAssigned,FSheetProtection.CheckAssigned);
  Inc(ElemsAssigned,FProtectedRanges.CheckAssigned);
  Inc(ElemsAssigned,FScenarios.CheckAssigned);
  Inc(ElemsAssigned,FAutoFilter.CheckAssigned);
  Inc(ElemsAssigned,FSortState.CheckAssigned);
  Inc(ElemsAssigned,FDataConsolidate.CheckAssigned);
  Inc(ElemsAssigned,FCustomSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FMergeCells.CheckAssigned);
  Inc(ElemsAssigned,FPhoneticPr.CheckAssigned);
  Inc(ElemsAssigned,FConditionalFormattingXpgList.CheckAssigned);
  Inc(ElemsAssigned,FDataValidations.CheckAssigned);
  Inc(ElemsAssigned,FHyperlinks.CheckAssigned);
  Inc(ElemsAssigned,FPrintOptions.CheckAssigned);
  Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  Inc(ElemsAssigned,FRowBreaks.CheckAssigned);
  Inc(ElemsAssigned,FColBreaks.CheckAssigned);
  Inc(ElemsAssigned,FCustomProperties.CheckAssigned);
  Inc(ElemsAssigned,FCellWatches.CheckAssigned);
  Inc(ElemsAssigned,FIgnoredErrors.CheckAssigned);
  Inc(ElemsAssigned,FSmartTags.CheckAssigned);
  Inc(ElemsAssigned,FDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawingHF.CheckAssigned);
  Inc(ElemsAssigned,FPicture.CheckAssigned);
  Inc(ElemsAssigned,FOleObjects.CheckAssigned);
  Inc(ElemsAssigned,FControls.CheckAssigned);
  Inc(ElemsAssigned,FWebPublishItems.CheckAssigned);
  Inc(ElemsAssigned,FTableParts.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Worksheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002DB: Result := FSheetPr;
    $000003C6: Result := FDimension;
    $00000427: Result := FSheetViews;
    $00000544: Result := FSheetFormatPr;
    $000001B1: Result := FColsXpgList.Add;
    $00000393: Result := FSheetData;
    $0000044E: Result := FSheetCalcPr;
    $00000640: Result := FSheetProtection;
    $0000062A: Result := FProtectedRanges;
    $000003C7: Result := FScenarios;
    $0000041F: Result := FAutoFilter;
    $000003C9: Result := FSortState;
    $0000060F: Result := FDataConsolidate;
    $000006A2: Result := FCustomSheetViews;
    $00000403: Result := FMergeCells;
    $0000041C: Result := FPhoneticPr;
    $000008AF: Result := FConditionalFormattingXpgList.Add;
    $00000618: Result := FDataValidations;
    $00000449: Result := FHyperlinks;
    $00000519: Result := FPrintOptions;
    $0000046E: Result := FPageMargins;
    $000003AE: Result := FPageSetup;
    $000004D8: Result := FHeaderFooter;
    $000003B0: Result := FRowBreaks;
    $00000396: Result := FColBreaks;
    $000006C8: Result := FCustomProperties;
    $0000046F: Result := FCellWatches;
    $00000565: Result := FIgnoredErrors;
    $000003B6: Result := FSmartTags;
    $000002EC: Result := FDrawing;
    $00000541: Result := FLegacyDrawing;
    $000005CF: Result := FLegacyDrawingHF;
    $000002FC: Result := FPicture;
    $0000040A: Result := FOleObjects;
    $00000374: Result := FControls;
    $00000617: Result := FWebPublishItems;
    $00000412: Result := FTableParts;
    $00000284: Result := FExtLst;
    else 
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Worksheet.Write(AWriter: TXpgWriteXML);
begin
  if FSheetPr.Assigned then
  begin
    FSheetPr.WriteAttributes(AWriter);
    if xaElements in FSheetPr.FAssigneds then
    begin
      AWriter.BeginTag('sheetPr');
      FSheetPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetPr');
  end;
  if FDimension.Assigned then 
  begin
    FDimension.WriteAttributes(AWriter);
    AWriter.SimpleTag('dimension');
  end;
  if FSheetViews.Assigned then 
    if xaElements in FSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('sheetViews');
      FSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetViews');
  if FSheetFormatPr.Assigned then 
  begin
    FSheetFormatPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetFormatPr');
  end;
  FColsXpgList.Write(AWriter,'cols');
  if FSheetData.Assigned then 
    if xaElements in FSheetData.FAssigneds then 
    begin
      AWriter.BeginTag('sheetData');
      FSheetData.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetData');
  if FSheetCalcPr.Assigned then 
  begin
    FSheetCalcPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetCalcPr');
  end;
  if FSheetProtection.Assigned then 
  begin
    FSheetProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetProtection');
  end;
  if FProtectedRanges.Assigned then 
    if xaElements in FProtectedRanges.FAssigneds then 
    begin
      AWriter.BeginTag('protectedRanges');
      FProtectedRanges.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('protectedRanges');
  if FScenarios.Assigned then 
  begin
    FScenarios.WriteAttributes(AWriter);
    if xaElements in FScenarios.FAssigneds then 
    begin
      AWriter.BeginTag('scenarios');
      FScenarios.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('scenarios');
  end;
  if FAutoFilter.Assigned then 
  begin
    FAutoFilter.WriteAttributes(AWriter);
    if xaElements in FAutoFilter.FAssigneds then 
    begin
      AWriter.BeginTag('autoFilter');
      FAutoFilter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('autoFilter');
  end;
  if FSortState.Assigned then 
  begin
    FSortState.WriteAttributes(AWriter);
    if xaElements in FSortState.FAssigneds then 
    begin
      AWriter.BeginTag('sortState');
      FSortState.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sortState');
  end;
  if FDataConsolidate.Assigned then 
  begin
    FDataConsolidate.WriteAttributes(AWriter);
    if xaElements in FDataConsolidate.FAssigneds then 
    begin
      AWriter.BeginTag('dataConsolidate');
      FDataConsolidate.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dataConsolidate');
  end;
  if FCustomSheetViews.Assigned then 
    if xaElements in FCustomSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('customSheetViews');
      FCustomSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customSheetViews');
  if FMergeCells.Assigned then 
  begin
    FMergeCells.WriteAttributes(AWriter);
    if xaElements in FMergeCells.FAssigneds then 
    begin
      AWriter.BeginTag('mergeCells');
      FMergeCells.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('mergeCells');
  end;
  if FPhoneticPr.Assigned then 
  begin
    FPhoneticPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('phoneticPr');
  end;
  FConditionalFormattingXpgList.Write(AWriter,'conditionalFormatting');
  if FDataValidations.Assigned then 
  begin
    FDataValidations.WriteAttributes(AWriter);
    if xaElements in FDataValidations.FAssigneds then 
    begin
      AWriter.BeginTag('dataValidations');
      FDataValidations.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dataValidations');
  end;
  if FHyperlinks.Assigned then 
    if xaElements in FHyperlinks.FAssigneds then 
    begin
      AWriter.BeginTag('hyperlinks');
      FHyperlinks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('hyperlinks');
  if FPrintOptions.Assigned then 
  begin
    FPrintOptions.WriteAttributes(AWriter);
    AWriter.SimpleTag('printOptions');
  end;
  if FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageMargins');
  end;
  if FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetup');
  end;
  if FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.FAssigneds then 
    begin
      AWriter.BeginTag('headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('headerFooter');
  end;
  if FRowBreaks.Assigned then 
  begin
    FRowBreaks.WriteAttributes(AWriter);
    if xaElements in FRowBreaks.FAssigneds then 
    begin
      AWriter.BeginTag('rowBreaks');
      FRowBreaks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rowBreaks');
  end;
  if FColBreaks.Assigned then 
  begin
    FColBreaks.WriteAttributes(AWriter);
    if xaElements in FColBreaks.FAssigneds then 
    begin
      AWriter.BeginTag('colBreaks');
      FColBreaks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colBreaks');
  end;
  if FCustomProperties.Assigned then 
    if xaElements in FCustomProperties.FAssigneds then 
    begin
      AWriter.BeginTag('customProperties');
      FCustomProperties.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customProperties');
  if FCellWatches.Assigned then 
    if xaElements in FCellWatches.FAssigneds then 
    begin
      AWriter.BeginTag('cellWatches');
      FCellWatches.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('cellWatches');
  if FIgnoredErrors.Assigned then 
    if xaElements in FIgnoredErrors.FAssigneds then 
    begin
      AWriter.BeginTag('ignoredErrors');
      FIgnoredErrors.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('ignoredErrors');
  if FSmartTags.Assigned then 
    if xaElements in FSmartTags.FAssigneds then 
    begin
      AWriter.BeginTag('smartTags');
      FSmartTags.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('smartTags');
  if FDrawing.Assigned then 
  begin
    FDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('drawing');
  end;
  if FLegacyDrawing.Assigned then 
  begin
    FLegacyDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawing');
  end;
  if FLegacyDrawingHF.Assigned then 
  begin
    FLegacyDrawingHF.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawingHF');
  end;
  if FPicture.Assigned then 
  begin
    FPicture.WriteAttributes(AWriter);
    AWriter.SimpleTag('picture');
  end;
// Don't write OleObjects. Causes corrupt file as rId is not correct.
//  if FOleObjects.Assigned then
//    if xaElements in FOleObjects.FAssigneds then
//    begin
//      AWriter.BeginTag('oleObjects');
//      FOleObjects.Write(AWriter);
//      AWriter.EndTag;
//    end
//    else
//      AWriter.SimpleTag('oleObjects');
// Don't write Controls. Causes corrupt file as rId is not correct.
//  if FControls.Assigned then
//    if xaElements in FControls.FAssigneds then
//    begin
//      AWriter.BeginTag('controls');
//      FControls.Write(AWriter);
//      AWriter.EndTag;
//    end
//    else
//      AWriter.SimpleTag('controls');
  if FWebPublishItems.Assigned then
  begin
    FWebPublishItems.WriteAttributes(AWriter);
    if xaElements in FWebPublishItems.FAssigneds then 
    begin
      AWriter.BeginTag('webPublishItems');
      FWebPublishItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('webPublishItems');
  end;
  if FTableParts.Assigned then 
  begin
    FTableParts.WriteAttributes(AWriter);
    if xaElements in FTableParts.FAssigneds then 
    begin
      AWriter.BeginTag('tableParts');
      FTableParts.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('tableParts');
  end;
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Worksheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 38;
  FAttributeCount := 0;
  FSheetPr := TCT_SheetPr.Create(FOwner);
  FDimension := TCT_SheetDimension.Create(FOwner);
  FSheetViews := TCT_SheetViews.Create(FOwner);
  FSheetFormatPr := TCT_SheetFormatPr.Create(FOwner);
  FColsXpgList := TCT_ColsXpgList.Create(FOwner);
  FSheetData := TCT_SheetData.Create(FOwner);
  FSheetCalcPr := TCT_SheetCalcPr.Create(FOwner);
  FSheetProtection := TCT_SheetProtection.Create(FOwner);
  FProtectedRanges := TCT_ProtectedRanges.Create(FOwner);
  FScenarios := TCT_Scenarios.Create(FOwner);
  FAutoFilter := TCT_AutoFilter.Create(FOwner);
  FSortState := TCT_SortState.Create(FOwner);
  FDataConsolidate := TCT_DataConsolidate.Create(FOwner);
  FCustomSheetViews := TCT_CustomSheetViews.Create(FOwner);
  FMergeCells := TCT_MergeCells.Create(FOwner);
  FPhoneticPr := TCT_PhoneticPr.Create(FOwner);
  FConditionalFormattingXpgList := TCT_ConditionalFormattingXpgList.Create(FOwner);
  FDataValidations := TCT_DataValidations.Create(FOwner);
  FHyperlinks := TCT_Hyperlinks.Create(FOwner);
  FPrintOptions := TCT_PrintOptions.Create(FOwner);
  FPageMargins := TCT_PageMargins.Create(FOwner);
  FPageSetup := TCT_PageSetup.Create(FOwner);
  FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
  FRowBreaks := TCT_PageBreak.Create(FOwner);
  FColBreaks := TCT_PageBreak.Create(FOwner);
  FCustomProperties := TCT_CustomProperties.Create(FOwner);
  FCellWatches := TCT_CellWatches.Create(FOwner);
  FIgnoredErrors := TCT_IgnoredErrors.Create(FOwner);
  FSmartTags := TCT_SmartTags.Create(FOwner);
  FDrawing := TCT_Drawing.Create(FOwner);
  FLegacyDrawing := TCT_LegacyDrawing.Create(FOwner);
  FLegacyDrawingHF := TCT_LegacyDrawing.Create(FOwner);
  FPicture := TCT_SheetBackgroundPicture.Create(FOwner);
  FOleObjects := TCT_OleObjects.Create(FOwner);
  FControls := TCT_Controls.Create(FOwner);
  FWebPublishItems := TCT_WebPublishItems.Create(FOwner);
  FTableParts := TCT_TableParts.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Worksheet.Destroy;
begin
  FSheetPr.Free;
  FDimension.Free;
  FSheetViews.Free;
  FSheetFormatPr.Free;
  FColsXpgList.Free;
  FSheetData.Free;
  FSheetCalcPr.Free;
  FSheetProtection.Free;
  FProtectedRanges.Free;
  FScenarios.Free;
  FAutoFilter.Free;
  FSortState.Free;
  FDataConsolidate.Free;
  FCustomSheetViews.Free;
  FMergeCells.Free;
  FPhoneticPr.Free;
  FConditionalFormattingXpgList.Free;
  FDataValidations.Free;
  FHyperlinks.Free;
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
  FDrawing.Free;
  FLegacyDrawing.Free;
  FLegacyDrawingHF.Free;
  FPicture.Free;
  FOleObjects.Free;
  FControls.Free;
  FWebPublishItems.Free;
  FTableParts.Free;
  FExtLst.Free;
end;

procedure TCT_Worksheet.Clear;
begin
  FAssigneds := [];
  FSheetPr.Clear;
  FDimension.Clear;
  FSheetViews.Clear;
  FSheetFormatPr.Clear;
  FColsXpgList.Clear;
  FSheetData.Clear;
  FSheetCalcPr.Clear;
  FSheetProtection.Clear;
  FProtectedRanges.Clear;
  FScenarios.Clear;
  FAutoFilter.Clear;
  FSortState.Clear;
  FDataConsolidate.Clear;
  FCustomSheetViews.Clear;
  FMergeCells.Clear;
  FPhoneticPr.Clear;
  FConditionalFormattingXpgList.Clear;
  FDataValidations.Clear;
  FHyperlinks.Clear;
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
  FDrawing.Clear;
  FLegacyDrawing.Clear;
  FLegacyDrawingHF.Clear;
  FPicture.Clear;
  FOleObjects.Clear;
  FControls.Clear;
  FWebPublishItems.Clear;
  FTableParts.Clear;
  FExtLst.Clear;
end;

{ TCT_Macrosheet }

function  TCT_Macrosheet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetPr.CheckAssigned);
  Inc(ElemsAssigned,FDimension.CheckAssigned);
  Inc(ElemsAssigned,FSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FSheetFormatPr.CheckAssigned);
  Inc(ElemsAssigned,FColsXpgList.CheckAssigned);
  Inc(ElemsAssigned,FSheetData.CheckAssigned);
  Inc(ElemsAssigned,FSheetProtection.CheckAssigned);
  Inc(ElemsAssigned,FAutoFilter.CheckAssigned);
  Inc(ElemsAssigned,FSortState.CheckAssigned);
  Inc(ElemsAssigned,FDataConsolidate.CheckAssigned);
  Inc(ElemsAssigned,FCustomSheetViews.CheckAssigned);
  Inc(ElemsAssigned,FPhoneticPr.CheckAssigned);
  Inc(ElemsAssigned,FConditionalFormattingXpgList.CheckAssigned);
  Inc(ElemsAssigned,FPrintOptions.CheckAssigned);
  Inc(ElemsAssigned,FPageMargins.CheckAssigned);
  Inc(ElemsAssigned,FPageSetup.CheckAssigned);
  Inc(ElemsAssigned,FHeaderFooter.CheckAssigned);
  Inc(ElemsAssigned,FRowBreaks.CheckAssigned);
  Inc(ElemsAssigned,FColBreaks.CheckAssigned);
  Inc(ElemsAssigned,FCustomProperties.CheckAssigned);
  Inc(ElemsAssigned,FDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawing.CheckAssigned);
  Inc(ElemsAssigned,FLegacyDrawingHF.CheckAssigned);
  Inc(ElemsAssigned,FPicture.CheckAssigned);
  Inc(ElemsAssigned,FOleObjects.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_Macrosheet.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $000002DB: Result := FSheetPr;
    $000003C6: Result := FDimension;
    $00000427: Result := FSheetViews;
    $00000544: Result := FSheetFormatPr;
    $000001B1: Result := FColsXpgList.Add;
    $00000393: Result := FSheetData;
    $00000640: Result := FSheetProtection;
    $0000041F: Result := FAutoFilter;
    $000003C9: Result := FSortState;
    $0000060F: Result := FDataConsolidate;
    $000006A2: Result := FCustomSheetViews;
    $0000041C: Result := FPhoneticPr;
    $000008AF: Result := FConditionalFormattingXpgList.Add;
    $00000519: Result := FPrintOptions;
    $0000046E: Result := FPageMargins;
    $000003AE: Result := FPageSetup;
    $000004D8: Result := FHeaderFooter;
    $000003B0: Result := FRowBreaks;
    $00000396: Result := FColBreaks;
    $000006C8: Result := FCustomProperties;
    $000002EC: Result := FDrawing;
    $00000541: Result := FLegacyDrawing;
    $000005CF: Result := FLegacyDrawingHF;
    $000002FC: Result := FPicture;
    $0000040A: Result := FOleObjects;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Macrosheet.Write(AWriter: TXpgWriteXML);
begin
  if FSheetPr.Assigned then 
  begin
    FSheetPr.WriteAttributes(AWriter);
    if xaElements in FSheetPr.FAssigneds then 
    begin
      AWriter.BeginTag('sheetPr');
      FSheetPr.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetPr');
  end;
  if FDimension.Assigned then 
  begin
    FDimension.WriteAttributes(AWriter);
    AWriter.SimpleTag('dimension');
  end;
  if FSheetViews.Assigned then 
    if xaElements in FSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('sheetViews');
      FSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetViews');
  if FSheetFormatPr.Assigned then 
  begin
    FSheetFormatPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetFormatPr');
  end;
  FColsXpgList.Write(AWriter,'cols');
  if FSheetData.Assigned then 
    if xaElements in FSheetData.FAssigneds then 
    begin
      AWriter.BeginTag('sheetData');
      FSheetData.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetData');
  if FSheetProtection.Assigned then 
  begin
    FSheetProtection.WriteAttributes(AWriter);
    AWriter.SimpleTag('sheetProtection');
  end;
  if FAutoFilter.Assigned then 
  begin
    FAutoFilter.WriteAttributes(AWriter);
    if xaElements in FAutoFilter.FAssigneds then 
    begin
      AWriter.BeginTag('autoFilter');
      FAutoFilter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('autoFilter');
  end;
  if FSortState.Assigned then 
  begin
    FSortState.WriteAttributes(AWriter);
    if xaElements in FSortState.FAssigneds then 
    begin
      AWriter.BeginTag('sortState');
      FSortState.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sortState');
  end;
  if FDataConsolidate.Assigned then 
  begin
    FDataConsolidate.WriteAttributes(AWriter);
    if xaElements in FDataConsolidate.FAssigneds then 
    begin
      AWriter.BeginTag('dataConsolidate');
      FDataConsolidate.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('dataConsolidate');
  end;
  if FCustomSheetViews.Assigned then 
    if xaElements in FCustomSheetViews.FAssigneds then 
    begin
      AWriter.BeginTag('customSheetViews');
      FCustomSheetViews.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customSheetViews');
  if FPhoneticPr.Assigned then 
  begin
    FPhoneticPr.WriteAttributes(AWriter);
    AWriter.SimpleTag('phoneticPr');
  end;
  FConditionalFormattingXpgList.Write(AWriter,'conditionalFormatting');
  if FPrintOptions.Assigned then 
  begin
    FPrintOptions.WriteAttributes(AWriter);
    AWriter.SimpleTag('printOptions');
  end;
  if FPageMargins.Assigned then 
  begin
    FPageMargins.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageMargins');
  end;
  if FPageSetup.Assigned then 
  begin
    FPageSetup.WriteAttributes(AWriter);
    AWriter.SimpleTag('pageSetup');
  end;
  if FHeaderFooter.Assigned then 
  begin
    FHeaderFooter.WriteAttributes(AWriter);
    if xaElements in FHeaderFooter.FAssigneds then 
    begin
      AWriter.BeginTag('headerFooter');
      FHeaderFooter.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('headerFooter');
  end;
  if FRowBreaks.Assigned then 
  begin
    FRowBreaks.WriteAttributes(AWriter);
    if xaElements in FRowBreaks.FAssigneds then 
    begin
      AWriter.BeginTag('rowBreaks');
      FRowBreaks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('rowBreaks');
  end;
  if FColBreaks.Assigned then 
  begin
    FColBreaks.WriteAttributes(AWriter);
    if xaElements in FColBreaks.FAssigneds then 
    begin
      AWriter.BeginTag('colBreaks');
      FColBreaks.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('colBreaks');
  end;
  if FCustomProperties.Assigned then 
    if xaElements in FCustomProperties.FAssigneds then 
    begin
      AWriter.BeginTag('customProperties');
      FCustomProperties.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('customProperties');
  if FDrawing.Assigned then 
  begin
    FDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('drawing');
  end;
  if FLegacyDrawing.Assigned then 
  begin
    FLegacyDrawing.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawing');
  end;
  if FLegacyDrawingHF.Assigned then 
  begin
    FLegacyDrawingHF.WriteAttributes(AWriter);
    AWriter.SimpleTag('legacyDrawingHF');
  end;
  if FPicture.Assigned then 
  begin
    FPicture.WriteAttributes(AWriter);
    AWriter.SimpleTag('picture');
  end;
  if FOleObjects.Assigned then 
    if xaElements in FOleObjects.FAssigneds then 
    begin
      AWriter.BeginTag('oleObjects');
      FOleObjects.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('oleObjects');
  if FExtLst.Assigned then 
    if xaElements in FExtLst.FAssigneds then 
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('extLst');
end;

constructor TCT_Macrosheet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 26;
  FAttributeCount := 0;
  FSheetPr := TCT_SheetPr.Create(FOwner);
  FDimension := TCT_SheetDimension.Create(FOwner);
  FSheetViews := TCT_SheetViews.Create(FOwner);
  FSheetFormatPr := TCT_SheetFormatPr.Create(FOwner);
  FColsXpgList := TCT_ColsXpgList.Create(FOwner);
  FSheetData := TCT_SheetData.Create(FOwner);
  FSheetProtection := TCT_SheetProtection.Create(FOwner);
  FAutoFilter := TCT_AutoFilter.Create(FOwner);
  FSortState := TCT_SortState.Create(FOwner);
  FDataConsolidate := TCT_DataConsolidate.Create(FOwner);
  FCustomSheetViews := TCT_CustomSheetViews.Create(FOwner);
  FPhoneticPr := TCT_PhoneticPr.Create(FOwner);
  FConditionalFormattingXpgList := TCT_ConditionalFormattingXpgList.Create(FOwner);
  FPrintOptions := TCT_PrintOptions.Create(FOwner);
  FPageMargins := TCT_PageMargins.Create(FOwner);
  FPageSetup := TCT_PageSetup.Create(FOwner);
  FHeaderFooter := TCT_HeaderFooter.Create(FOwner);
  FRowBreaks := TCT_PageBreak.Create(FOwner);
  FColBreaks := TCT_PageBreak.Create(FOwner);
  FCustomProperties := TCT_CustomProperties.Create(FOwner);
  FDrawing := TCT_Drawing.Create(FOwner);
  FLegacyDrawing := TCT_LegacyDrawing.Create(FOwner);
  FLegacyDrawingHF := TCT_LegacyDrawing.Create(FOwner);
  FPicture := TCT_SheetBackgroundPicture.Create(FOwner);
  FOleObjects := TCT_OleObjects.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_Macrosheet.Destroy;
begin
  FSheetPr.Free;
  FDimension.Free;
  FSheetViews.Free;
  FSheetFormatPr.Free;
  FColsXpgList.Free;
  FSheetData.Free;
  FSheetProtection.Free;
  FAutoFilter.Free;
  FSortState.Free;
  FDataConsolidate.Free;
  FCustomSheetViews.Free;
  FPhoneticPr.Free;
  FConditionalFormattingXpgList.Free;
  FPrintOptions.Free;
  FPageMargins.Free;
  FPageSetup.Free;
  FHeaderFooter.Free;
  FRowBreaks.Free;
  FColBreaks.Free;
  FCustomProperties.Free;
  FDrawing.Free;
  FLegacyDrawing.Free;
  FLegacyDrawingHF.Free;
  FPicture.Free;
  FOleObjects.Free;
  FExtLst.Free;
end;

procedure TCT_Macrosheet.Clear;
begin
  FAssigneds := [];
  FSheetPr.Clear;
  FDimension.Clear;
  FSheetViews.Clear;
  FSheetFormatPr.Clear;
  FColsXpgList.Clear;
  FSheetData.Clear;
  FSheetProtection.Clear;
  FAutoFilter.Clear;
  FSortState.Clear;
  FDataConsolidate.Clear;
  FCustomSheetViews.Clear;
  FPhoneticPr.Clear;
  FConditionalFormattingXpgList.Clear;
  FPrintOptions.Clear;
  FPageMargins.Clear;
  FPageSetup.Clear;
  FHeaderFooter.Clear;
  FRowBreaks.Clear;
  FColBreaks.Clear;
  FCustomProperties.Clear;
  FDrawing.Clear;
  FLegacyDrawing.Clear;
  FLegacyDrawingHF.Clear;
  FPicture.Clear;
  FOleObjects.Clear;
  FExtLst.Clear;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];

  if FCurrWriteClass.ClassName = FStyleSheet.ClassName then
    Inc(ElemsAssigned,FStyleSheet.CheckAssigned);

  if FCurrWriteClass.ClassName = FSst.ClassName then
    Inc(ElemsAssigned,FSst.CheckAssigned);

  if FCurrWriteClass.ClassName = FWorkbook.ClassName then
    Inc(ElemsAssigned,FWorkbook.CheckAssigned);

  if FCurrWriteClass.ClassName = FComments.ClassName then
    Inc(ElemsAssigned,FComments.CheckAssigned);

  if FCurrWriteClass.ClassName = FTable.ClassName then
    Inc(ElemsAssigned,FTable.CheckAssigned);

  if FCurrWriteClass.ClassName = FChartsheet.ClassName then
    Inc(ElemsAssigned,FChartsheet.CheckAssigned);

  if FCurrWriteClass.ClassName = FDialogsheet.ClassName then
    Inc(ElemsAssigned,FDialogsheet.CheckAssigned);

  if FCurrWriteClass.ClassName = FWorksheet.ClassName then
    Inc(ElemsAssigned,FWorksheet.CheckAssigned);

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
    $0000042A: Result := FStyleSheet;
    $0000015A: Result := FSst;
    $0000036E: Result := FWorkbook;
    $00000366: Result := FComments;
    $00000208: Result := FTable;
    $0000042B: Result := FChartsheet;
    $00000489: Result := FDialogsheet;
    $000003DC: Result := FWorksheet;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

// TODO Select element to write.
procedure T__ROOT__.Write(AWriter: TXpgWriteXML; AClassToWrite: TClass);
begin
  FCurrWriteClass := AClassToWrite;
  AWriter.Attributes := FRootAttributes.Text;  // TODO
  if (AClassToWrite.ClassName = FStyleSheet.ClassName) and FStyleSheet.Assigned then
    if xaElements in FStyleSheet.FAssigneds then
    begin
      AWriter.BeginTag('styleSheet');
      FStyleSheet.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('styleSheet');
  if (AClassToWrite.ClassName = FSst.ClassName) and FSst.Assigned then
  begin
    FSst.WriteAttributes(AWriter);
    if System.Assigned(FSst.OnWriteSi) then
    begin
      AWriter.BeginTag('sst');
      FSst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('sst');
  end;
  if (AClassToWrite.ClassName = FWorkbook.ClassName) and FWorkbook.Assigned then
    if xaElements in FWorkbook.FAssigneds then
    begin
      AWriter.BeginTag('workbook');
      FWorkbook.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('workbook');
  if (AClassToWrite.ClassName = FComments.ClassName) and FComments.Assigned then
    if xaElements in FComments.FAssigneds then
    begin
      AWriter.BeginTag('comments');
      FComments.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('comments');
  if (AClassToWrite.ClassName = FTable.ClassName) and FTable.Assigned then begin
    FTable.WriteAttributes(AWriter);
    if xaElements in FTable.FAssigneds then
    begin
      AWriter.BeginTag('table');
      FTable.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('table');
  end;
  if FChartsheet.Assigned then
    if xaElements in FChartsheet.FAssigneds then
    begin
      AWriter.BeginTag('chartsheet');
      FChartsheet.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('chartsheet');
  if FDialogsheet.Assigned then
    if xaElements in FDialogsheet.FAssigneds then
    begin
      AWriter.BeginTag('dialogsheet');
      FDialogsheet.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('dialogsheet');
  if (AClassToWrite.ClassName = FWorksheet.ClassName) and FWorksheet.Assigned then
    if xaElements in FWorksheet.FAssigneds then
    begin
      AWriter.BeginTag('worksheet');
      FWorksheet.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('worksheet');
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 7;
  FAttributeCount := 0;
  FStyleSheet := TCT_Stylesheet.Create(FOwner);
  FSst := TCT_Sst.Create(FOwner);
  FWorkbook := TCT_Workbook.Create(FOwner);
  FComments := TCT_Comments.Create(FOwner);
  FTable := TCT_Table.Create(FOwner);
  FChartsheet := TCT_Chartsheet.Create(FOwner);
  FDialogsheet := TCT_Dialogsheet.Create(FOwner);
  FWorksheet := TCT_Worksheet.Create(FOwner);
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  FStyleSheet.Free;
  FSst.Free;
  FWorkbook.Free;
  FComments.Free;
  FTable.Free;
  FChartsheet.Free;
  FDialogsheet.Free;
  FWorksheet.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  FStyleSheet.Clear;
  FSst.Clear;
  FWorkbook.Clear;
  FComments.Clear;
  FTable.Clear;
  FChartsheet.Clear;
  FDialogsheet.Clear;
  FWorksheet.Clear;
end;

{ TCT_XStringElement }

function  TCT_XStringElement.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FV <> '' then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
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

{ TXPGDocXLSX }

function  TXPGDocXLSX.GetStyleSheet: TCT_Stylesheet;
begin
  Result := FRoot.StyleSheet;
end;

function TXPGDocXLSX.GetTable: TCT_Table;
begin
  Result := FRoot.Table;
end;

function  TXPGDocXLSX.GetSst: TCT_Sst;
begin
  Result := FRoot.Sst;
end;

function  TXPGDocXLSX.GetWorkbook: TCT_Workbook;
begin
  Result := FRoot.Workbook;
end;

function  TXPGDocXLSX.GetComments: TCT_Comments;
begin
  Result := FRoot.Comments;
end;

function  TXPGDocXLSX.GetChartsheet: TCT_Chartsheet;
begin
  Result := FRoot.Chartsheet;
end;

function  TXPGDocXLSX.GetDialogsheet: TCT_Dialogsheet;
begin
  Result := FRoot.Dialogsheet;
end;

function  TXPGDocXLSX.GetWorksheet: TCT_Worksheet;
begin
  Result := FRoot.Worksheet;
end;

constructor TXPGDocXLSX.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocXLSX.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGDocXLSX.LoadFromFile(const AFilename: AxUCString);
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

procedure TXPGDocXLSX.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocXLSX.SaveToFile(const AFilename: AxUCString);
begin
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter,Nil);
end;

procedure TXPGDocXLSX.SaveToStream(AStream: TStream; AClassToWrite: TClass);
begin
  FWriter.SaveToStream(AStream);
  FRoot.FCurrWriteClass := AClassToWrite; // TODO
  FRoot.CheckAssigned;
  FRoot.Write(FWriter,AClassToWrite);
end;

{ TCT_TableStyleInfo }

function  TCT_TableStyleInfo.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then
    Inc(AttrsAssigned);
  if FShowFirstColumn <> False then
    Inc(AttrsAssigned);
  if FShowLastColumn <> False then
    Inc(AttrsAssigned);
  if FShowRowStripes <> False then
    Inc(AttrsAssigned);
  if FShowColumnStripes <> False then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
end;

procedure TCT_TableStyleInfo.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TableStyleInfo.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then
    AWriter.AddAttribute('name',FName);
  if FShowFirstColumn <> False then
    AWriter.AddAttribute('showFirstColumn',XmlBoolToStr(FShowFirstColumn));
  if FShowLastColumn <> False then
    AWriter.AddAttribute('showLastColumn',XmlBoolToStr(FShowLastColumn));
  if FShowRowStripes <> False then
    AWriter.AddAttribute('showRowStripes',XmlBoolToStr(FShowRowStripes));
  if FShowColumnStripes <> False then
    AWriter.AddAttribute('showColumnStripes',XmlBoolToStr(FShowColumnStripes));
end;

procedure TCT_TableStyleInfo.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001A1: FName := AAttributes.Values[i];
      $00000637: FShowFirstColumn := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005C3: FShowLastColumn := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005E3: FShowRowStripes := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000719: FShowColumnStripes := XmlStrToBoolDef(AAttributes.Values[i],False);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_TableStyleInfo.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 5;
end;

destructor TCT_TableStyleInfo.Destroy;
begin
end;

procedure TCT_TableStyleInfo.Clear;
begin
  FAssigneds := [];
  FName := '';
  FShowFirstColumn := False;
  FShowLastColumn := False;
  FShowRowStripes := False;
  FShowColumnStripes := False;
end;

{ TCT_XmlColumnPr }

function  TCT_XmlColumnPr.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FMapId <> 0 then
    Inc(AttrsAssigned);
  if FXpath <> '' then
    Inc(AttrsAssigned);
  if FDenormalized <> False then
    Inc(AttrsAssigned);
  if Integer(FXmlDataType) <> XPG_UNKNOWN_ENUM then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_XmlColumnPr.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'extLst' then
    Result := FExtLst
  else
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_XmlColumnPr.Write(AWriter: TXpgWriteXML);
begin
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('extLst');
end;

procedure TCT_XmlColumnPr.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('mapId',XmlIntToStr(FMapId));
  AWriter.AddAttribute('xpath',FXpath);
  if FDenormalized <> False then
    AWriter.AddAttribute('denormalized',XmlBoolToStr(FDenormalized));
  if Integer(FXmlDataType) <> XPG_UNKNOWN_ENUM then
    AWriter.AddAttribute('xmlDataType',StrTST_XmlDataType[Integer(FXmlDataType)]);
end;

procedure TCT_XmlColumnPr.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000001EB: FMapId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000225: FXpath := AAttributes.Values[i];
      $000004FE: FDenormalized := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000046D: FXmlDataType := TST_XmlDataType(StrToEnum('stxdt' + AAttributes.Values[i]));
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_XmlColumnPr.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FDenormalized := False;
  FXmlDataType := TST_XmlDataType(XPG_UNKNOWN_ENUM);
end;

destructor TCT_XmlColumnPr.Destroy;
begin
  FExtLst.Free;
end;

procedure TCT_XmlColumnPr.Clear;
begin
  FAssigneds := [];
  FExtLst.Clear;
  FMapId := 0;
  FXpath := '';
  FDenormalized := False;
  FXmlDataType := TST_XmlDataType(XPG_UNKNOWN_ENUM);
end;

{ TCT_TableFormula }

function  TCT_TableFormula.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FArray <> False then
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then
    FAssigneds := [xaAttributes];
  if FContent <> '' then
  begin
    FAssigneds := FAssigneds + [xaContent];
    Inc(Result);
  end;
end;

function  TCT_TableFormula.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_TableFormula.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_TableFormula.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FArray <> False then
    AWriter.AddAttribute('array',XmlBoolToStr(FArray));
end;

procedure TCT_TableFormula.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'array' then
    FArray := XmlStrToBoolDef(AAttributes.Values[0],False)
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TableFormula.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
  FArray := False;
end;

destructor TCT_TableFormula.Destroy;
begin
end;

procedure TCT_TableFormula.Clear;
begin
  FAssigneds := [];
  FArray := False;
end;

{ TCT_TableColumn }

function  TCT_TableColumn.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FId <> 0 then
    Inc(AttrsAssigned);
  if FUniqueName <> '' then
    Inc(AttrsAssigned);
  if FName <> '' then
    Inc(AttrsAssigned);
  if FTotalsRowFunction <> sttrfNone then
    Inc(AttrsAssigned);
  if FTotalsRowLabel <> '' then
    Inc(AttrsAssigned);
  if FQueryTableFieldId <> 0 then
    Inc(AttrsAssigned);
  if FHeaderRowDxfId <> '' then
    Inc(AttrsAssigned);
  if FDataDxfId <> '' then
    Inc(AttrsAssigned);
  if FTotalsRowDxfId <> '' then
    Inc(AttrsAssigned);
  if FHeaderRowCellStyle <> '' then
    Inc(AttrsAssigned);
  if FDataCellStyle <> '' then
    Inc(AttrsAssigned);
  if FTotalsRowCellStyle <> '' then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCalculatedColumnFormula.CheckAssigned);
  Inc(ElemsAssigned,FTotalsRowFormula.CheckAssigned);
  Inc(ElemsAssigned,FXmlColumnPr.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TableColumn.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $00000956: begin
      Result := FCalculatedColumnFormula;
      if AReader.HasText then
        TCT_TableFormula(Result).Content := AReader.Text;
    end;
    $000006A5: begin
      Result := FTotalsRowFormula;
      if AReader.HasText then
        TCT_TableFormula(Result).Content := AReader.Text;
    end;
    $00000481: Result := FXmlColumnPr;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_TableColumn.Write(AWriter: TXpgWriteXML);
begin
  if FCalculatedColumnFormula.Assigned then
  begin
    AWriter.Text := FCalculatedColumnFormula.Content;
    FCalculatedColumnFormula.WriteAttributes(AWriter);
    AWriter.SimpleTag('calculatedColumnFormula');
  end;
  if FTotalsRowFormula.Assigned then
  begin
    AWriter.Text := FTotalsRowFormula.Content;
    FTotalsRowFormula.WriteAttributes(AWriter);
    AWriter.SimpleTag('totalsRowFormula');
  end;
  if FXmlColumnPr.Assigned then
  begin
    FXmlColumnPr.WriteAttributes(AWriter);
    if xaElements in FXmlColumnPr.FAssigneds then
    begin
      AWriter.BeginTag('xmlColumnPr');
      FXmlColumnPr.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('xmlColumnPr');
  end;
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('extLst');
end;

procedure TCT_TableColumn.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  if FUniqueName <> '' then
    AWriter.AddAttribute('uniqueName',FUniqueName);
  AWriter.AddAttribute('name',FName);
  if FTotalsRowFunction <> sttrfNone then
    AWriter.AddAttribute('totalsRowFunction',StrTST_TotalsRowFunction[Integer(FTotalsRowFunction)]);
  if FTotalsRowLabel <> '' then
    AWriter.AddAttribute('totalsRowLabel',FTotalsRowLabel);
  if FQueryTableFieldId <> 0 then
    AWriter.AddAttribute('queryTableFieldId',XmlIntToStr(FQueryTableFieldId));
  if FHeaderRowDxfId <> '' then
    AWriter.AddAttribute('headerRowDxfId',FHeaderRowDxfId);
  if FDataDxfId <> '' then
    AWriter.AddAttribute('dataDxfId',FDataDxfId);
  if FTotalsRowDxfId <> '' then
    AWriter.AddAttribute('totalsRowDxfId',FTotalsRowDxfId);
  if FHeaderRowCellStyle <> '' then
    AWriter.AddAttribute('headerRowCellStyle',FHeaderRowCellStyle);
  if FDataCellStyle <> '' then
    AWriter.AddAttribute('dataCellStyle',FDataCellStyle);
  if FTotalsRowCellStyle <> '' then
    AWriter.AddAttribute('totalsRowCellStyle',FTotalsRowCellStyle);
end;

procedure TCT_TableColumn.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000418: FUniqueName := AAttributes.Values[i];
      $000001A1: FName := AAttributes.Values[i];
      $00000715: FTotalsRowFunction := TST_TotalsRowFunction(StrToEnum('sttrf' + AAttributes.Values[i]));
      $000005AF: FTotalsRowLabel := AAttributes.Values[i];
      $000006AF: FQueryTableFieldId := XmlStrToIntDef(AAttributes.Values[i],0);
      $00000570: FHeaderRowDxfId := AAttributes.Values[i];
      $00000369: FDataDxfId := AAttributes.Values[i];
      $0000059E: FTotalsRowDxfId := AAttributes.Values[i];
      $00000732: FHeaderRowCellStyle := AAttributes.Values[i];
      $0000052B: FDataCellStyle := AAttributes.Values[i];
      $00000760: FTotalsRowCellStyle := AAttributes.Values[i];
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_TableColumn.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 12;
  FCalculatedColumnFormula := TCT_TableFormula.Create(FOwner);
  FTotalsRowFormula := TCT_TableFormula.Create(FOwner);
  FXmlColumnPr := TCT_XmlColumnPr.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FTotalsRowFunction := sttrfNone;
end;

destructor TCT_TableColumn.Destroy;
begin
  FCalculatedColumnFormula.Free;
  FTotalsRowFormula.Free;
  FXmlColumnPr.Free;
  FExtLst.Free;
end;

procedure TCT_TableColumn.Clear;
begin
  FAssigneds := [];
  FCalculatedColumnFormula.Clear;
  FTotalsRowFormula.Clear;
  FXmlColumnPr.Clear;
  FExtLst.Clear;
  FId := 0;
  FUniqueName := '';
  FName := '';
  FTotalsRowFunction := sttrfNone;
  FTotalsRowLabel := '';
  FQueryTableFieldId := 0;
  FHeaderRowDxfId := '';
  FDataDxfId := '';
  FTotalsRowDxfId := '';
  FHeaderRowCellStyle := '';
  FDataCellStyle := '';
  FTotalsRowCellStyle := '';
end;

{ TCT_TableColumnXpgList }

function  TCT_TableColumnXpgList.GetItems(Index: integer): TCT_TableColumn;
begin
  Result := TCT_TableColumn(inherited Items[Index]);
end;

function  TCT_TableColumnXpgList.Add: TCT_TableColumn;
begin
  Result := TCT_TableColumn.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_TableColumnXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_TableColumnXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_TableColumns }

function  TCT_TableColumns.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FCount <> 0 then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FTableColumnXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_TableColumns.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'tableColumn' then
    Result := FTableColumnXpgList.Add
  else
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_TableColumns.Write(AWriter: TXpgWriteXML);
begin
  FTableColumnXpgList.Write(AWriter,'tableColumn');
end;

procedure TCT_TableColumns.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FCount <> 0 then
    AWriter.AddAttribute('count',XmlIntToStr(FCount));
end;

procedure TCT_TableColumns.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'count' then
    FCount := XmlStrToIntDef(AAttributes.Values[0],0)
  else
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_TableColumns.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FTableColumnXpgList := TCT_TableColumnXpgList.Create(FOwner);
end;

destructor TCT_TableColumns.Destroy;
begin
  FTableColumnXpgList.Free;
end;

procedure TCT_TableColumns.Clear;
begin
  FAssigneds := [];
  FTableColumnXpgList.Clear;
  FCount := 0;
end;

{ TCT_Table }

function  TCT_Table.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FId <> 0 then
    Inc(AttrsAssigned);
  if FName <> '' then
    Inc(AttrsAssigned);
  if FDisplayName <> '' then
    Inc(AttrsAssigned);
  if FComment <> '' then
    Inc(AttrsAssigned);
  if FRef <> '' then
    Inc(AttrsAssigned);
  if FTableType <> stttWorksheet then
    Inc(AttrsAssigned);
  if FHeaderRowCount <> 1 then
    Inc(AttrsAssigned);
  if FInsertRow <> False then
    Inc(AttrsAssigned);
  if FInsertRowShift <> False then
    Inc(AttrsAssigned);
  if FTotalsRowCount <> 0 then
    Inc(AttrsAssigned);
  if FTotalsRowShown <> True then
    Inc(AttrsAssigned);
  if FPublished <> False then
    Inc(AttrsAssigned);
  if FHeaderRowDxfId <> '' then
    Inc(AttrsAssigned);
  if FDataDxfId <> '' then
    Inc(AttrsAssigned);
  if FTotalsRowDxfId <> '' then
    Inc(AttrsAssigned);
  if FHeaderRowBorderDxfId <> '' then
    Inc(AttrsAssigned);
  if FTableBorderDxfId <> '' then
    Inc(AttrsAssigned);
  if FTotalsRowBorderDxfId <> '' then
    Inc(AttrsAssigned);
  if FHeaderRowCellStyle <> '' then
    Inc(AttrsAssigned);
  if FDataCellStyle <> '' then
    Inc(AttrsAssigned);
  if FTotalsRowCellStyle <> '' then
    Inc(AttrsAssigned);
  if FConnectionId <> 0 then
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FAutoFilter.CheckAssigned);
  Inc(ElemsAssigned,FSortState.CheckAssigned);
  Inc(ElemsAssigned,FTableColumns.CheckAssigned);
  Inc(ElemsAssigned,FTableStyleInfo.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_Table.HandleElement(AReader: TXpgReadXML): TXPGBase;
begin
  Result := Self;
  case AReader.QNameHashA of
    $0000041F: Result := FAutoFilter;
    $000003C9: Result := FSortState;
    $000004E9: Result := FTableColumns;
    $000005A5: Result := FTableStyleInfo;
    $00000284: Result := FExtLst;
    else
      FOwner.Errors.Error(xemUnknownElement,AReader.QName);
  end;
  if Result <> Self then
    Result.FAssigneds := [xaRead];
end;

procedure TCT_Table.Write(AWriter: TXpgWriteXML);
begin
  if FAutoFilter.Assigned then
  begin
    FAutoFilter.WriteAttributes(AWriter);
    if xaElements in FAutoFilter.FAssigneds then
    begin
      AWriter.BeginTag('autoFilter');
      FAutoFilter.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('autoFilter');
  end;
  if FSortState.Assigned then
  begin
    FSortState.WriteAttributes(AWriter);
    if xaElements in FSortState.FAssigneds then
    begin
      AWriter.BeginTag('sortState');
      FSortState.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('sortState');
  end;
  if FTableColumns.Assigned then
  begin
    FTableColumns.WriteAttributes(AWriter);
    if xaElements in FTableColumns.FAssigneds then
    begin
      AWriter.BeginTag('tableColumns');
      FTableColumns.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('tableColumns');
  end;
  if FTableStyleInfo.Assigned then
  begin
    FTableStyleInfo.WriteAttributes(AWriter);
    AWriter.SimpleTag('tableStyleInfo');
  end;
  if FExtLst.Assigned then
    if xaElements in FExtLst.FAssigneds then
    begin
      AWriter.BeginTag('extLst');
      FExtLst.Write(AWriter);
      AWriter.EndTag;
    end
    else
      AWriter.SimpleTag('extLst');
end;

procedure TCT_Table.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('id',XmlIntToStr(FId));
  if FName <> '' then
    AWriter.AddAttribute('name',FName);
  AWriter.AddAttribute('displayName',FDisplayName);
  if FComment <> '' then
    AWriter.AddAttribute('comment',FComment);
  AWriter.AddAttribute('ref',FRef);
  if FTableType <> stttWorksheet then
    AWriter.AddAttribute('tableType',StrTST_TableType[Integer(FTableType)]);
  if FHeaderRowCount <> 1 then
    AWriter.AddAttribute('headerRowCount',XmlIntToStr(FHeaderRowCount));
  if FInsertRow <> False then
    AWriter.AddAttribute('insertRow',XmlBoolToStr(FInsertRow));
  if FInsertRowShift <> False then
    AWriter.AddAttribute('insertRowShift',XmlBoolToStr(FInsertRowShift));
  if FTotalsRowCount <> 0 then
    AWriter.AddAttribute('totalsRowCount',XmlIntToStr(FTotalsRowCount));
  if FTotalsRowShown <> True then
    AWriter.AddAttribute('totalsRowShown',XmlBoolToStr(FTotalsRowShown));
  if FPublished <> False then
    AWriter.AddAttribute('published',XmlBoolToStr(FPublished));
  if FHeaderRowDxfId <> '' then
    AWriter.AddAttribute('headerRowDxfId',FHeaderRowDxfId);
  if FDataDxfId <> '' then
    AWriter.AddAttribute('dataDxfId',FDataDxfId);
  if FTotalsRowDxfId <> '' then
    AWriter.AddAttribute('totalsRowDxfId',FTotalsRowDxfId);
  if FHeaderRowBorderDxfId <> '' then
    AWriter.AddAttribute('headerRowBorderDxfId',FHeaderRowBorderDxfId);
  if FTableBorderDxfId <> '' then
    AWriter.AddAttribute('tableBorderDxfId',FTableBorderDxfId);
  if FTotalsRowBorderDxfId <> '' then
    AWriter.AddAttribute('totalsRowBorderDxfId',FTotalsRowBorderDxfId);
  if FHeaderRowCellStyle <> '' then
    AWriter.AddAttribute('headerRowCellStyle',FHeaderRowCellStyle);
  if FDataCellStyle <> '' then
    AWriter.AddAttribute('dataCellStyle',FDataCellStyle);
  if FTotalsRowCellStyle <> '' then
    AWriter.AddAttribute('totalsRowCellStyle',FTotalsRowCellStyle);
  if FConnectionId <> 0 then
    AWriter.AddAttribute('connectionId',XmlIntToStr(FConnectionId));
end;

procedure TCT_Table.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do
  begin
    case AAttributes.HashA[i] of
      $000000CD: FId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001A1: FName := AAttributes.Values[i];
      $00000477: FDisplayName := AAttributes.Values[i];
      $000002F3: FComment := AAttributes.Values[i];
      $0000013D: FRef := AAttributes.Values[i];
      $000003AA: FTableType := TST_TableType(StrToEnum('sttt' + AAttributes.Values[i]));
      $000005AA: FHeaderRowCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000003CD: FInsertRow := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005CB: FInsertRowShift := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000005D8: FTotalsRowCount := XmlStrToIntDef(AAttributes.Values[i],0);
      $000005DE: FTotalsRowShown := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003C0: FPublished := XmlStrToBoolDef(AAttributes.Values[i],False);
      $00000570: FHeaderRowDxfId := AAttributes.Values[i];
      $00000369: FDataDxfId := AAttributes.Values[i];
      $0000059E: FTotalsRowDxfId := AAttributes.Values[i];
      $000007CE: FHeaderRowBorderDxfId := AAttributes.Values[i];
      $00000635: FTableBorderDxfId := AAttributes.Values[i];
      $000007FC: FTotalsRowBorderDxfId := AAttributes.Values[i];
      $00000732: FHeaderRowCellStyle := AAttributes.Values[i];
      $0000052B: FDataCellStyle := AAttributes.Values[i];
      $00000760: FTotalsRowCellStyle := AAttributes.Values[i];
      $000004DD: FConnectionId := XmlStrToIntDef(AAttributes.Values[i],0);
      else
        FOwner.Errors.Error(xemUnknownAttribute,AAttributes[i]);
    end;
  end
end;

constructor TCT_Table.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 5;
  FAttributeCount := 22;
  FAutoFilter := TCT_AutoFilter.Create(FOwner);
  FSortState := TCT_SortState.Create(FOwner);
  FTableColumns := TCT_TableColumns.Create(FOwner);
  FTableStyleInfo := TCT_TableStyleInfo.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
  FTableType := stttWorksheet;
  FHeaderRowCount := 1;
  FInsertRow := False;
  FInsertRowShift := False;
  FTotalsRowCount := 0;
  FTotalsRowShown := True;
  FPublished := False;
end;

destructor TCT_Table.Destroy;
begin
  FAutoFilter.Free;
  FSortState.Free;
  FTableColumns.Free;
  FTableStyleInfo.Free;
  FExtLst.Free;
end;

procedure TCT_Table.Clear;
begin
  FAssigneds := [];
  FAutoFilter.Clear;
  FSortState.Clear;
  FTableColumns.Clear;
  FTableStyleInfo.Clear;
  FExtLst.Clear;
  FId := 0;
  FName := '';
  FDisplayName := '';
  FComment := '';
  FRef := '';
  FTableType := stttWorksheet;
  FHeaderRowCount := 1;
  FInsertRow := False;
  FInsertRowShift := False;
  FTotalsRowCount := 0;
  FTotalsRowShown := True;
  FPublished := False;
  FHeaderRowDxfId := '';
  FDataDxfId := '';
  FTotalsRowDxfId := '';
  FHeaderRowBorderDxfId := '';
  FTableBorderDxfId := '';
  FTotalsRowBorderDxfId := '';
  FHeaderRowCellStyle := '';
  FDataCellStyle := '';
  FTotalsRowCellStyle := '';
  FConnectionId := 0;
end;

initialization
{$ifdef DELPHI_5}
  FEnums := TStringList.Create;
{$else}
  FEnums := THashedStringList.Create;
{$endif}
  TXPGBase.AddEnums;
  TXPGBase.AddEnums;

finalization
  FEnums.Free;

end.
