unit Xc12DataWorkbook5;

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

uses Classes, SysUtils, Contnrs, IniFiles,
     xpgPSimpleDOM, xpgParserPivot,
     Xc12Utils5, Xc12Common5, Xc12FileData5,
     XLSUtils5, XLSFormulaTypes5, XLSClassFactory5;

type TXc12UpdateLinks = (x12ulUserSet,x12ulNever,x12ulAlways);
type TXc12CalcMode = (cmManual,cmAutomatic,cmAutoExTables);
type TXc12RefMode = (x12rmA1,x12rmR1C1);
type TXc12Comments_ = (x12cCommNone,x12cCommIndicator,x12cCommIndAndComment);
type TXc12Objects = (x12oAll,x12oPlaceholders,x12oNone);
type TXc12SmartTagShow = (x12smtAll,x12smtNone,x12smtNoIndicator);
type TXc12TargetScreenSize =(x12tss544x376,x12tss640x480,x12tss720x512,x12tss800x600,x12tss1024x768,x12tss1152x882,x12tss1152x900,x12tss1280x1024,x12tss1600x1200,x12tss1800x1440,x12tss1920x1200);

type TXc12FileVersion = class(TObject)
private
     procedure SetAppName(const Value: AxUCString);
     procedure SetCodeName(const Value: AxUCString);
     procedure SetLastEdited(const Value: AxUCString);
     procedure SetLowestEdited(const Value: AxUCString);
     procedure SetRupBuild(const Value: AxUCString);
protected
     FAppName: AxUCString;
     FLastEdited: AxUCString;
     FLowestEdited: AxUCString;
     FRupBuild: AxUCString;
     FCodeName: AxUCString;
public
     constructor Create;

     procedure Clear;

     property AppName: AxUCString read FAppName write SetAppName;
     property LastEdited: AxUCString read FLastEdited write SetLastEdited;
     property LowestEdited: AxUCString read FLowestEdited write SetLowestEdited;
     property RupBuild: AxUCString read FRupBuild write SetRupBuild;
     property CodeName: AxUCString read FCodeName write SetCodeName;
     end;

type TXc12FileSharing = class(TObject)
protected
     FReadOnlyReccomended: boolean;
     FUsername: AxUCString;
     FReservationPassword: integer;
public
     constructor Create;

     procedure Clear;

     property ReadOnlyReccomended: boolean read FReadOnlyReccomended write FReadOnlyReccomended;
     property Username: AxUCString read FUsername write FUsername;
     property ReservationPassword: integer read FReservationPassword write FReservationPassword;
     end;

type TXc12WorkbookPr = class(TObject)
protected
     FDate1904: boolean;
     FShowObjects: TXc12Objects;
     FShowBorderUnselectedTables: boolean;
     FFilterPrivacy: boolean;
     FPromptedSolutions: boolean;
     FShowInkAnnotation: boolean;
     FBackupFile: boolean;
     FSaveExternalLinkValues: boolean;
     FUpdateLinks: TXc12UpdateLinks;
     FCodeName: AxUCString;
     FHidePivotFieldList: boolean;
     FShowPivotChartFilter: boolean;
     FAllowRefreshQuery: boolean;
     FPublishItems: boolean;
     FCheckCompatibility: boolean;
     FAutoCompressPictures: boolean;
     FRefreshAllConnections: boolean;
     FDefaultThemeVersion: integer;
public
     constructor Create;

     procedure Clear;

     property Date1904: boolean read FDate1904 write FDate1904;
     property ShowObjects: TXc12Objects read FShowObjects write FShowObjects;
     property ShowBorderUnselectedTables: boolean read FShowBorderUnselectedTables write FShowBorderUnselectedTables;
     property FilterPrivacy: boolean read FFilterPrivacy write FFilterPrivacy;
     property PromptedSolutions: boolean read FPromptedSolutions write FPromptedSolutions;
     property ShowInkAnnotation: boolean read FShowInkAnnotation write FShowInkAnnotation;
     property BackupFile: boolean read FBackupFile write FBackupFile;
     property SaveExternalLinkValues: boolean read FSaveExternalLinkValues write FSaveExternalLinkValues;
     property UpdateLinks: TXc12UpdateLinks read FUpdateLinks write FUpdateLinks;
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

type TXc12WorkbookProtection = class(TObject)
protected
     FWorkbookPassword: integer;
     FRevisionPassword: integer;
     FLockStructure: boolean;
     FLockWindows: boolean;
     FLockRevision: boolean;
public
     constructor Create;

     procedure Clear;

     property WorkbookPassword: integer read FWorkbookPassword write FWorkbookPassword;
     property RevisionPassword: integer read FRevisionPassword write FRevisionPassword;
     property LockStructure: boolean read FLockStructure write FLockStructure;
     property LockWindows: boolean read FLockWindows write FLockWindows;
     property LockRevision: boolean read FLockRevision write FLockRevision;
     end;

type TXc12BookView = class(TObject)
protected
     FVisibility: TXc12Visibility;
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
public
     constructor Create;

     procedure Clear; // x

     property Visibility: TXc12Visibility read FVisibility write FVisibility;
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
     end;

type TXc12BookViews = class(TObjectList)
private
     function GetItems(Index: integer): TXc12BookView;
protected
public
     constructor Create;

     procedure Clear(AAddDefault: boolean); reintroduce;
     function Add: TXc12BookView;
     function Last: TXc12BookView;

     property Items[Index: integer]: TXc12BookView read GetItems; default;
     end;

type TXc12Sheet = class(TObject)
protected
     FName: AxUCString;
     FSheetId: integer;
     FState: TXc12Visibility;
     FRId: TXc12RId;
public
     constructor Create;

     procedure Clear;

     property Name: AxUCString read FName write FName;
     property SheetId: integer read FSheetId write FSheetId;
     property State: TXc12Visibility read FState write FState;
     property RId: TXc12RId read FRId write FRId;
     end;

type TXc12Sheets = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Sheet;
protected
public
     constructor Create;

     function Add: TXc12Sheet;

     property Items[Index: integer]: TXc12Sheet read GetItems; default;
     end;

type TXc12FunctionGroups = class(TObject)
private
     function GetItems(Index: integer): AxUCString;
protected
     FItems: TStringList;
     FBuiltInGroupCount: integer;
public
     constructor Create;
     destructor Destroy; override;

     procedure Add(AName: AxUCString);
     function  Count: integer;
     procedure Clear;

     property BuiltInGroupCount: integer read FBuiltInGroupCount write FBuiltInGroupCount;
     property Items[Index: integer]: AxUCString read GetItems; default;
     end;

type TXc12DefinedName = class(TObject)
private
     function GetCol1: integer;
     function GetCol2: integer;
     function GetRow1: integer;
     function GetRow2: integer;
     function GetSheetIndex: integer;
protected
     FName             : AxUCString;
     FComment          : AxUCString;
     FCustomMenu       : AxUCString;
     FDescription      : AxUCString;
     FHelp             : AxUCString;
     FStatusBar        : AxUCString;
     FLocalSheetId     : integer;
     FHidden           : boolean;
     FFunction         : boolean;
     FVbProcedure      : boolean;
     FXlm              : boolean;
     FFunctionGroupId  : integer;
     FShortcutKey      : AxUCString;
     FPublishToServer  : boolean;
     FWorkbookParameter: boolean;

     FBuiltIn          : TXc12BuiltInName;
     FContent          : AxUCString;
     FSimpleName       : TXc12SimpleNameType;
     FArea             : TXLS3dCellArea;
     FPtgs             : PXLSPtgs;
     FPtgsSz           : integer;

     // Name is deleted and shall not be written to file.
     FDeleted          : boolean;

     // The name is a placeholder for an unsupported Excel 97 built in name.
     FUnused97         : byte;

     procedure SetName(const Value: AxUCString);
     function  GetContent: AxUCString; virtual;
     procedure SetContent(const Value: AxUCString);
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure Assign(AName: TXc12DefinedName);

     property Name: AxUCString read FName write SetName;
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

     property Ptgs: PXLSPtgs read FPtgs write FPtgs;
     property PtgsSz: integer read FPtgsSz write FPtgsSz;

     property BuiltIn: TXc12BuiltInName read FBuiltIn write FBuiltIn;
     property Content: AxUCString read GetContent write SetContent;
     property OrigContent: AxUCString read FContent;

     property SimpleName: TXc12SimpleNameType read FSimpleName;
     //* The area of the name if it is a simple name. Please note that columns
     //* and rows may contain flags for absolute columns and rows.
     //* Use ROW_ABSFLAG and COL_ABSFLAG to get the value.
     //* Example: Col := SimpleArea.Col1 and not COL_ABSFLAG
     property SimpleArea: TXLS3dCellArea read FArea;
     property SheetIndex: integer read GetSheetIndex;
     property Col1: integer read GetCol1;
     property Row1: integer read GetRow1;
     property Col2: integer read GetCol2;
     property Row2: integer read GetRow2;

     property Deleted: boolean read FDeleted write FDeleted;

     property Unused97: byte read FUnused97 write FUnused97;
     end;

type TXc12DefinedNames = class(TObjectList)
private
{$ifdef DELPHI_5}
     FHash: TStringList;
{$else}
     FHash: THashedStringList;
{$endif}

     function GetItems(Index: integer): TXc12DefinedName;
protected
     FClassFactory: TXLSClassFactory;
     FHashValid: boolean;

     function  AddOrGetFree: TXc12DefinedName;
     procedure DoDelete(AIndex: integer);
     procedure BuildHash;
public
     constructor Create(AClassFactory: TXLSClassFactory);
     destructor Destroy; override;

     procedure Assign(ANames: TXc12DefinedNames);

     function  Add(const ABuiltIn: TXc12BuiltInName; const ALocalSheetId: integer): TXc12DefinedName; overload;
     function  Add(const AName: AxUCString; const ALocalSheetId: integer): TXc12DefinedName; overload;
     function  FindId(const AName: AxUCString; const ASheetId: integer): integer; overload;
     function  FindBuiltIn(const AId: TXc12BuiltInName; const ASheetId: integer): TXc12DefinedName;
     function  FindIdInSheet(const AName: AxUCString; const ASheetId: integer): integer;
     procedure Delete(const AName: AxUCString); overload;
     procedure Delete(const AName: AxUCString; const ALocalSheetId: integer); overload;
     procedure Delete(const AId: TXc12BuiltInName; const ALocalSheetId: integer); overload;

     property Items[Index: integer]: TXc12DefinedName read GetItems; default;
     end;

type TXc12CalcPr = class(TObject)
protected
     FCalcId: integer;
     FCalcMode: TXc12CalcMode;
     FFullCalcOnLoad: boolean;
     FRefMode: TXc12RefMode;
     FIterate: boolean;
     FIterateCount: integer;
     FIterateDelta: double;
     FFullPrecision: boolean;
     FCalcCompleted: boolean;
     FCalcOnSave: boolean;
     FConcurrentCalc: boolean;
     FConcurrentManualCount: integer;
     FForceFullCalc: boolean;
public
     constructor Create;

     procedure Clear;

     property CalcId: integer read FCalcId write FCalcId;
     property CalcMode: TXc12CalcMode read FCalcMode write FCalcMode;
     property FullCalcOnLoad: boolean read FFullCalcOnLoad write FFullCalcOnLoad;
     property RefMode: TXc12RefMode read FRefMode write FRefMode;
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

type TXc12CustomWorkbookView = class(TObject)
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
     FShowComments: TXc12Comments_;
     FShowObjects: TXc12Objects;
public
     constructor Create;

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
     property ShowComments: TXc12Comments_ read FShowComments write FShowComments;
     property ShowObjects: TXc12Objects read FShowObjects write FShowObjects;
     end;

type TXc12CustomWorkbookViews = class(TObjectList)
private
     function GetItems(Index: integer): TXc12CustomWorkbookView;
protected
public
     function Add: TXc12CustomWorkbookView;

     property Items[Index: integer]: TXc12CustomWorkbookView read GetItems; default;
     end;

type TXc12SmartTagPr = class(TObject)
protected
     FEmbed: boolean;
     FShow: TXc12SmartTagShow;
public
     constructor Create;

     procedure Clear;

     property Embed: boolean read FEmbed write FEmbed;
     property Show: TXc12SmartTagShow read FShow write FShow;
     end;

type TXc12SmartTagType = class(TObject)
protected
     FNamespaceUri: AxUCString;
     FName: AxUCString;
     FUrl: AxUCString;
public
     constructor Create;

     procedure Clear;

     property NamespaceUri: AxUCString read FNamespaceUri write FNamespaceUri;
     property Name: AxUCString read FName write FName;
     property Url: AxUCString read FUrl write FUrl;
     end;

type TXc12SmartTagTypes = class(TObjectList)
private
     function GetItems(Index: integer): TXc12SmartTagType;
protected
public
     function Add: TXc12SmartTagType;

     property Items[Index: integer]: TXc12SmartTagType read GetItems; default;
     end;

type TXc12WebPublishing = class(TObject)
protected
     FCss: boolean;
     FThicket: boolean;
     FLongFileNames: boolean;
     FVml: boolean;
     FAllowPng: boolean;
     FTargetScreenSize: TXc12TargetScreenSize;
     FDpi: integer;
     FCodePage: integer;
public
     constructor Create;

     procedure Clear;

     property Css: boolean read FCss write FCss;
     property Thicket: boolean read FThicket write FThicket;
     property LongFileNames: boolean read FLongFileNames write FLongFileNames;
     property Vml: boolean read FVml write FVml;
     property AllowPng: boolean read FAllowPng write FAllowPng;
     property TargetScreenSize: TXc12TargetScreenSize read FTargetScreenSize write FTargetScreenSize;
     property Dpi: integer read FDpi write FDpi;
     property CodePage: integer read FCodePage write FCodePage;
     end;

type TXc12FileRecoveryPr = class(TObject)
protected
     FAutoRecover: boolean;
     FCrashSave: boolean;
     FDataExtractLoad: boolean;
     FRepairLoad: boolean;
public
     constructor Create;

     procedure Clear;

     property AutoRecover: boolean read FAutoRecover write FAutoRecover;
     property CrashSave: boolean read FCrashSave write FCrashSave;
     property DataExtractLoad: boolean read FDataExtractLoad write FDataExtractLoad;
     property RepairLoad: boolean read FRepairLoad write FRepairLoad;
     end;

type TXc12FileRecoveryPrs = class(TObjectList)
private
     function GetItems(Index: integer): TXc12FileRecoveryPr;
protected
public
     constructor Create;

     function Add: TXc12FileRecoveryPr;

     property Items[Index: integer]: TXc12FileRecoveryPr read GetItems; default;
     end;

type TXc12WebPublishObject = class(TObject)
protected
     FId: integer;
     FDivId: AxUCString;
     FSourceObject: AxUCString;
     FDestinationFile: AxUCString;
     FTitle: AxUCString;
     FAutoRepublish: boolean;
public
     constructor Create;

     procedure Clear;

     property Id: integer read FId write FId;
     property DivId: AxUCString read FDivId write FDivId;
     property SourceObject: AxUCString read FSourceObject write FSourceObject;
     property DestinationFile: AxUCString read FDestinationFile write FDestinationFile;
     property Title: AxUCString read FTitle write FTitle;
     property AutoRepublish: boolean read FAutoRepublish write FAutoRepublish;
     end;

type TXc12WebPublishObjects = class(TObjectList)
private
     function GetItems(Index: integer): TXc12WebPublishObject;
protected
public
     function Add: TXc12WebPublishObject;

     property Items[Index: integer]: TXc12WebPublishObject read GetItems; default;
     end;

type TXc12DataWorkbook = class(TXc12Data)
protected
     FFileVersion: TXc12FileVersion;
     FFileSharing: TXc12FileSharing;
     FWorkbookPr: TXc12WorkbookPr;
     FWorkbookProtection: TXc12WorkbookProtection;
     FBookViews: TXc12BookViews;
     FSheets: TXc12Sheets;
     FFunctionGroups: TXc12FunctionGroups;
     FDefinedNames: TXc12DefinedNames;
     FCalcPr: TXc12CalcPr;
     FOleSize: TXLSCellArea;
     FCustomWorkbookViews: TXc12CustomWorkbookViews;
     FPivotCaches: TCT_pivotCacheDefinitions;
     FSmartTagPr: TXc12SmartTagPr;
     FSmartTagTypes: TXc12SmartTagTypes;
     FWebPublishing: TXc12WebPublishing;
     FFileRecoveryPrs: TXc12FileRecoveryPrs;
     FWebPublishObjects: TXc12WebPublishObjects;

     FConnections: TXpgSimpleDOM;

     // Excel 97 username.
     FUserName: AxUCString;

     FClassFactory: TXLSClassFactory;
public
     constructor Create(AClassFactory: TXLSClassFactory);
     destructor Destroy; override;

     procedure Clear;

     property UserName: AxUCString read FUserName write FUserName;

     property FileVersion: TXc12FileVersion read FFileVersion;
     property FileSharing: TXc12FileSharing read FFileSharing;
     property WorkbookPr: TXc12WorkbookPr read FWorkbookPr;
     property WorkbookProtection: TXc12WorkbookProtection read FWorkbookProtection;
     property BookViews: TXc12BookViews read FBookViews;
     property Sheets: TXc12Sheets read FSheets;
     property FunctionGroups: TXc12FunctionGroups read FFunctionGroups;
     property DefinedNames: TXc12DefinedNames read FDefinedNames;
     property CalcPr: TXc12CalcPr read FCalcPr;
     property OleSize: TXLSCellArea read FOleSize write FOleSize;
     property CustomWorkbookViews: TXc12CustomWorkbookViews read FCustomWorkbookViews;
     property PivotCaches: TCT_pivotCacheDefinitions read FPivotCaches;
     property SmartTagPr: TXc12SmartTagPr read FSmartTagPr;
     property SmartTagTypes: TXc12SmartTagTypes read FSmartTagTypes;
     property WebPublishing: TXc12WebPublishing read FWebPublishing;
     property FileRecoveryPrs: TXc12FileRecoveryPrs read FFileRecoveryPrs;
     property WebPublishObjects: TXc12WebPublishObjects read FWebPublishObjects;
     property Connections: TXpgSimpleDOM read FConnections;
     end;

implementation

{ TXc12DataWorkbook }

procedure TXc12DataWorkbook.Clear;
begin
  FUserName := '';

  FFileVersion.Clear;
  FFileSharing.Clear;
  FWorkbookPr.Clear;
  FWorkbookProtection.Clear;
  FBookViews.Clear(True);
  FSheets.Clear;
  FFunctionGroups.Clear;
  FDefinedNames.Clear;
  FCalcPr.Clear;
  FCustomWorkbookViews.Clear;
  FPivotCaches.Clear;
  FSmartTagPr.Clear;
  FSmartTagTypes.Clear;
  FWebPublishing.Clear;
  FFileRecoveryPrs.Clear;
  FWebPublishObjects.Clear;
  FConnections.Clear;
end;

constructor TXc12DataWorkbook.Create(AClassFactory: TXLSClassFactory);
begin
  FClassFactory := AClassFactory;

  FFileVersion := TXc12FileVersion.Create;
  FFileSharing := TXc12FileSharing.Create;
  FWorkbookPr := TXc12WorkbookPr.Create;
  FWorkbookProtection := TXc12WorkbookProtection.Create;
  FBookViews := TXc12BookViews.Create;
  FSheets := TXc12Sheets.Create;
  FFunctionGroups := TXc12FunctionGroups.Create;

  FDefinedNames := TXc12DefinedNames(FClassFactory.CreateAClass(xcftNames));

  FCalcPr := TXc12CalcPr.Create;
  FCustomWorkbookViews := TXc12CustomWorkbookViews.Create;
  FPivotCaches := TCT_pivotCacheDefinitions.Create;
  FSmartTagPr := TXc12SmartTagPr.Create;
  FSmartTagTypes := TXc12SmartTagTypes.Create;
  FWebPublishing := TXc12WebPublishing.Create;
  FFileRecoveryPrs := TXc12FileRecoveryPrs.Create;
  FWebPublishObjects := TXc12WebPublishObjects.Create;

  FConnections := TXpgSimpleDOM.Create;
end;

destructor TXc12DataWorkbook.Destroy;
begin
  FFileVersion.Free;
  FFileSharing.Free;
  FWorkbookPr.Free;
  FWorkbookProtection.Free;
  FBookViews.Free;
  FSheets.Free;
  FFunctionGroups.Free;
  FDefinedNames.Free;
  FCalcPr.Free;
  FCustomWorkbookViews.Free;
  FPivotCaches.Free;
  FSmartTagPr.Free;
  FSmartTagTypes.Free;
  FWebPublishing.Free;
  FFileRecoveryPrs.Free;
  FWebPublishObjects.Free;
  FConnections.Free;
  inherited;
end;

{ TXc12FileVersion }

constructor TXc12FileVersion.Create;
begin
  Clear;
end;

procedure TXc12FileVersion.SetAppName(const Value: AxUCString);
begin
  if Value <> '' then
    FAppName := Value;
end;

procedure TXc12FileVersion.SetCodeName(const Value: AxUCString);
begin
  if Value <> '' then
    FCodeName := Value;
end;

procedure TXc12FileVersion.SetLastEdited(const Value: AxUCString);
begin
  if Value <> '' then
    FLastEdited := Value;
end;

procedure TXc12FileVersion.SetLowestEdited(const Value: AxUCString);
begin
  if Value <> '' then
    FLowestEdited := Value;
end;

procedure TXc12FileVersion.SetRupBuild(const Value: AxUCString);
begin
  if Value <> '' then
    FRupBuild := Value;
end;

procedure TXc12FileVersion.Clear;
begin
  FAppName := 'xl';
  FLastEdited := '4';
  FLowestEdited := '4';
  FRupBuild := CurrentBuildNumber;  // Excel 2010 value is 4506.
  FCodeName := '';
end;

{ TXc12FileSharing }

procedure TXc12FileSharing.Clear;
begin
  FReadOnlyReccomended := False;
  FUsername := '';
  FReservationPassword := 0;
end;

constructor TXc12FileSharing.Create;
begin
  Clear;
end;

{ TXc12WorkbookPr }

procedure TXc12WorkbookPr.Clear;
begin
  FShowBorderUnselectedTables := True;
  FFilterPrivacy := False;
  FPromptedSolutions := False;
  FShowInkAnnotation := True;
  FBackupFile := False;
  FSaveExternalLinkValues := True;
  FUpdateLinks := x12ulUserSet;
  FCodeName := '';
  FHidePivotFieldList := False;
  FShowPivotChartFilter := False;
  FAllowRefreshQuery := False;
  FPublishItems := False;
  FCheckCompatibility := False;
  FAutoCompressPictures := True;
  FRefreshAllConnections := False;
  FDefaultThemeVersion := 124226;
end;

constructor TXc12WorkbookPr.Create;
begin
  Clear;
end;

{ TXc12WorkbookProtection }

constructor TXc12WorkbookProtection.Create;
begin
  Clear;
end;

procedure TXc12WorkbookProtection.Clear;
begin
  FWorkbookPassword := 0;
  FRevisionPassword := 0;
  FLockStructure := False;
  FLockWindows := False;
  FLockRevision := False;
end;

{ TXc12BookView }

procedure TXc12BookView.Clear;
begin
  FVisibility := x12vVisible;
  FMinimized := False;
  FShowHorizontalScroll := True;
  FShowVerticalScroll := True;
  FShowSheetTabs := True;
  FXWindow := 0;
  FYWindow := 0;
  FWindowWidth := 0;
  FWindowHeight := 0 ;
  FTabRatio := 600;
  FFirstSheet := 0;
  FActiveTab := 0;
  FAutoFilterDateGrouping := True;
end;

constructor TXc12BookView.Create;
begin
  Clear;
end;

{ TXc12BookViews }

function TXc12BookViews.Add: TXc12BookView;
begin
  Result := TXc12BookView.Create;
  Result.Clear;
  inherited Add(Result);
end;

constructor TXc12BookViews.Create;
begin
  inherited Create;

  Clear(True);
end;

function TXc12BookViews.GetItems(Index: integer): TXc12BookView;
begin
  Result := TXc12BookView(inherited Items[Index]);
end;

procedure TXc12BookViews.Clear(AAddDefault: boolean);
var
  Item: TXc12BookView;
begin
  inherited Clear;

  if AAddDefault then begin
    Item := Add;

    Item.XWindow := 240;
    Item.YWindow := 45;
    Item.WindowWidth := 25875;
    Item.WindowHeight := 15915;

    Item.ShowHorizontalScroll := True;
    Item.ShowVerticalScroll := True;
    Item.ShowSheetTabs := True;
    Item.TabRatio := 600;
    Item.AutoFilterDateGrouping := True;
  end;
end;

function TXc12BookViews.Last: TXc12BookView;
begin
  Result := TXc12BookView(inherited Last);
end;

{ TXc12FunctionGroups }

procedure TXc12FunctionGroups.Add(AName: AxUCString);
begin
  FItems.Add(AName);
end;

function TXc12FunctionGroups.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXc12FunctionGroups.Create;
begin
  FItems := TStringList.Create;
  Clear;
end;

destructor TXc12FunctionGroups.Destroy;
begin
  FItems.Free;
  inherited;
end;

function TXc12FunctionGroups.GetItems(Index: integer): AxUCString;
begin
  Result := FItems[Index];
end;

procedure TXc12FunctionGroups.Clear;
begin
  FItems.Clear;
  FBuiltInGroupCount := 16;
end;

{ TXc12CalcPr }

procedure TXc12CalcPr.Clear;
begin
  FCalcId := 125725;;
  FCalcMode := cmAutomatic;
  FFullCalcOnLoad := False;
  FRefMode := x12rmA1;
  FIterate := False;
  FIterateCount := 100;
  FIterateDelta := 0.001;
  FFullPrecision := True;
  FCalcCompleted := True;
  FCalcOnSave := True ;
  FConcurrentCalc := True;
  FConcurrentManualCount := 0;
  FForceFullCalc := False;
end;

constructor TXc12CalcPr.Create;
begin
  Clear;
end;

{ TXc12CustomWorkbookView }

procedure TXc12CustomWorkbookView.Clear;
begin
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
  FShowComments := x12cCommIndicator;
  FShowObjects := x12oAll;
end;

constructor TXc12CustomWorkbookView.Create;
begin
  Clear;
end;

{ TXc12CustomWorkbookViews }

function TXc12CustomWorkbookViews.Add: TXc12CustomWorkbookView;
begin
  Result := TXc12CustomWorkbookView.Create;
  Result.Clear;
  inherited Add(Result);
end;

function TXc12CustomWorkbookViews.GetItems(Index: integer): TXc12CustomWorkbookView;
begin
  Result := TXc12CustomWorkbookView(inherited Items[Index]);
end;

{ TXc12SmartTagPr }

procedure TXc12SmartTagPr.Clear;
begin
  FEmbed := False;
  FShow := x12smtAll;
end;

constructor TXc12SmartTagPr.Create;
begin
  Clear;
end;

{ TXc12SmartTagType }

procedure TXc12SmartTagType.Clear;
begin
  FNamespaceUri := '';
  FName := '';
  FUrl := '';
end;

constructor TXc12SmartTagType.Create;
begin
  Clear;
end;

{ TXc12SmartTagTypes }

function TXc12SmartTagTypes.Add: TXc12SmartTagType;
begin
  Result := TXc12SmartTagType.Create;
  Result.Clear;
  inherited Add(Result);
end;

function TXc12SmartTagTypes.GetItems(Index: integer): TXc12SmartTagType;
begin
  Result := TXc12SmartTagType(inherited Items[Index]);
end;

{ TXc12WebPublishing }

procedure TXc12WebPublishing.Clear;
begin
  FCss := True;
  FThicket := True;
  FLongFileNames := True;
  FVml := False;
  FAllowPng := False;
  FTargetScreenSize := x12tss800x600;
  FDpi := 96;
  FCodePage := 0;
end;

constructor TXc12WebPublishing.Create;
begin
  Clear;
end;

{ TXc12FileRecoveryPr }

procedure TXc12FileRecoveryPr.Clear;
begin
  FAutoRecover := True;
  FCrashSave := False;
  FDataExtractLoad := False;
  FRepairLoad := False;
end;

constructor TXc12FileRecoveryPr.Create;
begin
  Clear;
end;

{ TXc12WebPublishObject }

procedure TXc12WebPublishObject.Clear;
begin
  FId := 0;
  FDivId := '';
  FSourceObject := '';
  FDestinationFile := '';
  FTitle := '';
  FAutoRepublish := False;
end;

constructor TXc12WebPublishObject.Create;
begin
  Clear;
end;

{ TXc12WebPublishObjects }

function TXc12WebPublishObjects.Add: TXc12WebPublishObject;
begin
  Result := TXc12WebPublishObject.Create;
  Result.Clear;
  inherited Add(Result);
end;

function TXc12WebPublishObjects.GetItems(Index: integer): TXc12WebPublishObject;
begin
  Result := TXc12WebPublishObject(inherited Items[Index]);
end;

{ TXc12Sheet }

procedure TXc12Sheet.Clear;
begin
  FName := '';
  FSheetId := -1;
  FState := x12vVisible;
  FRId := '';
end;

constructor TXc12Sheet.Create;
begin
  Clear;
end;

{ TXc12Sheets }

function TXc12Sheets.Add: TXc12Sheet;
begin
  Result := TXc12Sheet.Create;
  Result.Clear;
  inherited Add(Result);
end;

constructor TXc12Sheets.Create;
begin
  inherited Create;
end;

function TXc12Sheets.GetItems(Index: integer): TXc12Sheet;
begin
  Result := TXc12Sheet(inherited Items[Index]);
end;

{ TXcDefinedName }

procedure TXc12DefinedName.Assign(AName: TXc12DefinedName);
begin
  FName := AName.FName;
  FComment := AName.FComment;
  FCustomMenu := AName.FCustomMenu;
  FDescription := AName.FDescription;
  FHelp := AName.FHelp;
  FStatusBar := AName.FStatusBar;
  FLocalSheetId := AName.FLocalSheetId;
  FHidden := AName.FHidden;
  FFunction := AName.FFunction;
  FVbProcedure := AName.FVbProcedure;
  FXlm := AName.FXlm;
  FFunctionGroupId := AName.FFunctionGroupId;
  FShortcutKey := AName.FShortcutKey;
  FPublishToServer := AName.FPublishToServer;
  FWorkbookParameter := AName.FWorkbookParameter;

  FBuiltIn := AName.FBuiltIn;
  FContent := AName.FContent;
  FSimpleName := AName.FSimpleName;
  FArea := AName.FArea;
  FPtgsSz := AName.FPtgsSz;
  if FPtgsSz > 0 then begin
    GetMem(FPtgs,FPtgsSz);
    System.Move(AName.FPtgs^,FPtgs^,FPtgsSz);
  end;

  FDeleted := AName.FDeleted;

  FUnused97 := AName.FUnused97;
end;

procedure TXc12DefinedName.Clear;
begin
  FName := '';
  FComment := '';
  FCustomMenu := '';
  FDescription := '';
  FHelp := '';
  FStatusBar := '';
  FLocalSheetId := -1;
  FHidden := False;
  FFunction := False;
  FVbProcedure := False;
  FXlm := False;
  FFunctionGroupId := 0;
  FShortcutKey := '';
  FPublishToServer := False;
  FWorkbookParameter := False;

  FBuiltIn := bnNone;
  FContent := '';
  FDeleted := False;
end;

constructor TXc12DefinedName.Create;
begin
  Clear;
end;

destructor TXc12DefinedName.Destroy;
begin
  if FPtgs <> Nil then
    FreeMem(FPtgs);
  inherited;
end;

function TXc12DefinedName.GetCol1: integer;
begin
  Result := FArea.Col1 and not COL_ABSFLAG;
end;

function TXc12DefinedName.GetCol2: integer;
begin
  Result := FArea.Col2 and not COL_ABSFLAG;
end;

function TXc12DefinedName.GetContent: AxUCString;
begin
  Result := FContent;
end;

function TXc12DefinedName.GetRow1: integer;
begin
  Result := FArea.Row1 and not ROW_ABSFLAG;
end;

function TXc12DefinedName.GetRow2: integer;
begin
  Result := FArea.Row2 and not ROW_ABSFLAG;
end;

function TXc12DefinedName.GetSheetIndex: integer;
begin
  Result := FArea.SheetIndex;
end;

procedure TXc12DefinedName.SetContent(const Value: AxUCString);
begin
  FContent := Value;
end;

procedure TXc12DefinedName.SetName(const Value: AxUCString);
begin
  FName := Value;
  if Copy(FName,1,6) = '_xlnm.' then begin
         if FName = '_xlnm.Print_Area'       then FBuiltIn := bnPrintArea
    else if FName = '_xlnm.Print_Titles'     then FBuiltIn := bnPrintTitles
    else if FName = '_xlnm.Criteria'         then FBuiltIn := bnCriteria
    else if FName = '_xlnm._FilterDatabase'  then FBuiltIn := bnFilterDatabase
    else if FName = '_xlnm.Extract'          then FBuiltIn := bnExtract
    else if FName = '_xlnm.Consolidate_Area' then FBuiltIn := bnConsolidateArea
    else if FName = '_xlnm.Database'         then FBuiltIn := bnDatabase
    else if FName = '_xlnm.Sheet_Title'      then FBuiltIn := bnSheetTitle;
    FName := '';
  end
  else
    FBuiltIn := bnNone;
end;

{ TXc12DefinedNames }

function TXc12DefinedNames.Add(const AName: AxUCString; const ALocalSheetId: integer): TXc12DefinedName;
begin
  Result := AddOrGetFree;
  Result.Clear;
  Result.Name := AName;
  Result.LocalSheetId := ALocalSheetId;
  FHashValid := False;
end;

function TXc12DefinedNames.AddOrGetFree: TXc12DefinedName;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Deleted then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := TXc12DefinedName(FClassFactory.CreateAClass(xcftNamesMember));
  inherited Add(Result);
end;

procedure TXc12DefinedNames.Assign(ANames: TXc12DefinedNames);
var
  i: integer;
  N: TXc12DefinedName;
begin
  for i := 0 to ANames.Count - 1 do begin
    if not ANames[i].Deleted then begin
      N := TXc12DefinedName(FClassFactory.CreateAClass(xcftNamesMember));
      N.Assign(ANames[i]);

      inherited Add(N);
    end;
  end;
end;

procedure TXc12DefinedNames.BuildHash;
var
  i: integer;
begin
  if not FHashValid then begin
    FHash.Clear;
    for i := 0 to Count - 1 do begin
      if Items[i].LocalSheetId = -1 then
        FHash.AddObject(AnsiUppercase(Items[i].Name),TObject(i));
    end;
    FHashValid := True;
  end;
end;

function TXc12DefinedNames.Add(const ABuiltIn: TXc12BuiltInName; const ALocalSheetId: integer): TXc12DefinedName;
begin
  Result := AddOrGetFree;
  Result.BuiltIn := ABuiltIn;
  Result.LocalSheetId := ALocalSheetId;
end;

constructor TXc12DefinedNames.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;
  FClassFactory := AClassFactory;
{$ifdef DELPHI_5}
  FHash := TStringList.Create;
{$else}
  FHash := THashedStringList.Create;
{$endif}

  // FHash becomes invalid when a name is changed.
  FHashValid := False;
end;

procedure TXc12DefinedNames.Delete(const AName: AxUCString);
var
  i: integer;
begin
  BuildHash;

  i := FHash.IndexOf(AnsiUppercase(AName));
  if i >= 0 then
    DoDelete(i);
end;

procedure TXc12DefinedNames.Delete(const AName: AxUCString; const ALocalSheetId: integer);
var
  i: integer;
  S: AxUCString;
begin
  S := AnsiUppercase(AName);
  for i := 0 to Count - 1 do begin
    if (Items[i].Name = S) and (Items[i].LocalSheetId = ALocalSheetId) then begin
      DoDelete(i);
      Exit;
    end;
  end;
end;

procedure TXc12DefinedNames.Delete(const AId: TXc12BuiltInName; const ALocalSheetId: integer);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].BuiltIn = AId) and (Items[i].LocalSheetId = ALocalSheetId) then begin
      DoDelete(i);
      Exit;
    end;
  end;
end;

destructor TXc12DefinedNames.Destroy;
begin
  FHash.Free;
  inherited;
end;

procedure TXc12DefinedNames.DoDelete(AIndex: integer);
var
  Name: TXc12DefinedName;
begin
  Name := Items[AIndex];
  Name.Name := '';
  Name.Deleted := True;
  Name.BuiltIn := bnNone;
  if Name.FPtgs <> Nil then begin
    FreeMem(Name.FPtgs);
    Name.FPtgs := Nil;
  end;
  FHashValid := False;
end;

function TXc12DefinedNames.FindIdInSheet(const AName: AxUCString; const ASheetId: integer): integer;
var
  S: AxUCString;
begin
  S := AnsiUppercase(AName);

  for Result := 0 to Count - 1 do begin
    if (S = AnsiUppercase(Items[Result].Name)) and (Items[Result].LocalSheetId = ASheetId) and not Items[Result].Deleted then
      Exit;
  end;
  Result := -1;
end;

function TXc12DefinedNames.FindBuiltIn(const AId: TXc12BuiltInName; const ASheetId: integer): TXc12DefinedName;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].BuiltIn = AId) and (Items[i].LocalSheetId = ASheetId) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

// Checks first if there is a name with local scope (ASheetId).
function TXc12DefinedNames.FindId(const AName: AxUCString; const ASheetId: integer): integer;
var
  S: AxUCString;
begin
  if ASheetId >= 0 then begin
    Result := FindIdInSheet(AName,ASheetId);
    if Result >= 0 then
      Exit;
  end;

  S := AnsiUppercase(AName);

  BuildHash;

  Result := FHash.IndexOf(S);
  if Result >= 0 then
    Result := Integer(FHash.Objects[Result]);

//  for Result := 0 to Count - 1 do begin
//    if (S = AnsiUppercase(Items[Result].Name)) and (Items[Result].LocalSheetId = -1) then
//      Exit;
//  end;
//  Result := -1;
end;

function TXc12DefinedNames.GetItems(Index: integer): TXc12DefinedName;
begin
  Result := TXc12DefinedName(inherited Items[Index]);
end;

{ TXc12FileRecoveryPrs }

function TXc12FileRecoveryPrs.Add: TXc12FileRecoveryPr;
begin
  Result := TXc12FileRecoveryPr.Create;
  inherited Add(Result);
end;

constructor TXc12FileRecoveryPrs.Create;
begin
  inherited Create;
end;

function TXc12FileRecoveryPrs.GetItems(Index: integer): TXc12FileRecoveryPr;
begin
  Result := TXc12FileRecoveryPr(inherited Items[Index]);
end;

end.
