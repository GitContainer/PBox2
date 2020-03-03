unit xpgParseXLinks;

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
// File created on 2011-12-24 17:36:21

{$I AxCompilers.inc}

{$MINENUMSIZE 2}
{$BOOLEVAL OFF}
{$RANGECHECKS OFF}

interface

uses Classes, SysUtils, Contnrs, IniFiles, xpgPUtils, xpgPLists, xpgPXMLUtils,
XLSUtils5,
xpgPXML;

type TST_DdeValueType =  (stdvtNil,stdvtB,stdvtN,stdvtE,stdvtStr);
const StrTST_DdeValueType: array[0..4] of AxUCString = ('nil','b','n','e','str');
type TST_ExternalCellType =  (stectB,stectN,stectE,stectS,stectStr,stectInlineStr);
const StrTST_ExternalCellType: array[0..5] of AxUCString = ('b','n','e','s','str','inlineStr');

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

     class procedure AddEnums;
     class function  StrToEnum(AValue: AxUCString): integer;
     class function  StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
     class function  TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
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

     TCT_ExternalCell = class(TXPGBase)
protected
     FR: AxUCString;
     FT: TST_ExternalCellType;
     FVm: integer;
     FV: AxUCString;

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
     property T: TST_ExternalCellType read FT write FT;
     property Vm: integer read FVm write FVm;
     property V: AxUCString read FV write FV;
     end;

     TCT_ExternalCellXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExternalCell;
public
     function  Add: TCT_ExternalCell;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ExternalCell read GetItems; default;
     end;

     TCT_DdeValue = class(TXPGBase)
protected
     FT: TST_DdeValueType;
     FVal: AxUCString;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property T: TST_DdeValueType read FT write FT;
     property Val: AxUCString read FVal write FVal;
     end;

     TCT_DdeValueXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DdeValue;
public
     function  Add: TCT_DdeValue;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_DdeValue read GetItems; default;
     end;

     TCT_ExternalRow = class(TXPGBase)
protected
     FR: integer;
     FCellXpgList: TCT_ExternalCellXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R: integer read FR write FR;
     property CellXpgList: TCT_ExternalCellXpgList read FCellXpgList;
     end;

     TCT_ExternalRowXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExternalRow;
public
     function  Add: TCT_ExternalRow;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ExternalRow read GetItems; default;
     end;

     TCT_DdeValues = class(TXPGBase)
protected
     FRows: integer;
     FCols: integer;
     FValueXpgList: TCT_DdeValueXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Rows: integer read FRows write FRows;
     property Cols: integer read FCols write FCols;
     property ValueXpgList: TCT_DdeValueXpgList read FValueXpgList;
     end;

     TCT_ExternalSheetName = class(TXPGBase)
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

     TCT_ExternalSheetNameXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExternalSheetName;
public
     function  Add: TCT_ExternalSheetName;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ExternalSheetName read GetItems; default;
     end;

     TCT_ExternalDefinedName = class(TXPGBase)
protected
     FName: AxUCString;
     FRefersTo: AxUCString;
     FSheetId: integer;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property RefersTo: AxUCString read FRefersTo write FRefersTo;
     property SheetId: integer read FSheetId write FSheetId;
     end;

     TCT_ExternalDefinedNameXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExternalDefinedName;
public
     function  Add: TCT_ExternalDefinedName;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ExternalDefinedName read GetItems; default;
     end;

     TCT_ExternalSheetData = class(TXPGBase)
protected
     FSheetId: integer;
     FRefreshError: boolean;
     FRowXpgList: TCT_ExternalRowXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetId: integer read FSheetId write FSheetId;
     property RefreshError: boolean read FRefreshError write FRefreshError;
     property RowXpgList: TCT_ExternalRowXpgList read FRowXpgList;
     end;

     TCT_ExternalSheetDataXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_ExternalSheetData;
public
     function  Add: TCT_ExternalSheetData;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_ExternalSheetData read GetItems; default;
     end;

     TCT_DdeItem = class(TXPGBase)
protected
     FName: AxUCString;
     FOle: boolean;
     FAdvise: boolean;
     FPreferPic: boolean;
     FValues: TCT_DdeValues;

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
     property Ole: boolean read FOle write FOle;
     property Advise: boolean read FAdvise write FAdvise;
     property PreferPic: boolean read FPreferPic write FPreferPic;
     property Values: TCT_DdeValues read FValues;
     end;

     TCT_DdeItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_DdeItem;
public
     function  Add: TCT_DdeItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_DdeItem read GetItems; default;
     end;

     TCT_OleItem = class(TXPGBase)
protected
     FName: AxUCString;
     FIcon: boolean;
     FAdvise: boolean;
     FPreferPic: boolean;

     function  CheckAssigned: integer; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property Name: AxUCString read FName write FName;
     property Icon: boolean read FIcon write FIcon;
     property Advise: boolean read FAdvise write FAdvise;
     property PreferPic: boolean read FPreferPic write FPreferPic;
     end;

     TCT_OleItemXpgList = class(TXPGBaseObjectList)
protected
     function  GetItems(Index: integer): TCT_OleItem;
public
     function  Add: TCT_OleItem;
     function  CheckAssigned: integer;
     procedure Write(AWriter: TXpgWriteXML; AName: AxUCString);
     property Items[Index: integer]: TCT_OleItem read GetItems; default;
     end;

     TCT_ExternalSheetNames = class(TXPGBase)
protected
     FSheetNameXpgList: TCT_ExternalSheetNameXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetNameXpgList: TCT_ExternalSheetNameXpgList read FSheetNameXpgList;
     end;

     TCT_ExternalDefinedNames = class(TXPGBase)
protected
     FDefinedNameXpgList: TCT_ExternalDefinedNameXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DefinedNameXpgList: TCT_ExternalDefinedNameXpgList read FDefinedNameXpgList;
     end;

     TCT_ExternalSheetDataSet = class(TXPGBase)
protected
     FSheetDataXpgList: TCT_ExternalSheetDataXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property SheetDataXpgList: TCT_ExternalSheetDataXpgList read FSheetDataXpgList;
     end;

     TCT_DdeItems = class(TXPGBase)
protected
     FDdeItemXpgList: TCT_DdeItemXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DdeItemXpgList: TCT_DdeItemXpgList read FDdeItemXpgList;
     end;

     TCT_OleItems = class(TXPGBase)
protected
     FOleItemXpgList: TCT_OleItemXpgList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property OleItemXpgList: TCT_OleItemXpgList read FOleItemXpgList;
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

     TCT_ExternalBook = class(TXPGBase)
protected
     FR_Id: AxUCString;
     FSheetNames: TCT_ExternalSheetNames;
     FDefinedNames: TCT_ExternalDefinedNames;
     FSheetDataSet: TCT_ExternalSheetDataSet;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     property SheetNames: TCT_ExternalSheetNames read FSheetNames;
     property DefinedNames: TCT_ExternalDefinedNames read FDefinedNames;
     property SheetDataSet: TCT_ExternalSheetDataSet read FSheetDataSet;
     end;

     TCT_DdeLink = class(TXPGBase)
protected
     FDdeService: AxUCString;
     FDdeTopic: AxUCString;
     FDdeItems: TCT_DdeItems;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property DdeService: AxUCString read FDdeService write FDdeService;
     property DdeTopic: AxUCString read FDdeTopic write FDdeTopic;
     property DdeItems: TCT_DdeItems read FDdeItems;
     end;

     TCT_OleLink = class(TXPGBase)
protected
     FR_Id: AxUCString;
     FProgId: AxUCString;
     FOleItems: TCT_OleItems;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

     procedure WriteAttributes(AWriter: TXpgWriteXML);
     procedure AssignAttributes(AAttributes: TXpgXMLAttributeList); override;
public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property R_Id: AxUCString read FR_Id write FR_Id;
     property ProgId: AxUCString read FProgId write FProgId;
     property OleItems: TCT_OleItems read FOleItems;
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

     TCT_ExternalLink = class(TXPGBase)
protected
     FExternalBook: TCT_ExternalBook;
     FDdeLink: TCT_DdeLink;
     FOleLink: TCT_OleLink;
     FExtLst: TCT_ExtensionList;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property ExternalBook: TCT_ExternalBook read FExternalBook;
     property DdeLink: TCT_DdeLink read FDdeLink;
     property OleLink: TCT_OleLink read FOleLink;
     property ExtLst: TCT_ExtensionList read FExtLst;
     end;

     T__ROOT__ = class(TXPGBase)
protected
     FRootAttributes: TStringXpgList;
     FExternalLink: TCT_ExternalLink;

     function  CheckAssigned: integer; override;
     function  HandleElement(AReader: TXpgReadXML): TXPGBase; override;
     procedure Write(AWriter: TXpgWriteXML);

public
     constructor Create(AOwner: TXPGDocBase);
     destructor Destroy; override;
     procedure Clear;

     property RootAttributes: TStringXpgList read FRootAttributes;
     property ExternalLink: TCT_ExternalLink read FExternalLink;
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

     TXPGDocXLink = class(TXPGDocBase)
protected
     FRoot: T__ROOT__;
     FReader: TXPGReader;
     FWriter: TXpgWriteXML;

     function  GetExternalLink: TCT_ExternalLink;
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromFile(AFilename: AxUCString);
     procedure LoadFromStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString);

     procedure SaveToStream(AStream: TStream);

     property Root: T__ROOT__ read FRoot;
     property ExternalLink: TCT_ExternalLink read GetExternalLink;
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
  FEnums.AddObject('stdvtNil',TObject(0));
  FEnums.AddObject('stdvtB',TObject(1));
  FEnums.AddObject('stdvtN',TObject(2));
  FEnums.AddObject('stdvtE',TObject(3));
  FEnums.AddObject('stdvtStr',TObject(4));
  FEnums.AddObject('stectB',TObject(0));
  FEnums.AddObject('stectN',TObject(1));
  FEnums.AddObject('stectE',TObject(2));
  FEnums.AddObject('stectS',TObject(3));
  FEnums.AddObject('stectStr',TObject(4));
  FEnums.AddObject('stectInlineStr',TObject(5));
end;

class function  TXPGBase.StrToEnum(AValue: AxUCString): integer;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then 
    Result := Integer(FEnums.Objects[i])
  else 
    Result := 0;
end;

class function  TXPGBase.StrToEnumDef(AValue: AxUCString; ADefault: integer): integer;
var
  i: integer;
begin
  i := FEnums.IndexOf(AValue);
  if i >= 0 then 
    Result := Integer(FEnums.Objects[i])
  else 
    Result := ADefault;
end;

class function  TXPGBase.TryStrToEnum(AValue: AxUCString; AText: AxUCString; AEnumNames: array of AxUCString; APtrInt: PInteger): boolean;
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
  FCurrent.AssignAttributes(Attributes);
end;

procedure TXPGReader.EndTag;
begin
  FCurrent := TXPGBase(FStack.Pop);
  FCurrent.AfterTag;
end;

{ TCT_ExternalCell }

function  TCT_ExternalCell.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> '' then 
    Inc(AttrsAssigned);
  if FT <> stectN then 
    Inc(AttrsAssigned);
  if FVm <> 0 then 
    Inc(AttrsAssigned);
  if FV <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ExternalCell.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'v' then 
    FV := AReader.Text
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalCell.Write(AWriter: TXpgWriteXML);
begin
  if FV <> '' then 
    AWriter.SimpleTextTag('v',FV);
end;

procedure TCT_ExternalCell.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FR <> '' then 
    AWriter.AddAttribute('r',FR);
  if FT <> stectN then 
    AWriter.AddAttribute('t',StrTST_ExternalCellType[Integer(FT)]);
  if FVm <> 0 then 
    AWriter.AddAttribute('vm',XmlIntToStr(FVm));
end;

procedure TCT_ExternalCell.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $00000072: FR := AAttributes.Values[i];
      $00000074: FT := TST_ExternalCellType(StrToEnum('stect' + AAttributes.Values[i]));
      $000000E3: FVm := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_ExternalCell.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 3;
  FT := stectN;
  FVm := 0;
end;

destructor TCT_ExternalCell.Destroy;
begin
end;

procedure TCT_ExternalCell.Clear;
begin
  FAssigneds := [];
  FV := '';
  FR := '';
  FT := stectN;
  FVm := 0;
end;

{ TCT_ExternalCellXpgList }

function  TCT_ExternalCellXpgList.GetItems(Index: integer): TCT_ExternalCell;
begin
  Result := TCT_ExternalCell(inherited Items[Index]);
end;

function  TCT_ExternalCellXpgList.Add: TCT_ExternalCell;
begin
  Result := TCT_ExternalCell.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExternalCellXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExternalCellXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DdeValue }

function  TCT_DdeValue.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FT <> stdvtN then 
    Inc(AttrsAssigned);
  if FVal <> '' then 
    Inc(ElemsAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DdeValue.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'val' then 
    FVal := AReader.Text
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DdeValue.Write(AWriter: TXpgWriteXML);
begin
  if FVal <> '' then 
    AWriter.SimpleTextTag('val',FVal);
end;

procedure TCT_DdeValue.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FT <> stdvtN then 
    AWriter.AddAttribute('t',StrTST_DdeValueType[Integer(FT)]);
end;

procedure TCT_DdeValue.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 't' then 
    FT := TST_DdeValueType(StrToEnum('stdvt' + AAttributes.Values[0]))
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_DdeValue.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FT := stdvtN;
end;

destructor TCT_DdeValue.Destroy;
begin
end;

procedure TCT_DdeValue.Clear;
begin
  FAssigneds := [];
  FVal := '';
  FT := stdvtN;
end;

{ TCT_DdeValueXpgList }

function  TCT_DdeValueXpgList.GetItems(Index: integer): TCT_DdeValue;
begin
  Result := TCT_DdeValue(inherited Items[Index]);
end;

function  TCT_DdeValueXpgList.Add: TCT_DdeValue;
begin
  Result := TCT_DdeValue.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DdeValueXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DdeValueXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ExternalRow }

function  TCT_ExternalRow.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR <> 0 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FCellXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ExternalRow.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'cell' then 
    Result := FCellXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalRow.Write(AWriter: TXpgWriteXML);
begin
  FCellXpgList.Write(AWriter,'cell');
end;

procedure TCT_ExternalRow.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r',XmlIntToStr(FR));
end;

procedure TCT_ExternalRow.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'r' then 
    FR := XmlStrToIntDef(AAttributes.Values[0],0)
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ExternalRow.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 1;
  FCellXpgList := TCT_ExternalCellXpgList.Create(FOwner);
end;

destructor TCT_ExternalRow.Destroy;
begin
  FCellXpgList.Free;
end;

procedure TCT_ExternalRow.Clear;
begin
  FAssigneds := [];
  FCellXpgList.Clear;
  FR := 0;
end;

{ TCT_ExternalRowXpgList }

function  TCT_ExternalRowXpgList.GetItems(Index: integer): TCT_ExternalRow;
begin
  Result := TCT_ExternalRow(inherited Items[Index]);
end;

function  TCT_ExternalRowXpgList.Add: TCT_ExternalRow;
begin
  Result := TCT_ExternalRow.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExternalRowXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExternalRowXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DdeValues }

function  TCT_DdeValues.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FRows <> 1 then 
    Inc(AttrsAssigned);
  if FCols <> 1 then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FValueXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DdeValues.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'value' then 
    Result := FValueXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DdeValues.Write(AWriter: TXpgWriteXML);
begin
  FValueXpgList.Write(AWriter,'value');
end;

procedure TCT_DdeValues.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FRows <> 1 then 
    AWriter.AddAttribute('rows',XmlIntToStr(FRows));
  if FCols <> 1 then 
    AWriter.AddAttribute('cols',XmlIntToStr(FCols));
end;

procedure TCT_DdeValues.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000001CB: FRows := XmlStrToIntDef(AAttributes.Values[i],0);
      $000001B1: FCols := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_DdeValues.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FValueXpgList := TCT_DdeValueXpgList.Create(FOwner);
  FRows := 1;
  FCols := 1;
end;

destructor TCT_DdeValues.Destroy;
begin
  FValueXpgList.Free;
end;

procedure TCT_DdeValues.Clear;
begin
  FAssigneds := [];
  FValueXpgList.Clear;
  FRows := 1;
  FCols := 1;
end;

{ TCT_ExternalSheetName }

function  TCT_ExternalSheetName.CheckAssigned: integer;
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

procedure TCT_ExternalSheetName.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ExternalSheetName.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FVal <> '' then 
    AWriter.AddAttribute('val',FVal);
end;

procedure TCT_ExternalSheetName.AssignAttributes(AAttributes: TXpgXMLAttributeList);
begin
  if AAttributes[0] = 'val' then 
    FVal := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ExternalSheetName.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 1;
end;

destructor TCT_ExternalSheetName.Destroy;
begin
end;

procedure TCT_ExternalSheetName.Clear;
begin
  FAssigneds := [];
  FVal := '';
end;

{ TCT_ExternalSheetNameXpgList }

function  TCT_ExternalSheetNameXpgList.GetItems(Index: integer): TCT_ExternalSheetName;
begin
  Result := TCT_ExternalSheetName(inherited Items[Index]);
end;

function  TCT_ExternalSheetNameXpgList.Add: TCT_ExternalSheetName;
begin
  Result := TCT_ExternalSheetName.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExternalSheetNameXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExternalSheetNameXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ExternalDefinedName }

function  TCT_ExternalDefinedName.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FRefersTo <> '' then 
    Inc(AttrsAssigned);
  if FSheetId <> 0 then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_ExternalDefinedName.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_ExternalDefinedName.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  if FRefersTo <> '' then 
    AWriter.AddAttribute('refersTo',FRefersTo);
  if FSheetId <> 0 then 
    AWriter.AddAttribute('sheetId',XmlIntToStr(FSheetId));
end;

procedure TCT_ExternalDefinedName.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000001A1: FName := AAttributes.Values[i];
      $0000034A: FRefersTo := AAttributes.Values[i];
      $000002C6: FSheetId := XmlStrToIntDef(AAttributes.Values[i],0);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_ExternalDefinedName.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 3;
end;

destructor TCT_ExternalDefinedName.Destroy;
begin
end;

procedure TCT_ExternalDefinedName.Clear;
begin
  FAssigneds := [];
  FName := '';
  FRefersTo := '';
  FSheetId := 0;
end;

{ TCT_ExternalDefinedNameXpgList }

function  TCT_ExternalDefinedNameXpgList.GetItems(Index: integer): TCT_ExternalDefinedName;
begin
  Result := TCT_ExternalDefinedName(inherited Items[Index]);
end;

function  TCT_ExternalDefinedNameXpgList.Add: TCT_ExternalDefinedName;
begin
  Result := TCT_ExternalDefinedName.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExternalDefinedNameXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExternalDefinedNameXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ExternalSheetData }

function  TCT_ExternalSheetData.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
//  if FSheetId <> 0 then
    Inc(AttrsAssigned);
  if FRefreshError <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FRowXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ExternalSheetData.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'row' then 
    Result := FRowXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalSheetData.Write(AWriter: TXpgWriteXML);
begin
  FRowXpgList.Write(AWriter,'row');
end;

procedure TCT_ExternalSheetData.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('sheetId',XmlIntToStr(FSheetId));
  if FRefreshError <> False then 
    AWriter.AddAttribute('refreshError',XmlBoolToStr(FRefreshError));
end;

procedure TCT_ExternalSheetData.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000002C6: FSheetId := XmlStrToIntDef(AAttributes.Values[i],0);
      $000004F9: FRefreshError := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_ExternalSheetData.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FRowXpgList := TCT_ExternalRowXpgList.Create(FOwner);
  FRefreshError := False;
end;

destructor TCT_ExternalSheetData.Destroy;
begin
  FRowXpgList.Free;
end;

procedure TCT_ExternalSheetData.Clear;
begin
  FAssigneds := [];
  FRowXpgList.Clear;
  FSheetId := 0;
  FRefreshError := False;
end;

{ TCT_ExternalSheetDataXpgList }

function  TCT_ExternalSheetDataXpgList.GetItems(Index: integer): TCT_ExternalSheetData;
begin
  Result := TCT_ExternalSheetData(inherited Items[Index]);
end;

function  TCT_ExternalSheetDataXpgList.Add: TCT_ExternalSheetData;
begin
  Result := TCT_ExternalSheetData.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_ExternalSheetDataXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_ExternalSheetDataXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_DdeItem }

function  TCT_DdeItem.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FOle <> False then 
    Inc(AttrsAssigned);
  if FAdvise <> False then 
    Inc(AttrsAssigned);
  if FPreferPic <> False then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FValues.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DdeItem.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'values' then 
    Result := FValues
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DdeItem.Write(AWriter: TXpgWriteXML);
begin
  if FValues.Assigned then 
  begin
    FValues.WriteAttributes(AWriter);
    if xaElements in FValues.FAssigneds then 
    begin
      AWriter.BeginTag('values');
      FValues.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('values');
  end;
end;

procedure TCT_DdeItem.WriteAttributes(AWriter: TXpgWriteXML);
begin
  if FName <> '' then 
    AWriter.AddAttribute('name',FName);
  if FOle <> False then 
    AWriter.AddAttribute('ole',XmlBoolToStr(FOle));
  if FAdvise <> False then 
    AWriter.AddAttribute('advise',XmlBoolToStr(FAdvise));
  if FPreferPic <> False then 
    AWriter.AddAttribute('preferPic',XmlBoolToStr(FPreferPic));
end;

procedure TCT_DdeItem.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000001A1: FName := AAttributes.Values[i];
      $00000140: FOle := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000027C: FAdvise := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003A0: FPreferPic := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_DdeItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 4;
  FValues := TCT_DdeValues.Create(FOwner);
  FName := '0';
  FOle := False;
  FAdvise := False;
  FPreferPic := False;
end;

destructor TCT_DdeItem.Destroy;
begin
  FValues.Free;
end;

procedure TCT_DdeItem.Clear;
begin
  FAssigneds := [];
  FValues.Clear;
  FName := '0';
  FOle := False;
  FAdvise := False;
  FPreferPic := False;
end;

{ TCT_DdeItemXpgList }

function  TCT_DdeItemXpgList.GetItems(Index: integer): TCT_DdeItem;
begin
  Result := TCT_DdeItem(inherited Items[Index]);
end;

function  TCT_DdeItemXpgList.Add: TCT_DdeItem;
begin
  Result := TCT_DdeItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_DdeItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_DdeItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_OleItem }

function  TCT_OleItem.CheckAssigned: integer;
var
  AttrsAssigned: integer;
begin
  AttrsAssigned := 0;
  FAssigneds := [];
  if FName <> '' then 
    Inc(AttrsAssigned);
  if FIcon <> False then 
    Inc(AttrsAssigned);
  if FAdvise <> False then 
    Inc(AttrsAssigned);
  if FPreferPic <> False then 
    Inc(AttrsAssigned);
  Result := 0;
  Inc(Result,AttrsAssigned);
  if AttrsAssigned > 0 then 
    FAssigneds := [xaAttributes];
end;

procedure TCT_OleItem.Write(AWriter: TXpgWriteXML);
begin
end;

procedure TCT_OleItem.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('name',FName);
  if FIcon <> False then 
    AWriter.AddAttribute('icon',XmlBoolToStr(FIcon));
  if FAdvise <> False then 
    AWriter.AddAttribute('advise',XmlBoolToStr(FAdvise));
  if FPreferPic <> False then 
    AWriter.AddAttribute('preferPic',XmlBoolToStr(FPreferPic));
end;

procedure TCT_OleItem.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000001A1: FName := AAttributes.Values[i];
      $000001A9: FIcon := XmlStrToBoolDef(AAttributes.Values[i],False);
      $0000027C: FAdvise := XmlStrToBoolDef(AAttributes.Values[i],False);
      $000003A0: FPreferPic := XmlStrToBoolDef(AAttributes.Values[i],False);
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_OleItem.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 0;
  FAttributeCount := 4;
  FIcon := False;
  FAdvise := False;
  FPreferPic := False;
end;

destructor TCT_OleItem.Destroy;
begin
end;

procedure TCT_OleItem.Clear;
begin
  FAssigneds := [];
  FName := '';
  FIcon := False;
  FAdvise := False;
  FPreferPic := False;
end;

{ TCT_OleItemXpgList }

function  TCT_OleItemXpgList.GetItems(Index: integer): TCT_OleItem;
begin
  Result := TCT_OleItem(inherited Items[Index]);
end;

function  TCT_OleItemXpgList.Add: TCT_OleItem;
begin
  Result := TCT_OleItem.Create(FOwner);
  inherited Add(Result);
end;

function  TCT_OleItemXpgList.CheckAssigned: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do 
    Inc(Result,Items[i].CheckAssigned);
  FAssigned := Result > 0;
end;

procedure TCT_OleItemXpgList.Write(AWriter: TXpgWriteXML; AName: AxUCString);
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

{ TCT_ExternalSheetNames }

function  TCT_ExternalSheetNames.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetNameXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExternalSheetNames.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'sheetName' then 
    Result := FSheetNameXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalSheetNames.Write(AWriter: TXpgWriteXML);
begin
  FSheetNameXpgList.Write(AWriter,'sheetName');
end;

constructor TCT_ExternalSheetNames.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FSheetNameXpgList := TCT_ExternalSheetNameXpgList.Create(FOwner);
end;

destructor TCT_ExternalSheetNames.Destroy;
begin
  FSheetNameXpgList.Free;
end;

procedure TCT_ExternalSheetNames.Clear;
begin
  FAssigneds := [];
  FSheetNameXpgList.Clear;
end;

{ TCT_ExternalDefinedNames }

function  TCT_ExternalDefinedNames.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FDefinedNameXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExternalDefinedNames.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'definedName' then 
    Result := FDefinedNameXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalDefinedNames.Write(AWriter: TXpgWriteXML);
begin
  FDefinedNameXpgList.Write(AWriter,'definedName');
end;

constructor TCT_ExternalDefinedNames.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FDefinedNameXpgList := TCT_ExternalDefinedNameXpgList.Create(FOwner);
end;

destructor TCT_ExternalDefinedNames.Destroy;
begin
  FDefinedNameXpgList.Free;
end;

procedure TCT_ExternalDefinedNames.Clear;
begin
  FAssigneds := [];
  FDefinedNameXpgList.Clear;
end;

{ TCT_ExternalSheetDataSet }

function  TCT_ExternalSheetDataSet.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FSheetDataXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExternalSheetDataSet.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'sheetData' then 
    Result := FSheetDataXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalSheetDataSet.Write(AWriter: TXpgWriteXML);
begin
  FSheetDataXpgList.Write(AWriter,'sheetData');
end;

constructor TCT_ExternalSheetDataSet.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FSheetDataXpgList := TCT_ExternalSheetDataXpgList.Create(FOwner);
end;

destructor TCT_ExternalSheetDataSet.Destroy;
begin
  FSheetDataXpgList.Free;
end;

procedure TCT_ExternalSheetDataSet.Clear;
begin
  FAssigneds := [];
  FSheetDataXpgList.Clear;
end;

{ TCT_DdeItems }

function  TCT_DdeItems.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FDdeItemXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_DdeItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'ddeItem' then 
    Result := FDdeItemXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DdeItems.Write(AWriter: TXpgWriteXML);
begin
  FDdeItemXpgList.Write(AWriter,'ddeItem');
end;

constructor TCT_DdeItems.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FDdeItemXpgList := TCT_DdeItemXpgList.Create(FOwner);
end;

destructor TCT_DdeItems.Destroy;
begin
  FDdeItemXpgList.Free;
end;

procedure TCT_DdeItems.Clear;
begin
  FAssigneds := [];
  FDdeItemXpgList.Clear;
end;

{ TCT_OleItems }

function  TCT_OleItems.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FOleItemXpgList.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_OleItems.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'oleItem' then 
    Result := FOleItemXpgList.Add
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_OleItems.Write(AWriter: TXpgWriteXML);
begin
  FOleItemXpgList.Write(AWriter,'oleItem');
end;

constructor TCT_OleItems.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 0;
  FOleItemXpgList := TCT_OleItemXpgList.Create(FOwner);
end;

destructor TCT_OleItems.Destroy;
begin
  FOleItemXpgList.Free;
end;

procedure TCT_OleItems.Clear;
begin
  FAssigneds := [];
  FOleItemXpgList.Clear;
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
    FUri := AAttributes.Values[0]
  else 
    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
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

{ TCT_ExternalBook }

function  TCT_ExternalBook.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FSheetNames.CheckAssigned);
  Inc(ElemsAssigned,FDefinedNames.CheckAssigned);
  Inc(ElemsAssigned,FSheetDataSet.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_ExternalBook.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $0000040D: Result := FSheetNames;
    $000004C3: Result := FDefinedNames;
    $000004BF: Result := FSheetDataSet;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalBook.Write(AWriter: TXpgWriteXML);
begin
  if FSheetNames.Assigned then 
    if xaElements in FSheetNames.FAssigneds then 
    begin
      AWriter.BeginTag('sheetNames');
      FSheetNames.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetNames');
  if FDefinedNames.Assigned then 
    if xaElements in FDefinedNames.FAssigneds then 
    begin
      AWriter.BeginTag('definedNames');
      FDefinedNames.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('definedNames');
  if FSheetDataSet.Assigned then 
    if xaElements in FSheetDataSet.FAssigneds then 
    begin
      AWriter.BeginTag('sheetDataSet');
      FSheetDataSet.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('sheetDataSet');
end;

procedure TCT_ExternalBook.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
end;

procedure TCT_ExternalBook.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
begin
  for i := 0 to AAttributes.Count - 1 do begin
    if AAttributes[i] = 'r:id' then
      FR_Id := AAttributes.Values[i];
  end
//  else
//    FOwner.Errors.Error(xemUnknownAttribute,AAttributes[0]);
end;

constructor TCT_ExternalBook.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 3;
  FAttributeCount := 1;
  FSheetNames := TCT_ExternalSheetNames.Create(FOwner);
  FDefinedNames := TCT_ExternalDefinedNames.Create(FOwner);
  FSheetDataSet := TCT_ExternalSheetDataSet.Create(FOwner);
end;

destructor TCT_ExternalBook.Destroy;
begin
  FSheetNames.Free;
  FDefinedNames.Free;
  FSheetDataSet.Free;
end;

procedure TCT_ExternalBook.Clear;
begin
  FAssigneds := [];
  FSheetNames.Clear;
  FDefinedNames.Clear;
  FSheetDataSet.Clear;
  FR_Id := '';
end;

{ TCT_DdeLink }

function  TCT_DdeLink.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FDdeService <> '' then 
    Inc(AttrsAssigned);
  if FDdeTopic <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FDdeItems.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_DdeLink.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'ddeItems' then 
    Result := FDdeItems
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_DdeLink.Write(AWriter: TXpgWriteXML);
begin
  if FDdeItems.Assigned then 
    if xaElements in FDdeItems.FAssigneds then 
    begin
      AWriter.BeginTag('ddeItems');
      FDdeItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('ddeItems');
end;

procedure TCT_DdeLink.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('ddeService',FDdeService);
  AWriter.AddAttribute('ddeTopic',FDdeTopic);
end;

procedure TCT_DdeLink.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $000003FE: FDdeService := AAttributes.Values[i];
      $0000032C: FDdeTopic := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_DdeLink.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FDdeItems := TCT_DdeItems.Create(FOwner);
end;

destructor TCT_DdeLink.Destroy;
begin
  FDdeItems.Free;
end;

procedure TCT_DdeLink.Clear;
begin
  FAssigneds := [];
  FDdeItems.Clear;
  FDdeService := '';
  FDdeTopic := '';
end;

{ TCT_OleLink }

function  TCT_OleLink.CheckAssigned: integer;
var
  ElemsAssigned: integer;
  AttrsAssigned: integer;
begin
  ElemsAssigned := 0;
  AttrsAssigned := 0;
  FAssigneds := [];
  if FR_Id <> '' then 
    Inc(AttrsAssigned);
  if FProgId <> '' then 
    Inc(AttrsAssigned);
  Inc(ElemsAssigned,FOleItems.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaElements];
  if AttrsAssigned > 0 then 
    FAssigneds := FAssigneds + [xaAttributes];
  Inc(Result,ElemsAssigned + AttrsAssigned);
end;

function  TCT_OleLink.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  if QName = 'oleItems' then 
    Result := FOleItems
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_OleLink.Write(AWriter: TXpgWriteXML);
begin
  if FOleItems.Assigned then 
    if xaElements in FOleItems.FAssigneds then 
    begin
      AWriter.BeginTag('oleItems');
      FOleItems.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('oleItems');
end;

procedure TCT_OleLink.WriteAttributes(AWriter: TXpgWriteXML);
begin
  AWriter.AddAttribute('r:id',FR_Id);
  AWriter.AddAttribute('progId',FProgId);
end;

procedure TCT_OleLink.AssignAttributes(AAttributes: TXpgXMLAttributeList);
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to AAttributes.Count - 1 do 
  begin
    S := AAttributes[i];
    case CalcHash_A(S) of
      $00000179: FR_Id := AAttributes.Values[i];
      $00000265: FProgId := AAttributes.Values[i];
      else 
        FOwner.Errors.Error(xemUnknownAttribute,S);
    end;
  end
end;

constructor TCT_OleLink.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 1;
  FAttributeCount := 2;
  FOleItems := TCT_OleItems.Create(FOwner);
end;

destructor TCT_OleLink.Destroy;
begin
  FOleItems.Free;
end;

procedure TCT_OleLink.Clear;
begin
  FAssigneds := [];
  FOleItems.Clear;
  FR_Id := '';
  FProgId := '';
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

{ TCT_ExternalLink }

function  TCT_ExternalLink.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FExternalBook.CheckAssigned);
  Inc(ElemsAssigned,FDdeLink.CheckAssigned);
  Inc(ElemsAssigned,FOleLink.CheckAssigned);
  Inc(ElemsAssigned,FExtLst.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  TCT_ExternalLink.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  QName: AxUCString;
begin
  Result := Self;
  QName := AReader.QName;
  case CalcHash_A(QName) of
    $000004EE: Result := FExternalBook;
    $000002BB: Result := FDdeLink;
    $000002CE: Result := FOleLink;
    $00000284: Result := FExtLst;
    else 
      FOwner.Errors.Error(xemUnknownElement,QName);
  end;
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure TCT_ExternalLink.Write(AWriter: TXpgWriteXML);
begin
  if FExternalBook.Assigned then 
  begin
    FExternalBook.WriteAttributes(AWriter);
    if xaElements in FExternalBook.FAssigneds then 
    begin
      AWriter.BeginTag('externalBook');
      FExternalBook.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('externalBook');
  end;
  if FDdeLink.Assigned then 
  begin
    FDdeLink.WriteAttributes(AWriter);
    if xaElements in FDdeLink.FAssigneds then 
    begin
      AWriter.BeginTag('ddeLink');
      FDdeLink.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('ddeLink');
  end;
  if FOleLink.Assigned then 
  begin
    FOleLink.WriteAttributes(AWriter);
    if xaElements in FOleLink.FAssigneds then 
    begin
      AWriter.BeginTag('oleLink');
      FOleLink.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('oleLink');
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

constructor TCT_ExternalLink.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FElementCount := 4;
  FAttributeCount := 0;
  FExternalBook := TCT_ExternalBook.Create(FOwner);
  FDdeLink := TCT_DdeLink.Create(FOwner);
  FOleLink := TCT_OleLink.Create(FOwner);
  FExtLst := TCT_ExtensionList.Create(FOwner);
end;

destructor TCT_ExternalLink.Destroy;
begin
  FExternalBook.Free;
  FDdeLink.Free;
  FOleLink.Free;
  FExtLst.Free;
end;

procedure TCT_ExternalLink.Clear;
begin
  FAssigneds := [];
  FExternalBook.Clear;
  FDdeLink.Clear;
  FOleLink.Clear;
  FExtLst.Clear;
end;

{ T__ROOT__ }

function  T__ROOT__.CheckAssigned: integer;
var
  ElemsAssigned: integer;
begin
  ElemsAssigned := 0;
  FAssigneds := [];
  Inc(ElemsAssigned,FExternalLink.CheckAssigned);
  Result := 0;
  if ElemsAssigned > 0 then 
    FAssigneds := [xaElements];
  Inc(Result,ElemsAssigned);
end;

function  T__ROOT__.HandleElement(AReader: TXpgReadXML): TXPGBase;
var
  i: integer;
  QName: AxUCString;
begin
  for i := 0 to AReader.Attributes.Count - 1 do 
    FRootAttributes.Add(AReader.Attributes.AsXmlText2(i));
  Result := Self;
  QName := AReader.QName;
  if QName = 'externalLink' then 
    Result := FExternalLink
  else 
    FOwner.Errors.Error(xemUnknownElement,QName);
  if Result <> Self then 
    Result.FAssigneds := [xaRead];
end;

procedure T__ROOT__.Write(AWriter: TXpgWriteXML);
begin
  AWriter.Attributes := FRootAttributes.Text;
  if FExternalLink.Assigned then 
    if xaElements in FExternalLink.FAssigneds then 
    begin
      AWriter.BeginTag('externalLink');
      FExternalLink.Write(AWriter);
      AWriter.EndTag;
    end
    else 
      AWriter.SimpleTag('externalLink');
end;

constructor T__ROOT__.Create(AOwner: TXPGDocBase);
begin
  FOwner := AOwner;
  FRootAttributes := TStringXpgList.Create;
  FElementCount := 1;
  FAttributeCount := 0;
  FExternalLink := TCT_ExternalLink.Create(FOwner);
end;

destructor T__ROOT__.Destroy;
begin
  FRootAttributes.Free;
  FExternalLink.Free;
end;

procedure T__ROOT__.Clear;
begin
  FRootAttributes.Clear;
  FAssigneds := [];
  FExternalLink.Clear;
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

{ TXPGDocument }

function  TXPGDocXLink.GetExternalLink: TCT_ExternalLink;
begin
  Result := FRoot.ExternalLink;
end;

constructor TXPGDocXLink.Create;
begin
  FRoot := T__ROOT__.Create(Self);
  FErrors := TXpgPErrors.Create;
  FReader := TXPGReader.Create(FErrors,FRoot);
  FWriter := TXpgWriteXML.Create;
end;

destructor TXPGDocXLink.Destroy;
begin
  FRoot.Free;
  FReader.Free;
  FWriter.Free;
  FErrors.Free;
  inherited Destroy;
end;

procedure TXPGDocXLink.LoadFromFile(AFilename: AxUCString);
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

procedure TXPGDocXLink.LoadFromStream(AStream: TStream);
begin
  FReader.LoadFromStream(AStream);
end;

procedure TXPGDocXLink.SaveToFile(AFilename: AxUCString);
begin
  FWriter.SaveToFile(AFilename);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

procedure TXPGDocXLink.SaveToStream(AStream: TStream);
begin
  FWriter.SaveToStream(AStream);
  FRoot.CheckAssigned;
  FRoot.Write(FWriter);
end;

initialization
{$ifdef DELPHI_5}
  FEnums := TStringList.Create;
{$else}
  FEnums := THashedStringList.Create;
{$endif}
  TXPGBase.AddEnums;

finalization
  FEnums.Free;

end.
