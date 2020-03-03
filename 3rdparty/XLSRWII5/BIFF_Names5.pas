unit BIFF_Names5;

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
     BIFF_Utils5, BIFF_RecsII5, BIFF_Stream5, BIFF_EncodeFormulaII5, BIFF_DecodeFormula5,
     BIFF_ExcelFuncII5,
     Xc12Utils5,
     XLSUtils5, XLSFormulaTypes5;

//* ~exclude
const CRN_NUM =  $01;
//* ~exclude
const CRN_STR =  $02;
//* ~exclude
const CRN_BOOL = $04;
//* ~exclude
const CRN_ERR =  $10;

//* Built in names that defines special purpose areas.
type TX97BuiltInName = (xbnConsolidateArea, //* Consolidate area
                        xbnAutoOpen,        //* Auto open
                        xbnAutoClose,       //* Auto close
                        xbnExtract,         //* Extract
                        xbnDatabase,        //* Database
                        xbnCriteria,        //* Criteria
                        xbnPrintArea,       //* Print area
                        xbnPrintTitles,     //* Print titles
                        xbnRecorder,        //* Recorder
                        xbnDataForm,        //* Data form
                        xbnAutoActivate,    //* Auto activate
                        xbnAutoDeactivate,  //* Auto deactivate
                        xbnSheetTitle,      //* Sheet title
                        xbnFilterDatabase,  //* Filter database
                        xbnNone             //* No built in name.
                        );

//* ~exclude
type PCRNValue = ^TCRNValue;
     TCRNValue = record
     FirstCol,LastCol: byte;
     Row: word;
     Value: PByteArray;
     Size: word;
     end;

//* ~exclude
type TCRNList = class(TList)
private
     function GetItems(Index: integer): PCRNValue;
public
     destructor Destroy; override;
     procedure Add(Value: PRecCRN; Size: word);
     procedure WriteRecords(Stream: TXLSStream; Index: integer);

     property Items[Index: integer]: PCRNValue read GetItems; default;
     end;

//* ~exclude
type TSupBookSheet = class(TObject)
private
     FName: XLS8String;
     FCRN: TCRNList;
     FNegXCT: boolean;

     function GetName: AxUCString;
     procedure SetName(const Value: AxUCString);
protected
     property RawName: XLS8String read FName;
public
     constructor Create(Name: PByteArray; NameLen: integer); overload;
     constructor Create(Sheetname: AxUCString); overload;
     destructor Destroy; override;
     procedure SetCRN(P: PRecCRN; Size: word);
     function GetCachedValue(Col,Row: integer): TFormulaValue;

     property Name: AxUCString read GetName write SetName;
     property CRN: TCRNList read FCRN;
     property NegXCT: boolean read FNegXCT write FNegXCT;
     end;

//* ~exclude
type TExtNameType = (entName,entDDE,entOLE);

//* ~exclude
type TExtName = class(TObject)
private
     FOptions: word;
     FName: XLS8String;

     function GetName: AxUCString;
public
     constructor Create(var P: PRecEXTERNNAME8);
     procedure WriteRecords(Stream: TXLSStream); virtual; abstract;
     function NameType: TExtNameType; virtual; abstract;
     property Name: AxUCString read GetName;
     end;

//* ~exclude
type TExtNameName = class(TExtName)
private
     FNameDef: TDynByteArray;
public
     constructor Create(P: PRecEXTERNNAME8);
     destructor Destroy; override;
     procedure WriteRecords(Stream: TXLSStream); override;
     function NameType: TExtNameType; override;

     property NameDef: TDynByteArray read FNameDef;
     end;

//* ~exclude
type TExtNameDDE = class(TExtName)
private
     FOPER: PByteArray;
public
     constructor Create(P: PRecEXTERNNAME8);
     destructor Destroy; override;
     procedure WriteRecords(Stream: TXLSStream); override;
     function NameType: TExtNameType; override;
     end;

//* ~exclude
type TExtNameOLE = class(TExtName)
private
     FOLE2Id: longword;
public
     constructor Create(P: PRecEXTERNNAME8);
     function NameType: TExtNameType; override;
     procedure WriteRecords(Stream: TXLSStream); override;
     property OLE2Id: longword read FOLE2Id;
     end;

//* ~exclude
type TExtNameList = class(TList)
private
     function GetItems(Index: integer): TExtName;
public
     destructor Destroy; override;
     procedure Add(P: PRecEXTERNNAME8);

     property Items[Index: integer]: TExtName read GetItems; default;
     end;

//* ~exclude
type TSupBook = class(TObject)
private
     FTabCount: integer;
     FEncoded: word;
     // Do not change to AxUCString.
     FFilename: XLS8String;
     FSheets: array of TSupBookSheet;
     FExtNames: TExtNameList;

     function GetSheets(Index: integer): TSupBookSheet;
     function GetExtNames(Index: integer): TExtName;
     function GetFilename: string;
     function GetCleanFilename: AxUCString;
public
     constructor Create(P: PRecSUPBOOK);
     constructor CreateEncoded(Tabs: integer; Code: word);
     destructor Destroy; override;
     procedure WriteRecords(Stream: TXLSStream);
     function  Count: integer;
     procedure SetEXTERNNAME(P: PRecEXTERNNAME8);
     procedure AddSheet(Sheetname: AxUCString);

     property Filename: string read GetFilename;
     property CleanFilename: AxUCString read GetCleanFilename;
     property Sheets[Index: integer]: TSupBookSheet read GetSheets; default;
     property ExtNames[Index: integer]: TExtName read GetExtNames;
     end;

//* ~exclude
type TSupBookList = class(TList)
private
     function GetItems(Index: integer): TSupBook;
public
     destructor Destroy; override;
     procedure Clear; override;
     procedure Add(P: PRecSUPBOOK); overload;
     procedure Add(Path,Filename,SheetName: AxUCString); overload;
     function AddEncodec(Tabs: integer; Code: word): integer;

     property Items[Index: integer]: TSupBook read GetItems; default;
     end;

//* ~exclude
type PExtSheetData = ^TExtSheetData;
     TExtSheetData = record
     SupBookIndex: integer;
     FirstTab,LastTab: word;
     end;

//* ~exclude
type TExtSheetList = class(TList)
private
     function GetItems(Index: integer): PExtSheetData;
public
     destructor Destroy; override;
     procedure Clear; override;
     procedure Add(Index,First,Last: integer);
     procedure WriteRecords(Stream: TXLSStream);

     property Items[Index: integer]: PExtSheetData read GetItems; default;
     end;

//* ~exclude
type TExternalNames = class(TObject)
private
     FExtSheets: TExtSheetList;
     FSupBooks: TSupBookList;
     FFilePath: AxUCString;

     function GetAsString(SheetIndex, NameIndex: integer): AxUCString;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     procedure UpdateIntSupbooks(TabCount: integer);
     function  IsSelf(ExtSheetIndex: integer): integer;
     function  AddSelf(SheetIndex,TabCount: integer): integer;
     function  FindIndex(SheetIndex: integer): integer;
     function  AddRef(Path,Filename,SheetName: AxUCString; var Index,NameIndex: integer): boolean;
     function  RefIndexByName(Path,Filename,Name: AxUCString; var Index,NameIndex: integer): boolean;
     function  NameIndexByName(Path,Filename,Name: AxUCString; var Index,NameIndex: integer): boolean;
     function  GetNameValue(SheetIndex, NameIndex: integer): TFormulaValue;
     function  GetValue(Index,Col,Row: integer): TFormulaValue;
     procedure DeleteSheet(SheetIndex: integer);
     procedure InsertSheet(SheetIndex: integer);
     function  Get3dSheets(const AIndex: integer; out ASheet1,ASheet2: integer): boolean;
     function  HasSUPBOOK: boolean;
     function  HasEXTERNSHEET: boolean;

     procedure SetSUPBOOK     (P: PRecSUPBOOK);
     procedure SetEXTERNSHEET (P: PByteArray);
     procedure SetCRN         (SheetIndex: integer; P: PRecCRN; Size: word);
     procedure SetEXTERNNAME  (P: PRecEXTERNNAME8);

     procedure WriteRecords   (Stream: TXLSStream);

     property AsString[SheetIndex,NameIndex: integer]: AxUCString read GetAsString;
     property FilePath: AxUCString read FFilePath write FFilePath;
     end;

type TNameOption = (noBuiltIn,noHidden,noFunction,noVBProcedure,noMacro,noXLM,
                    noPublishToServer,noWorkbookParameter,noComplex,noBig);
     TNameOptions = set of TNameOption;

type
//* Names that referrers to areas on the worksheets.
//* Use TInternalName to create a name that can be used in formulas.
//* TInternalName is also used to define source areas for built in names,
//* such as print areas.
    TInternalName = class(TObject)
private
     FParent: TObjectList;
     FKeyShortcut: byte;
     FUnused: word;
     FSheetIndex: word;
     // Do not change FName to AxUCString. Used to store coded names.
     FName: XLS8String;
     FNameDef: TDynByteArray;
     // Excel 2007
     FNameDefStr: AxUCString;
     FCustomMenu: AxUCString;
     FDescription: AxUCString;
     FHelpTopic: AxUCString;
     FStatusBar: AxUCString;
     FFormulaSize: integer;
     FLoadedFormFile: boolean;
     FComment: AxUCString;
     FNameOptions: TNameOptions;
     FFuncGroupId: integer;
     // Excel 2007
     FShortcutKey: AxUCString;

     function  GetIsVBProcedure: boolean;
     procedure SetIsVBProcedure(const Value: boolean);
     procedure IntSetDefinition(const Value: AxUCString);
     procedure SetDefinition(const Value: AxUCString);
     function  GetDefinition: AxUCString;
     function  GetName: AxUCString;
     procedure SetName(const Value: AxUCString);
     function  GetBuiltInName: TX97BuiltInName;
     procedure SetBuiltInName(const Value: TX97BuiltInName);
     function  GetCol1: integer;
     function  GetCol2: integer;
     function  GetRow1: integer;
     function  GetRow2: integer;
     function  GetExtSheet: integer;
     procedure SetRow1(const Value: integer);
     procedure SetRow2(const Value: integer);
     procedure SetCol1(const Value: integer);
     procedure SetCol2(const Value: integer);
public
     //* ~exclude
     constructor Create(Parent: TObjectList);
     //* ~exclude
     destructor Destroy; override;
     //* ~exclude
     procedure Set_NAME(P: PRecNAME);
     //* ~exclude
     procedure WriteRecords(Stream: TXLSStream);
     //* Checks if the name is a single cell.
     //* ~result True if the name is a single cell.
     function  NameIsCell: boolean;
     //* ~exclude
     function  NameIsArea(ListAccepted: boolean): boolean;

     //* ~exclude
     procedure SetError;
     //* ~exclude
     property NameDef: TDynByteArray read FNameDef write FNameDef;
     //* ~exclude
     property NameDefLen: integer read FFormulaSize;
     //* ~exclude
     property KeyShortcut: byte read FKeyShortcut write FKeyShortcut;
     //* ~exclude
     property SheetIndex: word read FSheetIndex write FSheetIndex;
     //* ~exclude
     property NameName: XLS8String read FName write FName;
     //* ~exclude
     property NameOptions: TNameOptions read FNameOptions write FNameOptions;
     //* Custom menu text.
     property CustomMenu: AxUCString read FCustomMenu write FCustomMenu;
     //* Description.
     property Description: AxUCString read FDescription write FDescription;
     //* Help topic.
     property HelpTopic: AxUCString read FHelpTopic write FHelpTopic;
     //* Status bar text.
     property StatusBar: AxUCString read FStatusBar write FStatusBar;

     property Comment: AxUCString read FComment write FComment;
     //* ~exclude
     property NameDefStr: AxUCString read FNameDefStr write FNameDefStr;
     //* ~exclude
     property FuncGroupId: integer read FFuncGroupId write FFuncGroupId;
     //* ~exclude
     property ShortcutKey: AxUCString read FShortcutKey write FShortcutKey;
//     property LoadedFormFile: boolean read FLoadedFormFile;
public
     //* The name of the internal name.
     property Name: AxUCString read GetName write SetName stored False;
     //* The definition of the name.
     //* The definition is in the same format as a formula, such as Sheet1!$D$4:$G$10
     //* It is important that cell references are absoult, otherwise there may be undesired consequences.
     property Definition: AxUCString read GetDefinition write SetDefinition stored False;
     //* If the name is a built in name.
     //* Use BuiltInName to set the type of built in name.
     property BuiltInName: TX97BuiltInName read GetBuiltInName write SetBuiltInName stored False;
     //* The left column of the name area.
     property Col1: integer read GetCol1  write SetCol1;
     //* The right column of the name area.
     property Col2: integer read GetCol2  write SetCol2;
     //* The top row of the name area.
     property Row1: integer read GetRow1 write SetRow1;
     //* The bottom row of the name area.
     property Row2: integer read GetRow2 write SetRow2;
     //* ~exclude
     property ExtSheet: integer read GetExtSheet;
     //* ~exclude
     property IsVBProcedure: boolean read GetIsVBProcedure write SetIsVBProcedure;
     end;

type
//* List of internal names.
     TInternalNames = class(TObjectList)
private
     FEncoder: TEncodeFormula;
     function GetItems(Index: integer): TInternalName;
protected
     FOwner: TPersistent;
     FGetNameMethod: TGetNameEvent;
     FNameDeletedEvent: TIntegerEvent;
     // FIsUpdating is by default True, in order to prevent sorting of the names.
     // Names can be sorted, as long as their index not is used. Excel will
     // however load files with sorted names *MUCH* faster. This may be of
     // interrest if there are 10000+ names in the file...
     FIsUpdating: boolean;

     function  FindUniqueName: string;
     procedure IntDeleteBuiltIn(BuiltInName: TX97BuiltInName; SheetIndex: integer; ExcludeItem: TInternalName);
public
     //* ~exclude
     constructor Create(AOwner: TPersistent; GetNameMethod: TGetNameEvent; Encoder: TEncodeFormula);
     //* ~exclude
     destructor Destroy; override;
     //* Add a new TInternalName object.
     //* ~result The new TInternalName object.
     function  Add: TInternalName; overload;
     function  Add(AName,ADefinition: AxUCString): TInternalName; overload;
     //* Deletes a name from the list.
     //* ~param Index Index of the name to delete.
     procedure Delete(Index: Integer);
     //* Find an internal name.
     //* Searches the list of names for the name given by AName. The search is
     //* case incensitive. If the name is found, FindName returnes the index of
     //* it. When not found, -1 is returned.
     //* ~param AName The name to search for.
     //* ~result Index of the found name. If not found, -1 is returned.
     function  FindName(AName: AxUCString): integer;
     //* ~exclude
     function  FindBuiltInName(BuiltInName: TX97BuiltInName; SheetIndex: integer): TInternalName;
     //* ~exclude
     function  AddBuiltInName(ABuiltInName: TX97BuiltInName; AUnique, AHidden: boolean; ASheetName: AxUCString; ASheetIndex,C1,R1,C2,R2: integer): TInternalName;
     //* ~exclude
     procedure DeleteBuiltInName(BuiltInName: TX97BuiltInName; SheetIndex: integer);
     //* ~exclude
     procedure SetNAME(P: PRecNAME);
     //* ~exclude
     procedure WriteRecords(Stream: TXLSStream);
     procedure BeginUpdate;
     procedure EndUpdate;

     //* The TInternalName items in the list.
     property Items[Index: integer]: TInternalName read GetItems; default;
     //* ~exclude
     property OnNameDeleted: TIntegerEvent read FNameDeletedEvent write FNameDeletedEvent;
     end;

implementation

{$ifdef ver130}
type PDouble = ^Double;
type PWord = ^word;
{$endif}

{ TExtSheetList }

procedure TExtSheetList.Add(Index, First, Last: integer);
var
  P: PExtSheetData;
begin
{
  if Index > Count then
    raise XLSRWException.Create('Invalid sheet index in EXTSHEET');
}
  New(P);
  P.SupBookIndex := Index;
  P.FirstTab := First;
  P.LastTab := Last;
  inherited Add(P);
end;

procedure TExtSheetList.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FreeMem(Items[i]);
  inherited Clear;
end;

destructor TExtSheetList.Destroy;
begin
  Clear;
  inherited;
end;

function TExtSheetList.GetItems(Index: integer): PExtSheetData;
begin
  Result := PExtSheetData(inherited Items[Index]);
end;

procedure TExtSheetList.WriteRecords(Stream: TXLSStream);
var
  i: integer;
begin
  if Count > 0 then begin
    Stream.WriteHeader(BIFFRECID_EXTERNSHEET,2 + Count * 6);
    Stream.WWord(Count);
    for i := 0 to Count - 1 do begin
      Stream.WWord(Items[i].SupBookIndex);
      Stream.WWord(Items[i].FirstTab);
      Stream.WWord(Items[i].LastTab);
    end;
  end;
end;

{ TExternalNames }

function TExternalNames.NameIndexByName(Path, Filename,Name: AxUCString; var Index, NameIndex: integer): boolean;
var
  i,j,SupBook: integer;
  S: AxUCString;
begin
  S := MyWideUppercase(Path + Filename);
  Name := MyWideUppercase(Name);
  for i := 0 to FExtSheets.Count - 1 do begin
    SupBook := FExtSheets[i].SupBookIndex;
    if MyWideUppercase(FSupBooks[SupBook].CleanFilename) = S then begin
//      for j := 0 to FSupBooks[SupBook].Sheets .FExtNames.Count - 1 do begin
      for j := 0 to FSupBooks[SupBook].FExtNames.Count - 1 do begin
        if MyWideUppercase(FSupBooks[SupBook].FExtNames[j].Name) = Name then begin
          Index := i;
          NameIndex := j;
          Result := True;
          Exit;
        end;
      end;
    end;
  end;
  Result := False;
end;

function TExternalNames.RefIndexByName(Path, Filename,Name: AxUCString; var Index, NameIndex: integer): boolean;
var
  i,j,SupBook: integer;
  S: AxUCString;
begin
  S := MyWideUppercase(Path + Filename);
  Name := MyWideUppercase(Name);
  for i := 0 to FExtSheets.Count - 1 do begin
    SupBook := FExtSheets[i].SupBookIndex;
    if MyWideUppercase(FSupBooks[SupBook].CleanFilename) = S then begin
      for j := 0 to FSupBooks[SupBook].Count - 1 do begin
        if MyWideUppercase(FSupBooks[SupBook].Sheets[j].Name) = Name then begin
          Index := i;
          NameIndex := j;
          Result := True;
          Exit;
        end;
      end;
    end;
  end;
  Result := False;
end;

function TExternalNames.AddRef(Path, Filename,SheetName: AxUCString; var Index,NameIndex: integer): boolean;
var
  i,j,SupBook: integer;
  S1,S2: AxUCString;
begin
  S1 := MyWideUppercase(Path + Filename);
  S2 := MyWideUppercase(SheetName);
  Result := False;
  for i := 0 to FExtSheets.Count - 1 do begin
    SupBook := FExtSheets[i].SupBookIndex;
    if MyWideUppercase(FSupBooks[SupBook].Filename) = S1 then begin
      for j := 0 to FSupBooks[SupBook].Count - 1 do begin
        if MyWideUppercase(FSupBooks[SupBook].Sheets[j].Name) = S2 then begin
          Index := i;
          Result := True;
          Exit;
        end;
      end;
    end;
  end;
  FSupBooks.Add(Path,Filename,Sheetname);
  NameIndex := FSupBooks.Count - 1;
  FExtSheets.Add(NameIndex,0,0);
  Index := FExtSheets.Count - 1;
end;

function TExternalNames.AddSelf(SheetIndex,TabCount: integer): integer;
var
  i,SelfIndex: integer;
begin
  SelfIndex := -1;
  for Result := 0 to FExtSheets.Count - 1 do begin
    i := FExtSheets[Result].SupBookIndex;
    if FSupBooks[i].FEncoded = $0401 then begin
      SelfIndex := i;
      if (SheetIndex <= FExtSheets[Result].FirstTab) and (SheetIndex >= FExtSheets[Result].LastTab) then
        Exit;
    end;
  end;
  if SelfIndex < 0 then
    SelfIndex := FSupBooks.AddEncodec(TabCount,$0401);
  FExtSheets.Add(SelfIndex,SheetIndex,SheetIndex);
  Result := FExtSheets.Count - 1;
end;

procedure TExternalNames.Clear;
begin
  FExtSheets.Clear;
  FSupBooks.Clear;
end;

constructor TExternalNames.Create;
begin
  FExtSheets := TExtSheetList.Create;
  FSupBooks := TSupBookList.Create;
end;

destructor TExternalNames.Destroy;
begin
  FExtSheets.Free;
  FSupBooks.Free;
  inherited;
end;

function TExternalNames.FindIndex(SheetIndex: integer): integer;
var
  i: integer;
begin
  for Result := 0 to FExtSheets.Count - 1 do begin
    i := FExtSheets[Result].SupBookIndex;
    if FSupBooks[i].FEncoded = $0401 then begin
      if (SheetIndex <= FExtSheets[Result].FirstTab) and (SheetIndex >= FExtSheets[Result].LastTab) then
        Exit;
    end;
  end;
  Result := -1;
end;

function TExternalNames.Get3dSheets(const AIndex: integer; out ASheet1, ASheet2: integer): boolean;
begin
  ASheet1 := FExtSheets[AIndex].FirstTab;
  ASheet2 := FExtSheets[AIndex].LastTab;
  Result := ASheet1 = ASheet2;
end;

function TExternalNames.GetAsString(SheetIndex, NameIndex: integer): AxUCString;
var
  i,j: integer;
  S: AxUCString;
begin
  i := FExtSheets[SheetIndex].SupBookIndex;
  Result := FSupBooks[i].Filename;
  if NameIndex >= 0 then
    Result := Result + FSupBooks[i].ExtNames[NameIndex - 1].Name
  else begin
    j := FExtSheets[SheetIndex].FirstTab;
    if (j >= 0) and (j <= $FF) then
      S := FSupBooks[i].Sheets[j].Name;
    Result := '''' + Result + S + '''!';
  end;
end;

function TExternalNames.GetNameValue(SheetIndex, NameIndex: integer): TFormulaValue;
var
  SupBook: integer;
  Ref: PPTGRef3d8;
  Name: TExtNameName;
begin
  if (SheetIndex < 0) or (SheetIndex >= FExtSheets.Count) then
    raise XLSRWException.Create('Extrensheet index out of range.');
  if (FExtSheets[SheetIndex].FirstTab <= $00FF) or (FExtSheets[SheetIndex].LastTab <= $00FF) then
    raise XLSRWException.Create('Extern name is a reference.');
  SupBook := FExtSheets[SheetIndex].SupBookIndex;
  case FSupBooks[SupBook].FExtNames[NameIndex - 1].NameType of
    entName : begin
      Name := TExtNameName(FSupBooks[SupBook].FExtNames[NameIndex - 1]);
      if (Length(Name.NameDef) > 0) and (Name.NameDef[0] = xptgRef3d97) then begin
        // There is a dog burried here... PPTGRef3d8 shall follows imidiatelly after ptgRef3d
        // but there are two extra bytes, that seems to repeat the index.
        Ref := PPTGRef3d8(@Name.NameDef[3]);
        Result := FSupBooks[SupBook].Sheets[Ref.Index].GetCachedValue(Ref.Col and $3FFF,Ref.Row);
      end
      else begin
        FVSetString(Result, Name.Name);
//        FVSetError(Result,errName);
//        raise XLSRWException.Create('Name def error for external value.');
      end;
    end;
    entDDE:
      raise XLSRWException.Create('Can not access DDE values.');
    entOLE:
      raise XLSRWException.Create('Can not access OLE values.');
  end;
end;

function TExternalNames.GetValue(Index, Col, Row: integer): TFormulaValue;
var
  SupBook: integer;
begin
  if (FExtSheets[Index].FirstTab > $00FF) or (FExtSheets[Index].LastTab > $00FF) then
    raise XLSRWException.Create('Extern name is a reference.');
  SupBook := FExtSheets[Index].SupBookIndex;
  Result := FSupBooks[SupBook].Sheets[FExtSheets[Index].FirstTab].GetCachedValue(Col and $3FFF,Row);
end;

function TExternalNames.HasEXTERNSHEET: boolean;
begin
  Result := FExtSheets.Count > 0;
end;

function TExternalNames.HasSUPBOOK: boolean;
begin
  Result := FSupbooks.Count > 0;
end;

function TExternalNames.IsSelf(ExtSheetIndex: integer): integer;
var
  i: integer;
begin
  i := FExtSheets[ExtSheetIndex].SupBookIndex;
  if (i >= 0) and (FSupBooks[i].FEncoded = $0401) then
    Result := FExtSheets[ExtSheetIndex].FirstTab
  else
    Result := -1;
end;

procedure TExternalNames.DeleteSheet(SheetIndex: integer);
var
  i,j: integer;
begin
  for i := 0 to FExtSheets.Count - 1 do begin
    j := FExtSheets[i].SupBookIndex;
    if (FSupBooks[j].FEncoded = $0401) and (FExtSheets[i].FirstTab >= SheetIndex) then begin
      if FExtSheets[i].FirstTab = SheetIndex then
        FExtSheets[i].FirstTab := $FFFF
      else if FExtSheets[i].FirstTab > SheetIndex then
        FExtSheets[i].FirstTab := FExtSheets[i].FirstTab - 1;

      if FExtSheets[i].LastTab = SheetIndex then
        FExtSheets[i].LastTab := $FFFF
      else if FExtSheets[i].LastTab > SheetIndex then
        FExtSheets[i].LastTab := FExtSheets[i].LastTab - 1;
    end;
  end;
end;

procedure TExternalNames.InsertSheet(SheetIndex: integer);
var
  i,j: integer;
begin
  for i := 0 to FExtSheets.Count - 1 do begin
    j := FExtSheets[i].SupBookIndex;
    if (FSupBooks[j].FEncoded = $0401) and (FExtSheets[i].FirstTab >= SheetIndex) then begin
      if FExtSheets[i].FirstTab < $F000 then
        FExtSheets[i].FirstTab := FExtSheets[i].FirstTab + 1;

      if (FExtSheets[i].LastTab >= SheetIndex) and (FExtSheets[i].LastTab < $F000) then
        FExtSheets[i].LastTab := FExtSheets[i].LastTab + 1;
    end;
  end;
end;

procedure TExternalNames.SetCRN(SheetIndex: integer; P: PRecCRN; Size: word);
begin
  if FSupBooks.Count <= 0 then
    raise XLSRWException.Create('No SUPBOOK for CRN');
  if SheetIndex >= FSupBooks[FSupBooks.Count - 1].Count then
    raise XLSRWException.Create('Invalid SUPBOOK Sheet Index');
  FSupBooks[FSupBooks.Count - 1].Sheets[SheetIndex].SetCRN(P,Size);
end;

procedure TExternalNames.SetEXTERNNAME(P: PRecEXTERNNAME8);
begin
  FSupBooks[FSupBooks.Count - 1].SetEXTERNNAME(P);
end;

procedure TExternalNames.SetEXTERNSHEET(P: PByteArray);
var
  i: integer;
begin
  for i := 0 to PRecEXTERNSHEET8(P).XTICount - 1 do
    FExtSheets.Add(PRecEXTERNSHEET8(P).XTI[i].SupBook,PRecEXTERNSHEET8(P).XTI[i].FirstTab,PRecEXTERNSHEET8(P).XTI[i].LastTab);
end;

procedure TExternalNames.SetSUPBOOK(P: PRecSUPBOOK);
begin
  FSupBooks.Add(P);
end;

procedure TExternalNames.UpdateIntSupbooks(TabCount: integer);
var
  i,j: integer;
begin
  if FExtSheets.Count > 0 then begin
    for i := 0 to FSupBooks.Count - 1 do begin
      j := FExtSheets[i].SupBookIndex;
      if FSupBooks[j].FEncoded = $0401 then
        FSupBooks[j].FTabCount := TabCount;
    end;
  end;
end;

procedure TExternalNames.WriteRecords(Stream: TXLSStream);
var
  i: integer;
begin
  for i := 0 to FSupBooks.Count - 1 do
    FSupBooks[i].WriteRecords(Stream);
  FExtSheets.WriteRecords(Stream);
end;

{ TCRNList }

procedure TCRNList.Add(Value: PRecCRN; Size: word);
var
  P: PCRNValue;
begin
  New(P);
  P.FirstCol := Value.FirstCol;
  P.LastCol := Value.LastCol;
  P.Row := Value.Row;
  P.Size := Size - 4;
  GetMem(P.Value,P.Size);
  System.Move(Pointer(@Value.OPER)^,P.Value^,P.Size);
  inherited Add(P);
end;

destructor TCRNList.Destroy;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    FreeMem(Items[i].Value);
    FreeMem(inherited Items[i]);
  end;
  inherited;
end;

function TCRNList.GetItems(Index: integer): PCRNValue;
begin
  Result := PCRNValue(inherited Items[Index]);
end;

procedure TCRNList.WriteRecords(Stream: TXLSStream; Index: integer);
var
  i: integer;
begin
  if Count > 0 then begin
    Stream.WriteHeader(BIFFRECID_XCT,SizeOf(TRecXCT));
    Stream.WWord(Count);
    Stream.WWord(Index);
    for i := 0 to Count - 1 do begin
      Stream.WriteHeader(BIFFRECID_CRN,Items[i].Size + 4);
      Stream.WByte(Items[i].LastCol);
      Stream.WByte(Items[i].FirstCol);
      Stream.WWord(Items[i].Row);
      Stream.Write(Items[i].Value^,Items[i].Size);
    end;
  end;
end;

{ TSupBookSheet }

constructor TSupBookSheet.Create(Name: PByteArray; NameLen: integer);
begin
  FName := ByteToCompressedStr(Name,NameLen);
  FCRN := TCRNList.Create;
end;

constructor TSupBookSheet.Create(Sheetname: AxUCString);
begin
  SetName(Sheetname);
  FCRN := TCRNList.Create;
end;

destructor TSupBookSheet.Destroy;
begin
  FCRN.Free;
  inherited;
end;

function TSupBookSheet.GetCachedValue(Col, Row: integer): TFormulaValue;
var
  i,L: integer;
  pD: PDouble;
begin
  for i := 0 to FCRN.Count - 1 do begin
    if (Row = FCRN[i].Row) and (Col >= FCRN[i].FirstCol) and (Col <= FCRN[i].LastCol) then begin
      case FCRN[i].Value[0] of
        $01: begin
          pD := @FCRN[i].Value[1];
          FVSetFloat(Result,pD^);
          Exit;
        end;
        $02: begin
          L := PWordArray(@FCRN[i].Value[1])[0];
          FVSetString(Result,ByteStrToWideString(@FCRN[i].Value[3],L));
          Exit;
        end;
        $04: begin
          FVSetBoolean(Result,FCRN[i].Value[1] <> 0);
          Exit;
        end;
        $10: begin
          FVSetError(Result,TXc12CellError(FCRN[i].Value[1]));
          Exit;
        end;
      end;
    end;
  end;
  FVSetError(Result,errValue);
end;

function TSupBookSheet.GetName: AxUCString;
begin
  Result := CompressedStrToWideString(FName);
end;

procedure TSupBookSheet.SetCRN(P: PRecCRN; Size: word);
begin
  FCRN.Add(P,Size);
end;

procedure TSupBookSheet.SetName(const Value: AxUCString);
begin
  SetLength(FName,Length(Value) * 2);
  Move(Pointer(Value)^,Pointer(FName)^,Length(Value) * 2);
  FName := #1 + FName;
end;

{ TSupBook }

procedure TSupBook.SetEXTERNNAME(P: PRecEXTERNNAME8);
begin
  FExtNames.Add(P);
end;

function TSupBook.Count: integer;
begin
  Result := Length(FSheets);
end;

constructor TSupBook.Create(P: PRecSUPBOOK);
var
  i,L: integer;
begin
  FExtNames := TExtNameList.Create;
  FTabCount := PWord(P)^;
  P := Pointer(NativeInt(P) + 2);
  if (PWordArray(P)[0] = $0401) or (PWordArray(P)[0] = $3A01) then begin
    FEncoded := PWordArray(P)[0];
    FFilename := '';
  end
  else begin
    SetLength(FSheets,FTabCount);
    L := PWord(P)^;
    P := Pointer(NativeInt(P) + 2);
    FFilename := ByteToCompressedStr(PByteArray(P),L);
    for i := 0 to FTabCount - 1 do begin
      if PByteArray(P)[0] = $01 then
        P := Pointer(NativeInt(P) + L * 2 + 1)
      else
        P := Pointer(NativeInt(P) + L + 1);
      L := PWord(P)^;
      P := Pointer(NativeInt(P) + 2);
      FSheets[i] := TSupBookSheet.Create(PByteArray(P),L);
    end;
  end;
end;

destructor TSupBook.Destroy;
var
  i: integer;
begin
  for i := 0 to High(FSheets) do
    FSheets[i].Free;
  SetLength(FSheets,0);
  FExtNames.Free;
  inherited;
end;

function TSupBook.GetSheets(Index: integer): TSupBookSheet;
begin
  Result := FSheets[Index];
end;

function TSupBook.GetExtNames(Index: integer): TExtName;
begin
  Result := FExtNames[Index];
end;

function TSupBook.GetFilename: string;
var
  UC,BracketAdded: boolean;
  i: integer;
  IsDDE: boolean;
begin
  UC := FFilename[1] = #1;
    IsDDE := Length(FSheets) <= 0;
  i := 2;
  if (Length(FFilename) > 2) and (FFilename[i] in [#00,#01,#02]) then
    Inc(i);
  BracketAdded := False;
  Result := '';
  while i <= Length(FFilename) do begin
    case FFilename[i] of
      #01: begin
        if FFilename[i + 1] = '@' then
          Result := Result + '\\'
        else
          Result := WideChar(FFilename[i + 1]) + Result + ':\';
        Inc(i);
      end;
      #02,#03: begin
        if IsDDE then
          Result := Result + '|'
        else begin
          Result := Result + '\[';
          BracketAdded := True;
        end;
      end;
      #04:
        Result := Result + '..\';
      #05,#06,#07,#08: ;
      else begin
        if UC then
          Result := Result + WideChar(Byte(FFilename[i]) + (Byte(FFilename[i + 1]) shl 8))
        else
          Result := Result + WideChar(FFilename[i]);
      end;
    end;
    if UC then
      Inc(i,2)
    else
      Inc(i);
  end;
  if IsDDE then
    Result := Result + '!'
  else if BracketAdded then
    Result := Result + ']';
end;

procedure TSupBook.WriteRecords(Stream: TXLSStream);
var
  Sz,i: integer;
begin
  if FEncoded <> 0 then begin
    Stream.WriteHeader(BIFFRECID_SUPBOOK,4);
    Stream.WWord(FTabCount);
    Stream.WWord(FEncoded);
  end
  else begin
    // Tabs + Length filename + Filename.
    Sz := 2 + 2 + Length(FFilename);
    for i := 0 to High(FSheets) do
      Inc(Sz,2 + Length(FSheets[i].FName));
    Stream.WriteHeader(BIFFRECID_SUPBOOK,Sz);
    Stream.WWord(FTabCount);
    Stream.WriteUnicodeStr16(FFilename);
    for i := 0 to High(FSheets) do
      Stream.WriteUnicodeStr16(FSheets[i].RawName);
  end;
  for i := 0 to FExtNames.Count - 1 do
    FExtNames[i].WriteRecords(Stream);
  for i := 0 to High(FSheets) do
    FSheets[i].CRN.WriteRecords(Stream,i);
end;

constructor TSupBook.CreateEncoded(Tabs: integer; Code: word);
begin
  FExtNames := TExtNameList.Create;
  FTabCount := Tabs;
  FEncoded := Code;
  FFilename := '';
end;

procedure TSupBook.AddSheet(Sheetname: AxUCString);
begin
  SetLength(FSheets,Length(FSheets) + 1);
  FSheets[High(FSheets)] := TSupBookSheet.Create(Sheetname);
end;

function TSupBook.GetCleanFilename: AxUCString;
var
  UC: boolean;
  i: integer;
begin
  UC := FFilename[1] = #1;
  i := 2;
  if (Length(FFilename) > 2) and (FFilename[i] in [#00,#01,#02]) then
    Inc(i);
  Result := '';
  while i <= Length(FFilename) do begin
    case FFilename[i] of
      #01: begin
        if FFilename[i + 1] = '@' then
          Result := Result + '\\'
        else
          Result := Char(FFilename[i + 1]) + Result + ':\';
        Inc(i);
      end;
      #02,#03: begin
        Result := Result + '\';
      end;
      #04:
        Result := Result + '..\';
      #05,#06,#07,#08: ;
      else begin
        if UC then
          Result := Result + WideChar(Byte(FFilename[i]) + (Byte(FFilename[i + 1]) shl 8))
        else
          Result := Result + Char(FFilename[i]);
      end;
    end;
    if UC then
      Inc(i,2)
    else
      Inc(i);
  end;
end;

{ TSupBookList }

procedure TSupBookList.Add(P: PRecSUPBOOK);
begin
  inherited Add(TSupBook.Create(P));
end;

procedure TSupBookList.Add(Path,Filename, SheetName: AxUCString);
var
  S1,S2: XLS8String;
  SupBook: TSupBook;
begin
  SupBook := TSupBook.CreateEncoded(1,0);

  if Copy(Path,Length(Path),1) = '\' then
    Path := Copy(Path,1,Length(Path) - 1);

  SetLength(S1,Length(Path) * 2);
  System.Move(Pointer(Path)^,Pointer(S1)^,Length(Path) * 2);

  SetLength(S2,Length(Filename) * 2);
  System.Move(Pointer(Filename)^,Pointer(S2)^,Length(Filename) * 2);

  SupBook.FFilename := #1 + S1 + #3 + #0 + S2;
  SupBook.AddSheet(Sheetname);
  inherited Add(SupBook);
end;

function TSupBookList.AddEncodec(Tabs: integer; Code: word): integer;
begin
  inherited Add(TSupBook.CreateEncoded(Tabs,Code));
  Result := Count - 1;
end;

procedure TSupBookList.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    TSupBook(inherited Items[i]).Free;
  inherited Clear;
end;

destructor TSupBookList.Destroy;
begin
  Clear;
  inherited;
end;

function TSupBookList.GetItems(Index: integer): TSupBook;
begin
  Result := TSupBook(inherited Items[Index]);
end;

{ TExtNameList }

procedure TExtNameList.Add(P: PRecEXTERNNAME8);
begin
  if (P.Options and $FFFE) = $0000 then
    inherited Add(TExtNameName.Create(P))
  else if (P.Options and $0010) = $0010 then
    inherited Add(TExtNameOLE.Create(P))
  else
    inherited Add(TExtNameDDE.Create(P));
end;

destructor TExtNameList.Destroy;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    TExtName(inherited Items[i]).Free;
  inherited;
end;

function TExtNameList.GetItems(Index: integer): TExtName;
begin
  Result := TExtName(inherited Items[Index]);
end;

{ TExtNameName }

constructor TExtNameName.Create(P: PRecEXTERNNAME8);
begin
  inherited Create(P);
  // After Create, P is no more a pointer to PRecEXTERNNAME8
  if PWordArray(P)[0] > 0 then begin
    SetLength(FNameDef,PWordArray(P)[0]);
    Move(PWordArray(P)[1],FNameDef[0],PWordArray(P)[0]);
  end;
end;

destructor TExtNameName.Destroy;
begin
  SetLength(FNameDef,0);
  inherited;
end;

function TExtNameName.NameType: TExtNameType;
begin
  Result := entName;
end;

procedure TExtNameName.WriteRecords(Stream: TXLSStream);
begin
  // Options(2) + Reserved(4) + Len name(1) + Name(x) + Len name def(2) + Name def(x)
  Stream.WriteHeader(BIFFRECID_EXTERNNAME,2 + 4 + 1 + Length(FName) + 2 + Length(FNameDef));
  Stream.WWord(FOptions);
  Stream.WWord(0);
  Stream.WWord(0);
  Stream.WriteUnicodeStr8(FName);
  Stream.WWord(Length(FNameDef));
  if Length(FNameDef) > 0 then
    Stream.Write(FNameDef[0],Length(FNameDef));
end;

{ TExtName }

constructor TExtName.Create(var P: PRecEXTERNNAME8);
begin
  FOptions := P.Options;
  FName := ByteToCompressedStr(@P.Data[0],P.LenName);
  if FName[1] = #1 then
    P := PRecEXTERNNAME8(NativeInt(P) + P.LenName * 2 + 8)
  else
    P := PRecEXTERNNAME8(NativeInt(P) + P.LenName + 8);
end;

function TExtName.GetName: AxUCString;
begin
  Result := CompressedStrToWideString(FName);
end;

{ TExtNameDDE }

constructor TExtNameDDE.Create(P: PRecEXTERNNAME8);
begin
  inherited Create(P);
end;

destructor TExtNameDDE.Destroy;
begin
  FreeMem(FOPER);
  inherited;
end;

function TExtNameDDE.NameType: TExtNameType;
begin
  Result := entDDE;
end;

procedure TExtNameDDE.WriteRecords(Stream: TXLSStream);
begin
  // Options(2) + Reserved(4) + Len name(1) + Name(x)
  Stream.WriteHeader(BIFFRECID_EXTERNNAME,2 + 4 + 1 + Length(FName));
  Stream.WWord(FOptions);
  Stream.WWord(0);
  Stream.WWord(0);
  Stream.WriteUnicodeStr8(FName);
end;

{ TExtNameOLE }

constructor TExtNameOLE.Create(P: PRecEXTERNNAME8);
begin
  inherited Create(P);
end;

function TExtNameOLE.NameType: TExtNameType;
begin
  Result := entOLE;
end;

procedure TExtNameOLE.WriteRecords(Stream: TXLSStream);
begin
  // Options(2) + Reserved(4) + Len name(1) + Name(x)
  Stream.WriteHeader(BIFFRECID_EXTERNNAME,2 + 4 + 1 + Length(FName));
  Stream.WWord(FOptions);
  Stream.Write(FOLE2Id,4);
  Stream.WriteUnicodeStr8(FName);
end;

{ TInternalNames }

function TInternalNameCompare(Item1, Item2: Pointer): Integer;
begin
{$ifdef DELPHI_5}
  Result := CompareStr(TInternalName(Item1).Name,TInternalName(Item2).Name);
{$else}
  Result := WideCompareStr(TInternalName(Item1).Name,TInternalName(Item2).Name);
{$endif}
end;

function TInternalNames.Add: TInternalName;
begin
  Result := TInternalName.Create(Self);
  inherited Add(Result);
end;

function TInternalNames.Add(AName, ADefinition: AxUCString): TInternalName;
begin
  Result := Add;
  Result.Name := AName;
  Result.Definition := ADefinition;
end;

function TInternalNames.AddBuiltInName(ABuiltInName: TX97BuiltInName; AUnique, AHidden: boolean; ASheetName: AxUCString; ASheetIndex,C1,R1,C2,R2: integer): TInternalName;
var
  N: TInternalName;

procedure SetData(N: TInternalName);
begin
  Result := N;
  with N do begin
    FNameOptions := FNameOptions + [noBuiltIn];
    FSheetIndex := ASheetIndex;
    FName := XLS8String(#0 + Char(ABuiltInName));
    IntSetDefinition('''' + ASheetName + '''!' + AreaToRefStr(C1,R1,C2,R2,True,True,True,True));
    if AHidden then
      FNameOptions := FNameOptions + [noHidden];
  end;
end;

begin
  Inc(ASheetIndex);
  if AUnique then begin
    N := FindBuiltInName(ABuiltInName,ASheetIndex);
    if N <> Nil then begin
      SetData(N);
      Exit;
    end;
  end;
  SetData(Add);
end;

procedure TInternalNames.BeginUpdate;
begin
  FIsUpdating := True;
end;

constructor TInternalNames.Create(AOwner: TPersistent; GetNameMethod: TGetNameEvent; Encoder: TEncodeFormula);
begin
  inherited Create;
  FIsUpdating := True;
  FGetNameMethod := GetNameMethod;
  FEncoder := Encoder;
end;

procedure TInternalNames.Delete(Index: Integer);
begin
  inherited Delete(Index);
  FNameDeletedEvent(Self,-1);
end;

procedure TInternalNames.DeleteBuiltInName(BuiltInName: TX97BuiltInName; SheetIndex: integer);
begin
  IntDeleteBuiltIn(BuiltInName,SheetIndex,Nil);
end;

procedure TInternalNames.IntDeleteBuiltIn(BuiltInName: TX97BuiltInName; SheetIndex: integer; ExcludeItem: TInternalName);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i] <> ExcludeItem) and (TX97BuiltInName(Items[i].FName[2]) = BuiltInName) and (Items[i].FSheetIndex = SheetIndex) then begin
      Delete(i);
      Exit;
    end;
  end;
end;

destructor TInternalNames.Destroy;
begin
  inherited;
end;

procedure TInternalNames.EndUpdate;
begin
  FIsUpdating := False;
  Sort(TInternalNameCompare)
end;

function TInternalNames.FindBuiltInName(BuiltInName: TX97BuiltInName; SheetIndex: integer): TInternalName;
var
  i: integer;
begin
  Inc(SheetIndex);
  for i := 0 to Count - 1 do begin
    if (TX97BuiltInName(Items[i].FName[2]) = BuiltInName) and (Items[i].FSheetIndex = SheetIndex) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil
end;

function TInternalNames.FindName(AName: AxUCString): integer;
var
  WS,WS2: AxUCString;
begin
{$ifdef ver130}
  WS := AnsiLowercase(AName);
{$else}
  WS := WideLowercase(AName);
{$endif}
  for Result := 0 to Count - 1 do begin
    WS2 := TInternalName(inherited Items[Result]).Name;
{$ifdef ver130}
    if AnsiLowercase(WS2) = WS then
{$else}
    if WideLowercase(WS2) = WS then
{$endif}
      Exit;
  end;
  Result := -1;
end;

function TInternalNames.FindUniqueName: string;
var
  N: integer;
begin
  N := Count + 1;
  repeat
    Result := #0 + 'Name' + IntToStr(N);
    Inc(N);
  until (FindName(Result) < 0);
end;

function TInternalNames.GetItems(Index: integer): TInternalName;
begin
  Result := TInternalName(inherited Items[Index]);
end;

procedure TInternalNames.SetNAME(P: PRecNAME);
begin
  Add.Set_NAME(P);
end;

procedure TInternalNames.WriteRecords(Stream: TXLSStream);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].WriteRecords(Stream);
end;

{ TInternalName }

constructor TInternalName.Create(Parent: TObjectList);
begin
  inherited Create;
  FParent := Parent;
  SetLength(FNameDef,0);
end;

destructor TInternalName.Destroy;
begin
  SetLength(FNameDef,0);
  inherited;
end;

function TInternalName.GetBuiltInName: TX97BuiltInName;
begin
  if (Length(FName) = 2) and (FName[2] in [#0..Char(Ord(bnNone) - 1)]) then
    Result := TX97BuiltInName(FName[2])
  else
    Result := xbnNone;
end;

function TInternalName.GetDefinition: AxUCString;
var
  P: PByteArray;
begin
  if Length(FNameDef) > 0 then begin
    GetMem(P,512);
    try
      Move(FNameDef[0],P^,Length(FNameDef));
      Result := DecodeFmla(P,FFormulaSize,0,0,0,TInternalNames(FParent).FGetNameMethod,FuncArgSeparator)
    finally
      FreeMem(P);
    end;
  end
  else
    Result := '';
end;

function TInternalName.GetName: AxUCString;
begin
  if GetBuiltInName <> xbnNone then
    Result := ''
  else begin
    if FName[1] = #0 then
      Result := Copy(WideString(FName),2,MAXINT)
    else begin
      SetLength(Result,(Length(FName) - 1) div 2);
      Move(FName[2],Pointer(Result)^,Length(FName) - 1);
    end;
  end;
end;

procedure TInternalName.SetBuiltInName(const Value: TX97BuiltInName);
begin
  if Value < xbnNone then begin
    FNameOptions := FNameOptions + [noBuiltIn];
    FName := XLS8String(#0 + Char(Value));
  end
  else
    FNameOptions := FNameOptions - [noBuiltIn];
end;

procedure TInternalName.SetDefinition(const Value: AxUCString);
begin
  if Length(FName) < 2 then
    raise XLSRWException.Create('Name must be set before definition.');
  IntSetDefinition(Value);
  if noBuiltIn in FNameOptions then begin
    case TX97BuiltInName(FName[2]) of
      xbnPrintTitles,xbnPrintArea: begin

        if not NameIsArea(True) then
          raise XLSRWException.Create('Name definition must be an area');

        FSheetIndex := TInternalNames(FParent).FEncoder.LastSheetIndex + 1;
        TInternalNames(FParent).IntDeleteBuiltIn(TX97BuiltInName(FName[2]),FSheetIndex,Self);
      end;
    end;
  end
end;

procedure TInternalName.SetName(const Value: AxUCString);
var
  S: AxUCString;
begin
  S := Value;
  if Copy(S,1,6) = '_xlnm.' then begin
    S := Copy(S,7,MAXINT);
    FNameOptions := FNameOptions + [noBuiltIn];
  end;

  if Length(S) < 1 then
    raise XLSRWException.Create('Name is empty');

  if not TInternalNames(FParent).FIsUpdating then begin
    if TInternalNames(FParent).FindName(S) >= 0 then
      raise XLSRWException.Create('There is allready an area named ' + S);
  end;

  SetLength(FName,Length(S) * 2);
  Move(Pointer(S)^,Pointer(FName)^,Length(S) * 2);
  FName := #1 + FName;
  if not TInternalNames(FParent).FIsUpdating then
    TInternalNames(FParent).Sort(TInternalNameCompare);
end;

procedure TInternalName.Set_NAME(P: PRecNAME);
var
  Ptr: PByteArray;
  FormulaSz: integer;

function GetString(var P: PByteArray; Len: integer): XLS8String;
var
  Sz: integer;
begin
  if P[0] = 0 then
    Sz := Len + 1
  else
    Sz := Len * 2 + 1;
  SetLength(Result,Sz);
  Move(P^,Pointer(Result)^,Sz);
  P := PByteArray(NativeInt(P) + Sz);
end;

function GetWideString(var P: PByteArray; Len: integer): AxUCString;
begin
  Result := ByteToWideString(P,Len);
  if Result <> '' then begin
    if P[0] = 0 then
      P := PByteArray(NativeInt(P) + Len + 1)
    else
      P := PByteArray(NativeInt(P) + (Len * 2) + 1);
  end;
end;

begin
  FNameOptions := [];
  if ((P.Options and $0001) <> 0) then FNameOptions := FNameOptions + [noHidden];
  if ((P.Options and $0002) <> 0) then FNameOptions := FNameOptions + [noFunction];
  if ((P.Options and $0004) <> 0) then FNameOptions := FNameOptions + [noVBProcedure];
  if ((P.Options and $0008) <> 0) then FNameOptions := FNameOptions + [noMacro];
  if ((P.Options and $0010) <> 0) then FNameOptions := FNameOptions + [noComplex];
  if ((P.Options and $0020) <> 0) then FNameOptions := FNameOptions + [noBuiltIn];
  if ((P.Options and $1000) <> 0) then FNameOptions := FNameOptions + [noBig];

  FFuncGroupId := (P.Options and $0FC0) shr 5;

  FKeyShortcut := P.KeyShortcut;
  FSheetIndex := P.LocalSheetIndex;
  FUnused := P.Unused;
  Ptr := @P.Data;
  FName := GetString(Ptr,P.LenName);
  FFormulaSize := P.LenNameDef;
  raise XLSRWException.Create('TODO Get size of expression');
  FormulaSz := 0;
//  FormulaSz := P.LenNameDef + GetExFormulaSize(Ptr,P.LenNameDef);
  SetLength(FNameDef,FormulaSz);
  Move(Ptr^,FNameDef[0],FormulaSz);
  Ptr := PByteArray(NativeInt(Ptr) + FormulaSz);
  if P.LenCustMenu > 0   then FCustomMenu  := GetWideString(Ptr,P.LenCustMenu);
  if P.LenDescText > 0   then FDescription := GetWideString(Ptr,P.LenDescText);
  if P.LenHelpText > 0   then FHelpTopic   := GetWideString(Ptr,P.LenHelpText);
  if P.LenStatusText > 0 then FStatusBar   := GetWideString(Ptr,P.LenStatusText);
  FLoadedFormFile := True;
end;

procedure TInternalName.WriteRecords(Stream: TXLSStream);
var
  Sz: integer;
  Opts: word;

function FieldLength(S: AxUCString): integer;
begin
  if S <> '' then
    Result := 1 + (Length(S) * 2)
  else
    Result := 0;
end;

begin
  if not FLoadedFormFile and (FName = '') {or (Length(FNameDef) <= 0)} then Exit;
  Sz := SizeOf(TRecNAME) - 256 +
        Length(FName) +
        Length(FNameDef) +
        FieldLength(FCustomMenu) +
        FieldLength(FDescription) +
        FieldLength(FHelpTopic) +
        FieldLength(FStatusBar);
  Stream.WriteHeader(BIFFRECID_NAME,Sz);

  Opts := $0000;
  if noHidden      in FNameOptions then Opts := Opts or $0001;
  if noFunction    in FNameOptions then Opts := Opts or $0002;
  if noVBProcedure in FNameOptions then Opts := Opts or $0004;
  if noMacro       in FNameOptions then Opts := Opts or $0008;
  if noComplex     in FNameOptions then Opts := Opts or $0010;
  if noBuiltIn     in FNameOptions then Opts := Opts or $0020;
  if noBig         in FNameOptions then Opts := Opts or $1000;

  Opts := Opts or (FFuncGroupId shl 5);
  Stream.WWord(Opts);

  Stream.WByte(FKeyShortcut);
  Stream.WByte(CompressedStringLen(FName));
  Stream.WWord(FFormulaSize);
  Stream.WWord(FUnused);
  Stream.WWord(FSheetIndex);
  Stream.WByte(Length(FCustomMenu));
  Stream.WByte(Length(FDescription));
  Stream.WByte(Length(FHelpTopic));
  Stream.WByte(Length(FStatusBar));
  Stream.Write(Pointer(FName)^,Length(FName));
  if Length(FNameDef) > 0 then
    Stream.Write(FNameDef[0],Length(FNameDef));
  if FCustomMenu  <> '' then  Stream.WriteWideString(FCustomMenu,False);
  if FDescription <> '' then  Stream.WriteWideString(FDescription,False);
  if FHelpTopic   <> '' then  Stream.WriteWideString(FHelpTopic,False);
  if FStatusBar   <> '' then  Stream.WriteWideString(FStatusBar,False);
end;

function TInternalName.NameIsArea(ListAccepted: boolean): boolean;
begin
  if Length(FNameDef) > 0 then begin
    Result := (Length(FNameDef) = (1 + SizeOf(TPTGArea3d8))) and (FNameDef[0] in [xptgArea3d97,xptgArea3dV97,xptgArea3dA97]);
    if not Result and ListAccepted then
      //
      Result := (Length(FNameDef) = ((1 + SizeOf(TPTGArea3d8)) * 2 + 1)) and (FNameDef[0] = xptgArea3d97) and (FNameDef[1 + SizeOf(TPTGArea3d8)] = xptgArea3d97);
  end
  else
    Result := False;
end;

function TInternalName.NameIsCell: boolean;
begin
  Result := (Length(FNameDef) = (1 + SizeOf(TPTGRef3d8))) and (FNameDef[0] = xptgRef3d97);
end;

// TODO: None of these will work with relative refs.
function TInternalName.GetCol1: integer;
begin
  if NameIsArea(False) then
    Result := PPTGArea3d8(@FNameDef[1]).Col1
  else if NameIsCell then
    Result := PPTGRef3d8(@FNameDef[1]).Col
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

function TInternalName.GetCol2: integer;
begin
  if NameIsArea(False) then
    Result := PPTGArea3d8(@FNameDef[1]).Col2
  else if NameIsCell then
    Result := PPTGRef3d8(@FNameDef[1]).Col
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

function TInternalName.GetRow1: integer;
begin
  if NameIsArea(False) then
    Result := PPTGArea3d8(@FNameDef[1]).Row1
  else if NameIsCell then
    Result := PPTGRef3d8(@FNameDef[1]).Row
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

function TInternalName.GetRow2: integer;
begin
  if NameIsArea(False) then
    Result := PPTGArea3d8(@FNameDef[1]).Row2
  else if NameIsCell then
    Result := PPTGRef3d8(@FNameDef[1]).Row
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

procedure TInternalName.IntSetDefinition(const Value: AxUCString);
var
  P: PbyteArray;
  Sz: integer;
begin
  GetMem(P,512);
  try
    Sz := TInternalNames(FParent).FEncoder.Encode(Value,TokenClassRef,P,512,fvExcel97);
    SetLength(FNameDef,Sz);
    Move(P^,FNameDef[0],Sz);
    FFormulaSize := Sz;
  finally
    FreeMem(P);
  end;
end;

function TInternalName.GetExtSheet: integer;
begin
  if NameIsArea(False) then
    Result := PPTGArea3d8(@FNameDef[1]).Index
  else if NameIsCell then
    Result := PPTGRef3d8(@FNameDef[1]).Index
  else
    Result := -1;
end;

function TInternalName.GetIsVBProcedure: boolean;
begin
  Result := noVBProcedure in FNameOptions;
end;

procedure TInternalName.SetRow1(const Value: integer);
begin
  if NameIsArea(False) then
    PPTGArea3d8(@FNameDef[1]).Row1 := Value
  else if NameIsCell then
    PPTGRef3d8(@FNameDef[1]).Row := Value
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

procedure TInternalName.SetRow2(const Value: integer);
begin
  if NameIsArea(False) then
    PPTGArea3d8(@FNameDef[1]).Row2 := Value
  else if NameIsCell then
    PPTGRef3d8(@FNameDef[1]).Row := Value
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

procedure TInternalName.SetCol1(const Value: integer);
begin
  if NameIsArea(False) then
    PPTGArea3d8(@FNameDef[1]).Col1 := Value
  else if NameIsCell then
    PPTGRef3d8(@FNameDef[1]).Col := Value
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

procedure TInternalName.SetCol2(const Value: integer);
begin
  if NameIsArea(False) then
    PPTGArea3d8(@FNameDef[1]).Col2 := Value
  else if NameIsCell then
    PPTGRef3d8(@FNameDef[1]).Col := Value
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

procedure TInternalName.SetError;
begin
  if NameIsArea(False) then
    FNameDef[0] := xptgAreaErr3d97
  else if NameIsCell then
    FNameDef[0] := xptgRefErr3d97
  else
    raise XLSRWException.Create('Name is not of requested type.');
end;

procedure TInternalName.SetIsVBProcedure(const Value: boolean);
begin
  if Value then
    FNameOptions := FNameOptions + [noVBProcedure]
  else
    FNameOptions := FNameOptions - [noVBProcedure];
end;

end.
