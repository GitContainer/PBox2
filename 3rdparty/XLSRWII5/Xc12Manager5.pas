unit Xc12Manager5;

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
{$ifdef MSWINDOWS}
     Windows,
{$endif}
     xpgPUtils, xpgPSimpleDOM,
{$ifdef XLS_BIFF}
     BIFF_Names5,
{$endif}
     xpgParseDrawing, xpgParserPivot,
     Xc12Utils5, Xc12Common5, Xc12DataSST5, Xc12DataStyleSheet5, Xc12DataWorkbook5,
     Xc12DataWorksheet5, Xc12DataXLinks5, Xc12FileData5, Xc12DataTable5, Xc12Graphics,
     XLSUtils5, XLSClassFactory5, XLSRelCells5;

type TXLSEventCell = class(TObject)
private
     function  GetBoolean: boolean;
     function  GetCellType: TXLSCellType;
     function  GetError: TXc12CellError;
     function  GetFloat: double;
     function  GetFormula: AxUCString;
     function  GetString: AxUCString;
     procedure SetBoolean(const Value: boolean);
     procedure SetError(const Value: TXc12CellError);
     procedure SetFloat(const Value: double);
     procedure SetFormula(const Value: AxUCString);
     procedure SetString(const Value: AxUCString);
     procedure SetRef(const Value: AxUCString);
     procedure SetRow(const Value: integer);
     procedure SetCol(const Value: integer);
     function  GetActive: boolean;
protected
     FPrevCol   : integer;
     FPrevRow   : integer;

     FActive    : boolean;
     FAborted   : boolean;
     FSheetIndex: integer;
     FSheetName : AxUCString;
     FCol       : integer;
     FRow       : integer;

     FTargetArea: TXLSCellArea;

     FCellType  : TXLSCellType;
     FIsFormula : boolean;

     FValBoolean: boolean;
     FValError  : TXc12CellError;
     FValFloat  : double;
     FValString : AxUCString;

     FFormula   : AxUCString;

     FXF        : TXc12XF;
public
     constructor Create;

     procedure Clear;
     procedure ClearValues;

     procedure SetTargetArea(const ASheetIndex,ACol1,ARow1,ACol2,ARow2: integer);

     procedure ColDone;
     procedure RowDone;

     procedure Abort;

     property Active    : boolean read GetActive write FActive;

     property SheetIndex: integer read FSheetIndex write FSheetIndex;
     // Sheet name when reading files. Not used when writing.
     property SheetName : AxUCString read FSheetName write FSheetName;
     property Col       : integer read FCol write SetCol;
     property Row       : integer read FRow write SetRow;
     property Ref       : AxUCString write SetRef;
     property TargetArea: TXLSCellArea read FTargetArea;

     property CellType  : TXLSCellType read GetCellType;

     property AsBoolean : boolean read GetBoolean write SetBoolean;
     property AsError   : TXc12CellError read GetError write SetError;
     property AsFloat   : double read GetFloat write SetFloat;
     property AsString  : AxUCString read GetString write SetString;

     property AsFormula : AxUCString read GetFormula write SetFormula;

     property XF        : TXc12XF read FXF write FXF;
     end;

type TCellReadWriteEvent = procedure(ACell: TXLSEventCell) of object;

type TXc12DocProps = class(TObject)
private
     function  GetCreator: AxUCString;
     function  GetDateCreated: TDateTime;
     function  GetDateModified: TDateTime;
     function  GetLastModifiedBy: AxUCString;
     procedure SetCreator(const Value: AxUCString);
     procedure SetDateCreated(const Value: TDateTime);
     procedure SetDateModified(const Value: TDateTime);
     procedure SetLastModifiedBy(const Value: AxUCString);
protected
     FDocPropsCore: TXpgSimpleDOM;
     FAutoUpdate  : boolean;
public
     constructor Create(ADocPropsCore : TXpgSimpleDOM);

     property AutoUpdate: boolean read FAutoUpdate write FAutoUpdate;
     property Creator: AxUCString read GetCreator write SetCreator;
     property LastModifiedBy: AxUCString read GetLastModifiedBy write SetLastModifiedBy;
     property DateCreated: TDateTime read GetDateCreated write SetDateCreated;
     property DateModified: TDateTime read GetDateModified write SetDateModified;
     end;

type TXc12Manager = class;

     TXc12GraphicManagerImpl = class(TXc12GraphicManager)
protected
     FManager: TXc12Manager;
public
     constructor Create(AManager: TXc12Manager; AErrors: TXLSErrorManager);

     function CreateRelativeCells(ARef: AxUCString): TXLSRelCells; override;
     end;

     TXc12Manager = class(TObject)
private
     function  GetFilenameAsXLS: AxUCString;
     function  GetUseAlternateZip: boolean;
     procedure SetUseAlternateZip(const Value: boolean);
     procedure SetVersion(const Value: TExcelVersion);
protected
     FErrors          : TXLSErrorManager;

     FVersion         : TExcelVersion;

     FFilename        : AxUCString;
     FPassword        : AxUCString;
     FPasswordEvent   : TStringEvent;

     FHasXSS          : boolean;

     FFileData        : TXc12FileData;
     FStyleSheet      : TXc12DataStyleSheet;
     FSST             : TXc12DataSST;
     FWorkbook        : TXc12DataWorkbook;
     FWorksheets      : TXc12DataWorksheets;
     FXLinks          : TXc12DataXLinks;
     FGrManager       : TXc12GraphicManager;
{$ifdef XLS_BIFF}
     FNames97         : TInternalNames;
     FExtNames97      : TExternalNames;
{$endif}
     FIsUpdating      : boolean;

     FPaletteChanged  : boolean;

     FColSeparator    : AxUCChar;
     FRowSeparator    : AxUCChar;

     FAborted         : boolean;

     FEventCell       : TXLSEventCell;
     FCellReadEvent   : TCellReadWriteEvent;
     FCellWriteEvent  : TCellReadWriteEvent;
     FDirectRead      : boolean;
     FDirectWrite     : boolean;
     FDefaultPaperSz  : TXc12PaperSize;

     FClassFactory    : TXLSClassFactory;

     FProgress        : TXLSProgressEvent;
     FProgressType    : TXLSProgressType;
     FProgressCount   : integer;

     FXSSAfterRead    : TNotifyEvent;

     FStrSERIES       : AxUCString;

     FDocProps        : TXc12DocProps;

     FVirtualCells    : TObjectList;

     procedure SetupLocale;
public
     constructor Create(AClassFactory: TXLSClassFactory; AErrors: TXLSErrorManager);
     destructor Destroy; override;

     procedure CreateMembers;

     procedure Clear;

     function  FindTable(const AName: AxUCString; out ATables: TXc12Tables; out ASheetIndex: integer): integer;
     function  FindPivotTable(ACache: TCT_pivotCacheDefinition): TCT_pivotTableDefinition;

     procedure BeforeRead;
     procedure AfterRead;
     procedure BeforeWrite;

     procedure BeginProgress(const AType: TXLSProgressType; const ACount: integer);
     procedure WorkProgress(AValue: integer);
     procedure EndProgress;

     procedure FireReadCellEvent; {$ifdef D2006PLUS} inline; {$endif}
     procedure FireWriteCellEvent(const ARowChanged: boolean); {$ifdef D2006PLUS} inline; {$endif}

     function  CheckDirectWrite: boolean;

     property Errors          : TXLSErrorManager read FErrors;

     property Version         : TExcelVersion read FVersion write SetVersion;

     property Filename        : AxUCString read FFilename write FFilename;
     property FilenameAsXLS   : AxUCString read GetFilenameAsXLS;
     property Password        : AxUCString read FPassword write FPassword;

     property HasXSS          : boolean read FHasXSS write FHasXSS;

     property ColSeparator    : AxUCChar read FColSeparator write FColSeparator;
     property RowSeparator    : AxUCChar read FRowSeparator write FRowSeparator;

     property FileData        : TXc12FileData read FFileData;
     property StyleSheet      : TXc12DataStyleSheet read FStyleSheet;
     property SST             : TXc12DataSST read FSST;
     property Workbook        : TXc12DataWorkbook read FWorkbook;
     property Worksheets      : TXc12DataWorksheets read FWorksheets;
     property XLinks          : TXc12DataXLinks read FXLinks;
     property GrManager       : TXc12GraphicManager read FGrManager;
{$ifdef XLS_BIFF}
     property Names97         : TInternalNames read FNames97 write FNames97;
     property _ExtNames97      : TExternalNames read FExtNames97 write FExtNames97;
{$endif}
     property EventCell       : TXLSEventCell read FEventCell;

     property Aborted         : boolean read FAborted write FAborted;

     property PaletteChanged  : boolean read FPaletteChanged write FPaletteChanged;

     property DirectRead      : boolean read FDirectRead write FDirectRead;
     property DirectWrite     : boolean read FDirectWrite write FDirectWrite;

     property DefaultPaperSz  : TXc12PaperSize read FDefaultPaperSz write FDefaultPaperSz;

     property UseAlternateZip : boolean read GetUseAlternateZip write SetUseAlternateZip;

     property StrSERIES       : AxUCString read FStrSERIES write FStrSERIES;

     property DocProps        : TXc12DocProps read FDocProps;

     property VirtualCells    : TObjectList read FVirtualCells;

     property OnPassword      : TStringEvent read FPasswordEvent write FPasswordEvent;
     property OnReadCell      : TCellReadWriteEvent read FCellReadEvent write FCellReadEvent;
     property OnWriteCell     : TCellReadWriteEvent read FCellWriteEvent write FCellWriteEvent;
     property OnProgress      : TXLSProgressEvent read FProgress write FProgress;

     property OnXSSAfterRead  : TNotifyEvent read FXSSAfterRead write FXSSAfterRead;
     end;

implementation

{ TXc12Manager }

procedure TXc12Manager.AfterRead;
begin
  FIsUpdating := False;
  FStyleSheet.AfterRead;
  FSST.AfterRead;

  if Assigned(FXSSAfterRead) then
    FXSSAfterRead(Self);
end;

procedure TXc12Manager.BeforeRead;
begin
  FIsUpdating := True;
end;

procedure TXc12Manager.BeforeWrite;
{$ifndef BABOON}
var
  YY,MM,DD: word;
  UTC: SystemTime;
{$endif}
begin
  FStyleSheet.BeforeWrite;
  FGrManager.BeforeWrite;

  if FDocProps.AutoUpdate then begin
{$ifndef BABOON}
    DecodeDate(FDocProps.DateCreated,YY,MM,DD);
    if (FDocProps.DateCreated = 0) or (YY <= 2000) then
      FDocProps.DateCreated := Now;
    GetSystemTime(UTC);
    FDocProps.DateModified := SystemTimeToDateTime(UTC);
{$endif}
  end;
end;

procedure TXc12Manager.BeginProgress(const AType: TXLSProgressType; const ACount: integer);
begin
  FProgressType := AType;
  FProgressCount := ACount;
  if Assigned(FProgress) then
    FProgress(FProgressType,xpsBegin,0);
end;

function TXc12Manager.CheckDirectWrite: boolean;
begin
  Result := (FEventCell.Col > FEventCell.FPrevCol) or (FEventCell.Row > FEventCell.FPrevRow);
  FEventCell.Active := Result;
  if Result then begin
    FEventCell.FPrevCol := FEventCell.FCol;
    FEventCell.FPrevRow := FEventCell.FRow;
  end;
end;

procedure TXc12Manager.Clear;
begin
  FXLinks.Clear;
  FWorksheets.Clear;
  FWorkbook.Clear;
  FSST.Clear;
  FStyleSheet.Clear;
  FFileData.Clear;
  FGrManager.Clear;
{$ifdef XLS_BIFF}
  FExtNames97 := Nil;
{$endif}
end;

constructor TXc12Manager.Create(AClassFactory: TXLSClassFactory; AErrors: TXLSErrorManager);
begin
  SetupLocale;

  FClassFactory := AClassFactory;
  FErrors := AErrors;

  FVirtualCells := TObjectList.Create(False);

  if FormatSettings.DecimalSeparator = ',' then
    FColSeparator := '\'
  else
    FColSeparator := ',';
  FRowSeparator := FormatSettings.ListSeparator;

  FStrSERIES := 'Series';

  FEventCell := TXLSEventCell.Create;
end;

procedure TXc12Manager.CreateMembers;
begin
  FFileData := TXc12FileData.Create;
  FStyleSheet := TXc12DataStyleSheet.Create;
  FSST := TXc12DataSST.Create(FStyleSheet);
  FWorkbook := TXc12DataWorkbook.Create(FClassFactory);
  FWorksheets := TXc12DataWorksheets.Create(FClassFactory,FSST,FStyleSheet);
  FXLinks := TXc12DataXLinks.Create;
  FGrManager := TXc12GraphicManagerImpl.Create(Self,FErrors);
  FDocProps := TXc12DocProps.Create(FFileData.DocPropsCore);
end;

destructor TXc12Manager.Destroy;
begin
  FXLinks.Free;
  FWorksheets.Free;
  FWorkbook.Free;
  FSST.Free;
  FStyleSheet.Free;
  FFileData.Free;
  FEventCell.Free;
  FGrManager.Free;
  FDocProps.Free;
  FVirtualCells.Free;
  inherited;
end;

procedure TXc12Manager.EndProgress;
begin
  if Assigned(FProgress) then
    FProgress(FProgressType,xpsEnd,1);
end;

function TXc12Manager.FindPivotTable(ACache: TCT_pivotCacheDefinition): TCT_pivotTableDefinition;
var
  i: integer;
begin
  for i := 0 to FWorksheets.Count - 1 do begin
    Result := FWorksheets[i].PivotTables.FindByCache(ACache);
    if Result <> Nil then
      Exit;
  end;
  Result := Nil;
end;

function TXc12Manager.FindTable(const AName: AxUCString; out ATables: TXc12Tables; out ASheetIndex: integer): integer;
var
  i: integer;
  S: AxUCString;
begin
  S := AnsiUppercase(AName);
  for i := 0 to FWorksheets.Count - 1 do begin
    for Result := 0 to FWorksheets[i].Tables.Count - 1 do begin
      if S = AnsiUppercase(FWorksheets[i].Tables[Result].Name) then begin
        ATables := FWorksheets[i].Tables;
        ASheetIndex := i;
        Exit;
      end;
    end;
  end;
  Result := -1;
end;

procedure TXc12Manager.FireReadCellEvent;
begin
  if Assigned(FCellReadEvent) then
    FCellReadEvent(FEventCell);
end;

procedure TXc12Manager.FireWriteCellEvent(const ARowChanged: boolean);
begin
  if Assigned(FCellWriteEvent) then begin
    FEventCell.Active := FEventCell.Row <= FEventCell.TargetArea.Row2;
    if FEventCell.Active then begin
      if FEventCell.Col >= FEventCell.TargetArea.Col2 then begin
        FEventCell.Col := FEventCell.TargetArea.Col1 - 1;
        FEventCell.Row := FEventCell.Row + 1;
        FEventCell.Active := False;
        Exit;
      end
      else begin
        if FEventCell.Row < 0 then
          FEventCell.Row := FEventCell.TargetArea.Row1;
        if not ARowChanged then
          FEventCell.Col := FEventCell.Col + 1;
      end;

      if not ARowChanged then
        FCellWriteEvent(FEventCell);
    end;
  end;
end;

function TXc12Manager.GetFilenameAsXLS: AxUCString;
begin
  if FFilename <> '' then
    Result := ExtractFilePath(FFilename) + '[' + ExtractFileName(FFilename) + ']'
  else
    Result := '';
end;

function TXc12Manager.GetUseAlternateZip: boolean;
begin
  Result := FFileData.UseAlternateZip;
end;

procedure TXc12Manager.SetupLocale;
const
  LOCALE_NAME_SYSTEM_DEFAULT = '!sys-default-locale' + #0;
  LOCALE_IPAPERSIZE          =   $0000100A;
  LOCALE_RETURN_NUMBER       =   $20000000;
{$ifndef BABOON}
var
  PB: PByteArray;
  Sz: integer;
{$endif}
begin
{$ifdef BABOON}
  FDefaultPaperSz := psA4;
{$else}
  Sz := GetLocaleInfo(GetUserDefaultLCID,LOCALE_IPAPERSIZE,Nil,0);
  GetMem(PB,Sz * 2);
  try
    if GetLocaleInfo(LOCALE_SYSTEM_DEFAULT,LOCALE_IPAPERSIZE,PChar(PB),Sz) <> 0 then begin
      case Char(PB[0]) of
        '1': FDefaultPaperSz := psLegal;
        '5': FDefaultPaperSz := psLetter;
        '8': FDefaultPaperSz := psA3;
        '9': FDefaultPaperSz := psA4;
        else FDefaultPaperSz := psA4;
      end;
    end;
  finally
     FreeMem(PB);
  end;
{$endif}
end;

procedure TXc12Manager.SetUseAlternateZip(const Value: boolean);
begin
  FFileData.UseAlternateZip := Value;
end;

procedure TXc12Manager.SetVersion(const Value: TExcelVersion);
begin
  FVersion := Value;
  FSST.IsExcel97 := FVersion < xvExcel2007;
end;

procedure TXc12Manager.WorkProgress(AValue: integer);
begin
  if Assigned(FProgress) and (FProgressCount > 0) then
    FProgress(FProgressType,xpsWork,AValue / FProgressCount);
end;

{ TXLSEventCell }

procedure TXLSEventCell.Abort;
begin
  FAborted := True;
end;

procedure TXLSEventCell.Clear;
begin
  FCol := FTargetArea.Col1 - 1;
  FRow := FTargetArea.Row1;
  FPrevCol := -1;
  FPrevRow := -1;
  ClearValues;
end;

procedure TXLSEventCell.ClearValues;
begin
  FCellType := xctNone;
  FFormula := '';
  FIsFormula := False;
  FXF := Nil;
end;

procedure TXLSEventCell.ColDone;
begin
  FPrevCol := FCol;
end;

constructor TXLSEventCell.Create;
begin
  FTargetArea.Col1 := 0;
  FTargetArea.Row1 := 0;
  FTargetArea.Col2 := 255;
  FTargetArea.Row2 := 65535;
end;

function TXLSEventCell.GetActive: boolean;
begin
  Result := FActive and not FAborted;
end;

function TXLSEventCell.GetBoolean: boolean;
begin
  Result := FValBoolean;
end;

function TXLSEventCell.GetCellType: TXLSCellType;
begin
  Result := FCellType;
end;

function TXLSEventCell.GetError: TXc12CellError;
begin
  Result := FValError;
end;

function TXLSEventCell.GetFloat: double;
begin
  Result := FValFloat;
end;

function TXLSEventCell.GetFormula: AxUCString;
begin
  Result := FFormula;
end;

function TXLSEventCell.GetString: AxUCString;
begin
  Result := FValString;
end;

procedure TXLSEventCell.RowDone;
begin
  FPrevRow := FRow;
end;

procedure TXLSEventCell.SetBoolean(const Value: boolean);
begin
  FValBoolean := Value;
  if FIsFormula then
    FCellType := xctBooleanFormula
  else
    FCellType := xctBoolean;
end;

procedure TXLSEventCell.SetCol(const Value: integer);
begin
  FCol := Value;
end;

procedure TXLSEventCell.SetError(const Value: TXc12CellError);
begin
  FValError := Value;
  if FIsFormula then
    FCellType := xctErrorFormula
  else
    FCellType := xctError;
end;

procedure TXLSEventCell.SetFloat(const Value: double);
begin
  FValFloat := Value;
  if FIsFormula then
    FCellType := xctFloatFormula
  else
    FCellType := xctFloat;
end;

procedure TXLSEventCell.SetFormula(const Value: AxUCString);
begin
  FFormula := Value;
  FIsFormula := True;
end;

procedure TXLSEventCell.SetRef(const Value: AxUCString);
begin
  RefStrToColRow(Value,FCol,FRow);
end;

procedure TXLSEventCell.SetRow(const Value: integer);
begin
  FRow := Value;
end;

procedure TXLSEventCell.SetString(const Value: AxUCString);
begin
  FValString := Value;
  if FIsFormula then
    FCellType := xctStringFormula
  else
    FCellType := xctString;
end;

procedure TXLSEventCell.SetTargetArea(const ASheetIndex,ACol1, ARow1, ACol2, ARow2: integer);
begin
  FSheetIndex := ASheetIndex;
  FTargetArea.Col1 := ACol1;
  FTargetArea.Row1 := ARow1;
  FTargetArea.Col2 := ACol2;
  FTargetArea.Row2 := ARow2;
end;

{ TXc12DocProps }

constructor TXc12DocProps.Create(ADocPropsCore: TXpgSimpleDOM);
begin
  FDocPropsCore := ADocPropsCore;
  FAutoUpdate := True;
end;

function TXc12DocProps.GetCreator: AxUCString;
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/dc:creator');
  if Node <> Nil then
    Result := Node.Content
  else
    Result := '';
end;

function TXc12DocProps.GetDateCreated: TDateTime;
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/dcterms:created');
  if Node <> Nil then
    Result := XmlStrToDateTime(Node.Content)
  else
    Result := 0;
end;

function TXc12DocProps.GetDateModified: TDateTime;
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/dcterms:modified');
  if Node <> Nil then
    Result := XmlStrToDateTime(Node.Content)
  else
    Result := 0;
end;

function TXc12DocProps.GetLastModifiedBy: AxUCString;
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/cp:lastModifiedBy');
  if Node <> Nil then
    Result := Node.Content
  else
    Result := '';
end;

procedure TXc12DocProps.SetCreator(const Value: AxUCString);
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/dc:creator');
  if Node <> Nil then
    Node.Content := Value;
end;

procedure TXc12DocProps.SetDateCreated(const Value: TDateTime);
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/dcterms:created');
  if Node <> Nil then
    Node.Content := XmlDateTimeToStr(Value);
end;

procedure TXc12DocProps.SetDateModified(const Value: TDateTime);
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/dcterms:modified');
  if Node <> Nil then
    Node.Content := XmlDateTimeToStr(Value);
end;

procedure TXc12DocProps.SetLastModifiedBy(const Value: AxUCString);
var
  Node: TXpgDOMNode;
begin
  Node := FDocPropsCore.Root.Find('cp:coreProperties/cp:lastModifiedBy');
  if Node <> Nil then
    Node.Content := Value;
end;

{ TXc12GraphicManagerImpl }

constructor TXc12GraphicManagerImpl.Create(AManager: TXc12Manager; AErrors: TXLSErrorManager);
begin
  inherited Create(AErrors);

  FManager := AManager;
end;

function TXc12GraphicManagerImpl.CreateRelativeCells(ARef: AxUCString): TXLSRelCells;
begin
  Result := TXLSRelCells(FManager.FClassFactory.CreateAClass(xcftVirtualCells));
  Result.__Ref := ARef;
  FManager.FVirtualCells.Add(Result);
end;

end.
