unit Xc12DataXLinks5;

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
     Xc12Utils5,
     XLSUtils5, Xc12Common5;

type TXc12XCellType = (x12ctNil,x12ctBoolean,x12ctError,x12ctFloat,x12ctString);
type TXc121DdeValueType = (x12dvtNil,x12dvtBoolean,x12dvtFloat,x12dvtError,x12dvtString);

type TXc12XCellData = record
     case CT: TXc12XCellType of
       x12ctBoolean: (valBoolean: boolean);
       x12ctError  : (valError: TXc12CellError);
       x12ctFloat  : (valFloat: double);
       x12ctString : (valString: AxPUCChar);
     end;

type TXc12XDefinedName = class(TObject)
protected
     FName: AxUCString;
     FRefersTo: AxUCString;
     FSheetId: integer;
public
     property Name: AxUCString read FName write FName;
     property RefersTo: AxUCString read FRefersTo write FRefersTo;
     property SheetId: integer read FSheetId write FSheetId;
     end;

type TXc12XDefinedNames = class(TObjectList)
private
     function GetItems(Index: integer): TXc12XDefinedName;
protected
public
     constructor Create;

     function Add: TXc12XDefinedName;

     property Items[Index: integer]: TXc12XDefinedName read GetItems; default;
     end;

type TXc12XCellValue = class(TObject)
private
     function  GetAsBoolean: boolean;
     function  GetAsError: TXc12CellError;
     function  GetAsFloat: double;
     function  GetAsString: AxUCString;
     procedure SetAsBoolean(const Value: boolean);
     procedure SetAsError(const Value: TXc12CellError);
     procedure SetAsFloat(const Value: double);
     procedure SetAsString(const Value: AxUCString);
     function  GetAsNil: boolean;
     procedure SetAsNil(const Value: boolean);
protected
     FData: TXc12XCellData;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property AsBoolean: boolean read GetAsBoolean write SetAsBoolean;
     property AsError  : TXc12CellError read GetAsError write SetAsError;
     property AsFloat  : double read GetAsFloat write SetAsFloat;
     property AsString : AxUCString read GetAsString write SetAsString;
     property AsNil    : boolean read GetAsNil write SetAsNil;
     property Type_    : TXc12XCellType read FData.CT;
     property Data     : TXc12XCellData read FData;
     end;

type TXc12XCell = class(TXc12XCellValue)
private
     function  GetCellRef: AxUCString;
     procedure SetCellRef(const Value: AxUCString);
protected
     FCol,FRow: integer;
     FValueMetadata: integer;
public
     property CellRef: AxUCString read GetCellRef write SetCellRef;
     property Col: integer read FCol write FCol;
     property Row: integer read FRow write FRow;
     property ValueMetadata: integer read FValueMetadata write FValueMetadata;
     end;

type TXc12XRow = class(TObjectList)
private
     function  GetCells(Index: integer): TXc12XCell;
protected
     FRow: integer;
public
     constructor Create;

     function Add: TXc12XCell;

     property Row: integer read FRow write FRow;
     property Cells[Index: integer]: TXc12XCell read GetCells; default;
     end;

type TXc12XSheetData = class(TObjectList)
private
     function GetRows(Index: integer): TXc12XRow;
protected
     FSheetId: integer;
     FRefreshError: boolean;
public
     constructor Create;

     function Add: TXc12XRow;
     function FindCell(ACol,ARow: integer; out AValue: TXc12XCellData): boolean;

     property SheetId: integer read FSheetId write FSheetId; // TODO Replace with id to sheet object.
     property RefreshError: boolean read FRefreshError write FRefreshError;
     property Rows[Index: integer]: TXc12XRow read GetRows; default;
     end;

type TXc12XSheetDataSet = class(TObjectList)
private
     function GetSheetData(Index: integer): TXc12XSheetData;
protected
public
     constructor Create;

     function Add: TXc12XSheetData;
     function FindSheet(ASheet: integer): TXc12XSheetData;

     property SheetData[Index: integer]: TXc12XSheetData read GetSheetData; default;
     end;

type TXc12XExternalBook = class(TObject)
protected
     FSheetNames  : TStringList;
     FDefinedNames: TXc12XDefinedNames;
     FSheetDataSet: TXc12XSheetDataSet;
     FFilename    : AxUCString;
public
     constructor Create;
     destructor Destroy; override;

     function  FindSheetName(AName: AxUCString): integer;

     property SheetNames: TStringList read FSheetNames;
     property DefinedNames: TXc12XDefinedNames read FDefinedNames;
     property SheetDataSet: TXc12XSheetDataSet read FSheetDataSet;
     property Filename: AxUCString read FFilename write FFilename;
     end;

type TXc12DdeValue = class(TXc12XCellValue)
protected
public
     end;

type TXc12DdeValues = class(TObjectList)
private
     function GetItems(Index: integer): TXc12DdeValue;
protected
     FCols: integer;
     FRows: integer;
public
     constructor Create;

     function Add: TXc12DdeValue;

     property Cols: integer read FCols write FCols;
     property Rows: integer read FRows write FRows;
     property Items[Index: integer]: TXc12DdeValue read GetItems; default;
     end;

type TXc12DdeItem = class(TObject)
private
protected
     FName: AxUCString;
     FOle: boolean;
     FAdvise: boolean;
     FPreferPic: boolean;
     FValues: TXc12DdeValues;
public
     constructor Create;
     destructor Destroy; override;

     property Name: AxUCString read FName write FName;
     property Ole: boolean read FOle write FOle;
     property Advise: boolean read FAdvise write FAdvise;
     property PreferPic: boolean read FPreferPic write FPreferPic;

     property Values: TXc12DdeValues read FValues;
     end;

type TXc12DdeLink = class(TObjectList)
private
     function GetItems(Index: integer): TXc12DdeItem;
protected
     FDdeService: AxUCString;
     FDdeTopic: AxUCString;
public
     constructor Create;

     function Add: TXc12DdeItem;

     property DdeService: AxUCString read FDdeService write FDdeService;
     property DdeTopic: AxUCString read FDdeTopic write FDdeTopic;
     property Items[Index: integer]: TXc12DdeItem read GetItems; default;
     end;

type TXc12OleItem = class(TObject)
protected
     FName: AxUCString;
     FIcon: boolean;
     FAdvice: boolean;
     FPreferPic: boolean;
public
     property Name: AxUCString read FName write FName;
     property Icon: boolean read FIcon write FIcon;
     property Advice: boolean read FAdvice write FAdvice;
     property PreferPic: boolean read FPreferPic write FPreferPic;
     end;

type TXc12OleLink = class(TObjectList)
private
     function GetItems(Index: integer): TXc12OleItem;
protected
     FFilename: AxUCString;
     FProgId: AxUCString;
public
     constructor Create;

     function Add: TXc12OleItem;

     property Filename: AxUCString read FFilename write FFilename;
     property ProgId: AxUCString read FProgId write FProgId;
     property Items[Index: integer]: TXc12OleItem read GetItems; default;
     end;

type TXc12ExternalLink = class(TObject)
protected
     FExternalBook: TXc12XExternalBook;
     FDdeLink: TXc12DdeLink;
     FOleLink: TXc12OleLink;
public
     constructor Create;
     destructor Destroy; override;

     property ExternalBook: TXc12XExternalBook read FExternalBook;
     property DdeLink: TXc12DdeLink read FDdeLink;
     property OleLink: TXc12OleLink read FOleLink;
     end;

type TXc12DataXLinks = class(TObjectList)
private
     function GetItems(Index: integer): TXc12ExternalLink;
protected
public
     constructor Create;

     function Add: TXc12ExternalLink;
     function AddXBook(AFilename: AxUCString; ASheets: array of AxUCString): TXc12ExternalLink;
     function FindBook(const AName: AxUCString): integer;
     function GetValue(AXBook,AXSheetIndex,ACol,ARow: integer; out AValue: TXc12XCellData): boolean;

     property Items[Index: integer]: TXc12ExternalLink read GetItems; default;
     end;

implementation


{ TXc12Cell }

procedure TXc12XCellValue.Clear;
begin
  if FData.CT = x12ctString then
    FreeMem(FData.valString);
  FData.CT := x12ctFloat;
  FData.valFloat := 0;
end;

constructor TXc12XCellValue.Create;
begin
  Clear;
end;

destructor TXc12XCellValue.Destroy;
begin
  Clear;
  inherited;
end;

function TXc12XCellValue.GetAsBoolean: boolean;
begin
  case FData.CT of
    x12ctBoolean: Result := FData.valBoolean;
    x12ctError  : Result := False;
    x12ctFloat  : Result := FData.valFloat <> 0;
    x12ctString : if Uppercase(FData.valString) = G_StrTRUE then Result := True else Result := False;
    else          Result := False;
  end;
end;

function TXc12XCellValue.GetAsError: TXc12CellError;
begin
  case FData.CT of
    x12ctBoolean: Result := errUnknown;
    x12ctError  : Result := FData.valError;
    x12ctFloat  : Result := errUnknown;
    x12ctString : Result := errUnknown;
    else          Result := errUnknown;
  end;
end;

function TXc12XCellValue.GetAsFloat: double;
begin
  Result := 0;
  case FData.CT of
    x12ctBoolean: if FData.valBoolean then Result := 1 else Result := 0;
    x12ctError  : Result := 0;
    x12ctFloat  : Result := FData.valFloat;
    x12ctString : StrToFloatDef(FData.valString,0);
    else          Result := 0;
  end;
end;

function TXc12XCellValue.GetAsNil: boolean;
begin
  Result := FData.CT = x12ctNil;
end;

function TXc12XCellValue.GetAsString: AxUCString;
begin
  case FData.CT of
    x12ctBoolean: if FData.valBoolean then Result := G_StrTRUE else Result := G_StrFalse;
    x12ctError  : Result := Xc12CellErrorNames[FData.valError];
    x12ctFloat  : Result := FloatToStr(FData.valFloat);
    x12ctString : Result := FData.valString;
    else          Result := '';
  end;
end;

procedure TXc12XCellValue.SetAsBoolean(const Value: boolean);
begin
  Clear;
  FData.CT := x12ctBoolean;
  FData.valBoolean := Value;
end;

procedure TXc12XCellValue.SetAsError(const Value: TXc12CellError);
begin
  Clear;
  FData.CT := x12ctError;
  FData.valError := Value;
end;

procedure TXc12XCellValue.SetAsFloat(const Value: double);
begin
  Clear;
  FData.CT := x12ctFloat;
  FData.valFloat := Value;
end;

procedure TXc12XCellValue.SetAsNil(const Value: boolean);
begin
  Clear;
  if Value then
    FData.CT := x12ctNil;
end;

procedure TXc12XCellValue.SetAsString(const Value: AxUCString);
begin
  Clear;
  FData.CT := x12ctString;
  GetMem(FData.valString,Length(Value) * 2 + 2);
  Move(Value[1],FData.valString^,Length(Value) * 2);
  FData.valString[Length(Value)] := #0;
end;

{ TXc12XRow }

function TXc12XRow.Add: TXc12XCell;
begin
  Result := TXc12XCell.Create;
  inherited Add(Result);
end;

constructor TXc12XRow.Create;
begin
  inherited Create;
end;

function TXc12XRow.GetCells(Index: integer): TXc12XCell;
begin
  Result := TXc12XCell(inherited Items[Index]);
end;

{ TXc12XSheetDataSet }

function TXc12XSheetDataSet.Add: TXc12XSheetData;
begin
  Result := TXc12XSheetData.Create;
  inherited Add(Result);
end;

constructor TXc12XSheetDataSet.Create;
begin
  inherited Create;
end;

function TXc12XSheetDataSet.FindSheet(ASheet: integer): TXc12XSheetData;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if SheetData[i].SheetId = ASheet then begin
      Result := SheetData[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXc12XSheetDataSet.GetSheetData(Index: integer): TXc12XSheetData;
begin
  Result := TXc12XSheetData(inherited Items[Index]);
end;

{ TXc12SXheetData }

function TXc12XSheetData.Add: TXc12XRow;
begin
  Result := TXc12XRow.Create;
  inherited Add(Result);
end;

constructor TXc12XSheetData.Create;
begin
  inherited Create;
end;

function TXc12XSheetData.GetRows(Index: integer): TXc12XRow;
begin
  Result := TXc12XRow(inherited Items[Index]);
end;

function TXc12XSheetData.FindCell(ACol, ARow: integer; out AValue: TXc12XCellData): boolean;
var
  i,j: integer;
begin
  for i := 0 to Count - 1 do begin
    if Rows[i].Row = ARow then begin
      for j := 0 to Rows[i].Count - 1 do begin
        if Rows[i].Cells[j].Col = ACol then begin
          AValue := Rows[i].Cells[j].Data;
          Result := True;
          Exit;
        end;
      end;
    end;
  end;

  Result := False;
end;

{ TXc12XDefinedNames }

function TXc12XDefinedNames.Add: TXc12XDefinedName;
begin
  Result := TXc12XDefinedName.Create;
  inherited Add(Result);
end;

constructor TXc12XDefinedNames.Create;
begin
  inherited Create;
end;

function TXc12XDefinedNames.GetItems(Index: integer): TXc12XDefinedName;
begin
  Result := TXc12XDefinedName(inherited Items[Index]);
end;

{ TXc12XExternalBook }

constructor TXc12XExternalBook.Create;
begin
  FSheetNames := TStringList.Create;
  FDefinedNames := TXc12XDefinedNames.Create;
  FSheetDataSet := TXc12XSheetDataSet.Create;
end;

destructor TXc12XExternalBook.Destroy;
begin
  FSheetNames.Free;
  FDefinedNames.Free;
  FSheetDataSet.Free;
  inherited;
end;

function TXc12XExternalBook.FindSheetName(AName: AxUCString): integer;
begin
  AName := Uppercase(AName);
  for Result := 0 to FSheetNames.Count - 1 do begin
    if Uppercase(FSheetNames[Result]) = AName then
      Exit;
  end;
  Result := -1;
end;

{ TXc12XCell }

function TXc12XCell.GetCellRef: AxUCString;
begin
  Result := ColRowToRefStr(FCol,FRow,False,False);
end;

procedure TXc12XCell.SetCellRef(const Value: AxUCString);
begin
  RefStrToColRow(Value,FCol,FRow);
end;

{ TXc12DdeValues }

function TXc12DdeValues.Add: TXc12DdeValue;
begin
  Result := TXc12DdeValue.Create;
  inherited Add(Result);
end;

constructor TXc12DdeValues.Create;
begin
  inherited Create;
end;

function TXc12DdeValues.GetItems(Index: integer): TXc12DdeValue;
begin
  Result := TXc12DdeValue(inherited Items[Index]);
end;

{ TXc12DdeItem }

constructor TXc12DdeItem.Create;
begin
  FValues := TXc12DdeValues.Create;
end;

destructor TXc12DdeItem.Destroy;
begin
  FValues.Free;
  inherited;
end;

{ TXc12DdeLink }

function TXc12DdeLink.Add: TXc12DdeItem;
begin
  Result := TXc12DdeItem.Create;
  inherited Add(Result);
end;

constructor TXc12DdeLink.Create;
begin
  inherited Create;
end;

function TXc12DdeLink.GetItems(Index: integer): TXc12DdeItem;
begin
  Result := TXc12DdeItem(inherited Items[Index]);
end;

{ TXc12OleLink }

function TXc12OleLink.Add: TXc12OleItem;
begin
  Result := TXc12OleItem.Create;
  inherited Add(Result);
end;

constructor TXc12OleLink.Create;
begin
  inherited Create;
end;

function TXc12OleLink.GetItems(Index: integer): TXc12OleItem;
begin
  Result := TXc12OleItem(inherited Items[Index]);
end;

{ TXc12ExternalLink }

constructor TXc12ExternalLink.Create;
begin
  FExternalBook := TXc12XExternalBook.Create;
  FDdeLink := TXc12DdeLink.Create;
  FOleLink := TXc12OleLink.Create;
end;

destructor TXc12ExternalLink.Destroy;
begin
  FExternalBook.Free;
  FDdeLink.Free;
  FOleLink.Free;
  inherited;
end;

{ TXc12DataXLinks }

function TXc12DataXLinks.Add: TXc12ExternalLink;
begin
  Result := TXc12ExternalLink.Create;
  inherited Add(Result);
end;

function TXc12DataXLinks.AddXBook(AFilename: AxUCString; ASheets: array of AxUCString): TXc12ExternalLink;
var
  i: integer;
begin
  Result := Add;
  Result.ExternalBook.Filename := AFilename;

  for i := 0 to High(ASheets) do
    Result.ExternalBook.SheetNames.Add(ASheets[i]);
end;

constructor TXc12DataXLinks.Create;
begin
  inherited Create;
end;

function TXc12DataXLinks.FindBook(const AName: AxUCString): integer;
var
  S: AxUCString;
begin
  if TryStrToInt(AName,Result) then begin
    if (Result >= 1) and (Result <= Count) then
      Exit;
  end
  else begin
    S := Uppercase(AName);
    for Result := 0 to Count - 1 do begin
      if (Items[Result].ExternalBook <> Nil) and (Uppercase(Items[Result].ExternalBook.Filename) = S) then
        Exit;
    end;
  end;
  Result := -1;
end;

function TXc12DataXLinks.GetItems(Index: integer): TXc12ExternalLink;
begin
  Result := TXc12ExternalLink(inherited Items[Index]);
end;

function TXc12DataXLinks.GetValue(AXBook, AXSheetIndex, ACol, ARow: integer; out AValue: TXc12XCellData): boolean;
var
  XBook: TXc12XExternalBook;
begin
  Result := (AXBook >= 0) and (AXBook < Count);
  if not Result then
    Exit;
  XBook := Items[AXBook].ExternalBook;
  Result :=  XBook.SheetDataSet.SheetData[AXSheetIndex].FindCell(ACol,ARow,AValue);
end;

end.
