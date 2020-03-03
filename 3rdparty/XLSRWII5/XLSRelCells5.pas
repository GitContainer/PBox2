unit XLSRelCells5;

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

uses Classes, SysUtils, Contnrs, Math,
     Xc12Utils5,
     XLSUtils5, XLSSharedItems5, XLSCmdFormatValues5;

type TXLSRelCells = class(TObject)
private
     function  GetCols: integer;
     function  GetRows: integer;
     procedure SetCols(const Value: integer);
     procedure SetRows(const Value: integer);
     function  GetRef: AxUCString;
     procedure SetRef(const Value: AxUCString);
     function  GetColCount: integer;
     function  GetRowCount: integer;
     procedure SetColCount(const Value: integer);
     procedure SetRowCount(const Value: integer);
     function  GetRefAbs: AxUCString;
     function  GetShortRef: AxUCString;
     function  GetAsStringRel(ACol, ARow: integer): AxUCString;
     procedure SetAsStringRel(ACol, ARow: integer; const Value: AxUCString);
     function  GetAsBooleanRel(ACol, ARow: integer): boolean;
     function  GetAsErrorRel(ACol, ARow: integer): TXc12CellError;
     function  GetAsFloatRel(ACol, ARow: integer): double;
     function  GetAsSharedItemRel(ACol, ARow: integer): TXLSSharedItemsValue;
     procedure SetAsBooleanRel(ACol, ARow: integer; const Value: boolean);
     procedure SetAsErrorRel(ACol, ARow: integer; const Value: TXc12CellError);
     procedure SetAsFloatRel(ACol, ARow: integer; const Value: double);
     procedure SetAsSharedItemRel(ACol, ARow: integer; const Value: TXLSSharedItemsValue);
     function  GetAsBooleanCurr: boolean;
     function  GetAsErrorCurr: TXc12CellError;
     function  GetAsFloatCurr: double;
     function  GetAsStringCurr: AxUCString;
     procedure SetAsBooleanCurr(const Value: boolean);
     procedure SetAsErrorCurr(const Value: TXc12CellError);
     procedure SetAsFloatCurr(const Value: double);
     procedure SetAsStringCurr(const Value: AxUCString);
     function  GetCellTypeRel(ACol, ARow: integer): TXLSCellType;
     procedure Set__FRef(const Value: AxUCString);
protected
     __FRef : AxUCString;

     FCol1,FRow1,
     FCol2,FRow2: integer;

     FCurrCol   : integer;
     FCurrRow   : integer;

     FAutoExtend: boolean;

     FExpand    : boolean;

     FCommand   : TXLSCmdFormatValues;

     procedure GetRelColRow(var ACol, ARow: integer); overload;
     procedure GetRelColRow(var ACol1, ARow1, ACol2, ARow2: integer); overload;

     procedure DoAutoExtend(ACol,ARow: integer);

     function  Clone: TXLSRelCells; virtual; abstract;
     function  GetAsBoolean(ACol, ARow: integer): boolean; virtual; abstract;
     function  GetAsFloat(ACol,ARow: integer): double; virtual; abstract;
     function  GetAsString(ACol,ARow: integer): AxUCString; virtual; abstract;
     function  GetAsError(ACol,ARow: integer): TXc12CellError; virtual; abstract;
     function  GetAsBlank(ACol,ARow: integer): boolean; virtual; abstract;
     function  GetAsEmpty(ACol,ARow: integer): boolean; virtual; abstract;
     function  GetAsFmtString(ACol,ARow: integer): AxUCString; virtual; abstract;
     function  GetAsSharedItem(ACol,ARow: integer): TXLSSharedItemsValue; virtual; abstract;
     function  GetCellType(ACol, ARow: integer): TXLSCellType; virtual; abstract;
     function  GetIsDateTime(Col, Row: integer): boolean; virtual; abstract;
     procedure SetAsBoolean(ACol, ARow: integer; const Value: boolean); virtual; abstract;
     procedure SetAsFloat(ACol, ARow: integer; const Value: double); virtual; abstract;
     procedure SetAsString(ACol, ARow: integer; const Value: AxUCString); virtual; abstract;
     procedure SetAsError(ACol, ARow: integer; const Value: TXc12CellError); virtual; abstract;
     procedure SetAsBlank(ACol, ARow: integer; const Value: boolean); virtual; abstract;
     procedure SetAsSharedItem(ACol,ARow: integer; Value: TXLSSharedItemsValue); virtual; abstract;
public
     constructor Create; overload;
     constructor Create(ACol,ARow: integer); overload;
     constructor Create(ACol1,ARow1,ACol2,ARow2: integer); overload;
     destructor Destroy; override;

     function  Valid: boolean;

     function   IsList: boolean;

     procedure SetArea(ACol1,ARow1,ACol2,ARow2: integer);

     procedure ClearAll(ACol1,ARow1,ACol2,ARow2: integer); overload; virtual; abstract;
     procedure ClearAll; overload;

     function  Hit(ACol, ARow: integer): boolean;

     function  CloneRow(ARow: integer): TXLSRelCells;
     function  CloneCol(ACol: integer): TXLSRelCells;

     function  IsArea: boolean;

     function  GetFloatArray(out AArray: TDynSingleArray; AMaxSize: integer): integer;
     function  GetStrArray(out AArray: TDynStringArray; AMaxSize: integer): integer;
     function  AsText: AxUCString;

     procedure AutoWidthCol(ACol: integer); virtual; abstract;
     procedure AutoWidthCols;

     function  RefTopRow: AxUCString;

     function  SheetName: AxUCString; virtual; abstract;
     function  SheetIndex: integer; virtual; abstract;

     procedure SettCurrent(ACol,ARow: integer);

     procedure IncCol(ABy: integer = 1);
     procedure IncRow(ABy: integer = 1);

     procedure IncCols(ABy: integer = 1);
     procedure IncRows(ABy: integer = 1);

     function  LastNonBlankRow: integer;

     function  Equal(ACol,ARow: integer; AValue: TXLSSharedItemsValue): boolean;
     function  EqualRel(ACol,ARow: integer; AValue: TXLSSharedItemsValue): boolean;

     procedure BeginCommand; virtual; abstract;
     procedure ApplyCommand; overload; virtual; abstract;
     procedure ApplyCommand(const ACol,ARow: integer); overload; virtual; abstract;
     procedure ApplyCommand(ACol1,ARow1,ACol2,ARow2: integer); overload; virtual; abstract;
     procedure ApplyCommandRel(ACol,ARow: integer); overload;
     procedure ApplyCommandRel(ACol1,ARow1,ACol2,ARow2: integer); overload;

     property Col1: integer read FCol1 write FCol1;
     property Row1: integer read FRow1 write FRow1;
     property Col2: integer read FCol2 write FCol2;
     property Row2: integer read FRow2 write FRow2;

     property Cols: integer read GetCols write SetCols;
     property Rows: integer read GetRows write SetRows;

     property __Ref   : AxUCString read __FRef write Set__FRef;
     property RefAbs  : AxUCString read GetRefAbs write SetRef;
     property Ref     : AxUCString read GetRef write SetRef;
     // No worksheet name.
     property ShortRef: AxUCString read GetShortRef write SetRef;

     property AutoExtend: boolean read FAutoExtend write FAutoExtend;

     property AsFloat[ACol,ARow: integer]: double read GetAsFloat write SetAsFloat;
     property AsString[ACol,ARow: integer]: AxUCString read GetAsString write SetAsString;
     property AsBoolean[ACol,ARow: integer]: boolean read GetAsBoolean write SetAsBoolean;
     property AsError[ACol,ARow: integer]: TXc12cellError read GetAsError write SetAsError;
     property AsBlank[ACol,ARow: integer]: boolean read GetAsBlank write SetAsBlank;
     // True if the cell is blank or there is no cell.
     property AsEmpty[ACol,ARow: integer]: boolean read GetAsEmpty;
     // Returned object must be destroyed.
     property AsSharedItem[ACol,ARow: integer]: TXLSSharedItemsValue read GetAsSharedItem write SetAsSharedItem;

     property AsBooleanCurr: boolean read GetAsBooleanCurr write SetAsBooleanCurr;
     property AsFloatCurr: double read GetAsFloatCurr write SetAsFloatCurr;
     property AsStringCurr: AxUCString read GetAsStringCurr write SetAsStringCurr;
     property AsErrorCurr: TXc12CellError read GetAsErrorCurr write SetAsErrorCurr;

     property AsBooleanRel[ACol,ARow: integer]: boolean read GetAsBooleanRel write SetAsBooleanRel;
     property AsFloatRel[ACol,ARow: integer]: double read GetAsFloatRel write SetAsFloatRel;
     property AsStringRel[ACol,ARow: integer]: AxUCString read GetAsStringRel write SetAsStringRel;
     property AsErrorRel[ACol,ARow: integer]: TXc12CellError read GetAsErrorRel write SetAsErrorRel;
     property AsSharedItemRel[ACol,ARow: integer]: TXLSSharedItemsValue read GetAsSharedItemRel write SetAsSharedItemRel;

     property CellType[ACol,ARow: integer]: TXLSCellType read GetCellType;
     property CellTypeRel[ACol,ARow: integer]: TXLSCellType read GetCellTypeRel;

     property IsDateTime[Col,Row: integer]: boolean read GetIsDateTime;
     property ColCount: integer read GetColCount write SetColCount;
     property RowCount: integer read GetRowCount write SetRowCount;

     property CurrCol : integer read FCurrCol write FCurrCol;
     property CurrRow : integer read FCurrRow write FCurrRow;

     property Expand  : boolean read FExpand write FExpand;

     property Command : TXLSCmdFormatValues read FCommand write FCommand;
     end;

implementation

{ TXLSRelCells }

constructor TXLSRelCells.Create;
begin
  FExpand := True;
end;

constructor TXLSRelCells.Create(ACol, ARow: integer);
begin
  Create;

  FCol1 := ACol;
  FRow1 := ARow;
  FCol2 := ACol;
  FRow2 := ARow;
end;

procedure TXLSRelCells.ApplyCommandRel(ACol, ARow: integer);
begin
  GetRelColRow(ACol,ARow);

  ApplyCommand(FCol1 + ACol, FRow1 + ARow);
end;

procedure TXLSRelCells.ApplyCommandRel(ACol1, ARow1, ACol2, ARow2: integer);
begin
  GetRelColRow(ACol1, ARow1, ACol2, ARow2);

  ApplyCommand(FCol1 + ACol1, FRow1 + ARow1, FCol1 + ACol2, FRow1 + ARow2);
end;

function TXLSRelCells.AsText: AxUCString;
var
  R,C: integer;
begin
  Result := '';

  for R := FRow1 to FRow2 do begin
    for C := FCol1 to FCol2 do
      Result := Result + GetAsFmtString(C,R) + ' ';
  end;

  Result := Trim(Result);
end;

function TXLSRelCells.CloneRow(ARow: integer): TXLSRelCells;
begin
  Result := Clone;

  Result.Col1 := FCol1;
  Result.Row1 := ARow;
  Result.Col2 := FCol2;
  Result.Row2 := ARow;
end;

procedure TXLSRelCells.AutoWidthCols;
var
  c: integer;
begin
  for c := FCol1 to FCol2 do
    AutoWidthCol(c);
end;

procedure TXLSRelCells.ClearAll;
begin
  ClearAll(FCol1,FRow1,FCol2,FRow2);
end;

function TXLSRelCells.CloneCol(ACol: integer): TXLSRelCells;
begin
  Result := Clone;

  Result.Col1 := ACol;
  Result.Row1 := FRow1;
  Result.Col2 := ACol;
  Result.Row2 := FRow2;
end;

constructor TXLSRelCells.Create(ACol1, ARow1, ACol2, ARow2: integer);
begin
  Create;

  FCol1 := ACol1;
  FRow1 := ARow1;
  FCol2 := ACol2;
  FRow2 := ARow2;
end;

destructor TXLSRelCells.Destroy;
begin
  inherited;
end;

procedure TXLSRelCells.DoAutoExtend(ACol, ARow: integer);
begin
  if ACol < FCol1 then
    FCol1 := ACol;
  if ACol > FCol2 then
    FCol2 := ACol;

  if ARow < FRow1 then
    FRow1 := ARow;
  if ARow > FRow2 then
    FRow2 := ARow;
end;

function TXLSRelCells.Equal(ACol, ARow: integer; AValue: TXLSSharedItemsValue): boolean;
begin
  Result := False;

  case GetCellType(ACol,ARow) of
    xctBlank         : Result := (AValue.Type_ = xsivtBlank);
    xctFloat,
    xctFloatFormula  : begin
      if AValue.Type_ = xsivtNumeric then
        Result := GetAsFloat(ACol,ARow) = AValue.AsNumeric
      else if AValue.Type_ = xsivtDate then
        Result := GetAsFloat(ACol,ARow) = AValue.AsDate;
    end;
    xctString,
    xctStringFormula : Result := (AValue.Type_ = xsivtString) and (GetAsString(ACol,ARow) = AValue.AsString);
    xctBoolean,
    xctBooleanFormula: Result := (AValue.Type_ = xsivtBoolean) and (GetAsBoolean(ACol,ARow) = AValue.AsBoolean);
    xctError,
    xctErrorFormula  : Result := (AValue.Type_ = xsivtError) and (GetAsError(ACol,ARow) = AValue.AsError);
  end;
end;

function TXLSRelCells.EqualRel(ACol, ARow: integer; AValue: TXLSSharedItemsValue): boolean;
begin
  Result := Equal(FCol1 + ACol,FRow1 + ARow,AValue);
end;

function TXLSRelCells.GetRef: AxUCString;
begin
  if IsArea then
    Result := SheetName + '!' + AreaToRefStr(FCol1,FRow1,FCol2,FRow2,False,False,False,False)
  else
    Result := SheetName + '!' + ColRowToRefStr(FCol1,FRow1,False,False);
end;

function TXLSRelCells.GetRefAbs: AxUCString;
begin
  if IsList then
    Result := __FRef
  else if IsArea then
    Result := SheetName + '!' + AreaToRefStr(FCol1,FRow1,FCol2,FRow2,True,True,True,True)
  else
    Result := SheetName + '!' + ColRowToRefStr(FCol1,FRow1,True,True);
end;

procedure TXLSRelCells.GetRelColRow(var ACol, ARow: integer);
begin
  if ACol = -1 then ACol := 0 else if ACol = -2 then ACol := FCol2 - FCol1;
  if ARow = -1 then ARow := 0 else if ARow = -2 then ARow := FRow2 - FRow1;
end;

procedure TXLSRelCells.GetRelColRow(var ACol1, ARow1, ACol2, ARow2: integer);
begin
  if ACol1 = -1 then ACol1 := 0 else if ACol1 = -2 then ACol1 := FCol2 - FCol1;
  if ARow1 = -1 then ARow1 := 0 else if ARow1 = -2 then ARow1 := FRow2 - FRow1;

  if ACol2 = -1 then ACol2 := 0 else if ACol2 = -2 then ACol2 := FCol2 - FCol1;
  if ARow2 = -1 then ARow2 := 0 else if ARow2 = -2 then ARow2 := FRow2 - FRow1;
end;

function TXLSRelCells.GetAsBooleanCurr: boolean;
begin
  Result := GetAsBoolean(FCurrCol,FCurrRow);
end;

function TXLSRelCells.GetAsBooleanRel(ACol, ARow: integer): boolean;
begin
  Result := GetAsBoolean(ACol + FCol1,ARow + FRow1);
end;

function TXLSRelCells.GetAsErrorCurr: TXc12CellError;
begin
  Result := GetAsError(FCurrCol,FCurrRow);
end;

function TXLSRelCells.GetAsErrorRel(ACol, ARow: integer): TXc12CellError;
begin
  Result := GetAsError(ACol + FCol1,ARow + FRow1);
end;

function TXLSRelCells.GetAsFloatCurr: double;
begin
  Result := GetAsFloat(FCurrCol,FCurrRow);
end;

function TXLSRelCells.GetAsFloatRel(ACol, ARow: integer): double;
begin
  Result := GetAsFloat(ACol + FCol1,ARow + FRow1);
end;

function TXLSRelCells.GetAsSharedItemRel(ACol,  ARow: integer): TXLSSharedItemsValue;
begin
  Result := GetAsSharedItem(ACol + FCol1,ARow + FRow1);
end;

function TXLSRelCells.GetAsStringCurr: AxUCString;
begin
  Result := GetAsString(FCurrCol,FCurrRow);
end;

function TXLSRelCells.GetAsStringRel(ACol, ARow: integer): AxUCString;
begin
  Result := GetAsString(ACol + FCol1,ARow + FRow1);
end;

function TXLSRelCells.GetCellTypeRel(ACol, ARow: integer): TXLSCellType;
begin
  Result := GetCellType(FCol1 + ACol,FRow1 + ARow);
end;

function TXLSRelCells.GetColCount: integer;
begin
  Result := FCol2 - FCol1 + 1;
end;

function TXLSRelCells.GetCols: integer;
begin
  Result := (FCol2 - FCol1) + 1;
end;

function TXLSRelCells.GetFloatArray(out AArray: TDynSingleArray; AMaxSize: integer): integer;
var
  i: integer;
begin
  if Cols > Rows then begin
    SetLength(AArray,Min(Cols,AMaxSize));

    for i := 0 to High(AArray) do
      AArray[i] := GetAsFloat(FCol1 + i,FRow1);
  end
  else begin
    SetLength(AArray,Min(Rows,AMaxSize));

    for i := 0 to High(AArray) do
      AArray[i] := GetAsFloat(FCol1,FRow1 + i);
  end;

  Result := Length(AArray);
end;

function TXLSRelCells.GetRowCount: integer;
begin
  Result := FRow2 - FRow1 + 1;
end;

function TXLSRelCells.GetRows: integer;
begin
  Result := (FRow2 - FRow1) + 1;
end;

function TXLSRelCells.GetShortRef: AxUCString;
begin
  if IsArea then
    Result := AreaToRefStr(FCol1,FRow1,FCol2,FRow2,False,False,False,False)
  else
    Result := ColRowToRefStr(FCol1,FRow1,False,False);
end;

function TXLSRelCells.GetStrArray(out AArray: TDynStringArray; AMaxSize: integer): integer;
var
  i: integer;
begin
  if Cols > Rows then begin
    SetLength(AArray,Min(Cols,AMaxSize));

    for i := 0 to High(AArray) do
      AArray[i] := GetAsFmtString(FCol1 + i,FRow1);
  end
  else begin
    SetLength(AArray,Min(Rows,AMaxSize));

    for i := 0 to High(AArray) do
      AArray[i] := GetAsFmtString(FCol1,FRow1 + i);
  end;

  Result := Length(AArray);
end;

function TXLSRelCells.Hit(ACol, ARow: integer): boolean;
begin
  Result := (ACol >= FCol1) and (ARow >= FRow1) and (ACol <= FCol2) and (ARow <= FRow2);
end;

procedure TXLSRelCells.IncCol(ABy: integer);
begin
  Inc(FCurrCol);

  if FExpand then
    FCol2 := Max(FCol2,FCurrCol);
end;

procedure TXLSRelCells.IncCols(ABy: integer);
begin
  Inc(FCol2,ABy);
end;

procedure TXLSRelCells.IncRow(ABy: integer);
begin
  Inc(FCurrRow);

  if FExpand then
    FRow2 := Max(FRow2,FCurrRow);
end;

procedure TXLSRelCells.IncRows(ABy: integer);
begin
  Inc(FRow2,ABy);
end;

function TXLSRelCells.IsArea: boolean;
begin
  Result := (FCol1 <> FCol2) or (FRow1 <> FRow2);
end;

function TXLSRelCells.IsList: boolean;
begin
  Result := (__FRef <> '') and (__FRef[1] = '(') and (__FRef[Length(__FRef)] = ')');
end;

function TXLSRelCells.LastNonBlankRow: integer;
var
  c,r: integer;
  Ok : boolean;
begin
  Result := -1;

  if Valid then begin
    for r := FRow1 to FRow2 do begin
      Ok := False;
      for c := FCol1 to FCol2 do begin
        if not GetAsEmpty(c,r) then
          Break;
        Ok := True;
      end;
      if Ok then begin
        Result := r;
        Exit;
      end;
    end;
  end;
end;

function TXLSRelCells.RefTopRow: AxUCString;
begin
  if IsArea then
    Result := SheetName + '!' + AreaToRefStr(FCol1,FRow1,FCol2,FRow1,True,True,True,True)
  else
    Result := SheetName + '!' + ColRowToRefStr(FCol1,FRow1,True,True);
end;

procedure TXLSRelCells.SetArea(ACol1, ARow1, ACol2, ARow2: integer);
begin
  FCol1 := ACol1;
  FRow1 := ARow1;
  FCol2 := ACol2;
  FRow2 := ARow2;
end;

procedure TXLSRelCells.SetAsBooleanCurr(const Value: boolean);
begin
  SetAsBoolean(FCurrCol,FCurrRow,Value);
end;

procedure TXLSRelCells.SetAsBooleanRel(ACol, ARow: integer; const Value: boolean);
begin
  SetAsBoolean(ACol + FCol1, ARow + FRow1,Value);
end;

procedure TXLSRelCells.SetAsErrorCurr(const Value: TXc12CellError);
begin
  SetAsError(FCurrCol,FCurrRow,Value);
end;

procedure TXLSRelCells.SetAsErrorRel(ACol, ARow: integer; const Value: TXc12CellError);
begin
  SetAsError(ACol + FCol1, ARow + FRow1,Value);
end;

procedure TXLSRelCells.SetAsFloatCurr(const Value: double);
begin
  SetAsFloat(FCurrCol,FCurrRow,Value);
end;

procedure TXLSRelCells.SetAsFloatRel(ACol, ARow: integer; const Value: double);
begin
  SetAsFloat(ACol + FCol1, ARow + FRow1,Value);
end;

procedure TXLSRelCells.SetAsSharedItemRel(ACol, ARow: integer; const Value: TXLSSharedItemsValue);
begin
  SetAsSharedItem(ACol + FCol1, ARow + FRow1,Value);
end;

procedure TXLSRelCells.SetAsStringCurr(const Value: AxUCString);
begin
  SetAsString(FCurrCol,FCurrRow,Value);
end;

procedure TXLSRelCells.SetAsStringRel(ACol, ARow: integer; const Value: AxUCString);
begin
  GetRelColRow(ACol,ARow);

  SetAsString(ACol + FCol1, ARow + FRow1,Value);
end;

procedure TXLSRelCells.SetColCount(const Value: integer);
begin
  FCol2 := FCol1 + Value - 1;
end;

procedure TXLSRelCells.SetCols(const Value: integer);
begin
  FCol2 := FCol1 + Value - 1;
end;

procedure TXLSRelCells.SetRef(const Value: AxUCString);
var
  C1,R1,C2,R2: integer;
begin
  if AreaStrToColRow(Value,C1,R1,C2,R2) then begin
    FCol1 := C1;
    FRow1 := R1;
    FCol2 := C2;
    FRow2 := R2;
  end;
end;

procedure TXLSRelCells.SetRowCount(const Value: integer);
begin
  FRow2 := FRow1 + Value - 1;
end;

procedure TXLSRelCells.SetRows(const Value: integer);
begin
  FRow2 := FRow1 + Value - 1;
end;

procedure TXLSRelCells.SettCurrent(ACol, ARow: integer);
begin
  if ACol < 0 then
    FCurrCol := FCol1
  else
    FCurrCol := FCol1 + ACol;

  if ARow < 0 then
    FCurrRow := FRow1
  else
    FCurrRow := FRow1 + ARow;
end;

procedure TXLSRelCells.Set__FRef(const Value: AxUCString);
begin
  __FRef := Value;
end;

function TXLSRelCells.Valid: boolean;
begin
  Result := (FCol1 >= 0) and (FRow1 >= 0) and (FCol2 >= FCol1) and (FRow2 >= FRow1) and (FCol2 <= XLS_MAXCOL) and (FRow2 <= XLS_MAXROW);
end;

end.
