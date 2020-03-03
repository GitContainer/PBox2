unit XLSDbRead5;

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

uses Classes, SysUtils, Contnrs, Db,
     Xc12Utils5, Xc12DataStylesheet5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSReadWriteII5;


//* Event fired when ~[link TXLSDbRead4] reads a cell.
//* ~param Sender The TXLSDbRead4 object.
//* ~param IsFieldNames True if the cell is a field name.
//* ~param Dataset The source data set.
//* ~param FieldIndex Index into the source data set's field list.
//* ~param Col The column in the destination XLSReadWriteII object where the cell is inserted.
//* ~param Row The row in the destination XLSReadWriteII object where the cell is inserted.
type TXLSDbCellEvent = procedure(Sender: TObject; IsFieldNames: boolean; Dataset: TDataSet; FieldIndex,Col,Row: integer) of object;

type TXLSDbRecordEvent = procedure(Sender: TObject; var ReadRecord: boolean) of object;

type TXLSDbColEvent = procedure(Sender: TObject; var AColumn: integer) of object;

//* ~exclude
type TXLSDataSetList = class(TObjectList)
private
     FDataSet: TDataSet;
     FIncludeFieldsInx: array of boolean;
     FExcludeFieldsInx: array of boolean;

     function GetItems(Index: integer): TXLSDataSetList;
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear; override;
     procedure GetDetail(DS: TDataSet; InclFields,ExclFields: TStrings);
     function  Add: TXLSDataSetList;

     property DataSet: TDataSet read FDataSet;
     property Items[Index: integer]: TXLSDataSetList read GetItems; default;
     end;

//* Component for importing data sets (TDataset) into a XLSReadWriteII component.
type TXLSDbRead5 = class(TComponent)
private
    FXLS               : TXLSReadWriteII5;
    FDataSet           : TDataSet;
    FCol               : byte;
    FRow               : word;
    FSheet             : integer;
    FIncludeFields     : TStrings;
    FExcludeFields     : TStrings;
    FIncludeFieldnames : boolean;
    FReadDetailTables  : boolean;
    FIndentDetailTables: boolean;
    FFormatCells       : boolean;
    FXLSDetails        : TXLSDataSetList;
    FDbCellEvent       : TXLSDbCellEvent;
    FDbRecordEvent     : TXLSDbRecordEvent;
    FDbColEvent        : TXLSDbColEvent;

    function  GetCellColors(Index: integer): TXc12IndexColor;
    procedure SetCellColors(Index: integer; const Value: TXc12IndexColor);

    procedure SetExcludeFields(const Value: TStrings);
    procedure SetIncludeFields(const Value: TStrings);
    procedure ReadDataSet(XLSDetails: TXLSDataSetList; var ARow: integer; Level: integer);
    procedure ShowFieldnames(XLSDetails: TXLSDataSetList; var ARow: integer; Level: integer);
    procedure SetXLS(const Value: TXLSReadWriteII5);
public
    //* ~exclude
    constructor Create(AOwner: TComponent); override;
    //* ~exclude
    destructor Destroy; override;
    //* Starts reading the source ~[link Dataset].
    procedure Read;

    //* Cell colors used when ~[link FormatCells] is on. The frist index (#0) is
    //* for the master table, the following for detail tables. There can be
    //* 8 levels, index 0..7
    property CellColors[Index: integer]: TXc12IndexColor read GetCellColors write SetCellColors;
published
    //* Current column where the cells are inserted.
    property Column: byte read FCol write FCol;
    //* Source data set.
    property Dataset: TDataset read FDataset write FDataset;
    //* List of field names in the data set that not shall be imported.
    property ExcludeFields: TStrings read FExcludeFields write SetExcludeFields;
    //* List of field names in the data set that shall be imported. If the list is empty, all fileds are included.
    property IncludeFields: TStrings read FIncludeFields write SetIncludeFields;
    //* Write the names of the fields on the first row.
    property IncludeFieldnames: boolean read FIncludeFieldnames write FIncludeFieldnames;
    //* When True, detail tables are shifted one column to the right for each level.
    property IndentDetailTables: boolean read FIndentDetailTables write FIndentDetailTables;
    //* When True, detail tables are read. When False, only the master table is read.
    property ReadDetailTables: boolean read FReadDetailTables write FReadDetailTables;
    //* When True, the color of the inserted cells are set according to the colors in ~[link CellColors]
    property FormatCells: boolean read FFormatCells write FFormatCells;
    //* Current row where the cells are inserted.
    property Row: word read FRow write FRow;
    //* Index to the current sheet where the cells are inserted.
    property Sheet: integer read FSheet write FSheet;
    //* The destination TXLSReadWriteII2 object.
    property XLS: TXLSReadWriteII5 read FXLS write SetXLS;
    //* Event fired when a cell is inserted. The event is fired after data is written to the cell.
    //* Use this event for user formatting of the cells, or changing the destination on where the cells are inserted.
    property OnDbCell: TXLSDbCellEvent read FDbCellEvent write FDbCellEvent;
    //* Event fired when a record is read.
    property OnDbRecord: TXLSDbRecordEvent read FDbRecordEvent write FDbRecordEvent;
    //* Event fired before a value is written to a cell. Use the event to change the column order in the worksheet.
    property OnDbColumn: TXLSDbColEvent read FDbColEvent write FDbColEvent;
    end;

implementation

var DetailCellColors: array [0..7] of TXc12IndexColor = (xcLightYellow,xcPaleTurquois,xcPaleGreen,xcLightBrown,xc31,xc31,xc31,xc31);

{ TXLSDbRead }

constructor TXLSDbRead5.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FReadDetailTables := True;
  FIndentDetailTables := True;
  FIncludeFieldnames := True;
  FFormatCells := False;
  FIncludeFields := TStringList.Create;
  FExcludeFields := TStringList.Create;
  FXLSDetails := TXLSDataSetList.Create;
end;

destructor TXLSDbRead5.Destroy;
begin
  FIncludeFields.Free;
  FExcludeFields.Free;
  FXLSDetails.Free;
  inherited;
end;

function TXLSDbRead5.GetCellColors(Index: integer): TXc12IndexColor;
begin
  if (Index < 0) or (Index > High(DetailCellColors)) then
    raise Exception.Create('Index out of range');
  Result := DetailCellColors[Index];
end;

procedure TXLSDbRead5.Read;
var
  ARow: integer;
begin
  if FXLS = Nil then
    raise Exception.Create('No TXLSReadWriteII defined');
  if FSheet >= FXLS.Count then
    raise Exception.Create('Sheet index out of range');
  if FDataset = Nil then
    raise Exception.Create('No Dataset defined');
  FXLSDetails.GetDetail(FDataSet,FIncludeFields,FExcludeFields);
  ARow := FRow;
  ReadDataSet(FXLSDetails,ARow,0);
  FXLS[FSheet].CalcDimensions;
end;

procedure TXLSDbRead5.ReadDataSet(XLSDetails: TXLSDataSetList; var ARow: integer; Level: integer);
var
  i,j: integer;
  ACol: integer;
  C: integer;
  ReadRecord: boolean;
begin
  ShowFieldnames(XLSDetails,ARow,Level);
  XLSDetails.DataSet.First;
  while not XLSDetails.DataSet.Eof do begin
    if Assigned(FDbRecordEvent) then begin
      ReadRecord := True;
      FDbRecordEvent(Self,ReadRecord);
      if not ReadRecord then
        Continue;
    end;

    ACol := FCol;
    if FIndentDetailTables then
      Inc(ACol,Level);
    for i := 0 to XLSDetails.DataSet.FieldCount - 1 do begin
      if XLSDetails.FIncludeFieldsInx[i] and not XLSDetails.FExcludeFieldsInx[i] then begin
        if not XLSDetails.DataSet.Fields[i].IsNull then begin
          C := ACol;
          if Assigned(FDbColEvent) then
            FDbColEvent(Self,C);
          case XLSDetails.DataSet.Fields[i].DataType of
            ftString,
            ftFixedChar,
            ftGuid:
                        FXLS[FSheet].AsString[C,ARow] := XLSDetails.DataSet.Fields[i].AsString;
            ftWideString:
                        FXLS[FSheet].AsString[C,ARow] := XLSDetails.DataSet.Fields[i].AsString;
            ftVariant:  FXLS[FSheet].AsVariant[C,ARow] := XLSDetails.DataSet.Fields[i].AsVariant;
//            ftMemo,
//            ftFmtMemo:
//                        FXLS[FSheet].AsString[C,ARow] := XLSDetails.DataSet.Fields[i].AsString;
{$ifdef DELPHI_2006_OR_LATER}
            ftWideMemo:
                        FXLS[FSheet].AsString[C,ARow] := XLSDetails.DataSet.Fields[i].AsString;
{$endif}
            ftSmallint,
            ftInteger,
            ftLargeInt,
            ftWord,
            ftCurrency,
            ftBCD,
{$ifndef DELPHI_5}
            ftFMTBcd,
{$endif}
            ftAutoInc,
            ftFloat:
                        FXLS[FSheet].AsFloat[C,ARow] := XLSDetails.DataSet.Fields[i].AsFloat;
            ftBoolean:
                        FXLS[FSheet].AsBoolean[C,ARow] := XLSDetails.DataSet.Fields[i].AsBoolean;
            ftGraphic: ;
{$ifndef DELPHI_5}
            ftTimestamp,
{$endif}
            ftDate,
            ftTime,
//            ftDateTime: FXLS[FSheet].AsFloat[C,ARow] := XLSDetails.DataSet.Fields[i].AsFloat;
            ftDateTime: FXLS[FSheet].AsDateTime[C,ARow] := XLSDetails.DataSet.Fields[i].AsFloat;
          end;
        end;
        if FFormatCells then begin
          if FXLS[FSheet].Cell[C,ARow] = Nil then
            FXLS[FSheet].AsBlank[C,ARow] := True;
          if Level <= High(DetailCellColors) then
            FXLS[FSheet].Cell[C,ARow].FillPatternForeColor := DetailCellColors[Level]
          else
            FXLS[FSheet].Cell[C,ARow].FillPatternForeColor := DetailCellColors[High(DetailCellColors)];
          if Assigned(FDbCellEvent) then
            FDbCellEvent(Self,False,XLSDetails.DataSet,i,C,ARow);
        end;
        Inc(ACol);
      end;
    end;
    Inc(ARow);
    if ARow >= FXLS.MaxRowCount then
      Break;
    if FReadDetailTables then begin
      for j := 0 to XLSDetails.Count - 1 do begin
        ReadDataSet(XLSDetails[j],ARow,Level + 1);
      end;
    end;
    XLSDetails.DataSet.Next;
  end;
end;

procedure TXLSDbRead5.SetCellColors(Index: integer; const Value: TXc12IndexColor);
begin
  if (Index < 0) or (Index > High(DetailCellColors)) then
    raise Exception.Create('Index out of range');
  DetailCellColors[Index] := Value;
end;

procedure TXLSDbRead5.SetExcludeFields(const Value: TStrings);
begin
  FExcludeFields.Assign(Value);
end;

procedure TXLSDbRead5.SetIncludeFields(const Value: TStrings);
begin
  FIncludeFields.Assign(Value);
end;

procedure TXLSDbRead5.SetXLS(const Value: TXLSReadWriteII5);
begin
  FXLS := Value;
end;

procedure TXLSDbRead5.ShowFieldnames(XLSDetails: TXLSDataSetList; var ARow: integer; Level: integer);
var
  i,ACol: integer;
  C: integer;
begin
  if FIncludeFieldnames then begin
    ACol := FCol;
    if FIndentDetailTables then
      Inc(ACol,Level);
    for i := 0 to XLSDetails.DataSet.FieldCount - 1 do begin
      if XLSDetails.FIncludeFieldsInx[i] and not XLSDetails.FExcludeFieldsInx[i] then begin
          C := ACol;
          if Assigned(FDbColEvent) then
            FDbColEvent(Self,C);
          FXLS[FSheet].AsString[C,ARow] := XLSDetails.DataSet.Fields[i].DisplayName;
          if FFormatCells then begin
            FXLS[FSheet].Cell[C,ARow].FillPatternForeColor := xcOrange;
            FXLS[FSheet].Cell[C,ARow].FontStyle := [xfsBold];
            if Assigned(FDbCellEvent) then
              FDbCellEvent(Self,True,DataSet,i,C,ARow);
          end;
        Inc(ACol);
      end;
    end;
    Inc(ARow);
  end;
end;

{ TXLSDataSetList }

function TXLSDataSetList.Add: TXLSDataSetList;
begin
  Result := TXLSDataSetList.Create;
  inherited Add(Result);
end;

procedure TXLSDataSetList.Clear;
begin
  inherited Clear;
end;

constructor TXLSDataSetList.Create;
begin

end;

destructor TXLSDataSetList.Destroy;
begin
  inherited;
end;

procedure TXLSDataSetList.GetDetail(DS: TDataSet; InclFields,ExclFields: TStrings);
var
  i: integer;
  List: TList;
  XDS: TXLSDataSetList;
  Field: TField;
begin
  SetLength(FIncludeFieldsInx,DS.FieldCount);
  SetLength(FExcludeFieldsInx,DS.FieldCount);
  for i := 0 to DS.FieldCount - 1 do begin
    FIncludeFieldsInx[i] := InclFields.Count <= 0;
    FExcludeFieldsInx[i] := False;
  end;
  for i := 0 to InclFields.Count - 1 do begin
    Field := DS.Fields.FindField(InclFields[i]);
    if Field <> Nil then
      FIncludeFieldsInx[Field.Index] := True;
  end;
  for i := 0 to ExclFields.Count - 1 do begin
    Field := DS.Fields.FindField(ExclFields[i]);
    if Field <> Nil then
      FExcludeFieldsInx[Field.Index] := True;
  end;
  List := TList.Create;
  try
    FDataSet := DS;
{$ifndef DELPHI_5}
{$WARN SYMBOL_DEPRECATED OFF}
{$endif}
    FDataSet.GetDetailDataSets(List);
{$ifndef DELPHI_5}
{$WARN SYMBOL_DEPRECATED ON}
{$endif}
    for i := 0 to List.Count - 1 do begin
      XDS := Add;
      XDS.GetDetail(List[i],InclFields,ExclFields);
    end;
  finally
    List.Free;
  end;
end;

function TXLSDataSetList.GetItems(Index: integer): TXLSDataSetList;
begin
  Result := TXLSDataSetList(inherited Items[Index]);
end;

end.

