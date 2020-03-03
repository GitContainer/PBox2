unit Xc12DataTable5;

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
     xpgPSimpleDOM,
     Xc12Utils5, Xc12Common5, Xc12DataAutofilter5,
     XLSUtils5, XLSClassFactory5;

type TXc12TableType = (x12ttWorksheet,x12ttXml,x12ttQueryTable);
type TXc12TotalsRowFunction = (x12trfNone,x12trfSum,x12trfMin,x12trfMax,x12trfAverage,x12trfCount,x12trfCountNums,x12trfStdDev,x12trfVar,x12trfCustom);
type TXc12XmlDataType =  (x12xdtString,x12xdtNormalizedString,x12xdtToken,x12xdtByte,x12xdtUnsignedByte,x12xdtBase64Binary,x12xdtHexBinary,x12xdtInteger,x12xdtPositiveInteger,x12xdtNegativeInteger,x12xdtNonPositiveInteger,x12xdtNonNegativeInteger,x12xdtInt,x12xdtUnsignedInt,x12xdtLong,x12xdtUnsignedLong,x12xdtShort,x12xdtUnsignedShort,x12xdtDecimal,x12xdtFloat,x12xdtDouble,x12xdtBoolean,x12xdtTime,x12xdtDateTime,x12xdtDuration,x12xdtDate,x12xdtGMonth,x12xdtGYear,x12xdtGYearMonth,x12xdtGDay,x12xdtGMonthDay,x12xdtName,x12xdtQName,x12xdtNCName,x12xdtAnyURI,x12xdtLanguage,x12xdtID,x12xdtIDREF,x12xdtIDREFS,x12xdtENTITY,x12xdtENTITIES,x12xdtNOTATION,x12xdtNMTOKEN,x12xdtNMTOKENS,x12xdtAnyType);

type TXc12TableFormula = class(TXc12Data)
protected
     FArray: boolean;
     FContent: AxUCString;
public
     constructor Create;

     procedure Clear;

     property Array_: boolean read FArray write FArray;
     property Content: AxUCString read FContent write FContent;
     end;

type TXc12XmlColumnPr = class(TXc12Data)
protected
     FMapId: integer;
     FXpath: AxUCString;
     FDenormalized: boolean;
     FXmlDataType: TXc12XmlDataType;
public
     constructor Create;

     procedure Clear;

     property MapId: integer read FMapId write FMapId;
     property Xpath: AxUCString read FXpath write FXpath;
     property Denormalized: boolean read FDenormalized write FDenormalized;
     property XmlDataType: TXc12XmlDataType read FXmlDataType write FXmlDataType;
     end;

type TXc12TableColumn = class(TXc12Data)
protected
     FId: integer;
     FUniqueName: AxUCString;
     FName: AxUCString;
     FTotalsRowFunction: TXc12TotalsRowFunction;
     FTotalsRowLabel: AxUCString;
     FQueryTableFieldId: integer;
     FHeaderRowDxfId: AxUCString;
     FDataDxfId: AxUCString;
     FTotalsRowDxfId: AxUCString;
     FHeaderRowCellStyle: AxUCString;
     FDataCellStyle: AxUCString;
     FTotalsRowCellStyle: AxUCString;
     FCalculatedColumnFormula: TXc12TableFormula;
     FTotalsRowFormula: TXc12TableFormula;
     FXmlColumnPr: TXc12XmlColumnPr;
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear;

     property Id: integer read FId write FId;
     property UniqueName: AxUCString read FUniqueName write FUniqueName;
     property Name: AxUCString read FName write FName;
     property TotalsRowFunction: TXc12TotalsRowFunction read FTotalsRowFunction write FTotalsRowFunction;
     property TotalsRowLabel: AxUCString read FTotalsRowLabel write FTotalsRowLabel;
     property QueryTableFieldId: integer read FQueryTableFieldId write FQueryTableFieldId;
     property HeaderRowDxfId: AxUCString read FHeaderRowDxfId write FHeaderRowDxfId;
     property DataDxfId: AxUCString read FDataDxfId write FDataDxfId;
     property TotalsRowDxfId: AxUCString read FTotalsRowDxfId write FTotalsRowDxfId;
     property HeaderRowCellStyle: AxUCString read FHeaderRowCellStyle write FHeaderRowCellStyle;
     property DataCellStyle: AxUCString read FDataCellStyle write FDataCellStyle;
     property TotalsRowCellStyle: AxUCString read FTotalsRowCellStyle write FTotalsRowCellStyle;
     property CalculatedColumnFormula: TXc12TableFormula read FCalculatedColumnFormula;
     property TotalsRowFormula: TXc12TableFormula read FTotalsRowFormula;
     property XmlColumnPr: TXc12XmlColumnPr read FXmlColumnPr;
     end;

type TXc12TableColumns = class(TObjectList)
private
     function GetItems(Index: integer): TXc12TableColumn;
protected
public
     constructor Create;

     function Add: TXc12TableColumn;

     property Items[Index: integer]: TXc12TableColumn read GetItems; default;
     end;

type TXc12TableStyleInfo = class(TXc12Data)
protected
     FName: AxUCString;
     FShowFirstColumn: boolean;
     FShowLastColumn: boolean;
     FShowRowStripes: boolean;
     FShowColumnStripes: boolean;
public
     constructor Create;

     procedure Clear;

     property Name: AxUCString read FName write FName;
     property ShowFirstColumn: boolean read FShowFirstColumn write FShowFirstColumn;
     property ShowLastColumn: boolean read FShowLastColumn write FShowLastColumn;
     property ShowRowStripes: boolean read FShowRowStripes write FShowRowStripes;
     property ShowColumnStripes: boolean read FShowColumnStripes write FShowColumnStripes;
     end;

type TXc12Table = class(TXc12Data)
protected
     FClassFactory: TXLSClassFactory;

     FId: integer;
     FName: AxUCString;
     FDisplayName: AxUCString;
     FComment: AxUCString;
     FRef: TXLSCellArea;
     FTableType: TXc12TableType;
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
     FAutoFilter: TXc12AutoFilter;
     FSortState: TXc12SortState;
     FTableColumns: TXc12TableColumns;
     FTableStyleInfo: TXc12TableStyleInfo;

     FQueryTable: TXpgSimpleDOM;
public
     constructor Create(AClassFactory: TXLSClassFactory);
     destructor Destroy; override;

     procedure Clear;

     function Hit(const ACol,ARow: integer): boolean;

     property Id: integer read FId write FId;
     property Name: AxUCString read FName write FName;
     property DisplayName: AxUCString read FDisplayName write FDisplayName;
     property Comment: AxUCString read FComment write FComment;
     property Ref: TXLSCellArea read FRef write FRef;
     property TableType: TXc12TableType read FTableType write FTableType;
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
     property AutoFilter: TXc12AutoFilter read FAutoFilter;
     property SortState: TXc12SortState read FSortState;
     property TableColumns: TXc12TableColumns read FTableColumns;
     property TableStyleInfo: TXc12TableStyleInfo read FTableStyleInfo;

     property QueryTable: TXpgSimpleDOM read FQueryTable;
     end;

type TXc12Tables = class(TObjectList)
private
     function GetItems(Index: integer): TXc12Table;
protected
     FClassFactory: TXLSClassFactory;

     function FixupColNameFromFormula(AName: AxUCString): AxUCString;
public
     constructor Create(AClassFactory: TXLSClassFactory);

     function  Add: TXc12Table;
     function  Find(const AName: AxUCString): integer; overload;
     function  Find(const ACol,ARow: integer): integer; overload;
     function  FindCol(const AName: AxUCString): integer;

     property Items[Index: integer]: TXc12Table read GetItems; default;
     end;

implementation

{ TXc12Tables }

function TXc12Tables.Add: TXc12Table;
begin
  Result := TXc12Table.Create(FClassFactory);
  inherited Add(Result);
end;

constructor TXc12Tables.Create(AClassFactory: TXLSClassFactory);
begin
  inherited Create;

  FClassFactory := AClassFactory;
end;

function TXc12Tables.Find(const AName: AxUCString): integer;
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

function TXc12Tables.Find(const ACol, ARow: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Hit(ACol,ARow) then
      Exit;
  end;
  Result := -1;
end;

function TXc12Tables.FindCol(const AName: AxUCString): integer;
var
  i: integer;
  S: AxUCString;
  T: TXc12Table;
begin
  S := FixupColNameFromFormula(AName);
  for i := 0 to Count - 1 do begin
    T := Items[i];
    for Result := 0 to T.TableColumns.Count - 1 do begin
      if S = AnsiUppercase(T.TableColumns[Result].Name) then
        Exit;
    end;
  end;
  Result := -1;
end;

function TXc12Tables.FixupColNameFromFormula(AName: AxUCString): AxUCString;
var
  i,j: integer;
begin
  // Remove single quotes as these are not stored in the col name.
  SetLength(Result,Length(AName));
  i := 1;
  j := 0;
  while i <= Length(AName) do begin
    if AName[i] <> '''' then begin
      Inc(j);
      Result[j] := AName[i];
    end
    else if (i < Length(AName)) and (AName[i + 1] = '''') then begin
      Inc(j);
      Result[j] := '''';
      Inc(i);
    end;
    Inc(i);
  end;
  SetLength(Result,j);
  Result := AnsiUppercase(Result);
end;

function TXc12Tables.GetItems(Index: integer): TXc12Table;
begin
  Result := TXc12Table(inherited Items[Index]);
end;

{ TXc12Table }

procedure TXc12Table.Clear;
begin
  FHeaderRowCount := 1;
  FInsertRow := False;
  FInsertRowShift := False;
  FTotalsRowCount := 0;
  FTotalsRowShown := True;
  FPublished := False;
  FTableType := x12ttWorksheet;
  FAutoFilter.Clear;
  FSortState.Clear;
  FTableColumns.Clear;
  FTableStyleInfo.Clear;
  FQueryTable.Clear;
end;

constructor TXc12Table.Create(AClassFactory: TXLSClassFactory);
begin
  FClassFactory := AClassFactory;

  FAutoFilter := TXc12AutoFilter.Create(FClassFactory);
  FSortState := TXc12SortState.Create;
  FTableColumns := TXc12TableColumns.Create;
  FTableStyleInfo := TXc12TableStyleInfo.Create;
  FQueryTable := TXpgSimpleDOM.Create;
  Clear;
end;

destructor TXc12Table.Destroy;
begin
  FAutoFilter.Free;
  FSortState.Free;
  FTableColumns.Free;
  FTableStyleInfo.Free;
  FQueryTable.Free;
  inherited;
end;

function TXc12Table.Hit(const ACol, ARow: integer): boolean;
begin
  Result := (ACol >= FRef.Col1) and (ACol <= FRef.Col2) and (ARow >= FRef.Row1) and (ARow <= FRef.Row2);
end;

{ TXc12TableColumns }

function TXc12TableColumns.Add: TXc12TableColumn;
begin
  Result := TXc12TableColumn.Create;
  inherited Add(Result);
end;

constructor TXc12TableColumns.Create;
begin
  inherited Create;
end;

function TXc12TableColumns.GetItems(Index: integer): TXc12TableColumn;
begin
  Result := TXc12TableColumn(inherited Items[Index]);
end;

{ TXc12TableColumn }

procedure TXc12TableColumn.Clear;
begin
  FCalculatedColumnFormula.Clear;
  FTotalsRowFormula.Clear;
  FXmlColumnPr.Clear;
  FId := 0;
  FUniqueName := '';
  FName := '';
  FTotalsRowFunction := x12trfNone;
  FTotalsRowLabel := '';
  FQueryTableFieldId := 0;
  FHeaderRowDxfId := '';
  FDataDxfId := '';
  FTotalsRowDxfId := '';
  FHeaderRowCellStyle := '';
  FDataCellStyle := '';
  FTotalsRowCellStyle := '';
end;

constructor TXc12TableColumn.Create;
begin
  FCalculatedColumnFormula := TXc12TableFormula.Create;
  FTotalsRowFormula := TXc12TableFormula.Create;
  FXmlColumnPr := TXc12XmlColumnPr.Create;
  Clear;
end;

destructor TXc12TableColumn.Destroy;
begin
  Clear;
  FCalculatedColumnFormula.Free;
  FTotalsRowFormula.Free;
  FXmlColumnPr.Free;
  inherited;
end;

{ TXc12TableStyleInfo }

procedure TXc12TableStyleInfo.Clear;
begin
  FName := '';
  FShowFirstColumn := False;
  FShowLastColumn := False;
  FShowRowStripes := False;
  FShowColumnStripes := False;
end;

constructor TXc12TableStyleInfo.Create;
begin
  Clear;
end;

{ TXc12TableFormula }

procedure TXc12TableFormula.Clear;
begin
  FArray := False;
  FContent := '';
end;

constructor TXc12TableFormula.Create;
begin
  Clear;
end;

{ TXc12XmlColumnPr }

procedure TXc12XmlColumnPr.Clear;
begin
  FMapId := 0;
  FXpath := '';
  FDenormalized := False;
  FXmlDataType := x12xdtString;
end;

constructor TXc12XmlColumnPr.Create;
begin
  Clear;
end;

end.
