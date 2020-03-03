unit XLSPivotTables5;

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

uses Classes, SysUtils, Contnrs, Math, IniFiles,
     xpgParserPivot,
     xc12Utils5, Xc12Manager5, Xc12DataWorksheet5,
     XLSUtils5, XLSSharedItems5, XLSRelCells5, XLSCmdFormatValues5;

type TXLSPivotFunc = (xpfSum,xpfCount,xpfAverage,xpfMax,xpfMin,xpfProduct,xpfCountNumbers,xpfStdDev,xpfStdDevp,xpfVar,xpfVarp,xpfDefault,xpfNone);

const XLSPivotFuncStr: array[TXLSPivotFunc] of string =
('Sum','Count','Average','Max','Min','Product','Count numbers','StdDev','StdDevp','Var','Varp','Default','None');

type TXLSPivotResultCellType = (xprctValue,xprctSumCol,xprctSumRow,xprctTotal);

type TXLSPivotResultCell = record
     Value   : double;
     CellType: TXLSPivotResultCellType;
     Level   : integer;
     end;

type TXLSPivotCompare = record
     Value: TXLSSharedItemsValue;
     X    : integer;
     end;

type TXLSPivotCondition = array of TXLSPivotCompare;

type TXLSPivotTableResult = class(TObject)
private
     function  GetCells(ACol, ARow: integer): double;
     procedure SetCells(ACol, ARow: integer; const Value: double);
protected
     FCells    : array of array of TXLSPivotResultCell;
     FVals     : array of double;
     FValsCount: integer;
     FColCount : integer;
     FRowCount : integer;
public
     constructor Create;

     procedure Clear;
     procedure ClearValues;

     procedure Resize(AColCount,ARowCount: integer);

     procedure Val(ACol,ARow: integer; AValue: double);

     procedure BeginCollect;
     procedure Collect(AVal: double);

     procedure Calc(AFunc: TXLSPivotFunc; ACol,ARow: integer; AValue: double);

     procedure CopyTo(ACol,ARow: integer; AColor: integer; ARef: TXLSRelCells);

     procedure ColTotal(ARow: integer);
     procedure RowSum(ACol: integer);

     procedure SumCells(AMaxRowLevel: integer; AColGrandTotals, ARowGrandTotals: boolean);

     property ColCount: integer read FColCount;
     property RowCount: integer read FRowCount;
     property Cells[ACol,ARow: integer]: double read GetCells write SetCells; default;
     end;

type TXLSPivotField = class(TIndexObject)
private
     function  GetName: AxUCString;
     function  GetFunction: TXLSPivotFunc;
     procedure SetFunction(const Value: TXLSPivotFunc);
     function  GetFunctionName: AxUCString;
     function  GetUserName: AxUCString;
     procedure SetUserName(const Value: AxUCString);
     function  GetDisplayName: AxUCString;
protected
     FCacheField: TCT_CacheField;
     FPivFld    : TCT_PivotField;
     FFilter    : TXLSUniqueSharedItemsValues;
public
     constructor Create;
     destructor Destroy; override;

     procedure Sort;

     property Name       : AxUCString read GetName;
     property UserName   : AxUCString read GetUserName write SetUserName;
     property DisplayName: AxUCString read GetDisplayName;
     property Func       : TXLSPivotFunc read GetFunction write SetFunction;
     property FuncName   : AxUCString read GetFunctionName;
     property CacheField : TCT_CacheField read FCacheField;
     property Filter     : TXLSUniqueSharedItemsValues read FFilter;
     end;

type TXLSPivotFields = class(TIndexObjectList)
private
     function  GetItems(Index: integer): TXLSPivotField;
protected
public
     constructor Create;

     function  Add: TXLSPivotField; overload;

     function  Find(AName: AxUCString): TXLSPivotField; overload;
     function  Find(AField: TXLSPivotField): integer; overload;

     property Items[Index: integer]: TXLSPivotField read GetItems; default;
     end;

type TXLSPivotFieldsDest = class(TObjectList)
private
     function GetItems(Index: integer): TXLSPivotField;
protected
public
     constructor Create;

     procedure Add(AField: TXLSPivotField); overload;

     function  Find(AField: TXLSPivotField): integer; overload;
     function  Find(AName: AxUCString): integer; overload;

     function  Remove(AField: TXLSPivotField): boolean; overload;
     function  Remove(AName: AxUCString): boolean; overload;

     property Items[Index: integer]: TXLSPivotField read GetItems; default;
     end;

type TXLSPivotRCItem = class(TObject)
protected
     FParent  : TXLSPivotRCItem;
     FValue   : TXLSSharedItemsValue;
     FCacheX  : integer;
     FExpanded: boolean;
     FChilds  : THashedStringList;

     function  GetItems(Index: integer): TXLSPivotRCItem;
     procedure DoMaxLevel(var ALevel, AMaxLevel: integer);
public
     constructor Create(AParent: TXLSPivotRCItem; AValue: TXLSSharedItemsValue; ACacheX: integer);
     destructor Destroy; override;

     procedure Clear;

     procedure Sort;

     function Count: integer;
     function TotalCount: integer;
     function MaxLevelCount: integer;
     function MaxLevel: integer;

     function Add(AValue: TXLSSharedItemsValue; ACacheX: integer): TXLSPivotRCItem;

     function Find(AValue: AxUCString): TXLSPivotRCItem;

     property Parent  : TXLSPivotRCItem read FParent;
     property CacheX  : integer read FCacheX;
     property Value   : TXLSSharedItemsValue read FValue;
     property Items[Index: integer]: TXLSPivotRCItem read GetItems; default;
     end;

type TXLSPivotTable = class(TObject)
private
     function  GetColsTotals: boolean;
     function  GetRowsTotals: boolean;
     procedure SetColsTotals(const Value: boolean);
     procedure SetRowsTotals(const Value: boolean);
     function  GetName: AxUCString;
     procedure SetName(const Value: AxUCString);
protected
     FTableDef    : TCT_PivotTableDefinition;

     FFields      : TXLSPivotFields;

     FReportFilter: TXLSPivotFieldsDest;
     FColumnLabels: TXLSPivotFieldsDest;
     FRowLabels   : TXLSPivotFieldsDest;
     FDataValues  : TXLSPivotFieldsDest;

     FSrcRef      : TXLSRelCells;
     FDestRef     : TXLSRelCells;

     FRows        : TXLSPivotRCItem;
     FCols        : TXLSPivotRCItem;

     FSubCols     : TList;

     FCalcRowGrand: boolean;

     FMaxRowLevel : integer;
     FMaxColLevel : integer;

     FPivResult   : TXLSPivotTableResult;

     FColorCols   : integer;
     FColorSubs   : integer;

     FFixupEdit   : boolean;

     FDirty       : boolean;

     procedure CheckUsed;

     procedure DoFields;
     procedure DoRowFields;
     procedure DoColFields;
     procedure DoDataValues;
     procedure DoRowColItems(ARow, AIndex: integer; ALabels: TXLSPivotFieldsDest; ARCItem: TXLSPivotRCItem; var AMaxLevel: integer); overload;
     procedure DoColValues(var AR,AC: integer; ACols: TXLSPivotRCItem; ALevel: integer; ASubtotals: boolean);
     procedure DoRowValues(var AR,AC: integer; ARows: TXLSPivotRCItem; ALevel: integer; ACondition: TXLSPivotCondition);
     procedure DoValues(ASubtotals: boolean);
     procedure DoColHeaders(var AIndex: integer; ALevel: integer; AItem: TXLSPivotRCItem; ASubtotals: boolean);
     procedure DoRowHeaders(var AIndex: integer; ALevel: integer; AItem: TXLSPivotRCItem);
     procedure DoWriteRowItems(ADefItems: TCT_RowColItems; ARCItems: TXLSPivotRCItem; ALevel: integer; AHasTotal: boolean);
     procedure DoWriteColItems(ADefItems: TCT_RowColItems; ARCItems: TXLSPivotRCItem; ALevel: integer; AHasTotal: boolean);
     procedure DoColDataValues;
     procedure DoFixupEdit;
     procedure DoEpilogue;

     function  ColHeaderRows: integer;

     function  HasLabels: boolean;
public
     constructor Create(ATblDef: TCT_PivotTableDefinition);
     destructor Destroy; override;

     procedure Clear;

     procedure ClearReport;

     function  Make: boolean;

     property Name           : AxUCString read GetName write SetName;

     property TableDef       : TCT_PivotTableDefinition read FTableDef;

     property Fields         : TXLSPivotFields read FFields;

     property ReportFilter   : TXLSPivotFieldsDest read FReportFilter;
     property ColumnLabels   : TXLSPivotFieldsDest read FColumnLabels;
     property RowLabels      : TXLSPivotFieldsDest read FRowLabels;
     property DataValues     : TXLSPivotFieldsDest read FDataValues;

     property RowsGrandTotals: boolean read GetRowsTotals write SetRowsTotals;
     property ColsGrandTotals: boolean read GetColsTotals write SetColsTotals;

     property PivResult      : TXLSPivotTableResult read FPivResult;

     property DestRef        : TXLSRelCells read FDestRef write FDestRef;

     property ColorCols      : integer read FColorCols write FColorCols;
     property ColorSubtotals : integer read FColorSubs write FColorSubs;

     // Not set by the class. Used by others
     // Initiated to True. Set to False when Make is called.
     property Dirty          : boolean read FDirty write FDirty;
     end;

type TXLSPivotTables = class(TObjectList)
private
     function  GetItems(Index: integer): TXLSPivotTable;
protected
     FManager : TXc12Manager;
     FOwnerRef: TXLSRelCells;
public
     constructor Create(AManager: TXc12Manager; AOwnerRef: TXLSRelCells);
     destructor Destroy; override;

     function  Add(ATblDef: TCT_PivotTableDefinition): TXLSPivotTable;

     function  Hit(ACol, ARow: integer): TXLSPivotTable;

     function  CreateTable(ASourceRef,ADestRef: TXLSRelCells): TXLSPivotTable;

     function  EditTable(ATableDef: TCT_pivotTableDefinition): TXLSPivotTable;

     property Items[Index: integer]: TXLSPivotTable read GetItems; default;
     end;

implementation

{ TXLSPivotFields }

function TXLSPivotFields.Add: TXLSPivotField;
begin
  Result := TXLSPivotField.Create;

  inherited Add(Result);
end;

constructor TXLSPivotFields.Create;
begin
  inherited Create;
end;

function TXLSPivotFields.Find(AName: AxUCString): TXLSPivotField;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].DisplayName = AName then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSPivotFields.Find(AField: TXLSPivotField): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result] = AField then
      Exit;
  end;
  Result := -1;
end;

function TXLSPivotFields.GetItems(Index: integer): TXLSPivotField;
begin
  Result := TXLSPivotField(inherited Items[Index]);
end;

{ TXLSPivotTables }

function TXLSPivotTables.Add(ATblDef: TCT_PivotTableDefinition): TXLSPivotTable;
begin
  Result := TXLSPivotTable.Create(ATblDef);

  inherited Add(Result);
end;

constructor TXLSPivotTables.Create(AManager: TXc12Manager; AOwnerRef: TXLSRelCells);
begin
  inherited Create;

  FManager := AManager;
  FOwnerRef := AOwnerRef;
end;

destructor TXLSPivotTables.Destroy;
begin
  FOwnerRef.Free;

  inherited;
end;

function TXLSPivotTables.EditTable(ATableDef: TCT_pivotTableDefinition): TXLSPivotTable;
var
  i,j: integer;
  Fld: TXLSPivotField;
begin
  ATableDef.Cache.Records.Clear;
  ATableDef.Cache.RecordCount := 0;

  Result := Add(ATableDef);

  for i := 0 to ATableDef.PivotFields.Count - 1 do begin
    Fld := Result.Fields.Add;
    Fld.FCacheField := ATableDef.Cache.CacheFields[i];
    Fld.FPivFld := Result.FTableDef.PivotFields[i];

    if not Result.FFixupEdit and (ATableDef.PivotFields[i].Items <> Nil) then begin
      for j := 0 to ATableDef.PivotFields[i].Items.Count - 1 do begin
        if ATableDef.PivotFields[i].Items[j].H then begin
          Result.FFixupEdit := True;
          Break;
        end;
      end;
    end;
  end;

  if ATableDef.RowFields <> Nil then begin
    for i := 0 to ATableDef.RowFields.Count - 1 do begin
      Fld := Result.Fields[ATableDef.RowFields[i].X];
      Result.RowLabels.Add(Fld);
    end;
  end;

  if ATableDef.ColFields <> Nil then begin
    for i := 0 to ATableDef.ColFields.Count - 1 do begin
      // Negative values (-2) seems to be used when there are no columns, but
      // the data values are used for columns instead. This is undocummented.
      if ATableDef.ColFields[i].x < 0 then begin
        ATableDef.Delete_ColFields;
        ATableDef.Delete_ColItems;
        Result.ColumnLabels.Clear;
        Break;
      end;
      Fld := Result.Fields[ATableDef.ColFields[i].X];
      Result.ColumnLabels.Add(Fld);
    end;
  end;

  if ATableDef.DataFields <> Nil then begin
    for i := 0 to ATableDef.DataFields.Count - 1 do begin
      Fld := Result.Fields[ATableDef.DataFields[i].Fld];
      Result.DataValues.Add(Fld);
    end;
  end;
end;

function TXLSPivotTables.GetItems(Index: integer): TXLSPivotTable;
begin
  Result := TXLSPivotTable(inherited Items[Index]);
end;

function TXLSPivotTables.Hit(ACol, ARow: integer): TXLSPivotTable;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].DestRef.Hit(ACol,ARow) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSPivotTables.CreateTable(ASourceRef,ADestRef: TXLSRelCells): TXLSPivotTable;
var
  i    : integer;
  Sheet: TXc12DataWorksheet;
  Cache: TCT_pivotCacheDefinition;
  Fld  : TXLSPivotField;
begin
  Result := Nil;

  Sheet := FManager.Worksheets[FOwnerRef.SheetIndex];

  Cache := FManager.Workbook.PivotCaches.Add;
  if Cache.AddFields(ASourceRef) then begin
    Result := Add(Sheet.PivotTables.Add);

    Result.FDestRef := ADestRef;

    Result.FTableDef.Cache := Cache;

    Result.FTableDef.Name := 'PivotTable' + IntToStr(Count + 1);

    Result.FTableDef.Create_Location;
    Result.FTableDef.Location.RCells := Result.FDestRef;
    Result.FTableDef.Location.FirstHeaderRow := 1;
    Result.FTableDef.Location.FirstDataRow := 1;
    Result.FTableDef.Location.FirstDataCol := 1;

    Result.FTableDef.Create_PivotFields;

    for i := 0 to Cache.CacheFields.Count - 1 do begin
      Fld := Result.Fields.Add;
      Fld.FCacheField := Cache.CacheFields[i];
      Fld.FPivFld := Result.FTableDef.PivotFields.Add;
    end;
  end
  else
    ADestRef.Free;
end;

{ TXLSPivotTable }

procedure TXLSPivotTable.CheckUsed;
var
  i: integer;
begin
  for i := 0 to FFields.Count - 1 do
    FFields[i].FCacheField.SharedItems.Values.Used := False;

  for i := 0 to FFields.Count - 1 do begin
    if not FFields[i].FCacheField.SharedItems.Values.Used then
      FFields[i].FCacheField.SharedItems.Values.Used := (FRowLabels.Find(FFields[i]) >= 0);
    if not FFields[i].FCacheField.SharedItems.Values.Used then
      FFields[i].FCacheField.SharedItems.Values.Used := (FColumnLabels.Find(FFields[i]) >= 0);
    if not FFields[i].FCacheField.SharedItems.Values.Used then
      FFields[i].FCacheField.SharedItems.Values.Used := (FReportFilter.Find(FFields[i]) >= 0);
//    if not FFields[i].FCacheField.SharedItems.Values.Used then
//      FFields[i].FCacheField.SharedItems.Values.Used := (FSumValues.Find(FFields[i]) >= 0);
  end;
end;

procedure TXLSPivotTable.Clear;
begin
  FRows.Clear;
  FCols.Clear;

  FSubCols.Clear;

  FPivResult.Clear;

  FMaxRowLevel := 0;
  FMaxColLevel := 0;
end;

procedure TXLSPivotTable.ClearReport;
begin
  FReportFilter.Clear;
  FColumnLabels.Clear;
  FRowLabels.Clear;
  FDataValues.Clear;

  FPivResult.Resize(0,0);
end;

function TXLSPivotTable.ColHeaderRows: integer;
begin
  Result := Max(FMaxColLevel,1) + 1;
  if (FMaxColLevel > 0) and (FDataValues.Count > 1) then
    Inc(Result);
end;

constructor TXLSPivotTable.Create(ATblDef: TCT_PivotTableDefinition);
begin
  FTableDef       := ATblDef;

  FFields       := TXLSPivotFields.Create;
  FReportFilter := TXLSPivotFieldsDest.Create;
  FColumnLabels := TXLSPivotFieldsDest.Create;
  FRowLabels    := TXLSPivotFieldsDest.Create;
  FDataValues    := TXLSPivotFieldsDest.Create;

  FRows         := TXLSPivotRCItem.Create(Nil,Nil,-1);
  FCols         := TXLSPivotRCItem.Create(Nil,Nil,-1);

  FSubCols      := TList.Create;

  FPivResult    := TXLSPivotTableResult.Create;

  FColorCols     := $00B9E6FD;
  FColorSubs     := $00C0C0C0;

  FDirty         := True;
end;

destructor TXLSPivotTable.Destroy;
begin
  FFields.Free;
  FReportFilter.Free;
  FColumnLabels.Free;
  FRowLabels.Free;
  FDataValues.Free;

  FDestRef.Free;

  FRows.Free;
  FCols.Free;

  FSubCols.Free;

  FPivResult.Free;

  inherited;
end;

procedure TXLSPivotTable.DoColDataValues;
var
  i : integer;
  iI: TCT_I;
begin
  FTableDef.Create_ColFields;
  FTableDef.ColFields.Add.X := -2;

  if FTableDef.ColItems = Nil then begin
    FTableDef.Create_ColItems;
    for i := 0 to FDataValues.Count - 1 do begin
      iI := FTableDef.ColItems.Add;
      iI.I := i;
      iI.Create_XXpgList;
      iI.XXpgList.Add.V := i;
    end;
  end;
end;

procedure TXLSPivotTable.DoColFields;
var
  i    : integer;
  Field: TXLSPivotField;
  Fld  : TCT_Field;
begin
  if FTableDef.ColFields <> Nil then
    FTableDef.ColFields.Clear
  else
    FTableDef.Create_ColFields;

  for i := 0 to FColumnLabels.Count - 1 do begin
    Field := FColumnLabels[i];

    Fld := FTableDef.ColFields.Add;
    Fld.X := Field.Index;

    FTableDef.Create_ColItems;
  end;
end;

procedure TXLSPivotTable.DoEpilogue;
begin
  FTableDef.Create_PivotTableStyleInfo;
  FTableDef.PivotTableStyleInfo.Name := 'PivotStyleLight16';
  FTableDef.PivotTableStyleInfo.ShowColHeaders := True;
  FTableDef.PivotTableStyleInfo.ShowRowHeaders := True;
  FTableDef.PivotTableStyleInfo.ShowLastColumn := True;
end;

procedure TXLSPivotTable.DoFields;
var
  i,j  : integer;
  Field: TXLSPivotField;
  Item : TCT_Item;
begin
  for i := 0 to FFields.Count - 1 do begin
    Field := FFields[i];
    if Field.FCacheField.SharedItems.Values.Used then begin
      Field.FPivFld.Create_Items;

      if FRowLabels.Find(Field) >= 0 then
        Field.FPivFld.Axis := staAxisRow
      else if FColumnLabels.Find(Field) >= 0 then
        Field.FPivFld.Axis := staAxisCol;

      Field.FPivFld.Items.Clear;
      for j := 0 to Field.FCacheField.SharedItems.Values.Count - 1 do begin
        Item := Field.FPivFld.Items.Add;
        Item.X := j;
      end;
      Item := Field.FPivFld.Items.Add;
      Item.T := stitDefault;

      Field.Sort;

      // Must do after Sort.
      if Field.Filter.Count > 0 then begin
        for j := 0 to Field.FCacheField.SharedItems.Values.ValueList.Count - 1 do begin
          Item := Field.FPivFld.Items[j];

          if not Field.Filter.Find(Field.FCacheField.SharedItems.Values.ValueList[j]) then
            Item.H := True;
        end;
      end;
    end
    else if FDataValues.Find(Field) >= 0 then
      Field.FPivFld.DataField := True;
  end;
end;

procedure TXLSPivotTable.DoFixupEdit;
var
  i,j : integer;
  Fld : TCT_PivotField;
  List: TXLSUniqueSharedItemsValues;
begin
  if FFixupEdit then begin
    for i := 0 to FTableDef.PivotFields.Count - 1 do begin
      Fld := FTableDef.PivotFields[i];
      if Fld.Items <> Nil then begin
        for j := 0 to Fld.Items.Count - 1 do begin
          List := FFields[i].FCacheField.SharedItems.Values.ValueList;
          if not Fld.Items[j].H and (Fld.Items[j].X < List.Count) then
            FFields[i].Filter.Add(List[Fld.Items[j].X].Clone);
        end;
      end;
    end;
    FFixupEdit := False;
  end;
end;

procedure TXLSPivotTable.DoRowFields;
var
  i    : integer;
  Field: TXLSPivotField;
  Fld  : TCT_Field;
begin
  if FTableDef.RowFields <> Nil then
    FTableDef.RowFields.Clear
  else
    FTableDef.Create_RowFields;


  for i := 0 to FRowLabels.Count - 1 do begin
    Field := FRowLabels[i];

    Fld := FTableDef.RowFields.Add;
    Fld.X := Field.Index;

    FTableDef.Create_RowItems;
  end;
end;

procedure TXLSPivotTable.DoRowHeaders(var AIndex: integer; ALevel: integer; AItem: TXLSPivotRCItem);
var
  i: integer;
begin
  for i := 0 to AItem.Count - 1 do begin
    FDestRef.AsSharedItemRel[0,AIndex] := AItem[i].Value;
    if (ALevel > 0) and (AItem[i].Value.Type_ = xsivtString) then begin
      FDestRef.BeginCommand;
      FDestRef.Command.Add(xcfcIndent,ALevel * 3);
      FDestRef.ApplyCommandRel(0,AIndex);
    end;
    if AItem[i].FChilds <> Nil then begin
      Inc(AIndex);
      DoRowHeaders(AIndex,ALevel + 1,AItem[i]);
    end;
    Inc(AIndex);
  end;
  Dec(AIndex);
end;

procedure TXLSPivotTable.DoRowValues(var AR,AC: integer; ARows: TXLSPivotRCItem; ALevel: integer; ACondition: TXLSPivotCondition);
var
  i,j : integer;
  r,c : integer;
  Ok  : boolean;
  Cnt : integer;
  Item: TXLSPivotRCItem;
  Val : TXLSSharedItemsValue;
begin
  Cnt := 0;
  if ALevel = FMaxRowLevel then begin
    i := 0;
    Item := ARows;
    while Item.Parent <> Nil do begin
      ACondition[FMaxColLevel + i].Value := Item.Value;
      ACondition[FMaxColLevel + i].X := Item.CacheX;

      Item := Item.Parent;
      Inc(i);
    end;

    for r := 0 to FSrcRef.RowCount - 1 do begin
      Ok := True;

      for c := 0 to FSrcRef.ColCount - 1 do begin
        if FFields[c].Filter.Count > 0 then begin
          Val := FSrcRef.AsSharedItemRel[c,r];
          try
            Ok := FFields[c].Filter.Find(Val);
            if not Ok then
              Break;
          finally
            Val.Free;
          end;
        end;
      end;

      if Ok then begin
        for i := 0 to High(ACondition) do begin
          if FFields[ACondition[i].X].Filter.Count > 0 then
            Ok := FFields[ACondition[i].X].Filter.Find(ACondition[i].Value)
          else
            Ok := True;

          if not FSrcRef.EqualRel(ACondition[i].X,r,ACondition[i].Value) then begin
            Ok := False;
            Break;
          end;
        end;
      end;

      if Ok then begin
        Inc(Cnt);

        for i := 0 to FDataValues.Count - 1 do begin
          case FDataValues[i].Func of
            xpfCount  : FPivResult.Val(AC + i,AR - 1,Cnt);
            xpfStdDev,
            xpfStdDevp,
            xpfVar,
            xpfVarp   : FPivResult.Collect(FSrcRef.AsFloatRel[FDataValues[i].Index,r]);
            else        FPivResult.Calc(FDataValues[i].Func,AC + i,AR - 1,FSrcRef.AsFloatRel[FDataValues[i].Index,r]);
          end;
        end;
      end;
    end;
    for i := 0 to FDataValues.Count - 1 do begin
      case FDataValues[i].Func of
        xpfStdDev,
        xpfStdDevp,
        xpfVar,
        xpfVarp: FPivResult.Calc(FDataValues[i].Func,AC + i,AR - 1,0);
      end;
    end;
  end
  else if ARows.Count > 0 then begin
    for i := 0 to ARows.Count - 1 do begin
      if (ALevel + 1) < FMaxRowLevel then begin
        for j := 0 to FDataValues.Count - 1 do begin
          FPivResult.FCells[AC + j,AR].CellType := xprctSumCol;
          FPivResult.FCells[AC + j,AR].Level := ALevel;

//          if ColsGrandTotals then begin
//            FPivResult.FCells[AC + j + 1,AR].CellType := xprctSumCol;
//            FPivResult.FCells[AC + j + 1,AR].Level := ALevel;
//          end;

        end;
      end;
      for j := 0 to FDataValues.Count - 1 do begin
        case FDataValues[j].Func of
          xpfMax    : FPivResult.Val(AC + j,AR,MINDOUBLE);
          xpfMin    : FPivResult.Val(AC + j,AR,MAXDOUBLE);
          xpfStdDev,
          xpfStdDevp,
          xpfVar,
          xpfVarp   : FPivResult.BeginCollect;
        end;
      end;
      Inc(AR);
      DoRowValues(AR,AC,ARows[i],ALevel + 1,ACondition);
    end;
  end;
end;

procedure TXLSPivotTable.DoRowColItems(ARow, AIndex: integer; ALabels: TXLSPivotFieldsDest; ARCItem: TXLSPivotRCItem; var AMaxLevel: integer);
var
  c   : integer;
  i   : integer;
  S   : AxUCString;
  Val : TXLSSharedItemsValue;
  Item: TXLSPivotRCItem;
begin
  if ALabels.Count < 1 then
    Exit;

  c := ALabels[AIndex].Index;

  S := FSrcRef.AsStringRel[c,ARow];

  Item := ARCItem.Find(S);

  if Item = Nil then begin
    i := -1;
    Val := FSrcRef.AsSharedItemRel[c,ARow];
    if Val <> Nil then begin
      try
        if FFields[c].Filter.Count > 0 then begin
          if FFields[c].Filter.Find(Val) then
            i := ALabels[AIndex].FCacheField.SharedItems.Values.ValueList.FindValue(Val)
          else
            i := -1;
        end
        else
          i := ALabels[AIndex].FCacheField.SharedItems.Values.ValueList.FindValue(Val);
      finally
        Val.Free;
      end;
    end;

    if i >= 0 then begin
      Val := ALabels[AIndex].FCacheField.SharedItems.Values.ValueList[i];
      if Val.Type_ <> xsivtBlank then
        Item := ARCItem.Add(Val,c);
    end;
  end;
  if (Item <> Nil) and (AIndex < (ALabels.Count - 1)) then begin
    DoRowColItems(ARow,AIndex + 1,ALabels,Item,AMaxLevel);

    AMaxLevel := Max(AMaxLevel,AIndex + 1);
  end;
end;

procedure TXLSPivotTable.DoValues(ASubtotals: boolean);
var
  R,C: integer;
begin
  R := 0;
  C := 0;
  DoColValues(R,C,FCols,1,ASubtotals);
end;

function TXLSPivotTable.GetColsTotals: boolean;
begin
  Result := FTableDef.ColGrandTotals;
end;

function TXLSPivotTable.GetName: AxUCString;
begin
  Result := FTableDef.Name;
end;

function TXLSPivotTable.GetRowsTotals: boolean;
begin
  Result := FTableDef.RowGrandTotals;
end;

function TXLSPivotTable.HasLabels: boolean;
begin
  Result := (FRowLabels.Count + FColumnLabels.Count) > 0;
end;

procedure TXLSPivotTable.DoColHeaders(var AIndex: integer; ALevel: integer; AItem: TXLSPivotRCItem; ASubtotals: boolean);
var
  i,j: integer;
begin
  for i := 0 to AItem.Count - 1 do begin
    FDestRef.AsSharedItemRel[AIndex,ALevel] := AItem[i].Value;
    if AItem[i].FChilds <> Nil then
      DoColHeaders(AIndex,ALevel + 1,AItem[i],ASubtotals);

    if (ALevel = 1) and ASubtotals then begin
      Inc(AIndex);
      FDestRef.AsStringRel[AIndex,ALevel] := AItem[i].Value.AsText + ' Total';
      FSubCols.Add(Pointer(AIndex));
    end;

    if FDataValues.Count > 1 then begin
      for j := 0 to FDataValues.Count - 1 do
        FDestRef.AsStringRel[AIndex + j,ALevel + 1] := FDataValues[j].DisplayName;
    end;

    Inc(AIndex,Max(1,FDataValues.Count));
  end;
  Dec(AIndex,Max(1,FDataValues.Count));
end;

procedure TXLSPivotTable.DoColValues(var AR,AC: integer; ACols: TXLSPivotRCItem; ALevel: integer; ASubtotals: boolean);
var
  i,j : integer;
  Item: TXLSPivotRCItem;
  Cond: array of TXLSPivotCompare;
begin
  SetLength(Cond,FMaxRowLevel + FMaxColLevel);

  if ACols.Count > 0 then begin
    for i := 0 to ACols.Count - 1 do begin
      AR := 0;
      if ALevel = FMaxColLevel then begin
        j := 0;
        Item := ACols[i];
        while Item.Parent <> Nil do begin
          Cond[j].Value := Item.Value;
          Cond[j].X := Item.CacheX;

          Item := Item.Parent;
          Inc(j);
        end;

        DoRowValues(AR,AC,FRows,0,TXLSPivotCondition(@Cond[0]));

        Inc(AC,FDataValues.Count);
      end
      else begin
        DoColValues(AR,AC,ACols[i],ALevel + 1,False);
      end;
      if (ALevel = 1) and ASubtotals then begin
        FPivResult.RowSum(AC);
        Inc(AC);
      end;
    end;
  end
  else if FDataValues.Count > 0 then begin
    DoRowValues(AR,AC,FRows,0,TXLSPivotCondition(@Cond[0]));
  end;
end;

procedure TXLSPivotTable.DoDataValues;
var
  i    : integer;
  Field: TXLSPivotField;
  Fld  : TCT_DataField;
begin
  if FTableDef.DataFields <> Nil then
    FTableDef.DataFields.Clear
  else
    FTableDef.Create_DataFields;

  for i := 0 to FDataValues.Count - 1 do begin
    Field := FDataValues[i];

    Fld := FTableDef.DataFields.Add;
    Fld.Fld := Field.Index;
  end;
end;

procedure TXLSPivotTable.DoWriteRowItems(ADefItems: TCT_RowColItems; ARCItems: TXLSPivotRCItem; ALevel: integer; AHasTotal: boolean);
var
  i : integer;
  iI: TCT_I;
begin
  for i := 0 to ARCItems.Count - 1 do begin
    iI := ADefItems.Add;
    iI.R := ALevel;
    iI.Create_XXpgList;
    iI.XXpgList.Add.V := i;

    if ARCItems[i].Count > 0 then
      DoWriteRowItems(ADefItems,ARCItems[i],ALevel + 1,AHasTotal);
  end;

  if AHasTotal then begin
    iI := ADefItems.Add;
    iI.T := stitGrand;
    iI.R := 0;
    iI.Create_XXpgList;
    iI.XXpgList.Add.V := 0;
  end;
end;

procedure TXLSPivotTable.DoWriteColItems(ADefItems: TCT_RowColItems; ARCItems: TXLSPivotRCItem; ALevel: integer; AHasTotal: boolean);
var
  i,j: integer;
  iI : TCT_I;
begin
  for i := 0 to ARCItems.Count - 1 do begin
    iI := ADefItems.Add;
    iI.R := ALevel;
    iI.Create_XXpgList;
    iI.XXpgList.Add.V := i;
    for j := 1 to FDataValues.Count - 1 do
      iI.XXpgList.Add.V := 0;

    for j := 1 to FDataValues.Count - 1 do begin
      iI := ADefItems.Add;
      iI.R := ALevel + 1;
      iI.I := ALevel + 1;
      iI.Create_XXpgList;
      iI.XXpgList.Add.V := j;
    end;

    if ARCItems[i].Count > 0 then
      DoWriteColItems(ADefItems,ARCItems[i],ALevel + 1,AHasTotal);
  end;

  if AHasTotal then begin
    iI := ADefItems.Add;
    iI.T := stitGrand;
    iI.R := 0;
    iI.Create_XXpgList;
    iI.XXpgList.Add.V := 0;
  end;
end;

function TXLSPivotTable.Make: boolean;
var
  r         : integer;
  i         : integer;
  Cnt       : integer;
  SubTot    : boolean;
  SubTotCols: integer;
begin
  Result := True;

  Clear;

  FTableDef.Cache.RefreshOnLoad := True;

  FDestRef.ClearAll;
  FDestRef.AutoExtend := True;
  FDestRef.Cols := 1;
  FDestRef.Rows := 1;

  FSrcRef := FTableDef.Cache.CacheSource.WorksheetSource.RCells;

  CheckUsed;

  DoFixupEdit;

  FTableDef.Cache.CacheValues;

  DoFields;

  DoRowFields;
  DoColFields;
  DoDataValues;

  FMaxRowLevel := 0;
  FMaxColLevel := 0;

  if HasLabels then begin
    for r := 1 to FSrcRef.RowCount - 1 do
      DoRowColItems(r,0,FColumnLabels,FCols,FMaxColLevel);

    for r := 1 to FSrcRef.RowCount - 1 do
      DoRowColItems(r,0,FRowLabels,FRows,FMaxRowLevel);

    Inc(FMaxRowLevel);
    if FCols.Count > 0 then begin
      Inc(FMaxColLevel);
      FCalcRowGrand := FTableDef.RowGrandTotals;
    end
    else
      FCalcRowGrand := False;

    SubTot := (FMaxColLevel > 1) and (FColumnLabels[0].Func <> xpfNone);

    FRows.Sort;
    FCols.Sort;

    i := 1;
    DoColHeaders(i,1,FCols,SubTot);

    i := ColHeaderRows;
    DoRowHeaders(i,0,FRows);

    if SubTot then
      SubTotCols := FCols.Count
    else
      SubTotCols := 0;

    if FCols.Count > 0 then
      Cnt := FCols.MaxLevelCount
    else
      Cnt := 1;

    FPivResult.Resize(Cnt * Max(FDataValues.Count,1) + SubTotCols + Integer(FCalcRowGrand),
                      FRows.TotalCount + Integer(FTableDef.ColGrandTotals));

    FDestRef.ColCount := 1 + FCols.MaxLevelCount + SubTotCols + Integer(FCalcRowGrand);
    FDestRef.RowCount := 1 + FRows.TotalCount + Max(FMaxColLevel,1) + Integer(FTableDef.ColGrandTotals);

    DoValues(SubTot);

    if FTableDef.RowItems <> Nil then begin
      FTableDef.RowItems.Clear;
      DoWriteRowItems(FTableDef.RowItems,FRows,0,FCalcRowGrand);
    end;

    if FTableDef.ColItems <> Nil then begin
      FTableDef.ColItems.Clear;
      DoWriteColItems(FTableDef.ColItems,FCols,0,FTableDef.ColGrandTotals);
    end;

    if FDataValues.Count > 1 then
      DoColDataValues;

    if FCols.Count <= 0 then begin
      for i := 0 to FDataValues.Count - 1 do
        FDestRef.AsStringRel[i + 1,1] := FDataValues[i].FuncName + ' of ' + FDataValues[i].DisplayName;
    end
    else if FDataValues.Count = 1 then
      FDestRef.AsStringRel[0,0] := FDataValues[0].FuncName + ' of ' + FDataValues[0].DisplayName;


    FDestRef.AsStringRel[0,1] := 'Row Labels';
    if FCols.Count > 0 then
      FDestRef.AsStringRel[1,0] := 'Column Labels'
    else if FDataValues.Count > 0 then
      FDestRef.AsStringRel[1,0] := 'Values';

    FTableDef.Location.Ref := FDestRef.ShortRef;

    FPivResult.ColTotal(-1);
    FPivResult.SumCells(FMaxRowLevel,FTableDef.ColGrandTotals,FCalcRowGrand);

    FPivResult.CopyTo(1,ColHeaderRows,FColorCols,FDestRef);

    if FCalcRowGrand then begin
      FDestRef.AsStringRel[-2,1] := 'Grand Total';
    end;

    FDestRef.BeginCommand;
    FDestRef.Command.Add(xcfcCellColor,FColorCols);
    FDestRef.Command.Add(xcfcFontBold);
    FDestRef.ApplyCommandRel(0,0,-2,ColHeaderRows - 1);

    if FTableDef.ColGrandTotals then begin
      FDestRef.AsStringRel[0,-2] := 'Grand Total';
      FDestRef.Command.Add(xcfcCellColor,FColorCols);
      FDestRef.Command.Add(xcfcFontBold);
      FDestRef.ApplyCommandRel(0,-2,-2,-2);
    end;

    for i := 0 to FSubCols.Count - 1 do begin
      FDestRef.BeginCommand;
      FDestRef.Command.Add(xcfcCellColor,FColorSubs);
      FDestRef.ApplyCommandRel(Integer(FSubCols[i]),0,Integer(FSubCols[i]),FDestRef.RowCount - 2);
    end;

    FDestRef.AutoWidthCols;
  end
  else begin
    FDestRef.AsStringRel[0,0] := 'Pivot table';
    FDestRef.IncCols;
    FDestRef.IncRows;
    FDestRef.BeginCommand;
    FDestRef.Command.Add(xcfcCellColor,FColorCols);
    FDestRef.ApplyCommandRel(0,0,-2,-2);
  end;

  DoEpilogue;

  FDirty := False;
end;

procedure TXLSPivotTable.SetColsTotals(const Value: boolean);
begin
  FTableDef.ColGrandTotals := Value;
end;

procedure TXLSPivotTable.SetName(const Value: AxUCString);
begin
  FTableDef.Name := Value;
end;

procedure TXLSPivotTable.SetRowsTotals(const Value: boolean);
begin
  FTableDef.RowGrandTotals := Value;
end;

{ TXLSPivotFieldsDest }

procedure TXLSPivotFieldsDest.Add(AField: TXLSPivotField);
begin
  inherited Add(AField);
end;

constructor TXLSPivotFieldsDest.Create;
begin
  inherited Create(False);
end;

function TXLSPivotFieldsDest.Find(AName: AxUCString): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].DisplayName = AName then
      Exit;
  end;

  Result := -1;
end;

function TXLSPivotFieldsDest.Find(AField: TXLSPivotField): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result] = AField then
      Exit;
  end;
  Result := -1;
end;

function TXLSPivotFieldsDest.GetItems(Index: integer): TXLSPivotField;
begin
  Result := TXLSPivotField(inherited Items[Index]);
end;

function TXLSPivotFieldsDest.Remove(AName: AxUCString): boolean;
var
  i: integer;
begin
  i := Find(AName);
  Result := i >= 0;

  if Result then
    Delete(i);
end;

function TXLSPivotFieldsDest.Remove(AField: TXLSPivotField): boolean;
var
  i: integer;
begin
  i := IndexOf(AField);
  Result := i >= 0;

  if Result then
    Delete(i);
end;

{ TXLSPivotField }

constructor TXLSPivotField.Create;
begin
  FFilter := TXLSUniqueSharedItemsValues.Create;
end;

destructor TXLSPivotField.Destroy;
begin
  FFilter.Free;

  inherited;
end;

function TXLSPivotField.GetDisplayName: AxUCString;
begin
  if UserName <> '' then
    Result := UserName
  else
    Result := Name;
end;

function TXLSPivotField.GetFunction: TXLSPivotFunc;
begin
       if FPivFld.SumSubtotal     then Result := xpfSum
  else if FPivFld.CountASubtotal  then Result := xpfCountNumbers
  else if FPivFld.AvgSubtotal     then Result := xpfAverage
  else if FPivFld.MaxSubtotal     then Result := xpfMax
  else if FPivFld.MinSubtotal     then Result := xpfMin
  else if FPivFld.ProductSubtotal then Result := xpfProduct
  else if FPivFld.CountSubtotal   then Result := xpfCount
  else if FPivFld.StdDevSubtotal  then Result := xpfStdDev
  else if FPivFld.StdDevPSubtotal then Result := xpfStdDevp
  else if FPivFld.VarSubtotal     then Result := xpfVar
  else if FPivFld.VarPSubtotal    then Result := xpfVarp
  else if FPivFld.DefaultSubtotal then Result := xpfDefault
  else                                 Result := xpfDefault;
end;

function TXLSPivotField.GetFunctionName: AxUCString;
begin
  if Func = xpfDefault then begin
    if FCacheField.SharedItems.Values.IsNumeric then
      Result := XLSPivotFuncStr[xpfSum]
    else
      Result := XLSPivotFuncStr[xpfCount];
  end
  else
    Result := XLSPivotFuncStr[GetFunction];
end;

function TXLSPivotField.GetName: AxUCString;
begin
  Result := FCacheField.Name;
end;

function TXLSPivotField.GetUserName: AxUCString;
begin
  Result := FPivFld.Name;
end;

procedure TXLSPivotField.SetFunction(const Value: TXLSPivotFunc);
begin
  FPivFld.SumSubtotal := False;
  FPivFld.CountASubtotal := False;
  FPivFld.AvgSubtotal := False;
  FPivFld.MaxSubtotal := False;
  FPivFld.MinSubtotal := False;
  FPivFld.ProductSubtotal := False;
  FPivFld.CountSubtotal := False;
  FPivFld.StdDevSubtotal := False;
  FPivFld.StdDevPSubtotal := False;
  FPivFld.VarSubtotal := False;
  FPivFld.VarPSubtotal := False;
  FPivFld.DefaultSubtotal := False;

  case Value of
    xpfSum         : FPivFld.SumSubtotal := True;
    xpfCount       : FPivFld.CountSubtotal := True;
    xpfAverage     : FPivFld.AvgSubtotal := True;
    xpfMax         : FPivFld.MaxSubtotal := True;
    xpfMin         : FPivFld.MinSubtotal := True;
    xpfProduct     : FPivFld.ProductSubtotal := True;
    xpfCountNumbers: FPivFld.CountASubtotal := True;
    xpfStdDev      : FPivFld.StdDevSubtotal := True;
    xpfStdDevp     : FPivFld.StdDevpSubtotal := True;
    xpfVar         : FPivFld.VarSubtotal := True;
    xpfVarp        : FPivFld.VarpSubtotal := True;
    xpfDefault     : FPivFld.DefaultSubtotal := True;
  end;
end;

procedure TXLSPivotField.SetUserName(const Value: AxUCString);
begin
  FPivFld.Name := Value;
end;

procedure TXLSPivotField.Sort;
begin
  if FCacheField.SharedItems.Values.CanSort then
    FCacheField.SharedItems.Values.ValueList.Sort;
end;

{ TXLSPivotRCItem }

function TXLSPivotRCItem.Add(AValue: TXLSSharedItemsValue; ACacheX: integer): TXLSPivotRCItem;
begin
  if FChilds = Nil then
    FChilds := THashedStringList.Create;

  Result := TXLSPivotRCItem.Create(Self,AValue,ACacheX);

  FChilds.AddObject(AValue.AsIdString,Result);
end;

procedure TXLSPivotRCItem.Clear;
var
  i: integer;
begin
  if FChilds <> Nil then begin
    for i := 0 to FChilds.Count - 1 do begin
      if FChilds.Objects[i] <> Nil then
        TXLSPivotRCItem(FChilds.Objects[i]).Free;
    end;
    FChilds.Free;

    FChilds := Nil;
  end;
end;

function TXLSPivotRCItem.Count: integer;
begin
  if FChilds <> Nil then
    Result := FChilds.Count
  else
    Result := 0;
end;

constructor TXLSPivotRCItem.Create(AParent: TXLSPivotRCItem; AValue: TXLSSharedItemsValue; ACacheX: integer);
begin
  FParent := AParent;
  FValue := AValue;
  FCacheX := ACacheX;
end;

destructor TXLSPivotRCItem.Destroy;
begin
  Clear;

  inherited;
end;

procedure TXLSPivotRCItem.DoMaxLevel(var ALevel,AMaxLevel: integer);
var
  i: integer;
begin
  if FChilds <> Nil then begin
    Inc(ALevel);
    AMaxLevel := Max(AMaxLevel,ALevel);
    for i := 0 to FChilds.Count - 1 do
      Items[i].DoMaxLevel(ALevel,AMaxLevel);
    Dec(ALevel);
  end;
end;

function TXLSPivotRCItem.Find(AValue: AxUCString): TXLSPivotRCItem;
var
  i: integer;
begin
  Result := Nil;

  if FChilds <> Nil then begin
    i := FChilds.IndexOf(AValue);
    if i >= 0 then
      Result := TXLSPivotRCItem(FChilds.Objects[i]);
  end;
end;

function TXLSPivotRCItem.GetItems(Index: integer): TXLSPivotRCItem;
begin
  Result := TXLSPivotRCItem(FChilds.Objects[Index]);
end;

function TXLSPivotRCItem.MaxLevel: integer;
var
  l: integer;
begin
  Result := 0;
  l := 0;

  DoMaxLevel(l,Result);
end;

function TXLSPivotRCItem.MaxLevelCount: integer;
var
  i: integer;
begin
  Result := 0;

  for i := 0 to Count - 1 do begin
    if Items[i].FChilds = Nil then begin
      Inc(Result,Count);
      Break;
    end
    else
      Inc(Result,Items[i].MaxLevelCount);
  end;
end;

procedure TXLSPivotRCItem.Sort;
var
  i: integer;
begin
  if (FChilds <> Nil) and (FChilds.Count > 0) then begin
    if Items[0].Value.Type_ in [xsivtBoolean,xsivtDate,xsivtError,xsivtString] then
      FChilds.Sort;

    for i := 0 to FChilds.Count - 1 do
      Items[i].Sort;
  end;
end;

function TXLSPivotRCItem.TotalCount: integer;
var
  i: integer;
begin
  Result := Count;

  for i := 0 to Count - 1 do
    Inc(Result,Items[i].TotalCount);
end;

{ TXLSPivotTableResult }

procedure TXLSPivotTableResult.BeginCollect;
begin
  SetLength(FVals,256);

  FValsCount := 0;
end;

procedure TXLSPivotTableResult.Calc(AFunc: TXLSPivotFunc; ACol, ARow: integer; AValue: double);
begin
  if (ACol >= FColCount) or (ARow >= FRowCount) then
    raise XLSRWException.Create('Pivot cell out or range');

  case AFunc of
    xpfSum         : FCells[ACol,ARow].Value := FCells[ACol,ARow].Value + AValue;
    xpfCount       : FCells[ACol,ARow].Value := AValue;
    xpfAverage     : ;
    xpfMax         : begin
      if AValue > FCells[ACol,ARow].Value then
        FCells[ACol,ARow].Value := AValue;
    end;
    xpfMin         : begin
      if AValue < FCells[ACol,ARow].Value then
        FCells[ACol,ARow].Value := AValue;
    end;
    xpfProduct     : begin
      if FCells[ACol,ARow].Value <> 0 then
        FCells[ACol,ARow].Value := FCells[ACol,ARow].Value * AValue
      else
        FCells[ACol,ARow].Value := AValue;
    end;
    xpfCountNumbers: ;
    xpfStdDev      : begin
      if FValsCount > 1 then begin
        SetLength(FVals,FValsCount);
        FCells[ACol,ARow].Value := StdDev(FVals);
      end;
    end;
    xpfStdDevp: begin
      if FValsCount > 1 then begin
        SetLength(FVals,FValsCount);
        FCells[ACol,ARow].Value := PopnStdDev(FVals);
      end;
    end;
    xpfVar: begin
      if FValsCount > 1 then begin
        SetLength(FVals,FValsCount);
        FCells[ACol,ARow].Value := Variance(FVals);
      end;
    end;
    xpfVarp: begin
      if FValsCount > 1 then begin
        SetLength(FVals,FValsCount);
        FCells[ACol,ARow].Value := PopnVariance(FVals);
      end;
    end;
    xpfDefault: begin
      FCells[ACol,ARow].Value := FCells[ACol,ARow].Value + AValue;
    end;
  end;
end;

procedure TXLSPivotTableResult.Clear;
begin
  Resize(0,0);

  FColCount := 0;
  FRowCount := 0;
end;

procedure TXLSPivotTableResult.ClearValues;
var
  c,r: integer;
begin
  for r := 0 to FRowCount - 1 do begin
    for c := 0 to FColCount - 1 do begin
      FCells[c,r].Value := 0;
      FCells[c,r].CellType := xprctValue;
      FCells[c,r].Level := -1;
    end;
  end;
end;

procedure TXLSPivotTableResult.Collect(AVal: double);
begin
  if FValsCount >= Length(FVals) then
    SetLength(FVals,Length(FVals) * 256);

  FVals[FValsCount] := AVal;
  Inc(FValsCount);
end;

procedure TXLSPivotTableResult.ColTotal(ARow: integer);
var
  c: integer;
begin
  if ARow < 0 then
    ARow := FRowCount - 1;

  for c := 0 to FColCount - 1 do
    FCells[c,ARow].CellType := xprctTotal;
end;

procedure TXLSPivotTableResult.CopyTo(ACol,ARow: integer; AColor: integer; ARef: TXLSRelCells);
var
  c,r: integer;
begin
  for r := 0 to FRowCount - 1 do begin
    for c := 0 to FColCount - 1 do begin
      if FCells[c,r].Value <> 0 then
        ARef.AsFloatRel[ACol + c,ARow + r] := FCells[c,r].Value;
      if FCells[c,r].CellType = xprctSumCol then begin
        ARef.BeginCommand;
        ARef.Command.Add(xcfcFontBold);
        if FCells[c,r].Level = 0 then
          ARef.Command.Add(xcfcBorderBottomColor,AColor);

        ARef.ApplyCommandRel(-1,ARow + r,-2,ARow + r);
      end;
    end;
  end;

//  for r := 0 to FRowCount - 1 do begin
//    for c := 0 to FColCount - 1 do begin
//      case FCells[c,r].CellType of
//        xprctValue : ARef.AsStringRel[ACol + c,ARow + r] := 'Value';
//        xprctSumCol: ARef.AsStringRel[ACol + c,ARow + r] := 'Sum col';
//        xprctSumRow: ARef.AsStringRel[ACol + c,ARow + r] := 'Sum row';
//        xprctTotal : ARef.AsStringRel[ACol + c,ARow + r] := 'Total';
//      end;
//    end;
//  end;
end;

constructor TXLSPivotTableResult.Create;
begin

end;

function TXLSPivotTableResult.GetCells(ACol, ARow: integer): double;
begin
  if (ACol >= FColCount) or (ARow >= FRowCount) then
    raise XLSRWException.Create('Pivot cell ourt or range');

  Result := FCells[ACol,ARow].Value;
end;

procedure TXLSPivotTableResult.Resize(AColCount, ARowCount: integer);
begin
  FColCount := AColCount;
  FRowCount := ARowCount;

  SetLength(FCells,FColCount,FRowCount);

  ClearValues;
end;

procedure TXLSPivotTableResult.RowSum(ACol: integer);
var
  r: integer;
begin
  if ACol < 0 then
    ACol := FColCount - 1;

  for r := 0 to FRowCount - 1 do
    FCells[ACol,r].CellType := xprctSumRow;
end;

procedure TXLSPivotTableResult.SetCells(ACol, ARow: integer; const Value: double);
begin
  if (ACol >= FColCount) or (ARow >= FRowCount) then
    raise XLSRWException.Create('Pivot cell ourt or range');

  FCells[ACol,ARow].Value := Value;
end;

procedure TXLSPivotTableResult.Val(ACol, ARow: integer; AValue: double);
begin
  if (ACol >= FColCount) or (ARow >= FRowCount) then
    raise XLSRWException.Create('Pivot cell ourt or range');

  FCells[ACol,ARow].Value := AValue;
  FCells[ACol,ARow].CellType := xprctValue;
  FCells[ACol,ARow].Level := -1;
end;

procedure TXLSPivotTableResult.SumCells(AMaxRowLevel: integer; AColGrandTotals, ARowGrandTotals: boolean);
var
  i            : integer;
  c,r          : integer;
  Sums         : array of double;
  Sum          : double;
  SubSum       : double;
  GrandTotalRow: integer;
begin
  if ARowGrandTotals then begin
    for r := 0 to FRowCount - 1 do begin
      Sum := 0;
      SubSum := 0;
      for c := 0 to FColCount - 2 do begin
        case FCells[c,r].CellType of
          xprctValue: begin
            // Min or Max cells that never was assigned.
            if not (FCells[c,r].Value = MINDOUBLE) or (FCells[c,r].Value = MAXDOUBLE) then begin
              Sum := Sum + FCells[c,r].Value;
              SubSum := SubSum + FCells[c,r].Value;
            end;
          end;
          xprctSumRow: begin
            FCells[c,r].Value := SubSum;
            SubSum := 0;
          end;
        end;
      end;
      FCells[FColCount - 1,r].Value := Sum;
    end;
  end;

  SetLength(Sums,AMaxRowLevel);

  GrandTotalRow := -1;
  for c := 0 to FColCount - 1 do begin
    for i := 0 to High(Sums) do
      Sums[i] := 0;

    for r := FRowCount - 1 downto 0 do begin
      case FCells[c,r].CellType of
        xprctValue,
        xprctSumRow: begin
          // Min or Max cells that never was assigned.
          if (FCells[c,r].Value = MINDOUBLE) or (FCells[c,r].Value = MAXDOUBLE) then
            FCells[c,r].Value := 0
          else
            Sums[AMaxRowLevel - 1] := Sums[AMaxRowLevel - 1] + FCells[c,r].Value;
        end;
        xprctSumCol: begin
          FCells[c,r].Value := Sums[FCells[c,r].Level + 1];
          Sums[FCells[c,r].Level] := Sums[FCells[c,r].Level] + Sums[FCells[c,r].Level + 1];
          Sums[FCells[c,r].Level + 1] := 0;
        end;
        xprctTotal: GrandTotalRow := r;
      end;
    end;

    if AColGrandTotals and (GrandTotalRow >= 0) then
      FCells[c,GrandTotalRow].Value := Sums[0];
  end;
end;

end.
