unit XLSRow5;

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
     Xc12Utils5, Xc12Common5, Xc12DataStylesheet5, Xc12Manager5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSFormattedObj5;

type TXLSRows = class;

     TXLSRow = class(TXLSFormattedObj)
private
     function  GetOutlineLevel: integer;
     procedure SetOutlineLevel(const Value: integer);
     function  GetCollapsedOutline: boolean;
     procedure SetCollapsedOutline(const Value: boolean);
     function  GetPixelHeight: integer;
     procedure SetPixelHeight(const Value: integer);
     function  GetHeightPt: double;
     procedure SetHeightPt(const Value: double);
     function  GetCustomHeight: boolean;
protected
     FOwner  : TXLSRows;
     FIndex  : integer;
     FRowItem: PXLSMMURowHeader;

     procedure SetHeight(AValue: integer);
     function  GetHeight: integer;
     procedure SetHidden(AValue: boolean);
     function  GetHidden: boolean;

     function  DefaultHeight: integer;

     procedure Changed(AIndex: integer);
     procedure StyleChanged; override;
public
     constructor Create(AOwner: TXLSRows);
     destructor Destroy; override;

     function Assigned: boolean;

     property Height: integer read GetHeight write SetHeight;
     property HeightPt: double read GetHeightPt write SetHeightPt;
     property PixelHeight: integer read GetPixelHeight write SetPixelHeight;
     property Hidden: boolean read GetHidden write SetHidden;
     property OutlineLevel: integer read GetOutlineLevel write SetOutlineLevel;
     property CollapsedOutline: boolean read GetCollapsedOutline write SetCollapsedOutline;
     property CustomHeight: boolean read GetCustomHeight;
     end;

     TXLSRows = class(TObject)
private
protected
     FStyles: TXc12DataStyleSheet;
     FTheRow: TXLSRow;
     FCells : TXLSCellMMU;
     FApplyStyleOnCells: boolean;
     FLastVisibleRow: integer;

     function GetItems(Index: integer): TXLSRow;
public
     constructor Create(ACells : TXLSCellMMU; AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure Clear;
     procedure ToggleGrouped(ARow: integer);
     procedure ToggleGroupedAll(ALevel: integer);

     function  DefaultPixelHeight: integer;

     // Returns the last unhidden row. Calculation starts at row 65535. I all rows from row n to the end of the sheet
     // are hidden, excel writes max 65536 rows.
     function CalcLastVisibleRow: integer;

     // When True (default) changes to a rows style will be applied to the cells as well.
     // Before version 5.20.83 this was not true but is changed as this is how Excel works.
     property ApplyStyleOnCells: boolean read FApplyStyleOnCells write FApplyStyleOnCells;
     property Items[Index: integer]: TXLSRow read GetItems; default;
     end;

implementation

{ TXLSRow }

function TXLSRow.Assigned: boolean;
begin
  Result := FRowItem <> Nil;
end;

procedure TXLSRow.Changed(AIndex: integer);
begin
  FIndex := AIndex;

  // FLastVisibleRow is used by XLSSpreadSheet to handle the situation when all rows (except for top rows with values) are hidden.
  if (FOwner.FLastVisibleRow > 0) and (AIndex > FOwner.FLastVisibleRow) then
    FRowItem := Nil
  else
    FRowItem := FOwner.FCells.FindRow(FIndex);
  if FRowItem <> Nil then
    FXF := FOwner.FStyles.XFs[FRowItem.Style]
  else
    FXF := FOwner.FStyles.XFs[XLS_STYLE_DEFAULT_XF];
end;

constructor TXLSRow.Create(AOwner: TXLSRows);
begin
  inherited Create(AOwner.FStyles);

  FOwner := AOwner;
  FXF := FOwner.FStyles.XFs[XLS_STYLE_DEFAULT_XF];
end;

function TXLSRow.DefaultHeight: integer;
begin
  Result := XLS_DEFAULT_ROWHEIGHT;
end;

destructor TXLSRow.Destroy;
begin
  FXF := Nil;
  inherited;
end;

function TXLSRow.GetCollapsedOutline: boolean;
begin
  if FRowItem <> Nil then
    Result := xroCollapsed in FRowItem.Options
  else
    Result := False;
end;

function TXLSRow.GetCustomHeight: boolean;
begin
  Result := xroCustomHeight in FRowItem.Options;
end;

function TXLSRow.GetHeight: integer;
begin
  if FRowItem <> Nil then begin
    Result := FRowItem.Height;
    if Result = XLS_DEFAULT_ROWHEIGHT_FLAG then
      Result := DefaultHeight;
  end
  else
    Result := DefaultHeight;
end;

function TXLSRow.GetHeightPt: double;
begin
  Result := GetHeight / 20;
end;

function TXLSRow.GetHidden: boolean;
begin
  if FRowItem <> Nil then
    Result := xroHidden in FRowItem.Options
  else
    Result := False;
end;

function TXLSRow.GetOutlineLevel: integer;
begin
  if FRowItem <> Nil then
    Result := FRowItem.OutlineLevel
  else
    Result := 0;
end;

function TXLSRow.GetPixelHeight: integer;
var
  PPIX,PPIY: integer;
begin
  FStyles.PixelsPerInchXY(PPIX,PPIY);

  Result := Round((Height / 20) / (72 / PPIY));
end;

procedure TXLSRow.SetCollapsedOutline(const Value: boolean);
begin
  if FRowItem <> Nil then begin
    if Value then
      FRowItem.Options := FRowItem.Options + [xroCollapsed]
    else
      FRowItem.Options := FRowItem.Options - [xroCollapsed];
  end
  else if Value then begin
    FRowItem := FOwner.FCells.AddRow(FIndex,XLS_STYLE_DEFAULT_XF);
    FRowItem.Options := FRowItem.Options + [xroCollapsed]
  end;
end;

procedure TXLSRow.SetHeight(AValue: integer);
begin
  if FRowItem <> Nil then
    FRowItem.Height := AValue
  else if AValue <> XLS_DEFAULT_ROWHEIGHT then begin
    FRowItem := FOwner.FCells.AddRow(FIndex,XLS_STYLE_DEFAULT_XF);
    FRowItem.Height := AValue;
  end;
  if FRowItem <> Nil then begin
    if AValue <> XLS_DEFAULT_ROWHEIGHT then
      FRowItem.Options := FRowItem.Options + [xroCustomHeight]
    else
      FRowItem.Options := FRowItem.Options - [xroCustomHeight];
  end;
end;

procedure TXLSRow.SetHeightPt(const Value: double);
begin
  SetHeight(Round(Value * 20));
end;

procedure TXLSRow.SetHidden(AValue: boolean);
var
  R: PXLSMMURowHeader;
begin
  R := FOwner.FCells.FindRow(FIndex);
  if R <> Nil then begin
    if AValue then
      R.Options := R.Options + [xroHidden]
    else
      R.Options := R.Options - [xroHidden];
  end
  else if AValue then begin
    R := FOwner.FCells.AddRow(FIndex,XLS_STYLE_DEFAULT_XF);
    R.Options := R.Options + [xroHidden]
  end;
end;

procedure TXLSRow.SetOutlineLevel(const Value: integer);
begin
  if Value < 0 then
    Exit;

  if FRowItem <> Nil then
    FRowItem.OutlineLevel := Value
  else begin
    FRowItem := FOwner.FCells.AddRow(FIndex,XLS_STYLE_DEFAULT_XF);
    FRowItem.OutlineLevel := Value;
  end;
end;

procedure TXLSRow.SetPixelHeight(const Value: integer);
var
  PPIX,PPIY: integer;
begin
  FStyles.PixelsPerInchXY(PPIX,PPIY);

  SetHeight(Round(((Value * 72) / PPIY) * 20));
end;

procedure TXLSRow.StyleChanged;
var
  C,C1,C2: integer;
  Cell: TXLSCellItem;
begin
  FOwner.FCells.SetRowStyle(FIndex,FXF.Index);

  if FOwner.FApplyStyleOnCells then begin
    FOwner.FCells.CalcRowDimensions(FIndex,C1,C2);
    if (C1 >= 0) and (C2 >= 0) then begin
       for C := C1 to C2 do begin
         if FOwner.FCells.FindCell(C,FIndex,Cell) then
           FOwner.FCells.SetStyle(C,FIndex,FXF.Index);
       end;
    end;
  end;
end;

{ TXLSRows }

procedure TXLSRows.Clear;
begin
  inherited;

  FLastVisibleRow := 0;
//  FCells.Clear;
end;

constructor TXLSRows.Create(ACells : TXLSCellMMU; AStyles: TXc12DataStyleSheet);
begin
  FCells := ACells;
  FStyles := AStyles;
  FTheRow := TXLSRow.Create(Self);

  FApplyStyleOnCells := True;
end;

function TXLSRows.DefaultPixelHeight: integer;
var
  PPIX,PPIY: integer;
begin
  FStyles.PixelsPerInchXY(PPIX,PPIY);

  Result := Round((XLS_DEFAULT_ROWHEIGHT / 20) / (72 / PPIY));
end;

destructor TXLSRows.Destroy;
begin
  FTheRow.Free;
  inherited;
end;

function TXLSRows.GetItems(Index: integer): TXLSRow;
begin
  FTheRow.Changed(Index);
  Result := FTheRow;
end;

function TXLSRows.CalcLastVisibleRow: integer;
var
  Row: PXLSMMURowHeader;
begin
  Row := FCells.FindRow(65535);
  if (Row <> Nil) and (xroHidden in Row.Options) then begin
    Result := 65535;
    while (Row <> Nil) and (xroHidden in Row.Options) do begin
      Dec(Result);
      Row := FCells.FindRow(Result);
    end;
  end
  else
    Result := XLS_MAXROW;

  FLastVisibleRow := Result;
end;

procedure TXLSRows.ToggleGrouped(ARow: integer);
var
  pRow: PXLSMMURowHeader;
  UnGroup: boolean;
  Level: integer;

procedure SkipCollapsed(var ARow: integer);
var
  pRow: PXLSMMURowHeader;
  Level: integer;
begin
  Dec(ARow);
  pRow := FCells.FindRow(ARow);
  if pRow <> Nil then
    Level := pRow.OutlineLevel
  else
    Level := -1;
  while (ARow >= 0) and (pRow <> Nil) and (pRow.OutlineLevel >= Level) do begin
    if xroCollapsed in pRow.Options then
      SkipCollapsed(ARow);
    Dec(ARow);
    pRow := FCells.FindRow(ARow);
  end;
  if ARow > 0 then
    Inc(ARow);
end;

begin
  if (ARow < 0) or (ARow > XLS_MAXROW) then
    Exit;
  pRow := FCells.FindRow(ARow - 1);
  if (pRow = Nil) or (pRow.OutlineLevel <= 0) then
    Exit;
  Level := pRow.OutlineLevel;
  pRow := FCells.FindRow(ARow);
  if pRow <> Nil then begin
    UnGroup := xroCollapsed in pRow.Options;
    if UnGroup then
      pRow.Options := pRow.Options - [xroCollapsed]
    else
      pRow.Options := pRow.Options + [xroCollapsed];
  end
  else begin
    UnGroup := False;
    pRow := FCells.AddRow(ARow,XLS_STYLE_DEFAULT_XF);
    pRow.Options := pRow.Options + [xroCollapsed];
  end;
  Dec(ARow);
  while ARow >= 0 do begin
    pRow := FCells.FindRow(ARow);
    if (pRow = Nil) or (pRow.OutlineLevel < Level) then
      Break
    else if pRow.OutlineLevel >= Level then begin
      if UnGroup then
        pRow.Options := pRow.Options - [xroHidden]
      else
        pRow.Options := pRow.Options + [xroHidden];
      if xroCollapsed in pRow.Options then
        SkipCollapsed(ARow)
    end;
    Dec(ARow);
  end;
end;

procedure TXLSRows.ToggleGroupedAll(ALevel: integer);
var
  i: integer;
  pRow: PXLSMMURowHeader;

procedure CheckButton(const ARow,ALevel: integer; const Value: boolean);
var
  pRow: PXLSMMURowHeader;
begin
  pRow := FCells.FindRow(ARow);
  if (pRow <> Nil) and (pRow.OutlineLevel >= ALevel) then
    Exit;
  if pRow = Nil then
    pRow := FCells.AddRow(ARow,XLS_STYLE_DEFAULT_XF);
  if pRow.OutlineLevel < ALevel then begin
    if Value then
      pRow.Options := pRow.Options + [xroCollapsed]
    else
      pRow.Options := pRow.Options - [xroCollapsed];
  end;
end;

begin
  for i := FCells.Dimension.Row2 downto FCells.Dimension.Row1 do begin
    pRow := FCells.FindRow(i);
    if (pRow <> Nil) and (pRow.OutlineLevel > 0) then begin
      if pRow.OutlineLevel >= ALevel then
        pRow.Options := pRow.Options + [xroHidden]
      else
        pRow.Options := pRow.Options - [xroHidden];
      CheckButton(i + 1,pRow.OutlineLevel,xroHidden in pRow.Options);
    end;
  end;
end;

end.
