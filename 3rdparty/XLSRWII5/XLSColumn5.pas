unit XLSColumn5;

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
     Xc12Utils5, Xc12Common5, Xc12DataStylesheet5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSFormattedObj5;

type TXLSColumns = class;

     TXLSColumn = class(TXLSFormattedObj)
protected
     FOwner   : TXLSColumns;

     FColIndex: integer;
     FCol     : TXc12Column;

     function  GetCollapsedOutline: boolean;
     function  GetOutlineLevel: integer;
     function  GetCharWidth: double;
     function  GetPixelWidth: integer;
     function  GetPixelWidth2: integer;
     function  GetHidden: boolean;
     function  GetWidth: integer;
     procedure SetOutlineLevel(const Value: integer);
     procedure SetCollapsedOutline(const Value: boolean);
     procedure SetCharWidth(const AValue: double);
     procedure SetPixelWidth(const AValue: integer);
     procedure SetHidden(const AValue: boolean);
     procedure SetWidth(const AValue: integer);
     function  GetWidthPt: double;
     procedure SetWidthPt(const Value: double);

     procedure Changed(AIndex: integer);
     procedure StyleChanged; override;

     procedure SetXc12ColWidth(ACol: TXc12Column; AWidth: integer); overload;
     procedure SetXc12ColWidth(ACol: TXc12Column; AWidth: double); overload;
     function  DefaultCharWidth: double; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure SetDefault;
     procedure Split;
public
     constructor Create(AOwner: TXLSColumns; AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     property CharWidth: double read GetCharWidth write SetCharWidth;
     // Getter is not equal to setter, PixelWidth = PixelWidth = False;
     property PixelWidth: integer read GetPixelWidth write SetPixelWidth;
     // Use this for equal values of getter/setter, PixelWidth2 = PixelWidth2 = True;
     property PixelWidth2: integer read GetPixelWidth2 write SetPixelWidth;
     property Hidden: boolean read GetHidden write SetHidden;
     property Width: integer read GetWidth write SetWidth;
     property WidthPt: double read GetWidthPt write SetWidthPt;
     property OutlineLevel: integer read GetOutlineLevel write SetOutlineLevel;
     property CollapsedOutline: boolean read GetCollapsedOutline write SetCollapsedOutline;
     end;

     //* Methods for manipulating columns. Please note that these methods has
     //* no affect on the cell values. They operate only on the columns heading.
     TXLSColumns = class(TObject)
protected
     FCols        : TXc12Columns;
     FTheColumn   : TXLSColumn;
     FStyles      : TXc12DataStyleSheet;
     FDefaultWidth: integer;
     FChangeEvent : TTwoIntegerEvent;

     procedure SetDefaultWidth(const Value: integer);
     function  GetItems(Index: integer): TXLSColumn;

     procedure CheckValidColumns(ACol1,ACol2: integer);
     function  FindCol(ACol: integer): TXc12Column;
     function  FindColIndex(ACol: integer): integer;
     function  Cut(AColIndex,ACol1,ACol2: integer): TXc12Column;
     function  Copy(AColIndex,ACol1,ACol2: integer): TXc12Column;
     procedure Delete(AColIndex,ACol1,ACol2: integer);
     procedure FillGaps(AList: TXc12Columns; ACol1,ACol2: integer);
     function  NewColumn(ACol1,ACol2, AWidth: integer): TXc12Column;
     function  FindRange(ACol1,ACol2: integer; var AIndex1,AIndex2: integer): boolean;
     procedure DeleteHits(ACol1,ACol2: integer);
     procedure Sort;
     procedure Compact;
public
     constructor Create(AColumns : TXc12Columns; AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure Assign(AColumns: TXLSColumns);

     procedure Clear;

     //* The method creates a TXc12Columns list and copies the the columns
     //* between ACol1,ACol2 to the list. The columns between ACol1,ACol2 is
     //* then deleted. The object must be destroyed by the caller.
     //* This method is mainly for internal use.
     function  CutHitList(ACol1,ACol2: integer): TXc12Columns;
     //* The method creates a TXc12Columns list and copies the the columns
     //* between ACol1,ACol2 to the list. The object must be destroyed by
     //* the caller. This methid is mainly for internal use.
     function  CopyHitList(ACol1,ACol2: integer): TXc12Columns;
     //* Offset the columns in AList. For use with lists created with @link(CutHitList)
     //* or @link(CopyHitList).
     procedure OffsetList(AList: TXc12Columns; const AOffset: integer);
     //* Inserts a column list. For use with lists created with @link(CutHitList)
     //* or @link(CopyHitList).
     procedure InsertList(AList: TXc12Columns);

     function  GetColWidth(ACol: integer): integer;
     function  GetColWidthPixels(ACol: integer): integer;
     procedure SetColWidth(ACol1,ACol2: integer; AValue: integer);
     function  BeginSetStyle(ACol1,ACol2: integer; AStyle: TXc12XF): TXc12Columns;
     procedure EndSetStyle(AList: TXc12Columns);
     //* Delets all columns between Col1 and Col2. Columns to the right of
     //* Col2 will be shifted left.
     procedure DeleteColumns(ACol1,ACol2: integer);
     //* Delets all columns between Col1 and Col2.
     procedure ClearColumns(ACol1,ACol2: integer);
     procedure CopyColumns(ACol1,ACol2,ADestCol: integer);
     procedure MoveColumns(ACol1,ACol2,ADestCol: integer);
     procedure InsertColumns(ACol,AColCount: integer);
     procedure InsertColumnsKeepFmt(ACol,AColCount: integer);
     procedure ToggleGrouped(ACol: integer);
     procedure ToggleGroupedAll(ALevel: integer);

     procedure AddIfNone(ACol,ACount: integer); {$ifdef DELPHI_2007_OR_LATER} deprecated 'Not required anymore'; {$endif}

     function  DefaultPixelWidth: integer;

     property DefaultWidth: integer read FDefaultWidth write SetDefaultWidth;
     property Styles: TXc12DataStyleSheet read FStyles;
     property OnColWidthChange: TTwoIntegerEvent read FChangeEvent write FChangeEvent;
     property Items[Index: integer]: TXLSColumn read GetItems; default;
     end;

implementation

{ TXLSColumn }

procedure TXLSColumn.Changed(AIndex: integer);
begin
  FColIndex := AIndex;
  FCol := FOwner.FindCol(FColIndex);
  if FCol <> Nil then
    FXF := FOwner.FStyles.XFs[FCol.Style.Index]
  else
    FXF := FOwner.FStyles.XFs[XLS_STYLE_DEFAULT_XF];
end;

constructor TXLSColumn.Create(AOwner: TXLSColumns; AStyles: TXc12DataStyleSheet);
begin
  inherited Create(AStyles);
  FOwner := AOwner;
  FXF := AStyles.XFs[XLS_STYLE_DEFAULT_XF];
end;

function TXLSColumn.DefaultCharWidth: double;
begin
  Result := FOwner.DefaultWidth / 256;
end;

destructor TXLSColumn.Destroy;
begin
  FXF := Nil;
  inherited;
end;

function TXLSColumn.GetCharWidth: double;
begin
  if FCol <> Nil then
    Result := FCol.Width
  else
    Result := DefaultCharWidth;
end;

function TXLSColumn.GetCollapsedOutline: boolean;
begin
  if FCol <> Nil then
    Result := xcoCollapsed in FCol.Options
  else
    Result := False;
end;

function TXLSColumn.GetHidden: boolean;
begin
  if FCol <> Nil then
    Result := xcoHidden in FCol.Options
  else
    Result := False;
end;

function TXLSColumn.GetOutlineLevel: integer;
begin
  if FCol <> Nil then
    Result := FCol.OutlineLevel
  else
    Result := 0;
end;

// 8:   6  56 8.50   8
// 9:   6  56 8.50   8
// 10:  7  64 8.43   8
// 11:  7  64 8.43   8
// 12:  8  72 8.38   8
// 14: 10  88 8.10   8
// 16: 11  96 8.09   8
// 18: 12 104 8.08   8
// 20: 14 128 8.50  16
// 22: 15 136 8.47  16
// 24: 16 144 8.44  16

function TXLSColumn.GetPixelWidth: integer;
begin
  Result := Trunc(((GetWidth + Trunc(128 / FOwner.FStyles.StdFontWidth)) / 256) * FOwner.FStyles.StdFontWidth);
end;

function TXLSColumn.GetPixelWidth2: integer;
begin
  Result := Round((((GetWidth / 256) * 100 - 0.5) / 100) * FOwner.FStyles.StdFontWidth + 5);
end;

function TXLSColumn.GetWidth: integer;
begin
  if FCol <> Nil then
    Result := Round(FCol.Width * 256)
  else
    Result := FOwner.DefaultWidth;
end;

function TXLSColumn.GetWidthPt: double;
var
  PPIX,PPIY: integer;
begin
  FOwner.Styles.PixelsPerInchXY(PPIX,PPIY);

  Result := (GetPixelWidth / PPIX) * 72;
end;

procedure TXLSColumn.SetCharWidth(const AValue: double);
begin
  if (FCol = Nil) and (AValue <> DefaultCharWidth) then begin
    FCol := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
    SetDefault;
  end
  else if AValue <> FCol.Width then
    Split;
  FCol.Width := AValue;
  SetXc12ColWidth(FCol,AValue);
  FOwner.Sort;
  FOwner.Compact;
end;

procedure TXLSColumn.SetCollapsedOutline(const Value: boolean);
begin
  if FCol = Nil then begin
    if not Value then
      Exit;
    FCol := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
    SetDefault;
  end
  else if Value = (xcoHidden in FCol.Options) then
    Exit
  else
    Split;

  if Value then
    FCol.Options := FCol.Options + [xcoCollapsed]
  else
    FCol.Options := FCol.Options - [xcoCollapsed];

  FOwner.Sort;
  FOwner.Compact;

  if Assigned(FOwner.FChangeEvent) then
    FOwner.FChangeEvent(FOwner,FCol.Min,FCol.Max);
end;

procedure TXLSColumn.SetDefault;
begin
  FCol.Min := FColIndex;
  FCol.Max := FColIndex;
  FCol.Width := DefaultCharWidth;
end;

procedure TXLSColumn.SetHidden(const AValue: boolean);
begin
  if FCol = Nil then begin
    if not AValue then
      Exit;
    FCol := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
    SetDefault;
  end
  else if AValue = (xcoHidden in FCol.Options) then
    Exit
  else
    Split;

  if AValue then
    FCol.Options := FCol.Options + [xcoHidden]
  else
    FCol.Options := FCol.Options - [xcoHidden];

  FOwner.Sort;
  FOwner.Compact;

  if Assigned(FOwner.FChangeEvent) then
    FOwner.FChangeEvent(FOwner,FCol.Min,FCol.Max);
end;

procedure TXLSColumn.SetOutlineLevel(const Value: integer);
begin
  if Value < 0 then
    Exit;
  if FCol = Nil then begin
    FCol := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
    SetDefault;
  end
  else if Value = FCol.OutlineLevel then
    Exit
  else
    Split;

  FCol.OutlineLevel := Value;
  FOwner.Sort;
  FOwner.Compact;

  if Assigned(FOwner.FChangeEvent) then
    FOwner.FChangeEvent(FOwner,FCol.Min,FCol.Max);
end;

procedure TXLSColumn.SetPixelWidth(const AValue: integer);
begin
  SetWidth(Round(Trunc((AValue - 5) / FOwner.FStyles.StdFontWidth * 100 + 0.5) / 100) * 256);
end;

procedure TXLSColumn.SetWidth(const AValue: integer);
begin
  if (FCol = Nil) and (AValue = FOwner.DefaultWidth) then
    Exit;
  if (FCol <> Nil) and (FCol.Min = FColIndex) and (FCol.Max = FColIndex) then
    SetXc12ColWidth(FCol,AValue)
  else begin
    if FCol = Nil then begin
      FCol := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
      SetDefault;
    end
    else if (AValue / 256) <> FCol.Width then
      Split;
    SetXc12ColWidth(FCol,AValue);
    FOwner.Sort;
    FOwner.Compact;
  end;

  if Assigned(FOwner.FChangeEvent) then
    FOwner.FChangeEvent(FOwner,FCol.Min,FCol.Max);
end;

procedure TXLSColumn.SetWidthPt(const Value: double);
var
  PPIX,PPIY: integer;
begin
  FOwner.Styles.PixelsPerInchXY(PPIX,PPIY);

  SetPixelWidth(Round((Value / 72) * PPIX));
end;

procedure TXLSColumn.SetXc12ColWidth(ACol: TXc12Column; AWidth: double);
begin
  if AWidth <> DefaultCharWidth then
    ACol.Options := ACol.Options + [xcoCustomWidth]
  else
    ACol.Options := ACol.Options - [xcoCustomWidth];
end;

procedure TXLSColumn.SetXc12ColWidth(ACol: TXc12Column; AWidth: integer);
begin
  ACol.Width := AWidth / 256;
  if AWidth <> FOwner.DefaultWidth then
    ACol.Options := ACol.Options + [xcoCustomWidth]
  else
    ACol.Options := ACol.Options - [xcoCustomWidth];
end;

procedure TXLSColumn.Split;
var
  Col: TXc12Column;
  C1,C2: TXc12Column;
begin
  // Nothing to split
  if (FColIndex = FCol.Min) and (FColIndex = FCol.Max) then
    Exit;

  Col := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
  Col.Assign(FCol);

  if FColIndex = FCol.Min then
    FCol.Min := FCol.Min + 1
  else if FColIndex = FCol.Max then
    FCol.Max := FCol.Max - 1
  else begin
    C1 := FCol;
    C2 := FOwner.FCols.Add(FOwner.FStyles.XFEditor.UseDefault);
    C2.Assign(C1);
    C1.Max := FColIndex - 1;
    C2.Min := FColIndex + 1;
  end;
  FCol := Col;
  FCol.Min := FColIndex;
  FCol.Max := FColIndex;
end;

procedure TXLSColumn.StyleChanged;
begin
  if ((FCol = Nil) and (FXF.Index = XLS_STYLE_DEFAULT_XF)) or ((FCol <> Nil) and (FCol.Style = FXF)) then
    Exit;

  if FCol = Nil then begin
    FCol := FOwner.FCols.Add(FXF);
    SetDefault;
  end
  else begin
    FOwner.FStyles.XFEditor.FreeStyle(FCol.Style.Index);
    FCol.Style := FXF;

    Split;
  end;

  FOwner.FStyles.XFEditor.UseStyle(FXF);

//  FOwner.Sort;
//  FOwner.Compact;

  if Assigned(FOwner.FChangeEvent) then
    FOwner.FChangeEvent(FOwner,FCol.Min,FCol.Max);
end;

{ TXLSColumns }

procedure TXLSColumns.AddIfNone(ACol, ACount: integer);
begin

end;

procedure TXLSColumns.Clear;
begin
  FCols.Clear;
end;

procedure TXLSColumns.ClearColumns(ACol1, ACol2: integer);
begin
  CheckValidColumns(ACol1,ACol2);

  DeleteHits(ACol1,ACol2);
end;

procedure TXLSColumns.Compact;
var
  i: integer;
begin
  i := 1;
  while i < FCols.Count do begin
    if ((FCols[i - 1].Max + 1) = FCols[i].Min) and FCols[i - 1].EqualProps(FCols[i]) then begin
      FCols[i- 1].Max := FCols[i].Max;
      FCols.Delete(i);
    end
    else
      Inc(i);
  end;
end;

//================================
//   1----2                        xchLess
//          +------+
//================================
//       1----2....x               xchCol2 (can hit the entire target)
//          +------+
//================================
//          1----2                 xchCol1Match
//          +------+
//================================
//          1------2               xchMatch
//          +------+
//================================
//         1--------2              xchTargetInside
//          +------+
//================================
//            1----2               xchCol2Match
//          +------+
//================================
//           1----2                xchInside
//          +------+
//================================
//          x....1----2            xchCol1 (can hit the entire target)
//          +------+
//================================
//                   1----2        xchGreater
//          +------+
//================================

function TXLSColumns.Copy(AColIndex, ACol1, ACol2: integer): TXc12Column;
var
  H: TXc12ColumnHit;
  C: TXc12Column;
begin
  C := FCols[AColIndex];
  H := C.Hit(ACol1,ACol2);
  if H in [xchLess,xchGreater] then begin
    Result := Nil;
    Exit;
  end;

  Result := TXc12Column.Create;
  Result.Assign(C);

  case H of
    xchCol2        : Result.Max := ACol2;
    xchCol1Match   : Result.Max := ACol2;
    xchMatch       : ;
    xchTargetInside: ;
    xchCol2Match   : Result.Min := ACol1;
    xchInside      : begin
      Result.Min := ACol1;
      Result.Max := ACol2;
    end;
    xchCol1        : Result.Min := ACol1;
  end;
end;

procedure TXLSColumns.CopyColumns(ACol1, ACol2, ADestCol: integer);
var
  List: TXc12Columns;
begin
  CheckValidColumns(ACol1,ACol2);

  List := CopyHitList(ACol1,ACol2);
  try
    DeleteHits(ADestCol,ADestCol + (ACol2 - ACol1));

    OffsetList(List,ADestCol - ACol1);

    InsertList(List);
  finally
    List.Free;
  end;
end;

function TXLSColumns.CopyHitList(ACol1, ACol2: integer): TXc12Columns;
var
  i: integer;
  iC1,iC2: integer;
  C: TXc12Column;
begin
  Result := TXc12Columns.Create;
  Result.OwnsObjects := False;

  if not FindRange(ACol1,ACol2,iC1,iC2) then
    Exit;

  for i := iC1 to iC2 do begin
    C := Copy(i,ACol1,ACol2);
    if C <> Nil then
      Result.Add(C);
  end;
end;

constructor TXLSColumns.Create(AColumns: TXc12Columns; AStyles: TXc12DataStyleSheet);
begin
  FCols := AColumns;
  FStyles := AStyles;
  FTheColumn := TXLSColumn.Create(Self,AStyles);
  if FCols.DefColWidth <> 0 then
    FDefaultWidth := Round(FCols.DefColWidth * 256)
  else
    FDefaultWidth := XLS_DEFAULT_COLWIDTH;
end;

function TXLSColumns.Cut(AColIndex, ACol1, ACol2: integer): TXc12Column;
var
  H: TXc12ColumnHit;
  C,C2: TXc12Column;
begin
  C := FCols[AColIndex];
  H := C.Hit(ACol1,ACol2);
  if H in [xchLess,xchGreater] then begin
    Result := Nil;
    Exit;
  end;

  Result := TXc12Column.Create;
  Result.Assign(C);

  case H of
    xchCol2     : begin
      Result.Max := ACol2;
      if ACol2 = C.Max then
        FCols.NilAndDelete(AColIndex)
      else
        C.Min := ACol2 + 1;
    end;
    xchCol1Match: begin
      Result.Max := ACol2;
      C.Min := ACol2 + 1
    end;
    xchMatch: begin
      FCols.NilAndDelete(AColIndex);
    end;
    xchTargetInside: begin
      FCols.NilAndDelete(AColIndex);
    end;
    xchCol2Match: begin
      Result.Min := ACol1;
      C.Max := ACol1 - 1
    end;
    xchInside: begin
      Result.Min := ACol1;
      Result.Max := ACol2;

      C2 := FCols.Add(FStyles.XFEditor.UseDefault);
      C2.Assign(C);
      FStyles.XFEditor.UseStyle(C2.Style);

      C.Max := ACol1 - 1;
      C2.Min := ACol2 + 1;
    end;
    xchCol1: begin
      Result.Min := ACol1;
      if ACol1 = C.Min then
        FCols.NilAndDelete(AColIndex)
      else
        C.Max := ACol1 - 1;
    end;
  end;
end;

function TXLSColumns.DefaultPixelWidth: integer;
begin
  Result := Round((((FDefaultWidth / 256) * 100 - 0.5) / 100) * FStyles.StdFontWidth + 5);
end;

procedure TXLSColumns.Delete(AColIndex, ACol1, ACol2: integer);
var
  H: TXc12ColumnHit;
  C,C2: TXc12Column;
begin
  C := FCols[AColIndex];
  H := C.Hit(ACol1,ACol2);
  if H in [xchLess,xchGreater] then
    Exit;

  case H of
    xchCol2        : begin
      if ACol2 = C.Max then
        FCols.NilAndDelete(AColIndex)
      else
        C.Min := ACol2 + 1;
    end;
    xchCol1Match   : C.Min := ACol2 + 1;
    xchMatch       : FCols.NilAndDelete(AColIndex);
    xchTargetInside: FCols.NilAndDelete(AColIndex);
    xchCol2Match   : C.Max := ACol1 - 1;
    xchInside      : begin
      C2 := FCols.Add(FStyles.XFEditor.UseDefault);
      C2.Assign(C);

      C.Max := ACol1 - 1;
      C2.Min := ACol2 + 1;
    end;
    xchCol1        : begin
      if ACol1 = C.Min then
        FCols.NilAndDelete(AColIndex)
      else
        C.Max := ACol1 - 1;
    end;
  end;
end;

procedure TXLSColumns.DeleteColumns(ACol1, ACol2: integer);
var
  i: integer;
  D: integer;
  iC1,iC2: integer;
begin
  CheckValidColumns(ACol1,ACol2);

  if not FindRange(ACol1,ACol2,iC1,iC2) then
    Exit;

  for i := iC1 to iC2 do
    Delete(i,ACol1,ACol2);

  D := ACol2 - ACol1 + 1;
  for i := iC2 to FCols.Count - 1 do begin
    if FCols[i] <> Nil then begin
      FCols[i].Min := FCols[i].Min - D;
      FCols[i].Max := FCols[i].Max - D;
    end;
  end;

  FCols.Pack;
end;

procedure TXLSColumns.DeleteHits(ACol1, ACol2: integer);
var
  i: integer;
  iC1,iC2: integer;
begin
  if not FindRange(ACol1,ACol2,iC1,iC2) then
    Exit;

  for i := iC2 downto iC1 do
    Delete(i,ACol1,ACol2);

  FCols.Pack;
end;

destructor TXLSColumns.Destroy;
begin
  FTheColumn.Free;
  inherited;
end;

procedure TXLSColumns.EndSetStyle(AList: TXc12Columns);
var
  i: integer;
begin
  for i := 0 to AList.Count - 1 do
    FCols.Add(AList[i]);

  Compact;
  Sort;
end;

procedure TXLSColumns.FillGaps(AList: TXc12Columns; ACol1, ACol2: integer);
var
  i: integer;
  Cnt: integer;
begin
  Cnt := AList.Count;
  if Cnt <= 0 then
    Exit;

  for i := 1 to Cnt - 1 do begin
    if (AList[i].Min - AList[i - 1].Max) > 1 then
      AList.Add(NewColumn(AList[i - 1].Max + 1,AList[i].Min - 1,FDefaultWidth));
  end;
  if ACol1 < AList[0].Min then
    AList.Add(NewColumn(ACol1,AList[0].Min - 1,FDefaultWidth));
  if ACol2 > AList[Cnt - 1].Max then
    AList.Add(NewColumn(AList[Cnt - 1].Max + 1,ACol2,FDefaultWidth));
end;

function TXLSColumns.FindCol(ACol: integer): TXc12Column;
var
  i: integer;
begin
  i := FindColIndex(ACol);
  if i >= 0 then
    Result := FCols[i]
  else
    Result := Nil;
end;

function TXLSColumns.FindColIndex(ACol: integer): integer;
var
  First,Last: integer;
begin
  First := 0;
  Last := FCols.Count - 1;
  while First <= Last do begin
    Result := (First + Last) div 2;
    if (ACol >= FCols[Result].Min) and (ACol <= FCols[Result].Max) then
      Exit
    else if FCols[Result].Max > ACol then
      Last := Result - 1
    else
      First := Result + 1;
  end;
  Result := -1;
end;

function TXLSColumns.FindRange(ACol1, ACol2: integer; var AIndex1, AIndex2: integer): boolean;
var
  i: integer;
begin
  AIndex1 := -1;
  AIndex2 := -1;
  for i := 0 to FCols.Count - 1 do begin
    case FCols[i].Hit(ACol1) of
      xchLess  : AIndex1 := i;
      xchInside: begin
        AIndex1 := i;
        Break;
      end;
    end;
  end;

  Result := AIndex1 >= 0;
  if not Result then
    Exit;

  if ACol1 = ACol2 then begin
    AIndex2 := AIndex1;
    Exit;
  end;

  for i := AIndex1 + 1 to FCols.Count - 1 do begin
    case FCols[i].Hit(ACol2) of
      xchInside : begin
        AIndex2 := i;
        Break;
      end;
      xchGreater: AIndex2 := i;
    end;
  end;

  Result := AIndex2 >= 0;
end;

function TXLSColumns.GetColWidth(ACol: integer): integer;
begin
  Result := Items[ACol].Width;
end;

function TXLSColumns.GetColWidthPixels(ACol: integer): integer;
begin
  Result := Items[ACol].PixelWidth;
end;

function TXLSColumns.CutHitList(ACol1, ACol2: integer): TXc12Columns;
var
  i,iC1,iC2: integer;
  C: TXc12Column;
begin
  Result := TXc12Columns.Create;
  Result.OwnsObjects := False;

  if not FindRange(ACol1,ACol2,iC1,iC2) then
    Exit;

  for i := iC1 to iC2 do begin
    C := Cut(i,ACol1,ACol2);
    if C <> Nil then
      Result.Add(C);
  end;

  FCols.Pack;
end;

function TXLSColumns.GetItems(Index: integer): TXLSColumn;
begin
  FTheColumn.Changed(Index);
  Result := FTheColumn;
end;

procedure TXLSColumns.ToggleGrouped(ACol: integer);
var
  XCol: TXLSColumn;
  UnGroup: boolean;
  Level: integer;

procedure SkipCollapsed(var Col: integer);
var
  XCol: TXLSColumn;
  Level: integer;
begin
  Dec(Col);
  XCol := Items[Col];
  if XCol <> Nil then
    Level := XCol.OutlineLevel
  else
    Level := -1;
  while (Col >= 0) and (XCol <> Nil) and (XCol.OutlineLevel >= Level) do begin
    if XCol.CollapsedOutline then
      SkipCollapsed(Col);
    Dec(Col);
    XCol := Items[Col];
  end;
  if Col > 0 then
    Inc(Col);
end;

begin
  if (ACol < 0) or (ACol > XLS_MAXCOL) then
    Exit;
  XCol := Items[ACol - 1];
  if (XCol = Nil) or (XCol.OutlineLevel <= 0) then
    Exit;
  Level := XCol.OutlineLevel;
  XCol := Items[ACol];
  if XCol <> Nil then begin
    UnGroup := XCol.CollapsedOutline;
    XCol.CollapsedOutline := not XCol.CollapsedOutline;
  end
  else begin
    UnGroup := False;
    Items[ACol].CollapsedOutline := True;
  end;
  Dec(ACol);
  while ACol >= 0 do begin
    XCol := Items[ACol];
    if (XCol = Nil) or (XCol.OutlineLevel < Level) then
      Break
    else if XCol.OutlineLevel >= Level then begin
      XCol.Hidden := not UnGroup;
      if XCol.CollapsedOutline then
        SkipCollapsed(ACol)
    end;
    Dec(ACol);
  end;
end;

procedure TXLSColumns.ToggleGroupedAll(ALevel: integer);
var
  i: integer;
  XCol: TXLSColumn;

procedure CheckButton(Index: integer; Value: boolean);
var
  XC: TXLSColumn;
begin
  XC := Items[Index];
  if (XC <> Nil) and (XC.OutlineLevel >= XCol.OutlineLevel) then
    Exit;
  XC := Items[Index];
  if (XC <> Nil) and (XC.OutlineLevel < XCol.OutlineLevel) then
    XC.CollapsedOutline := Value;
end;

begin
  for i := XLS_MAXCOL downto 0 do begin
    XCol := Items[i];
    if (XCol <> Nil) and (XCol.OutlineLevel > 0) then begin
      if XCol.OutlineLevel >= ALevel then
        XCol.Hidden := True
      else
        XCol.Hidden := False;
      CheckButton(i + 1,XCol.Hidden);
    end;
  end;
end;

procedure TXLSColumns.InsertColumns(ACol, AColCount: integer);
var
  C2: integer;
begin
  C2 := ACol + AColCount - 1;
  if C2 > XLS_MAXCOL then
    C2 := XLS_MAXCOL;
  CheckValidColumns(ACol,C2);

  MoveColumns(ACol,XLS_MAXCOL,C2 + 1);
end;

procedure TXLSColumns.InsertColumnsKeepFmt(ACol, AColCount: integer);
var
  Xc12Col,Xc12ColNew: TXc12Column;
begin
  InsertColumns(ACol,AColCount);
  if ACol > 0 then begin
    Xc12Col := FindCol(ACol - 1);
    if (Xc12Col <> Nil) and ((Xc12Col.Style <> FStyles.XFs.DefaultXF) or (Xc12Col.Width <> (FDefaultWidth / 256))) then begin
      FStyles.XFEditor.UseStyle(Xc12Col.Style);
      Xc12ColNew := FCols.Add(Xc12Col.Style);
      Xc12ColNew.Min := ACol;
      Xc12ColNew.Max := ACol + AColCount - 1;
      Xc12ColNew.Width := Xc12Col.Width;

      Sort;
      Compact;
    end;
  end;
end;

procedure TXLSColumns.InsertList(AList: TXc12Columns);
var
  i: integer;
  C: TXc12Column;
begin
  for i := 0 to AList.Count - 1 do begin
    C := AList[i];
    FCols.Add(C);
  end;

  Sort;
  Compact;
end;

procedure TXLSColumns.MoveColumns(ACol1, ACol2, ADestCol: integer);
var
  List: TXc12Columns;
begin
  CheckValidColumns(ACol1,ACol2);

  List := CutHitList(ACol1,ACol2);
  try
    List.OwnsObjects := False;

    DeleteHits(ADestCol,ADestCol + (ACol2 - ACol1));
    DeleteHits(ACol1,ACol2);

    OffsetList(List,ADestCol - ACol1);

    InsertList(List);
  finally
    List.Free;
  end;
end;

function TXLSColumns.NewColumn(ACol1, ACol2, AWidth: integer): TXc12Column;
begin
  Result := TXc12Column.Create;
  Result.Min := ACol1;
  Result.Max := ACol2;
  Result.Style := FStyles.XFEditor.UseDefault;
  FTheColumn.SetXc12ColWidth(Result,AWidth);
end;

procedure TXLSColumns.OffsetList(AList: TXc12Columns; const AOffset: integer);
var
  i: integer;
begin
  for i := 0 to AList.Count - 1 do  begin
    AList[i].Min := AList[i].Min + AOffset;
    AList[i].Max := AList[i].Max + AOffset;

    case AList[i].Hit(0) of
      xchLess   : ;
      xchInside : AList[i].Min := 0;
      xchGreater: AList.NilAndDelete(i);
    end;

    case AList[i].Hit(XLS_MAXCOL) of
      xchLess   : AList.NilAndDelete(i);
      xchInside : AList[i].Max := XLS_MAXCOL;
      xchGreater: ;
    end;
  end;
  AList.Pack;
end;

procedure TXLSColumns.SetColWidth(ACol1, ACol2, AValue: integer);
var
  i: integer;
  List: TXc12Columns;
begin
  CheckValidColumns(ACol1,ACol2);

  // Simple case, new columns added after the last.
  if (FCols.Count <= 0) or (FCols[FCols.Count - 1].Max < ACol1) then begin
    if AValue = FDefaultWidth then
      Exit;

    FCols.Add(NewColumn(ACol1,ACol2,AValue),FStyles.XFEditor.UseDefault);
  end
  else begin
    List := CutHitList(ACol1,ACol2);
    try
      if AValue <> FDefaultWidth then
        FillGaps(List,ACol1,ACol2);

      for i := 0 to List.Count - 1 do begin
        List[i].Width := AValue / 256;
        FCols.Add(List[i]);
      end;

      Compact;
      Sort;
    finally
      List.Free;
    end;
  end;
end;

procedure TXLSColumns.SetDefaultWidth(const Value: integer);
begin
  FDefaultWidth := Value;
  FCols.DefColWidth := FDefaultWidth / 256;
end;

procedure TXLSColumns.Assign(AColumns: TXLSColumns);
var
  i: integer;
begin
  Clear;

  for i := 0 to AColumns.FCols.Count - 1 do
    FCols.Add.Assign(AColumns.FCols[i]);
end;

function TXLSColumns.BeginSetStyle(ACol1, ACol2: integer; AStyle: TXc12XF): TXc12Columns;
begin
  CheckValidColumns(ACol1,ACol2);

  // Simple case, new columns added after the last.
  if (FCols.Count <= 0) or (FCols[FCols.Count - 1].Max < ACol1) then begin
    FStyles.XFEditor.UseStyle(AStyle);
    FCols.Add(NewColumn(ACol1,ACol2,FDefaultWidth),AStyle);
    Result := Nil;
  end
  else begin
    Result := CutHitList(ACol1,ACol2);

    FillGaps(Result,ACol1,ACol2);
  end;
end;

procedure TXLSColumns.Sort;
var
  iI, iJ, iK,
  iSize: integer;
begin
  // Shell sort. Never use Quicksort on almost (or sorted) lists.

  // Sort on two items don't work.
  if FCols.Count = 2 then begin
    if FCols[0].Min > FCols[1].Max then
      FCols.Exchange(0,1);
  end
  else begin
    iSize := FCols.Count - 1;
    iK := iSize shr 1;
    while iK > 0 do begin
      for iI := 0 to iSize - iK do begin
        iJ := iI;
        while (iJ >= 0) and (FCols[iJ].Min > FCols[iJ + iK].Max) do begin
          FCols.Exchange(iJ,iJ + iK);
          if iJ > iK then
            Dec(iJ, iK)
          else
            iJ := 0
        end;
      end;
      iK := iK shr 1;
    end;
  end;
end;

procedure TXLSColumns.CheckValidColumns(ACol1, ACol2: integer);
begin
  if not ((ACol1 >= 0) and (ACol2 >= 0) and (ACol1 <= XLS_MAXCOL) and (ACol2 <= XLS_MAXCOL) and (ACol2 >= ACol1)) then
    raise XLSRWException.Create('Invalid columns');
end;

end.
