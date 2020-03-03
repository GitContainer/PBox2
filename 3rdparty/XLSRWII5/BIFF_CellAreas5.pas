unit BIFF_CellAreas5;

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
{$ifdef BABOON}
{$else}
     Windows,
{$endif}
     BIFF_RecsII5, BIFF_Utils5,
     XLSUtils5;

type TCellEdge = (ceLeft,ceTop,ceRight,ceBottom);
     TCellEdges = set of TCellEdge;

type
//* Represents a cell area.
    TCellArea97 = class(TObject)
private
     function  GetAsRecArea: TRecCellAreaI;
     procedure SetAsRecArea(const Value: TRecCellAreaI);
     function  GetEdges(Col, Row: integer): TCellEdges;
protected
     FId: integer;
     FCol1,FRow1,FCol2,FRow2: integer;

     function FullInside(Col,Row: integer): boolean;
public
     //* Normalizes the area; Checks so that Col1 <= Col2 and Row1 <= Row2.
     procedure Normalize;
     //* Assignes another TCellArea to this area.
     //* ~param Source The source TCellArea.
     procedure Assign(Source: TCellArea97);
     //* Sets the size of the area. Same as setting the ~[link Col1], ~[link Row1], ~[link Col2] and ~[link Row2] properties.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     procedure SetSize(C1,R1,C2,R2: integer);
     //* Returns the coordinates of the area as an TRecCellArea object.
     property AsRecArea: TRecCellAreaI read GetAsRecArea write SetAsRecArea;

     //* Returns the edges in the area that the position given by Col and Row borders to.
     //* Col and Row must be inside the area.
     property Edges[Col,Row: integer]: TCellEdges read GetEdges;
     //* Check if a position is in the area.
     //* ~param Col Column of the porition.
     //* ~param Row Row of the porition.
     //* ~result True if the cell is in the area.
     function  CellInArea(Col,Row: integer): boolean;
     //* Tests if an area is equal to the coordinates of this area.
     //* ~param C1 Left column of the test area.
     //* ~param R1 Top row of the test area.
     //* ~param C2 Right column of the test area.
     //* ~param R2 Bottom row of the test area.
     //* ~result True is the areas are equal.
     function  Equal(C1,R1,C2,R2: integer): boolean;

     //* First column in the area.
     property Col1: integer read FCol1 write FCol1;
     //* First row in the area.
     property Row1: integer read FRow1 write FRow1;
     //* Last column in the area.
     property Col2: integer read FCol2 write FCol2;
     //* Last row in the area.
     property Row2: integer read FRow2 write FRow2;

     property Id: integer read FId;
     end;

type
//* Virtual base class for objects that manipulates cell areas.
    TBaseCellAreas97 = class(TObjectList)
private
     function GetItems(Index: integer): TCellArea97;
protected
     FIdCount: integer;
     FFoundAreaEvent: TIntegerEvent;

     function  IntersectArea(Source1,Source2: TRecCellAreaI; var Dest: TRecCellAreaI): boolean;
     function  Split(Index: integer; Area: TRecCellAreaI; var SplitAreas: array of TRecCellAreaI): integer;
     function  NewId: integer;
public
     //* ~exclude
     constructor Create;
     //* Assigns a TBaseCellAreas97 to this areas.
     procedure Assign(Areas: TBaseCellAreas97);
     //* Add a new cell area.
     //* ~result The added area.
     function  Add: TCellArea97; overload;
     //* ~exclude
     function  Add(RecArea: PRecCellArea): TCellArea97; overload;
     //* Add a new cell area.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     //* ~result The added area.
     function  Add(C1,R1,C2,R2: integer): TCellArea97; overload;
     //* Returns the smallest area that all areas in the list fits in.
     //* ~result The area.
     function  TotExtent: TRecCellAreaI;
     //* Normalizes (i.e. checks that C1 < C2 and R1 < R2).
     procedure NormalizeAll;
     //* Checks if an area intersects any of the areas.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~result True if the area given by Col1,Row1,Col2,Row2 intersects any
     //* of the areas in the list.
     function  AreaInAreas(Col1,Row1,Col2,Row2: integer): boolean;
     //* Searches for an area that is equal to Col1,Row1,Col2,Row2. If a area is
     //* found, the index to the area is returned. If not found, -1 is returned.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~result The index of the found area. If not found, -1 is returned.
     function  FindArea(Col1,Row1,Col2,Row2: integer): integer;
     //* Check if a cell is in any of the areas.
     //* Returns the index of the area that the cell address given by Col and
     //* Row is within. If no match is found, -1 is returned.
     //* ~param Col Column of the cell.
     //* ~param Row Row of the cell.
     //* ~result The index of the found area. If not found, -1 is returned.
     function  CellInAreas(Col,Row: integer): integer;
     //* ~exclude
     procedure SendAllInArea(Col1,Row1,Col2,Row2: integer);
     //* ~exclude
     procedure Copy(Col1,Row1,Col2,Row2: word; DeltaCol,DeltaRow: integer); virtual;
     //* ~exclude
     procedure Delete(Col1,Row1,Col2,Row2: integer); overload; virtual; abstract;
     //* ~exclude
     procedure Include(Col1,Row1,Col2,Row2: integer); virtual; abstract;
     //* ~exclude
     procedure Move(DeltaCol,DeltaRow: integer); overload; virtual; abstract;
     //* ~exclude
     procedure Move(Col1,Row1,Col2,Row2: integer; DeltaCol,DeltaRow: integer); overload; virtual; abstract;

     //* The areas in the list.
     property Items[Index: integer]: TCellArea97 read GetItems; default;
     //* ~exclude
     property OnFoundArea: TIntegerEvent read FFoundAreaEvent write FFoundAreaEvent;
     end;

type
//* Base class for objects that manipulates cell areas.
    TCellAreas97 = class(TBaseCellAreas97)
private
protected
     function  Combine(C1,R1,C2,R2: word): boolean;
public
     //* Copies cell areas.
     //* Copies the cell areas that intersects the area given by Col1,Row1,
     //* Col2,Row2. The areas are copied DeltaCol and DeltaRow offset from
     //* their current position. DeltaCol and DeltaRow can be negative for
     //* copying left/up.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be copied.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be copied.
     procedure Copy(Col1,Row1,Col2,Row2: word; DeltaCol,DeltaRow: integer); override;
     //* Delete areas.
     //* Delete the areas that intersects the area given by Col1,Row1,Col2,Row2.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure Delete(Col1,Row1,Col2,Row2: integer); overload; override;
     //* Includes areas.
     //* All areas that intersects the area given by Col1,Row1,Col2,Row2 are
     //* kept. All others are deleted.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure Include(Col1,Row1,Col2,Row2: integer); override;
     //* Moves areas.
     //* Move all areas by the offset DeltaCol,DeltaRow. DeltaCol and DeltaRow
     //* can be negative for moving left/up.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be moved.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be moved.
     procedure Move(DeltaCol,DeltaRow: integer); overload; override;
     //* Moves areas.
     //* Move all areas that intersects the area given by Col1,Row1,Col2,Row2,
     //* by the offset DeltaCol,DeltaRow. DeltaCol and DeltaRow can be negative
     //* for moving left/up.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be moved.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be moved.
     procedure Move(Col1,Row1,Col2,Row2: integer; DeltaCol,DeltaRow: integer); overload; override;
     end;

type
//* Base class for objects that manipulates cell areas. Solid areas are areas
//* that not can be splitted by delete/insert/move operations, such as merged
//* cells.
    TSolidCellAreas97 = class(TBaseCellAreas97)
private
protected
public
     //* Copies cell areas.
     //* Copies the cell areas that intersects the area given by Col1,Row1,
     //* Col2,Row2. The areas are copied DeltaCol and DeltaRow offset from
     //* their current position. DeltaCol and DeltaRow can be negative for
     //* copying left/up.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be copied.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be copied.
     procedure Copy(Col1,Row1,Col2,Row2: word; DeltaCol,DeltaRow: integer); override;
     //* Delete areas.
     //* Delete the areas that intersects the area given by Col1,Row1,Col2,Row2.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure Delete(Col1,Row1,Col2,Row2: integer); overload; override;
     //* Deprecated. Solid areas can not be included with other areas.
     //* ~param Col1 Deprecated. Solid areas can not be included with other areas..
     //* ~param Row1 Deprecated. Solid areas can not be included with other areas..
     //* ~param Col2 Deprecated. Solid areas can not be included with other areas..
     //* ~param Row2 Deprecated. Solid areas can not be included with other areas..
     procedure Include(Col1,Row1,Col2,Row2: integer); override;
     //* Moves areas.
     //* Move all areas by the offset DeltaCol,DeltaRow. DeltaCol and DeltaRow
     //* can be negative for moving left/up.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be moved.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be moved.
     procedure Move(DeltaCol,DeltaRow: integer); overload; override;
     //* Moves areas.
     //* Move all areas that intersects the area given by Col1,Row1,Col2,Row2,
     //* by the offset DeltaCol,DeltaRow. DeltaCol and DeltaRow can be negative
     //* for moving left/up.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be moved.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be moved.
     procedure Move(Col1,Row1,Col2,Row2: integer; DeltaCol,DeltaRow: integer); overload; override;
     end;

implementation

{ TBaseCellAreas }

function TBaseCellAreas97.Add: TCellArea97;
begin
  Result := TCellArea97.Create;
  Result.FId := NewId;
  inherited Add(Result);
end;

function TBaseCellAreas97.Add(RecArea: PRecCellArea): TCellArea97;
begin
  Result := TCellArea97.Create;
  Result.FCol1 := RecArea.Col1;
  Result.FRow1 := RecArea.Row1;
  Result.FCol2 := RecArea.Col2;
  Result.FRow2 := RecArea.Row2;
  Result.FId := NewId;
  inherited Add(Result);
end;

constructor TBaseCellAreas97.Create;
begin
  inherited Create;
end;

function TBaseCellAreas97.FindArea(Col1, Row1, Col2, Row2: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Equal(Col1, Row1, Col2, Row2) then
      Exit;
  end;
  Result := -1;
end;

function TBaseCellAreas97.GetItems(Index: integer): TCellArea97;
begin
  Result := TCellArea97(inherited Items[Index]);
end;

function TBaseCellAreas97.AreaInAreas(Col1, Row1, Col2, Row2: integer): boolean;
var
  i: integer;
  SelectArea,NewArea: TRecCellAreaI;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  for i := 0 to Count - 1 do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

procedure TBaseCellAreas97.Assign(Areas: TBaseCellAreas97);
var
  i: integer;
begin
  Clear;
  for i := 0 to Areas.Count - 1 do
    Add.Assign(Areas[i]);
end;

function TBaseCellAreas97.NewId: integer;
begin
  Result := FIdCount;
  Inc(FIdCount);
end;

procedure TBaseCellAreas97.NormalizeAll;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Normalize;
end;

function TBaseCellAreas97.TotExtent: TRecCellAreaI;
var
  i: integer;
begin
  Result.Row1 := MAXINT;
  Result.Row2 := 0;
  Result.Col1 := MAXINT;
  Result.Col2 := 0;
  for i := 0 to Count - 1 do begin
    Result.Col1 := Min(Result.Col1,Items[i].FCol1);
    Result.Row1 := Min(Result.Row1,Items[i].FRow1);
    Result.Col2 := Max(Result.Col2,Items[i].FCol2);
    Result.Row2 := Max(Result.Row2,Items[i].FRow2);
  end;
end;

function TBaseCellAreas97.Add(C1, R1, C2, R2: integer): TCellArea97;
begin
  Result := Add;
  Result.FCol1 := C1;
  Result.FRow1 := R1;
  Result.FCol2 := C2;
  Result.FRow2 := R2;
end;


function TBaseCellAreas97.IntersectArea(Source1,Source2: TRecCellAreaI; var Dest: TRecCellAreaI): boolean;
begin
  if Source1.Col1 > Source2.Col1 then Dest.Col1 := Source1.Col1 else Dest.Col1 := Source2.Col1;
  if Source1.Row1 > Source2.Row1 then Dest.Row1 := Source1.Row1 else Dest.Row1 := Source2.Row1;
  if Source1.Col2 < Source2.Col2 then Dest.Col2 := Source1.Col2 else Dest.Col2 := Source2.Col2;
  if Source1.Row2 < Source2.Row2 then Dest.Row2 := Source1.Row2 else Dest.Row2 := Source2.Row2;
  Result := (Dest.Row1 <= Dest.Row2) and (Dest.Col1 <= Dest.Col2);
end;

function TBaseCellAreas97.CellInAreas(Col, Row: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if (Col >= Items[Result].FCol1) and (Col <= Items[Result].FCol2) and (Row >= Items[Result].FRow1) and (Row <= Items[Result].FRow2) then
      Exit;
  end;
  Result := -1;
end;

function TBaseCellAreas97.Split(Index: integer; Area: TRecCellAreaI; var SplitAreas: array of TRecCellAreaI): integer;
var
  TmpR1,TmpR2: word;
begin
  Result := 0;
  TmpR1 := Items[Index].FRow1;
  TmpR2 := Items[Index].FRow2;

  if Area.Row1 > Items[Index].FRow1 then begin
    SplitAreas[Result].Col1 := Items[Index].FCol1;
    SplitAreas[Result].Row1 := Items[Index].FRow1;
    SplitAreas[Result].Col2 := Items[Index].FCol2;
    SplitAreas[Result].Row2 := Area.Row1 - 1;
    Inc(Result);
    TmpR1 := Area.Row1;
  end;

  if Area.Row2 < Items[Index].FRow2 then begin
    SplitAreas[Result].Col1 := Items[Index].FCol1;
    SplitAreas[Result].Row1 := Area.Row2 + 1;
    SplitAreas[Result].Col2 := Items[Index].FCol2;
    SplitAreas[Result].Row2 := Items[Index].FRow2;
    Inc(Result);
    TmpR2 := Area.Row2;
  end;

  if Area.Col1 > Items[Index].FCol1 then begin
    SplitAreas[Result].Col1 := Items[Index].Col1;
    SplitAreas[Result].Row1 := TmpR1;
    SplitAreas[Result].Col2 := Area.Col1 - 1;
    SplitAreas[Result].Row2 := TmpR2;
    Inc(Result);
  end;

  if Area.Col2 < Items[Index].FCol2 then begin
    SplitAreas[Result].Col1 := Area.Col2 + 1;
    SplitAreas[Result].Row1 := TmpR1;
    SplitAreas[Result].Col2 := Items[Index].Col2;
    SplitAreas[Result].Row2 := TmpR2;
    Inc(Result);
  end;
end;

procedure TBaseCellAreas97.Copy(Col1, Row1, Col2, Row2: word; DeltaCol, DeltaRow: integer);
begin

end;

procedure TBaseCellAreas97.SendAllInArea(Col1, Row1, Col2, Row2: integer);
var
  i: integer;
begin
  if not Assigned(FFoundAreaEvent) then
    Exit;
  for i := 0 to Count - 1 do begin
    if Items[i].FCol2 < Col1 then
      Continue;
    if Items[i].FCol1 > Col2 then
      Continue;
    if Items[i].FRow2 < Row1 then
      Continue;
    if Items[i].FRow1 > Row2 then
      Continue;
    FFoundAreaEvent(Self,i);
  end;
end;

{ TCellAreas }

procedure TCellAreas97.Delete(Col1, Row1, Col2, Row2: integer);
var
  i,j,Cnt,SplitCount: integer;
  SelectArea,NewArea: TRecCellAreaI;
  SplitAreas: array[0..3] of TRecCellAreaI;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  i := 0;
  // Save Count as number of elements may increase;
  Cnt := Count;
  while i < Cnt do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
      SplitCount := Split(i,SelectArea,SplitAreas);
      for j := 0 to SplitCount - 1 do
        Add(@SplitAreas[i]);
      Delete(i);
      Dec(Cnt);
    end
    else
      Inc(i);
  end;
end;

procedure TCellAreas97.Copy(Col1, Row1, Col2, Row2: word; DeltaCol,DeltaRow: integer);
var
  i,Cnt: integer;
  SelectArea,NewArea: TRecCellAreaI;

procedure DoCopy(C1,R1,C2,R2: integer);
begin
  Inc(C1,DeltaCol);
  Inc(R1,DeltaRow);
  Inc(C2,DeltaCol);
  Inc(R2,DeltaRow);
  if ClipAreaToSheet(C1,R1,C2,R2) and not Combine(C1,R1,C2,R2) then begin
    Delete(C1,R1,C2,R2);
    Add(C1,R1,C2,R2);
  end;
end;

begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  Cnt := Count;
  i := 0;
  while i < Cnt do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then
      DoCopy(NewArea.Col1,NewArea.Row1,NewArea.Col2,NewArea.Row2);
    Inc(i);
  end;
end;

procedure TCellAreas97.Move(DeltaCol, DeltaRow: integer);
var
  i,C1,R1,C2,R2: integer;
begin
  i := 0;
  while i < Count do begin
    C1 := Items[i].FCol1 + DeltaCol;
    R1 := Items[i].FRow1 + DeltaRow;
    C2 := Items[i].FCol2 + DeltaCol;
    R2 := Items[i].FRow2 + DeltaRow;
    if not ClipAreaToSheet(C1,R1,C2,R2) then
      Delete(i)
    else begin
      Items[i].SetSize(C1,R1,C2,R2);
      Inc(i);
    end;
  end;
end;

procedure TCellAreas97.Move(Col1, Row1, Col2, Row2: integer; DeltaCol, DeltaRow: integer);
var
  i,j,Cnt,SplitCount: integer;
  SelectArea,NewArea: TRecCellAreaI;
  SplitAreas: array[0..3] of TRecCellAreaI;

procedure DoMove(C1,R1,C2,R2: integer);
begin
  Delete(i);
  Dec(Cnt);
  Inc(C1,DeltaCol);
  Inc(R1,DeltaRow);
  Inc(C2,DeltaCol);
  Inc(R2,DeltaRow);
  if ClipAreaToSheet(C1,R1,C2,R2) and not Combine(C1,R1,C2,R2) then begin
    Delete(C1,R1,C2,R2);
    Add(C1,R1,C2,R2);
  end;
end;

begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  i := 0;
  // Save Count as number of elements may increase;
  Cnt := Count;
  while i < Cnt do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
      SplitCount := Split(i,SelectArea,SplitAreas);
      for j := 0 to SplitCount - 1 do
        Add(@SplitAreas[i]);
      DoMove(NewArea.Col1,NewArea.Row1,NewArea.Col2,NewArea.Row2);
    end
    else
      Inc(i);
  end;
end;

// TODO: Will not handle the case where the new area fills the space between
// two existing areas. If included, check dependent functions, as this will
// require that one area is deleted.
function TCellAreas97.Combine(C1, R1, C2, R2: word): boolean;
var
  i: integer;
begin
  Result := True;
  for i := 0 to Count - 1 do begin
    with Items[i] do begin
      // Entirly inside the area.
      if (C1 >= FCol1) and (R1 >= FRow1) and (C2 <= FCol2) and (R2 <= FRow2) then
        Exit;
      if (FCol1 = C1) and (FCol2 = C2) then begin
        if (R2 >= (FRow1 - 1)) and (R2 <= FRow2) then begin
          FRow1 := R1;
          Exit;
        end
        else if (R1 >= FRow1) and (R1 <= (FRow2 + 1)) then begin
          FRow2 := R2;
          Exit;
        end;
      end
      else if (FRow1 = R1) and (FRow2 = R2) then begin
        if (C2 >= (FCol1 - 1)) and (C2 <= FCol2) then begin
          FCol1 := C1;
          Exit;
        end
        else if (C1 >= FCol1) and (C1 <= (FCol2 + 1)) then begin
          FCol2 := C2;
          Exit;
        end;
      end;
    end;
  end;
  Result := False;
end;

procedure TCellAreas97.Include(Col1, Row1, Col2, Row2: integer);
var
  i,Cnt: integer;
  SelectArea,NewArea: TRecCellAreaI;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  Cnt := Count;
  i := 0;
  while i < Cnt do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
      Items[i].AsRecArea := NewArea;
      Inc(i);
    end
    else begin
      Delete(i);
      Dec(Cnt);
    end;
  end;
end;

{ TCellArea }

procedure TCellArea97.Assign(Source: TCellArea97);
begin
  FRow1 := TCellArea97(Source).FRow1;
  FRow2 := TCellArea97(Source).FRow2;
  FCol1 := TCellArea97(Source).FCol1;
  FCol2 := TCellArea97(Source).FCol2;
end;

function TCellArea97.CellInArea(Col, Row: integer): boolean;
begin
  Result := (Col >= Col1) and (Row >= Row1) and (Col <= Col2) and (Row <= Row2);
end;

function TCellArea97.Equal(C1, R1, C2, R2: integer): boolean;
begin
  Result := (C1 = FCol1) and (R1 = FRow1) and (C2 = FCol2) and (R2 = FRow2);
end;

function TCellArea97.FullInside(Col, Row: integer): boolean;
begin
  Result := (Col > FCol1) and (Row > FRow1) and (Col < FCol2) and (Row < FRow2);
end;

function TCellArea97.GetAsRecArea: TRecCellAreaI;
begin
  Result.Col1 := FCol1;
  Result.Row1 := FRow1;
  Result.Col2 := FCol2;
  Result.Row2 := FRow2;
end;

function TCellArea97.GetEdges(Col, Row: integer): TCellEdges;
begin
  Result := [];
  if (Col >= FCol1) and (Col <= FCol2) and (Row >= FRow1) and (Row <= FRow2) then begin
    if Col = FCol1 then Result := Result + [ceLeft];
    if Row = FRow1 then Result := Result + [ceTop];
    if Col = FCol2 then Result := Result + [ceRight];
    if Row = FRow2 then Result := Result + [ceBottom];
  end;
end;

procedure TCellArea97.Normalize;

procedure Swap(var W1,W2: integer);
var
  T: Word;
begin
  T := W1;
  W1 := W2;
  W2 := T;
end;

begin
  if FCol1 > FCol2 then
    Swap(FCol1,FCol2);
  if FRow1 > FRow2 then
    Swap(FRow1,FRow2);
end;

procedure TCellArea97.SetAsRecArea(const Value: TRecCellAreaI);
begin
  FCol1 := Value.Col1;
  FRow1 := Value.Row1;
  FCol2 := Value.Col2;
  FRow2 := Value.Row2;
end;

procedure TCellArea97.SetSize(C1, R1, C2, R2: integer);
begin
  FCol1 := C1;
  FRow1 := R1;
  FCol2 := C2;
  FRow2 := R2;
end;

{ TSolidCellAreas }

procedure TSolidCellAreas97.Copy(Col1, Row1, Col2, Row2: word; DeltaCol, DeltaRow: integer);
begin

end;

procedure TSolidCellAreas97.Delete(Col1, Row1, Col2, Row2: integer);
var
  i,SplitCount: integer;
  SelectArea,NewArea: TRecCellAreaI;
  SplitAreas: array[0..3] of TRecCellAreaI;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  i := 0;
  while i < Count do begin
    if (Items[i].FCol1 >= Col1) and (Items[i].FCol1 <= Col2) and (Items[i].FRow1 >= Row1) and (Items[i].FRow1 <= Row2) and
       (Items[i].FCol2 >= Col1) and (Items[i].FCol2 <= Col2) and (Items[i].FRow2 >= Row1) and (Items[i].FRow2 <= Row2) then
      Delete(i)
    else begin
      if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
        SplitCount := Split(i,SelectArea,SplitAreas);
        if SplitCount > 1 then
          raise XLSRWException.Create('Can not change part of a merged cell')
        else if SplitCount = 1 then
          Items[i].AsRecArea := SplitAreas[0];
      end;
      Inc(i);
    end;
  end;
end;

procedure TSolidCellAreas97.Include(Col1, Row1, Col2, Row2: integer);
begin

end;

procedure TSolidCellAreas97.Move(Col1, Row1, Col2, Row2: integer; DeltaCol, DeltaRow: integer);
var
  i: integer;
  R1,C1,R2,C2: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FullInside(Col1,Row1) or
       Items[i].FullInside(Col2,Row1) or
       Items[i].FullInside(Col1,Row2) or
       Items[i].FullInside(Col2,Row2) then
      raise XLSRWException.Create('Can not change part of a merged cell');
  end;
  i := 0;
  while i < Count do begin
    C1 := Items[i].FCol1;
    R1 := Items[i].FRow1;
    C2 := Items[i].FCol2;
    R2 := Items[i].FRow2;
    if (C1 >= Col1) and (C1 <= Col2) then
      Inc(C1,DeltaCol);
    if (R1 >= Row1) and (R1 <= Row2) then
      Inc(R1,DeltaRow);
    if (C2 >= Col1) and (C2 <= Col2) then
      Inc(C2,DeltaCol);
    if (R2 >= Row1) and (R2 <= Row2) then
      Inc(R2,DeltaRow);
    if (C2 < C1) or (R2 < R1) then
      Delete(i)
    else begin
      Items[i].FCol1 := C1;
      Items[i].FRow1 := R1;
      Items[i].FCol2 := C2;
      Items[i].FRow2 := R2;
      Inc(i);
    end;
  end;
end;

procedure TSolidCellAreas97.Move(DeltaCol, DeltaRow: integer);
begin
  raise XLSRWException.Create('TSolidCellAreas.Move not implemented');
end;

end.
