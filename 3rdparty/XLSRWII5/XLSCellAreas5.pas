unit XLSCellAreas5;

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
     XLSUtils5,
     Xc12Utils5, Xc12Common5;

type TCellEdge = (ceLeft,ceTop,ceRight,ceBottom);
     TCellEdges = set of TCellEdge;

type TCellRef = class(TObject)
private
     function  GetAsString: AxUCString;
     procedure SetAsString(const Value: AxUCString);
protected
     FCol,FRow: integer;
public
     procedure Clear;

     property Col: integer read FCol write FCol;
     property Row: integer read FRow write FRow;

     property AsString: AxUCString read GetAsString write SetAsString;
     end;

type TCellRefs = class(TObjectList)
private
     function GetItems(const Index: integer): TCellRef;
protected
public
     procedure Add(const ARef: AxUCString);

     property Items[const Index: integer]: TCellRef read GetItems;
     end;

type
//* Represents a cell area.
    TCellArea = class(TObject)
private
     function  GetAsRecArea: TXLSCellArea;
     procedure SetAsRecArea(const Value: TXLSCellArea);
     function  GetAsStringAbs: AxUCString;
protected
     FTag: integer;
     // Any Object asociated with this area.
     FObject: TObject;
//     FCol1,FRow1,FCol2,FRow2: integer;

     function  GetEdges(const Col, Row: integer): TCellEdges;
     function  GetAsString: AxUCString;
     function  GetCol1: integer; virtual; abstract;
     function  GetRow1: integer; virtual; abstract;
     function  GetCol2: integer; virtual; abstract;
     function  GetRow2: integer; virtual; abstract;
     procedure SetCol1(const AValue: integer); virtual; abstract;
     procedure SetRow1(const AValue: integer); virtual; abstract;
     procedure SetCol2(const AValue: integer); virtual; abstract;
     procedure SetRow2(const AValue: integer); virtual; abstract;
     function  FullInside(const Col,Row: integer): boolean;
public
     //* Normalizes the area; Checks so that Col1 <= Col2 and Row1 <= Row2.
     procedure Normalize;
     //* Assignes another TCellArea to this area.
     //* ~param Source The source TCellArea.
     procedure Assign(Source: TCellArea);
     //* Sets the size of the area. Same as setting the ~[link Col1], ~[link Row1], ~[link Col2] and ~[link Row2] properties.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     procedure SetSize(const C,R: integer); overload;
     procedure SetSize(const C1,R1,C2,R2: integer); overload;

     //* True if the area is fully inside C1,R1,C2,R2.
     function  InArea(const C1,R1,C2,R2: integer): boolean;
     //* True if any of the area corners is inside C1,R1,C2,R2.
     function  CornerInArea(const C1,R1,C2,R2: integer): boolean; overload;
     function  CornerInArea(AArea: TCellArea): boolean; overload;

     function  CellCount: integer;
     function  Width: integer;
     function  Height: integer;

     function  IsColumn: boolean;
     function  IsRow: boolean;

     //* Returns the coordinates of the area as an TRecCellArea object.
     property AsRecArea: TXLSCellArea read GetAsRecArea write SetAsRecArea;

     //* Returns the edges in the area that the position given by Col and Row borders to.
     //* Col and Row must be inside the area.
     property Edges[const Col,Row: integer]: TCellEdges read GetEdges;
     //* Check if a position is in the area.
     //* ~param Col Column of the porition.
     //* ~param Row Row of the porition.
     //* ~result True if the cell is in the area.
     function  CellInArea(const Col,Row: integer): boolean;
     //* Tests if an area is equal to the coordinates of this area.
     //* ~param C1 Left column of the test area.
     //* ~param R1 Top row of the test area.
     //* ~param C2 Right column of the test area.
     //* ~param R2 Bottom row of the test area.
     //* ~result True is the areas are equal.
     function  Equal(const C1,R1,C2,R2: integer): boolean;

     //* First column in the area.
     property Col1: integer read GetCol1 write SetCol1;
     //* First row in the area.
     property Row1: integer read GetRow1 write SetRow1;
     //* Last column in the area.
     property Col2: integer read GetCol2 write SetCol2;
     //* Last row in the area.
     property Row2: integer read GetRow2 write SetRow2;

     property AsString: AxUCString read GetAsString;
     property AsStringAbs: AxUCString read GetAsStringAbs;

     property Tag: integer read FTag write FTag;
     property Obj: TObject read FObject write FObject;
     end;

type TCellAreaImpl = class(TCellArea)
protected
     FCol1,FRow1,FCol2,FRow2: integer;

     function  GetCol1: integer; override;
     function  GetRow1: integer; override;
     function  GetCol2: integer; override;
     function  GetRow2: integer; override;
     procedure SetCol1(const AValue: integer); override;
     procedure SetRow1(const AValue: integer); override;
     procedure SetCol2(const AValue: integer); override;
     procedure SetRow2(const AValue: integer); override;
public
     end;

type
//* Virtual base class for objects that manipulates cell areas.
     TBaseCellAreas = class(TObjectList)
private
     function  GetDelimitedText: AxUCString;
     procedure SetDelimitedText(const Value: AxUCString);
     function  GetItems(const Index: integer): TCellArea;
protected
     FTag      : integer;
     FDelimiter: AxUCChar;

     function  IntersectArea(Source1,Source2: TXLSCellArea; var Dest: TXLSCellArea): boolean;
     function  Split(const Index: integer; Area: TXLSCellArea; var SplitAreas: array of TXLSCellArea): integer;
     function  CreateObject: TCellArea; virtual;
public
     //* ~exclude
     constructor Create;
     //* Assigns a TBaseCellAreas to this areas.
     procedure Assign(Areas: TBaseCellAreas);
     //* Add a new cell area.
     //* ~result The added area.
     function  Add: TCellArea; overload;
     //* Add a new cell area.
     //* ~param C1 Left column.
     //* ~param R1 Top row.
     //* ~param C2 Right column.
     //* ~param R2 Bottom row.
     //* ~result The added area.
     function  Add(const ARef: AxUCString): TCellArea; overload;
     function  Add(const AArea: PXLSCellArea): TCellArea; overload;
     function  Add(const C,R: integer): TCellArea; overload;
     function  Add(const C1,R1,C2,R2: integer): TCellArea; overload;
     procedure Add(const C1,R1,C2,R2: integer; AAreas: TBaseCellAreas); overload;

     function  CellCount: integer;
     //* Returns the smallest area that all areas in the list fits in.
     //* ~result The area.
     function  TotExtent: TXLSCellArea;
     //* Normalizes (i.e. checks that C1 < C2 and R1 < R2).
     procedure NormalizeAll;
     //* Max height (rows) of the area with most rows in the list.
     function  MaxWidth: integer;
     //* Max width (columns) of the area with most columns in the list.
     function  MaxHeight: integer;
     //* Checks if an area intersects any of the areas.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~result True if the area given by Col1,Row1,Col2,Row2 intersects any
     //* of the areas in the list.
     function  AreaInAreas(AArea: TCellArea): boolean; overload;
     function  AreaInAreas(const Col1,Row1,Col2,Row2: integer): boolean; overload;
     //* Copies all areas that intersects Col1,Row1,Col2,Row2.
     //* Areas are clipped to Col1,Row1,Col2,Row2.
     function  FillAndClipHitList(const Col1,Row1,Col2,Row2: integer; AList: TBaseCellAreas; AObject: TObject = Nil): boolean;
     //* Copies all areas that are inside Col1,Row1,Col2,Row2.
     function  FillHitList(const Col1,Row1,Col2,Row2: integer; AList: TBaseCellAreas): boolean;
     //* Searches for an area that is equal to Col1,Row1,Col2,Row2. If a area is
     //* found, the index to the area is returned. If not found, -1 is returned.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     //* ~result The index of the found area. If not found, -1 is returned.
     function  FindArea(const Col1,Row1,Col2,Row2: integer): integer; overload;
     function  FindArea(const Col,Row: integer): integer; overload;
     //* Check if a cell is in any of the areas.
     //* Returns the index of the area that the cell address given by Col and
     //* Row is within. If no match is found, -1 is returned.
     //* ~param Col Column of the cell.
     //* ~param Row Row of the cell.
     //* ~result The index of the found area. If not found, -1 is returned.
     function  CellInAreas(const Col,Row: integer): integer;
     //* ~exclude
     procedure Copy(const Col1,Row1,Col2,Row2: integer; const DeltaCol,DeltaRow: integer); virtual;
     //* ~exclude
     procedure Delete(const Col1,Row1,Col2,Row2: integer); overload; virtual; abstract;
     //* ~exclude
     procedure DeleteAndAdjust(const Col1,Row1,Col2,Row2: integer); virtual; abstract;
     //* ~exclude
     function  Include(const Col1,Row1,Col2,Row2: integer): TCellArea; virtual; abstract;
     //* ~exclude
     procedure Move(const DeltaCol,DeltaRow: integer); overload; virtual; abstract;
     //* ~exclude
     procedure Move(const Col1,Row1,Col2,Row2: integer; const DeltaCol,DeltaRow: integer); overload; virtual; abstract;

     //* The areas in the list.
     property Items[const Index: integer]: TCellArea read GetItems; default;

     property Delimiter: AxUCChar read FDelimiter;
     property DelimitedText: AxUCString read GetDelimitedText write SetDelimitedText;

     property Tag: integer read FTag write FTag;
     end;

type
//* Base class for objects that manipulates cell areas.
    TCellAreas = class(TBaseCellAreas)
private
protected
     function  CheckAndClip(var C1,R1,C2,R2: integer): boolean;
     function  Combine(const C1,R1,C2,R2: integer): boolean;
public
     function Last: TCellArea;

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
     procedure Copy(const Col1,Row1,Col2,Row2: integer; const DeltaCol,DeltaRow: integer); override;
     //* Delete areas.
     //* Delete the areas that intersects the area given by Col1,Row1,Col2,Row2.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure Delete(const Col1,Row1,Col2,Row2: integer); overload; override;
     //* Includes areas.
     //* All areas that intersects the area given by Col1,Row1,Col2,Row2 are
     //* kept. All others are deleted.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure DeleteAndAdjust(const Col1,Row1,Col2,Row2: integer); override;
     function  Include(const Col1,Row1,Col2,Row2: integer): TCellArea; override;
     //* Moves areas.
     //* Move all areas by the offset DeltaCol,DeltaRow. DeltaCol and DeltaRow
     //* can be negative for moving left/up.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be moved.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be moved.
     procedure Move(const DeltaCol,DeltaRow: integer); overload; override;
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
     procedure Move(const Col1,Row1,Col2,Row2: integer; const DeltaCol,DeltaRow: integer); overload; override;
     end;

type
//* Base class for objects that manipulates cell areas. Solid areas are areas
//* that not can be splitted by delete/insert/move operations, such as merged
//* cells.
     TSolidCellAreas = class(TBaseCellAreas)
private
protected
     FCopyFailed: boolean;
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
     procedure Copy(const Col1,Row1,Col2,Row2: integer; const DeltaCol,DeltaRow: integer); override;
     //* Delete areas.
     //* Delete the areas that intersects the area given by Col1,Row1,Col2,Row2.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure Delete(const Col1,Row1,Col2,Row2: integer); overload; override;
     //* All areas that intersects the area given by Col1,Row1,Col2,Row2 are
     //* merged into one single area, where all arwas fits.
     //* ~param Col1 Left column.
     //* ~param Row1 Top row.
     //* ~param Col2 Right column.
     //* ~param Row2 Bottom row.
     procedure DeleteAndAdjust(const Col1,Row1,Col2,Row2: integer); override;
     function  Include(const Col1,Row1,Col2,Row2: integer): TCellArea; override;
     //* Moves areas.
     //* Move all areas by the offset DeltaCol,DeltaRow. DeltaCol and DeltaRow
     //* can be negative for moving left/up.
     //* ~param DeltaCol How many columns to the right (positive) or left (negative) the areas shall be moved.
     //* ~param DeltaRow How many rows down (positive) or up (negative) the areas shall be moved.
     procedure Move(const DeltaCol,DeltaRow: integer); overload; override;
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
     procedure Move(const Col1,Row1,Col2,Row2: integer; const DeltaCol,DeltaRow: integer); overload; override;

     //* True if a Copy of Move operation failed because one of the areas was
     //* to be copied over an existing area. Equivalen to the Excel error
     //* message "Can not change part of a merged cell".
     property CopyFailed: boolean read FCopyFailed;
     end;

implementation

{ TBaseCellAreas }

function TBaseCellAreas.Add: TCellArea;
begin
  Result := CreateObject;
  inherited Add(Result);
end;

constructor TBaseCellAreas.Create;
begin
  inherited Create;
  FDelimiter := ' ';
end;

function TBaseCellAreas.CreateObject: TCellArea;
begin
  Result := TCellAreaImpl.Create;
end;

function TBaseCellAreas.FillAndClipHitList(const Col1, Row1, Col2, Row2: integer; AList: TBaseCellAreas; AObject: TObject = Nil): boolean;
var
  i: integer;
  Area: TCellArea;
  SelectArea,NewArea: TXLSCellArea;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  for i := 0 to Count - 1 do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
      Area := AList.Add(@NewArea);
      if AObject <> Nil then begin
        Area.Obj := AObject;
        Area.Tag := i;
      end;
    end;
  end;
  Result := AList.Count > 0;
end;

function TBaseCellAreas.FillHitList(const Col1, Row1, Col2, Row2: integer; AList: TBaseCellAreas): boolean;
var
  i: integer;
  Area: TCellArea;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].InArea(Col1, Row1, Col2, Row2) then begin
      Area := AList.Add;
      Area.Assign(Items[i]);
    end;
  end;
  Result := AList.Count > 0;
end;

function TBaseCellAreas.FindArea(const Col, Row: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].CellInArea(Col, Row) then
      Exit;
  end;
  Result := -1;
end;

function TBaseCellAreas.FindArea(const Col1, Row1, Col2, Row2: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Equal(Col1, Row1, Col2, Row2) then
      Exit;
  end;
  Result := -1;
end;

function TBaseCellAreas.GetDelimitedText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to Count - 1 do
    Result := Result + Items[i].AsString + ' ';
  Result := Trim(Result);
end;

function TBaseCellAreas.GetItems(const Index: integer): TCellArea;
begin
  Result := TCellArea(inherited Items[Index]);
end;

function TBaseCellAreas.Add(const C, R: integer): TCellArea;
begin
  Result := Add;
  Result.Col1 := C;
  Result.Row1 := R;
  Result.Col2 := C;
  Result.Row2 := R;
end;

procedure TBaseCellAreas.Add(const C1, R1, C2, R2: integer; AAreas: TBaseCellAreas);
var
  i: integer;
begin
  for i := 0 to AAreas.Count - 1 do begin
    if AAreas[i].InArea(C1,R1,C2,R2) then
      Add(C1,R1,C2,R2);
  end;
end;

function TBaseCellAreas.AreaInAreas(AArea: TCellArea): boolean;
begin
  Result := AreaInAreas(AArea.Col1,AArea.Row1,AArea.Col2,AArea.Row2);
end;

function TBaseCellAreas.AreaInAreas(const Col1, Row1, Col2, Row2: integer): boolean;
var
  i: integer;
  SelectArea,NewArea: TXLSCellArea;
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

procedure TBaseCellAreas.Assign(Areas: TBaseCellAreas);
var
  i: integer;
begin
  Clear;
  for i := 0 to Areas.Count - 1 do
    Add.Assign(Areas[i]);
end;

procedure TBaseCellAreas.NormalizeAll;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Normalize;
end;

function TBaseCellAreas.TotExtent: TXLSCellArea;
var
  i: integer;
begin
  Result.Row1 := MAXINT;
  Result.Row2 := 0;
  Result.Col1 := MAXINT;
  Result.Col2 := 0;
  for i := 0 to Count - 1 do begin
    Result.Col1 := Min(Result.Col1,Items[i].Col1);
    Result.Row1 := Min(Result.Row1,Items[i].Row1);
    Result.Col2 := Max(Result.Col2,Items[i].Col2);
    Result.Row2 := Max(Result.Row2,Items[i].Row2);
  end;
end;

function TBaseCellAreas.Add(const C1, R1, C2, R2: integer): TCellArea;
begin
  Result := Add;
  Result.Col1 := C1;
  Result.Row1 := R1;
  Result.Col2 := C2;
  Result.Row2 := R2;
end;

function TBaseCellAreas.Add(const ARef: AxUCString): TCellArea;
var
  p: integer;
  C1, R1, C2, R2: integer;
begin
  p := CPos(':',ARef);
  if p > 1 then
    AreaStrToColRow(ARef,C1,R1,C2,R2)
  else begin
    RefStrToColRow(ARef,C1, R1);
    C2 := C1;
    R2 := R1;
  end;
  Result := Add(C1,R1,C2,R2);
end;

function TBaseCellAreas.IntersectArea(Source1,Source2: TXLSCellArea; var Dest: TXLSCellArea): boolean;
begin
  if Source1.Col1 > Source2.Col1 then Dest.Col1 := Source1.Col1 else Dest.Col1 := Source2.Col1;
  if Source1.Row1 > Source2.Row1 then Dest.Row1 := Source1.Row1 else Dest.Row1 := Source2.Row1;
  if Source1.Col2 < Source2.Col2 then Dest.Col2 := Source1.Col2 else Dest.Col2 := Source2.Col2;
  if Source1.Row2 < Source2.Row2 then Dest.Row2 := Source1.Row2 else Dest.Row2 := Source2.Row2;
  Result := (Dest.Row1 <= Dest.Row2) and (Dest.Col1 <= Dest.Col2);
end;

function TBaseCellAreas.MaxHeight: integer;
var
  i: integer;
begin
  Result := 0;

  for i := 0 to Count - 1 do
    Result := Max(Result,Items[i].Height);
end;

function TBaseCellAreas.MaxWidth: integer;
var
  i: integer;
begin
  Result := 0;

  for i := 0 to Count - 1 do
    Result := Max(Result,Items[i].Width);
end;

function TBaseCellAreas.CellCount: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,Items[i].CellCount);
end;

function TBaseCellAreas.CellInAreas(const Col, Row: integer): integer;
begin
  for Result := 0 to Count - 1 do begin
    if (Col >= Items[Result].Col1) and (Col <= Items[Result].Col2) and (Row >= Items[Result].Row1) and (Row <= Items[Result].Row2) then
      Exit;
  end;
  Result := -1;
end;

function TBaseCellAreas.Split(const Index: integer; Area: TXLSCellArea; var SplitAreas: array of TXLSCellArea): integer;
var
  TmpR1,TmpR2: word;
begin
  Result := 0;
  TmpR1 := Items[Index].Row1;
  TmpR2 := Items[Index].Row2;

  if Area.Row1 > Items[Index].Row1 then begin
    SplitAreas[Result].Col1 := Items[Index].Col1;
    SplitAreas[Result].Row1 := Items[Index].Row1;
    SplitAreas[Result].Col2 := Items[Index].Col2;
    SplitAreas[Result].Row2 := Area.Row1 - 1;
    Inc(Result);
    TmpR1 := Area.Row1;
  end;

  if Area.Row2 < Items[Index].Row2 then begin
    SplitAreas[Result].Col1 := Items[Index].Col1;
    SplitAreas[Result].Row1 := Area.Row2 + 1;
    SplitAreas[Result].Col2 := Items[Index].Col2;
    SplitAreas[Result].Row2 := Items[Index].Row2;
    Inc(Result);
    TmpR2 := Area.Row2;
  end;

  if Area.Col1 > Items[Index].Col1 then begin
    SplitAreas[Result].Col1 := Items[Index].Col1;
    SplitAreas[Result].Row1 := TmpR1;
    SplitAreas[Result].Col2 := Area.Col1 - 1;
    SplitAreas[Result].Row2 := TmpR2;
    Inc(Result);
  end;

  if Area.Col2 < Items[Index].Col2 then begin
    SplitAreas[Result].Col1 := Area.Col2 + 1;
    SplitAreas[Result].Row1 := TmpR1;
    SplitAreas[Result].Col2 := Items[Index].Col2;
    SplitAreas[Result].Row2 := TmpR2;
    Inc(Result);
  end;
end;

procedure TBaseCellAreas.Copy(const Col1, Row1, Col2, Row2: integer; const DeltaCol, DeltaRow: integer);
begin

end;

procedure TBaseCellAreas.SetDelimitedText(const Value: AxUCString);
var
  S1,S2: AxUCString;
begin
  Clear;
  S1 := Value;
  while S1 <> '' do begin
    S2 := SplitAtChar(FDelimiter,S1);
    Add(S2);
  end;
end;

function TBaseCellAreas.Add(const AArea: PXLSCellArea): TCellArea;
begin
  Result := Add;
  Result.Col1 := AArea.Col1;
  Result.Row1 := AArea.Row1;
  Result.Col2 := AArea.Col2;
  Result.Row2 := AArea.Row2;
end;

{ TCellAreas }

procedure TCellAreas.Delete(const Col1, Row1, Col2, Row2: integer);
var
  i,j,Cnt,SplitCount: integer;
  SelectArea,NewArea: TXLSCellArea;
  SplitAreas: array[0..3] of TXLSCellArea;
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

procedure TCellAreas.DeleteAndAdjust(const Col1, Row1, Col2, Row2: integer);
begin

end;

procedure TCellAreas.Copy(const Col1, Row1, Col2, Row2: integer; const DeltaCol,DeltaRow: integer);
var
  i,Cnt: integer;
  SelectArea,NewArea: TXLSCellArea;

procedure DoCopy(C1,R1,C2,R2: integer);
begin
  Inc(C1,DeltaCol);
  Inc(R1,DeltaRow);
  Inc(C2,DeltaCol);
  Inc(R2,DeltaRow);
  if CheckAndClip(C1,R1,C2,R2) and not Combine(C1,R1,C2,R2) then begin
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

procedure TCellAreas.Move(const DeltaCol, DeltaRow: integer);
var
  i,C1,R1,C2,R2: integer;
begin
  i := 0;
  while i < Count do begin
    C1 := Items[i].Col1 + DeltaCol;
    R1 := Items[i].Row1 + DeltaRow;
    C2 := Items[i].Col2 + DeltaCol;
    R2 := Items[i].Row2 + DeltaRow;
    if not CheckAndClip(C1,R1,C2,R2) then
      Delete(i)
    else begin
      Items[i].SetSize(C1,R1,C2,R2);
      Inc(i);
    end;
  end;
end;

procedure TCellAreas.Move(const Col1, Row1, Col2, Row2: integer; const DeltaCol, DeltaRow: integer);
var
  i,j,Cnt,SplitCount: integer;
  SelectArea,NewArea: TXLSCellArea;
  SplitAreas: array[0..3] of TXLSCellArea;

procedure DoMove(C1,R1,C2,R2: integer);
begin
  Delete(i);
  Dec(Cnt);
  Inc(C1,DeltaCol);
  Inc(R1,DeltaRow);
  Inc(C2,DeltaCol);
  Inc(R2,DeltaRow);
  if (Count > 0) and CheckAndClip(C1,R1,C2,R2) and not Combine(C1,R1,C2,R2) then begin
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
function TCellAreas.CheckAndClip(var C1, R1, C2, R2: integer): boolean;
begin
  if (C1 > XLS_MAXCOL) or (R1 > XLS_MAXROW) or (C2 < 0) or (R2 < 0) then
    Result := False
  else begin
    C1 := Max(C1,0);
    R1 := Max(R1,0);
    C2 := Min(C2,XLS_MAXCOL);
    R2 := Min(R2,XLS_MAXROW);
    Result := True;
  end;
end;

function TCellAreas.Combine(const C1, R1, C2, R2: integer): boolean;
var
  i: integer;
begin
  Result := True;
  for i := 0 to Count - 1 do begin
    with Items[i] do begin
      // Entirly inside the area.
      if (C1 >= Col1) and (R1 >= Row1) and (C2 <= Col2) and (R2 <= Row2) then
        Exit;
      if (Col1 = C1) and (Col2 = C2) then begin
        if (R2 >= (Row1 - 1)) and (R2 <= Row2) then begin
          Row1 := R1;
          Exit;
        end
        else if (R1 >= Row1) and (R1 <= (Row2 + 1)) then begin
          Row2 := R2;
          Exit;
        end;
      end
      else if (Row1 = R1) and (Row2 = R2) then begin
        if (C2 >= (Col1 - 1)) and (C2 <= Col2) then begin
          Col1 := C1;
          Exit;
        end
        else if (C1 >= Col1) and (C1 <= (Col2 + 1)) then begin
          Col2 := C2;
          Exit;
        end;
      end;
    end;
  end;
  Result := False;
end;

function TCellAreas.Include(const Col1, Row1, Col2, Row2: integer): TCellArea;
var
  i,Cnt: integer;
  SelectArea,NewArea: TXLSCellArea;
begin
  Result := Nil;
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  Cnt := Count;
  i := 0;
  while i < Cnt do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then begin
      Items[i].AsRecArea := NewArea;
      Result := Items[i];
      Inc(i);
    end
    else begin
      Delete(i);
      Dec(Cnt);
    end;
  end;
end;

function TCellAreas.Last: TCellArea;
begin
  Result := Items[Count - 1];
end;

{ TCellArea }

procedure TCellArea.Assign(Source: TCellArea);
begin
  Row1 := TCellArea(Source).Row1;
  Row2 := TCellArea(Source).Row2;
  Col1 := TCellArea(Source).Col1;
  Col2 := TCellArea(Source).Col2;
  FObject := TCellArea(Source).FObject;
end;

function TCellArea.CellCount: integer;
begin
  Result := (Col2 - Col1 + 1) * (Row2 - Row1 + 1);
end;

function TCellArea.CellInArea(const Col, Row: integer): boolean;
begin
  Result := (Col >= Col1) and (Row >= Row1) and (Col <= Col2) and (Row <= Row2);
end;

function TCellArea.CornerInArea(AArea: TCellArea): boolean;
begin
  Result := CornerInArea(AArea.Col1,AArea.Row1,AArea.Col2,AArea.Row2);
end;

function TCellArea.CornerInArea(const C1, R1, C2, R2: integer): boolean;
begin
  Result := (GetCol1 >= C1) and (GetCol1 <= C2) and (GetRow1 >= R1) and (GetRow1 <= R2);
  if not Result then
    Result := (GetCol2 >= C1) and (GetCol2 <= C2) and (GetRow1 >= R1) and (GetRow1 <= R2);
    if not Result then
      Result := (GetCol1 >= C1) and (GetCol1 <= C2) and (GetRow2 >= R1) and (GetRow2 <= R2);
      if not Result then
        Result := (GetCol2 >= C1) and (GetCol2 <= C2) and (GetRow2 >= R1) and (GetRow2 <= R2);
end;

function TCellArea.Equal(const C1, R1, C2, R2: integer): boolean;
begin
  Result := (C1 = Col1) and (R1 = Row1) and (C2 = Col2) and (R2 = Row2);
end;

function TCellArea.FullInside(const Col, Row: integer): boolean;
begin
  Result := (Col > Col1) and (Row > Row1) and (Col < Col2) and (Row < Row2);
end;

function TCellArea.GetAsRecArea: TXLSCellArea;
begin
  Result.Col1 := Col1;
  Result.Row1 := Row1;
  Result.Col2 := Col2;
  Result.Row2 := Row2;
end;

function TCellArea.GetAsString: AxUCString;
begin
  if (Col1 = Col2) and (Row1 = Row2) then
    Result := ColRowToRefStr(Col1,Row1,False,False)
  else
    Result := AreaToRefStr(Col1,Row1,Col2,Row2,False,False,False,False);
end;

function TCellArea.GetAsStringAbs: AxUCString;
begin
  if (Col1 = Col2) and (Row1 = Row2) then
    Result := ColRowToRefStr(Col1,Row1,True,True)
  else
    Result := AreaToRefStr(Col1,Row1,Col2,Row2,True,True,True,True);
end;

function TCellArea.GetEdges(const Col, Row: integer): TCellEdges;
begin
  Result := [];
  if (Col >= Col1) and (Col <= Col2) and (Row >= Row1) and (Row <= Row2) then begin
    if Col = Col1 then Result := Result + [ceLeft];
    if Row = Row1 then Result := Result + [ceTop];
    if Col = Col2 then Result := Result + [ceRight];
    if Row = Row2 then Result := Result + [ceBottom];
  end;
end;

function TCellArea.Height: integer;
begin
  Result := Row2 - Row1 + 1;
end;

function TCellArea.InArea(const C1, R1, C2, R2: integer): boolean;
begin
  Result := (GetCol1 >= C1) and (GetCol2 <= C2) and (GetRow1 >= R1) and (GetRow2 <= R2);
end;

function TCellArea.IsColumn: boolean;
begin
  Result := (Row1 = 0) and (Row2 = XLS_MAXROW);
end;

function TCellArea.IsRow: boolean;
begin
  Result := (Col1 = 0) and (Col2 = XLS_MAXCOL);
end;

procedure TCellArea.Normalize;
var
  T: integer;
begin
  if Col1 > Col2 then begin
    T := Col1;
    Col1 := Col2;
    Col2 := T;
  end;

  if Row1 > Row2 then begin
    T := Row1;
    Row1 := Row2;
    Row2 := T;
  end;
end;

procedure TCellArea.SetAsRecArea(const Value: TXLSCellArea);
begin
  Col1 := Value.Col1;
  Row1 := Value.Row1;
  Col2 := Value.Col2;
  Row2 := Value.Row2;
end;

procedure TCellArea.SetSize(const C, R: integer);
begin
  Col1 := C;
  Row1 := R;
  Col2 := C;
  Row2 := R;
end;

procedure TCellArea.SetSize(const C1, R1, C2, R2: integer);
begin
  Col1 := C1;
  Row1 := R1;
  Col2 := C2;
  Row2 := R2;
end;

function TCellArea.Width: integer;
begin
  Result := Col2 - Col1 + 1;
end;

{ TSolidCellAreas }

procedure TSolidCellAreas.Copy(const Col1, Row1, Col2, Row2: integer; const DeltaCol, DeltaRow: integer);
var
  i: integer;
  Cnt: integer;
  Areas: array of TCellArea;
begin
  if (DeltaCol = 0) and (DeltaRow = 0) then
    Exit;

  FCopyFailed := False;

  SetLength(Areas,$FF);
  Cnt := 0;
  for i := 0 to Count - 1 do begin
    if Items[i].InArea(Col1,Row1,Col2,Row2) and AreaInsideSheet(Items[i].Col1 + DeltaCol,Items[i].Row1 + DeltaRow,Items[i].Col2 + DeltaCol,Items[i].Row2 + DeltaRow) then begin
      Areas[Cnt] := Items[i];
      Inc(Cnt);
      if Cnt > Length(Areas) then
        SetLength(Areas,Length(Areas) + $FF);
    end;
  end;
  SetLength(Areas,Cnt);

  for i := 0 to High(Areas) do begin
    if AreaInAreas(Areas[i].Col1 + DeltaCol,Areas[i].Row1 + DeltaRow,Areas[i].Col2 + DeltaCol,Areas[i].Row2 + DeltaRow) then begin
      FCopyFailed := True;
      Exit;
    end;
  end;

  for i := 0 to High(Areas) do
    Add(Areas[i].Col1 + DeltaCol,Areas[i].Row1 + DeltaRow,Areas[i].Col2 + DeltaCol,Areas[i].Row2 + DeltaRow);
end;

procedure TSolidCellAreas.Delete(const Col1, Row1, Col2, Row2: integer);
var
  i,SplitCount: integer;
  SelectArea,NewArea: TXLSCellArea;
  SplitAreas: array[0..3] of TXLSCellArea;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  i := 0;
  while i < Count do begin
    if Items[i].InArea(Col1, Row1, Col2, Row2) then
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

procedure TSolidCellAreas.DeleteAndAdjust(const Col1, Row1, Col2, Row2: integer);
var
  i    : integer;
  Area : TCellArea;
  R1,C1,
  R2,C2: integer;
begin
  for i := Count - 1 downto 0 do begin
    Area := Items[i];

    C1 := Area.Col1;
    R1 := Area.Row1;
    C2 := Area.Col2;
    R2 := Area.Row2;

    if AdjustDeleteRowsCols(R1,R2,Row1,Row2) then begin
      Area.Row1 := R1;
      Area.Row2 := R2;

      if AdjustDeleteRowsCols(C1,C2,Col1,Col2) then begin
        Area.Col1 := C1;
        Area.Col2 := C2;
      end
      else
        Delete(i);
    end
    else
      Delete(i);
  end;
end;

function TSolidCellAreas.Include(const Col1, Row1, Col2, Row2: integer): TCellArea;
var
  i: integer;
  List: TIntegerList;
  Area,Temp: TXLSCellArea;
  MinC,MinR,MaxC,MaxR: integer;
begin
  List := TIntegerList.Create;
  try
    Area.Col1 := Col1;
    Area.Row2 := Row2;
    Area.Col2 := Col2;
    Area.Row1 := Row1;
    MinC := Col1;
    MinR := Row1;
    MaxC := Col2;
    MaxR := Row2;
    for i := 0 to Count - 1 do begin
      if IntersectArea(Area,Items[i].GetAsRecArea,Temp) then begin
        List.Add(i);
        if Items[i].Col1 < MinC then MinC := Items[i].Col1;
        if Items[i].Row1 < MinR then MinR := Items[i].Row1;
        if Items[i].Col2 < MaxC then MaxC := Items[i].Col2;
        if Items[i].Row2 < MaxR then MaxR := Items[i].Row2;
      end;
    end;
    if List.Count > 0 then begin
      for i := List.Count - 1 downto 0 do
        Delete(List[i]);
      Result := Add(MinC,MinR,MaxC,MaxR);
    end
    else
      Result := Add(Col1,Row1,Col2,Row2);
  finally
    List.Free;
  end;
end;

procedure TSolidCellAreas.Move(const Col1, Row1, Col2, Row2: integer; const DeltaCol, DeltaRow: integer);
var
  i: integer;
  R1,C1,R2,C2: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].FullInside(Col1,Row1) or
       Items[i].FullInside(Col2,Row1) or
       Items[i].FullInside(Col1,Row2) or
       Items[i].FullInside(Col2,Row2) then
      raise Exception.Create('Can not change part of a merged cell');
  end;
  i := 0;
  while i < Count do begin
    C1 := Items[i].Col1;
    R1 := Items[i].Row1;
    C2 := Items[i].Col2;
    R2 := Items[i].Row2;
    if (C1 >= Col1) and (C1 <= Col2) then
      Inc(C1,DeltaCol);
    if (R1 >= Row1) and (R1 <= Row2) then
      Inc(R1,DeltaRow);

    if (C2 >= Col1) and (C2 <= Col2) then
      Inc(C2,DeltaCol)
    else if (C2 = (Col1 + DeltaCol)) and (C2 <= Col2) then
      Inc(C2,DeltaCol);

    if (R2 >= Row1) and (R2 <= Row2) then
      Inc(R2,DeltaRow)
    else if (R2 = (Row1 + DeltaRow)) and (R2 <= Row2) then
      Inc(R2,DeltaRow);

    if (C2 < C1) or (R2 < R1) then
      Delete(i)
    else begin
      Items[i].Col1 := C1;
      Items[i].Row1 := R1;
      Items[i].Col2 := C2;
      Items[i].Row2 := R2;
      Inc(i);
    end;
  end;
end;

procedure TSolidCellAreas.Move(const DeltaCol, DeltaRow: integer);
begin
  raise XLSRWException.Create('TSolidCellAreas.Move not implemented');
end;

{ TCellRef }

procedure TCellRef.Clear;
begin
  FCol := 0;
  FRow := 0;
end;

function TCellRef.GetAsString: AxUCString;
begin
  Result := ColRowToRefStr(FCol,FRow,False,False);
end;

procedure TCellRef.SetAsString(const Value: AxUCString);
begin
  RefStrToColRow(Value,FCol,FRow);
end;

{ TCellRefs }

procedure TCellRefs.Add(const ARef: AxUCString);
var
  R: TCellRef;
begin
  R := TCellRef.Create;
  R.AsString := ARef;
  inherited Add(R);
end;

function TCellRefs.GetItems(const Index: integer): TCellRef;
begin
  Result := TCellRef(inherited Items[Index]);
end;

{ TCellAreaImpl }

function TCellAreaImpl.GetCol1: integer;
begin
  Result := FCol1;
end;

function TCellAreaImpl.GetCol2: integer;
begin
  Result := FCol2;
end;

function TCellAreaImpl.GetRow1: integer;
begin
  Result := FRow1;
end;

function TCellAreaImpl.GetRow2: integer;
begin
  Result := FRow2;
end;

procedure TCellAreaImpl.SetCol1(const AValue: integer);
begin
  FCol1 := AValue;
end;

procedure TCellAreaImpl.SetCol2(const AValue: integer);
begin
  FCol2 := AValue;
end;

procedure TCellAreaImpl.SetRow1(const AValue: integer);
begin
  FRow1 := AValue;
end;

procedure TCellAreaImpl.SetRow2(const AValue: integer);
begin
  FRow2 := AValue;
end;

end.
