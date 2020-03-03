unit BIFF_MergedCells5;

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

uses Classes, SysUtils,
     BIFF_RecsII5, BIFF_Utils5, BIFF_CellAreas5,
     XLSUtils5;

type
//* An area that makes up a merged cell.
//* The cell in the upper left corner is the "master" cell, and it is the value
//* in that cell that is displayed in the merged area. It is also the formatting
//* of this cell that is used for the merged cells, with the exception of the
//* cell borders. If you want a border around the cells, each cell has to be
//* formatted individually, or use the ~[link SheetData4.TSheet.Range TSheet.Range] property.  
    TMergedCell = class(TCellArea97)
private
protected
     //* ~exclude
//     function GetDisplayName: string; override;
public
     end;

type
//* List of TMergedCell objects.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//* // Create a merged cell area. C4 is the master cell.
//* XLS.Sheets[0].AsString[3,3] := 'Hello,world';
//* XLS.Sheets[0].Cell[3,3].HorizAlignment := chaCenter;
//* XLS.Sheets[0].Cell[3,3].VertAlignment := cvaCenter;
//* XLS.Sheets[0].Cell[3,3].FillPatternForeColor := xcYellow;
//* Set the area of the merged cells.
//* with XLS.Sheets[0].MergedCells.Add do begin
//*   Col1 := 3;
//*   Row1 := 3;
//*   Col2 := 6;
//*   Row2 := 8;
//* end;
//* ]
     TMergedCells = class(TSolidCellAreas97)
private
protected
public
     //* ~exclude
     procedure CopyList(List: TList; Col1,Row1,Col2,Row2: integer);
     //* ~exclude
     procedure InsertList(List: TList; Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     //* ~exclude
     procedure _DeleteList(List: TList; Col1,Row1,Col2,Row2: integer);
     //* ~exclude
     procedure Copy(Col1,Row1,Col2,Row2: word; DeltaCol,DeltaRow: integer); override;
     end;

implementation

{ TMergedCell }

{ TMergedCells }

procedure TMergedCells.Copy(Col1, Row1, Col2, Row2: word; DeltaCol, DeltaRow: integer);
var
  List: TList;
begin
  List := TList.Create;
  try
    CopyList(List,Col1, Row1, Col2, Row2);
    InsertList(List,Col1, Row1, Col2, Row2,Col1 + DeltaCol,Row1 + DeltaRow);
  finally
    List.Free;
  end;
end;

procedure TMergedCells.CopyList(List: TList; Col1, Row1, Col2, Row2: integer);
var
  i: integer;
  SelectArea,NewArea: TRecCellAreaI;
begin
  SelectArea.Col1 := Col1;
  SelectArea.Row2 := Row2;
  SelectArea.Col2 := Col2;
  SelectArea.Row1 := Row1;
  for i := 0 to Count - 1 do begin
    if IntersectArea(Items[i].AsRecArea,SelectArea,NewArea) then
      List.Add(Items[i]);
  end;
end;

procedure TMergedCells._DeleteList(List: TList; Col1, Row1, Col2, Row2: integer);
var
  i: integer;

function Find(Item: TCellArea97): boolean;
var
  i: integer;
begin
  for i := 0 to Count -1 do begin
    if Items[i].Id = Item.Id then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

begin
  for i := 0 to List.Count - 1 do begin
    if Find(List[i]) then
      Delete(Col1,Row1,Col2,Row2);
  end;
end;

procedure TMergedCells.InsertList(List: TList; Col1, Row1, Col2, Row2, DestCol, DestRow: integer);
var
  i,C1,R1,C2,R2: integer;
  Item: TMergedCell;
begin
  for i := 0 to List.Count -1 do begin
    C1 := TMergedCell(List[i]).Col1 + DestCol - Col1;
    R1 := TMergedCell(List[i]).Row1 + DestRow - Row1;
    C2 := TMergedCell(List[i]).Col2 + DestCol - Col1;
    R2 := TMergedCell(List[i]).Row2 + DestRow - Row1;
    if ClipAreaToSheet(C1,R1,C2,R2) then begin
      Item := TMergedCell(Add);
      Item.Assign(TMergedCell(List[i]));
      Item.Col1 := C1;
      Item.Row1 := R1;
      Item.Col2 := C2;
      Item.Row2 := R2;
    end;
  end;
end;

end.
