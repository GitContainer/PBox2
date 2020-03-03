unit BIFF_MoveCopy5;

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

uses Classes, SysUtils;

type TCollectionItemMoveCopy = class(TCollectionItem)
protected
     function  Intersect(Col1,Row1,Col2,Row2: integer): boolean; virtual; abstract;
     procedure Copy(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); virtual; abstract;
     procedure Delete(Col1,Row1,Col2,Row2: integer); virtual; abstract;
     procedure Include(Col1,Row1,Col2,Row2: integer); virtual; abstract;
     procedure Move(DeltaCol,DeltaRow: integer); overload; virtual; abstract;
     procedure Move(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); overload; virtual; abstract;
public
     end;

type TCollectionMoveCopy = class(TCollection)
private
     function GetItems(Index: integer): TCollectionItemMoveCopy;
protected

public
     procedure MoveLocal(Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     procedure DeleteLocal(Col1,Row1,Col2,Row2: integer);
     procedure CopyLocal(Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     procedure CopyList(List: TList; Col1,Row1,Col2,Row2: integer);
     procedure InsertList(List: TList; Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     procedure DeleteList(List: TList; Col1,Row1,Col2,Row2: integer);

     property Items[Index: integer]: TCollectionItemMoveCopy read GetItems; default;
     end;

implementation

{ TCollectionMoveCopy }

procedure TCollectionMoveCopy.CopyList(List: TList; Col1, Row1, Col2, Row2: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do begin
    if Items[i].Intersect(Col1, Row1, Col2, Row2) then
      List.Add(Items[i]);
  end;
end;

procedure TCollectionMoveCopy.CopyLocal(Col1, Row1, Col2, Row2, DestCol, DestRow: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do
    Items[i].Copy(Col1, Row1, Col2, Row2, DestCol - Col1,DestRow - Row1);
end;

procedure TCollectionMoveCopy.DeleteList(List: TList; Col1,Row1,Col2,Row2: integer);
var
  i: integer;

function Find(Item: TCollectionItem): boolean;
var
  i: integer;
begin
  for i := 0 to Count -1 do begin
    if Items[i].Id = Item.ID then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

begin
  for i := 0 to List.Count - 1 do begin
    if Find(List[i]) then
      TCollectionItemMoveCopy(List[i]).Delete(Col1,Row1,Col2,Row2);
  end;
end;

procedure TCollectionMoveCopy.DeleteLocal(Col1, Row1, Col2, Row2: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do
    Items[i].Delete(Col1,Row1,Col2,Row2);
end;

function TCollectionMoveCopy.GetItems(Index: integer): TCollectionItemMoveCopy;
begin
  Result := TCollectionItemMoveCopy(inherited Items[Index]);
end;

procedure TCollectionMoveCopy.InsertList(List: TList; Col1,Row1,Col2,Row2,DestCol, DestRow: integer);
var
  i: integer;
  Item: TCollectionItemMoveCopy;
begin
  for i := 0 to List.Count -1 do begin
    Item := TCollectionItemMoveCopy(Add);
    Item.Assign(TCollectionItemMoveCopy(List[i]));
    Item.Include(Col1,Row1,Col2,Row2);
    Item.Move(DestCol - Col1,DestRow - Row1);
  end;
end;

procedure TCollectionMoveCopy.MoveLocal(Col1, Row1, Col2, Row2, DestCol, DestRow: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do
    Items[i].Move(Col1,Row1,Col2,Row2,DestCol - Col1,DestRow - Row1);
end;

end.
