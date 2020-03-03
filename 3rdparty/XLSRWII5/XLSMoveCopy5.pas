unit XLSMoveCopy5;

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
     Xc12Common5;

type TXLSMoveCopyItem = class(TObject)
protected
     function  Intersect(Col1,Row1,Col2,Row2: integer): boolean; virtual; abstract;
     procedure Copy(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); virtual; abstract;
     procedure Delete(Col1,Row1,Col2,Row2: integer); virtual; abstract;
     procedure Include(Col1,Row1,Col2,Row2: integer); virtual; abstract;
     procedure Move(DeltaCol,DeltaRow: integer); overload; virtual; abstract;
     procedure Move(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); overload; virtual; abstract;
public
     procedure Assign(ASource: TXLSMoveCopyItem); virtual; abstract;
     end;

type TXLSMoveCopyList = class(TObjectList)
private
     function GetItems(Index: integer): TXLSMoveCopyItem;
protected
public
     function Add: TXLSMoveCopyItem; overload; virtual; abstract;

     procedure MoveLocal(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer);
     procedure DeleteLocal(Col1,Row1,Col2,Row2: integer);
     procedure CopyLocal(Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     procedure CopyList(List: TList; Col1,Row1,Col2,Row2: integer);
     procedure InsertList(List: TList; Col1,Row1,Col2,Row2,DestCol,DestRow: integer);
     procedure DeleteList(List: TList; Col1,Row1,Col2,Row2: integer);

     property Items[Index: integer]: TXLSMoveCopyItem read GetItems; default;
     end;

implementation

{ TCollectionMoveCopy }

procedure TXLSMoveCopyList.CopyList(List: TList; Col1, Row1, Col2, Row2: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do begin
    if Items[i].Intersect(Col1, Row1, Col2, Row2) then
      List.Add(Items[i]);
  end;
end;

procedure TXLSMoveCopyList.CopyLocal(Col1, Row1, Col2, Row2, DestCol, DestRow: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do
    Items[i].Copy(Col1, Row1, Col2, Row2, DestCol - Col1,DestRow - Row1);
end;

procedure TXLSMoveCopyList.DeleteList(List: TList; Col1,Row1,Col2,Row2: integer);
var
  i: integer;

function Find(Item: TXLSMoveCopyItem): boolean;
var
  i: integer;
begin
  for i := 0 to Count -1 do begin
    if Items[i] = Item then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

begin
  for i := 0 to List.Count - 1 do begin
    if Find(List[i]) then
      TXLSMoveCopyItem(List[i]).Delete(Col1,Row1,Col2,Row2);
  end;
end;

procedure TXLSMoveCopyList.DeleteLocal(Col1, Row1, Col2, Row2: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do
    Items[i].Delete(Col1,Row1,Col2,Row2);
end;

function TXLSMoveCopyList.GetItems(Index: integer): TXLSMoveCopyItem;
begin
  Result := TXLSMoveCopyItem(inherited Items[Index]);
end;

procedure TXLSMoveCopyList.InsertList(List: TList; Col1,Row1,Col2,Row2,DestCol, DestRow: integer);
var
  i: integer;
  Item: TXLSMoveCopyItem;
begin
  for i := 0 to List.Count -1 do begin
    Item := TXLSMoveCopyItem(Add);
    Item.Assign(TXLSMoveCopyItem(List[i]));
    Item.Include(Col1,Row1,Col2,Row2);
    Item.Move(DestCol - Col1,DestRow - Row1);
  end;
end;

procedure TXLSMoveCopyList.MoveLocal(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer);
var
  i: integer;
begin
  for i := 0 to Count -1 do
    Items[i].Move(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow);
end;

end.
