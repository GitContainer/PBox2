unit XLSMergedCells5;

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
     Xc12Utils5, Xc12Common5, Xc12DataWorksheet5,
     XLSUtils5, XLSCellAreas5, XLSClassFactory5;

type TXLSMergedCell = class(TXc12MergedCell)
protected
     // Used by XLSSpreadSheet.
     FPainted: boolean;
public
     property Painted: boolean read FPainted write FPainted;
     end;

type TXLSMergedCells = class(TXc12MergedCells)
private
     function GetItems(Index: integer): TXLSMergedCell;
protected
public
     destructor Destroy; override;

     function  Add: TXLSMergedCell; overload;
     function  Add(const ARef: TXLSCellArea): TXLSMergedCell; overload;

     procedure ClearPainted;

     property Items[Index: integer]: TXLSMergedCell read GetItems; default;
     end;

implementation

{ TXLSMergedCellList }

function TXLSMergedCells.Add: TXLSMergedCell;
begin
  Result := TXLSMergedCell(CreateMember);
  inherited Add(Result);
end;

function TXLSMergedCells.Add(const ARef: TXLSCellArea): TXLSMergedCell;
begin
  Result := Add;
  Result.SetSize(ARef.Col1,ARef.Row1,ARef.Col2,ARef.Row2);
end;

procedure TXLSMergedCells.ClearPainted;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Painted := False;
end;

destructor TXLSMergedCells.Destroy;
begin
  inherited;
end;

function TXLSMergedCells.GetItems(Index: integer): TXLSMergedCell;
begin
  Result := TXLSMergedCell(inherited Items[Index]);
end;

end.
