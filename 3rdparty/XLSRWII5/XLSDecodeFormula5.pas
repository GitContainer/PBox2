unit XLSDecodeFormula5;

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

uses SysUtils, Classes, Contnrs,
     XLSUtils5;

type TFormulaNameType = (ntName,ntExternName,ntExternSheet,ntCurrBook,ntCellValue);

type TGetNameEvent = function(NameType: TFormulaNameType; SheetIndex,NameIndex,Col,Row: integer): AxUCString of object;

function  DecodeFormula(Buf: Pointer; Len: integer; SheetIndex,ACol,ARow: integer; GetNameMethod: TGetNameEvent; FuncArgSep: WideChar): AxUCString;

implementation

function  DecodeFormula(Buf: Pointer; Len: integer; SheetIndex,ACol,ARow: integer; GetNameMethod: TGetNameEvent; FuncArgSep: WideChar): AxUCString;
begin
  raise Exception.Create('TODO');
end;

end.
