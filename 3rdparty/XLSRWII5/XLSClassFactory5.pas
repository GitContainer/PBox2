unit XLSClassFactory5;

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

type TXLSClassFactoryType = (xcftNames,xcftNamesMember,
                             xcftMergedCells,xcftMergedCellsMember,
                             xcftHyperlinks,xcftHyperlinksMember,
                             xcftDataValidations,xcftDataValidationsMember,
                             xcftConditionalFormat,xcftConditionalFormats,
                             xcftAutofilter,xcftAutofilterColumn,
                             xcftDrawing,xcftVirtualCells);

type TXLSClassFactory = class(TObject)
protected
public
     function CreateAClass(AClassType: TXLSClassFactoryType; AOwner: TObject = Nil): TObject; virtual; abstract;
//     function GetAClass(AClassType: TXLSClassFactoryType): TObject; virtual; abstract;
     end;

implementation

end.
