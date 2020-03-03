unit XLSReadWriteReg5;

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
{$R-}

{$I AxCompilers.inc}
{$I XLSRWII.inc}

{$R XLSReadWriteReg5.dcr}

interface

uses Classes,
     XLSReadWriteII5, XLSExportHTML5, XLSDbRead5, XLSGrid5;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('XLS', [TXLSReadWriteII5,TXLSExportHTML5,TXLSDbRead5,TXLSGrid]);
//  RegisterPropertyEditor(TypeInfo(TExcelColor), nil, '', TExcelColorsProperty);
//  RegisterPropertyEditor(TypeInfo(string), TSheetPicture, 'PictureName', TPictureNameProperty);
//  RegisterPropertyEditor(TypeInfo(string), TXLSPicture, 'Filename', TFilenameProperty);
end;

end.
