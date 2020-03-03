unit XLSBookReg2;

{-
********************************************************************************
******* XLSSpreadSheet V3.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}

{$I AxCompilers.inc}

{$R XLSBookReg2.dcr}

interface

uses { Delphi  } Classes, SysUtils,
{$ifdef DELPHI_XE3_OR_LATER}
     Actions,
{$endif}
     XLSBook2, XBookPrint2, ExcelColorPicker, XBookStdComponents;

procedure Register;

implementation

procedure Register;
begin
  RegisterComponents('XLS', [TXLSSpreadSheet,
                             TXLSBookPrint,
                             TXLSBookPrintPreview,
                             TExcelColorPicker,
                             TXBookColorPicker,
                             TXBookColorComboBox,
                             TXBookCellBorderPicker,
                             TXBookCellBorderStylePicker]);
end;

end.
