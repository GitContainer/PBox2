unit XLSValidate5;

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
     Xc12DataWorksheet5,
     XLSUtils5, XLSCellAreas5, XLSMoveCopy5;

type
    //* Validation options.
    TXLSValidationOption = (voAllowEmptyCells, //* Allow empty cells in the validation areas.
                            voSupressDropDown, //* When ValidationType is vtList, prevent the combo box to be shown.
                            voShowPromptBox,   //* Show a hint text to the user.
                            voShowErrorBox     //* Show an error message when an invalid value is entered.
                            );
    //* Set of validation options.
    TXLSValidationOptions = set of TXLSValidationOption;

type TXLSDataValidation = class(TXc12DataValidation)
private
     function  GetSqref: TCellAreas;
protected
     FData: TXc12DataValidation;

     function  Intersect(Col1,Row1,Col2,Row2: integer): boolean; override;
     procedure Copy(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); override;
     procedure Delete(Col1,Row1,Col2,Row2: integer); override;
     procedure Include(Col1,Row1,Col2,Row2: integer); override;
     procedure Move(DeltaCol,DeltaRow: integer); override;
     procedure Move(Col1,Row1,Col2,Row2,DeltaCol,DeltaRow: integer); override;
public
     property Areas: TCellAreas read GetSqref;
     end;

type TXLSDataValidations = class(TXc12DataValidations)
private
     function GetItems(Index: integer): TXLSDataValidation;
protected
public
     destructor Destroy; override;

     function  Add: TXLSDataValidation; reintroduce;

     property Items[Index: integer]: TXLSDataValidation read GetItems;
     end;

implementation

{ TXLSDataValidationList }

function TXLSDataValidations.Add: TXLSDataValidation;
begin
  Result := TXLSDataValidation(CreateMember);
  inherited Add(Result);
end;

destructor TXLSDataValidations.Destroy;
begin
  inherited;
end;

function TXLSDataValidations.GetItems(Index: integer): TXLSDataValidation;
begin
  Result := TXLSDataValidation(inherited Items[Index]);
end;

{ TXLSDataValidation }

procedure TXLSDataValidation.Copy(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin
  Areas.Copy(Col1, Row1, Col2, Row2,DeltaCol,DeltaRow);
end;

procedure TXLSDataValidation.Delete(Col1, Row1, Col2, Row2: integer);
begin
  Areas.Delete(Col1, Row1, Col2, Row2);
end;

function TXLSDataValidation.GetSqref: TCellAreas;
begin
  Result := Sqref;
end;

procedure TXLSDataValidation.Include(Col1, Row1, Col2, Row2: integer);
begin
  inherited;

end;

function TXLSDataValidation.Intersect(Col1, Row1, Col2, Row2: integer): boolean;
begin
  Result := Areas.AreaInAreas(Col1, Row1, Col2, Row2);
end;

procedure TXLSDataValidation.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow: integer);
begin
  Areas.Move(Col1, Row1, Col2, Row2, DeltaCol, DeltaRow);
end;

procedure TXLSDataValidation.Move(DeltaCol, DeltaRow: integer);
begin
  Areas.Move(DeltaCol, DeltaRow);
end;

end.
