unit XLSExportCSV5;

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
     Xc12Utils5,
     XLSUtils5, XLSSheetData5, XLSReadWriteII5;

//* Options when exporting to CSV files.
type TExportCSVOption = (ecoQuotedStrings, //* Put the content of string cells withing quotes.
                         ecoFormatValues,  //* Apply number formatting to number cells. When not used, a number value is converted to a string with the FloatToStr function.
                         ecoSelection,     //* Export only the current selection. See ~[link SheetData4.TSheet.SelectedAreas TSheet.SelectedAreas]. If there is more than one selected area, all of them will be written to the file.
                         ecoFormulas,      //* Write the formulas to the file, instead of the values (result) of the formulas.
                         ecoWriteUTF8      //* Write the CSV file UTF8 encoded.
                         );
     TExportCSVOptions = set of TExportCSVOption;

//* This routin will export the content of a worksheet to a text (CSV) file.
//* ~param XLS The XLSReadWriteII object to export.
//* ~param SheetIndex Index of the work sheet to export.
//* ~param QuoteChar Quote character for text strings.
//* ~param SeparatorChar The character used to separate the cell values (fields) in the file. Common used characters are comma, semicolon and tab.
//* ~param Options File export options.
//* ~param Filename The name of the file that the cells shall be exported to.
procedure ExportCSV(XLS: TXLSReadWriteII5; SheetIndex: integer; SeparatorChar, QuoteChar: WideChar; Options: TExportCSVOptions; Filename: AxUCString);

implementation

procedure ExportCSV(XLS: TXLSReadWriteII5; SheetIndex: integer; SeparatorChar, QuoteChar: WideChar; Options: TExportCSVOptions; Filename: AxUCString);
const
  BOM_UTF8: array[0..2] of byte = ($EF, $BB, $BF);
var
  i: integer;
  b: byte;
  Stream: TFileStream;
  Sheet: TXLSWorksheet;

procedure ExportRange(Col1,Row1,Col2,Row2: integer);
var
  S: AxUCString;
  S2: XLS8String;
  Col,Row: integer;
begin
  for Row := Row1 to Row2 do begin
    S := '';
    for Col := Col1 to Col2 do begin
      if Sheet.Cell[Col,Row] = Nil then begin
        S := S + SeparatorChar;
      end
      else begin
        case Sheet.CellType[Col,Row] of
          xctFloat: begin
            if ecoFormatValues in Options then
              S := S + Sheet.AsFmtString[Col,Row] + SeparatorChar
            else
              S := S + Sheet.AsString[Col,Row] + SeparatorChar;
          end;
          xctString: begin
            if ecoQuotedStrings in Options then
              S := S + QuoteChar + StripCRLF(Sheet.AsString[Col,Row]) + QuoteChar + SeparatorChar
            else
              S := S + StripCRLF(Sheet.AsString[Col,Row]) + SeparatorChar;
          end;
          xctBoolean: begin
            if Sheet.AsBoolean[Col,Row] then
              S := S + G_StrTRUE + SeparatorChar
            else
              S := S + G_StrFALSE + SeparatorChar;
          end;
          xctError: begin
            if ecoQuotedStrings in Options then
              S := S + QuoteChar + Sheet.AsString[Col,Row] + QuoteChar + SeparatorChar
            else
              S := S + Sheet.AsString[Col,Row] + SeparatorChar;
          end;
          xctFloatFormula: begin
            if ecoFormulas in Options then
              S := S + Sheet.AsFormula[Col,Row] + SeparatorChar
            else begin
              if ecoFormatValues in Options then
                S := S + Sheet.AsFmtString[Col,Row] + SeparatorChar
              else
                S := S + Sheet.AsString[Col,Row] + SeparatorChar;
            end;
          end;
          xctStringFormula: begin
            if ecoFormulas in Options then
              S := S + Sheet.AsFormula[Col,Row] + SeparatorChar
            else begin
              if ecoQuotedStrings in Options then
                S := S + QuoteChar + Sheet.AsStrFormulaValue[Col,Row] + QuoteChar + SeparatorChar
              else
                S := S + Sheet.AsStrFormulaValue[Col,Row] + SeparatorChar;
            end;
          end;
          xctBooleanFormula: begin
            if ecoFormulas in Options then
              S := S + Sheet.AsFormula[Col,Row] + SeparatorChar
            else begin
              if Sheet.AsBoolFormulaValue[Col,Row] then
                S := S + G_StrTRUE + SeparatorChar
              else
                S := S + G_StrFALSE + SeparatorChar;
            end;
          end;
          xctErrorFormula: begin
            if ecoFormulas in Options then
              S := S + Sheet.AsFormula[Col,Row] + SeparatorChar
            else begin
              if ecoQuotedStrings in Options then
                S := S + QuoteChar + Sheet.AsString[Col,Row] + QuoteChar + SeparatorChar
              else
                S := S + Sheet.AsString[Col,Row] + SeparatorChar;
            end;
          end;
          else
            S := S + SeparatorChar;
        end;
      end;
    end;
    while (S <> '') and (S[Length(S)] = SeparatorChar) do
      S := Copy(S,1,Length(S) - 1);
    S := S + WideChar(#13) + WideChar(#10);
    if ecoWriteUTF8 in Options then
      S := AxUCString(XLSUTF8Encode(S));

    S2 := XLS8String(S);
    Stream.Write(Pointer(S2)^,Length(S2));
  end;
end;

begin
  Stream := TFileStream.Create(Filename,fmCreate);
  try
     if ecoWriteUTF8 in Options then begin
     
       b := $EF;
       Stream.Write(b,1);
       b := $BB;
       Stream.Write(b,1);
       b := $BF;
       Stream.Write(b,1);
     end;

    Sheet := XLS[SheetIndex];
    Sheet.CalcDimensions;
    if ecoSelection in Options then begin
      for i := 0 to Sheet.SelectedAreas.Count - 1 do
        ExportRange(Sheet.SelectedAreas[i].Col1,Sheet.SelectedAreas[i].Row1,Sheet.SelectedAreas[i].Col2,Sheet.SelectedAreas[i].Row2);
    end
    else
      ExportRange(Sheet.FirstCol,Sheet.FirstRow,Sheet.LastCol,Sheet.LastRow);
  finally
    Stream.Free;
  end;
end;

end.
