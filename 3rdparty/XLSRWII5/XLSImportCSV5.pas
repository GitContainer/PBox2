unit XLSImportCSV5;

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
     XLSUtils5, XLSReadWriteII5;

//* This routine will import a text (CSV) file into a worksheet. Any existing cells on the worksheet will be overwritten.
//* Correct value for decimal separator is detected automatically. See alse ~[link ImportCSVAuto]
//* ~param XLS The target XLSReadWriteII object.
//* ~param SheetIndex The index to the worksheet where the file will be imported to.
//* ~param Col Left column where the imported data will be written.
//* ~param Row Top row where the imported data will be written.
//* ~param Filename The name of the text file that shall be imported.
//* ~param SepChar Separator character for the fields (cells) in the source file.
//* ~param HasQuoteChar Set this to True if strings in the source file are quoted.
//* ~result The number of cells that where imported.
function ImportCSV(XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; Filename: AxUCString; SepChar: AxUCChar; HasQuoteChar: boolean): boolean; overload;
function ImportCSV(XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; List: TStrings; SepChar: AxUCChar; HasQuoteChar: boolean): boolean; overload;
//* This routine will import a text (CSV) file into a worksheet. Any existing cells on the worksheet will be overwritten.
//* The routine will try to detect correct values for separator character, quote character and decimal separator. See alse ~[link ImportCSV]
//* ~param XLS The target XLSReadWriteII object.
//* ~param SheetIndex The index to the worksheet where the file will be imported to.
//* ~param Col Left column where the imported data will be written.
//* ~param Row Top row where the imported data will be written.
//* ~param Filename The name of the text file that shall be imported.
//* ~result The number of cells that where imported.
function ImportCSVAuto(XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; Filename: AxUCString): boolean;

implementation

var SepChars: array[0..3] of AxUCChar = (',',';',':',#9);

function CPosSkipQuote(C: AXUCChar; S: AXUCString): integer;
var
  InQuote: boolean;
begin
  InQuote := False;
  for Result := 1 to Length(S) do begin
    if CharInSet(S[Result],['"','''']) then
      InQuote := not InQuote;
    if not InQuote and (S[Result] = C) then
      Exit;
  end;
  Result := -1;
end;

function StrTokenSkipQuote(var S: AXUCString; Token: AXUCChar): AXUCString;
var
  p: integer;
begin
  p := CPosSkipQuote(Token,S);
  if p >= 1 then begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

function StrToken(var S: AXUCString; Token: AXUCChar): AXUCString;
var
  p: integer;
begin
  p := CPos(Token,S);
  if p >= 1 then begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

function ImportCSVInt(Lines: TStrings; XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; SepChar: AxUCChar; HasQuoteChar: boolean): integer;
var
  i,C,L   : integer;
  V       : double;
  D       : TDateTime;
  S,Token : AXUCString;
  TempDS  : AXUCChar;
begin
  Result := 0;

  TempDS := FormatSettings.DecimalSeparator;
  try
    for i := 0 to Lines.Count - 1 do begin
      S := Lines[i];
      C := 0;
      while S <> '' do begin
        if HasQuoteChar then
          Token := Trim(StrTokenSkipQuote(S,SepChar))
        else
          Token := Trim(StrToken(S,SepChar));
        if Token <> '' then begin
          Inc(Result);
          L := Length(Token);
          if HasQuoteChar and (((Token[1] = '"') and (Token[L] = '"')) or ((Token[1] = '''') and (Token[L] = ''''))) then begin
            Token := Copy(Token,2,L - 2);
            XLS[SheetIndex].AsString[Col + C,Row + i] := Token;
          end
          else begin
            if Uppercase(Token) = G_StrTRUE then
              XLS[SheetIndex].AsBoolean[Col + C,Row + i] := True
            else if Uppercase(Token) = G_StrFALSE then
              XLS[SheetIndex].AsBoolean[Col + C,Row + i] := False
            else begin
              if TryStrToFloat(Token,V) then
                XLS[SheetIndex].AsFloat[Col + C,Row + i] := V
              else begin
                if FormatSettings.DecimalSeparator = '.' then
                  FormatSettings.DecimalSeparator := ','
                else
                  FormatSettings.DecimalSeparator := '.';
                if TryStrToFloat(Token,V) then
                  XLS[SheetIndex].AsFloat[Col + C,Row + i] := V
                else begin
                  if CharInSet(Token[1],['0'..'9']) and TryStrToDateTime(Token,D) then
                    XLS[SheetIndex].AsDateTime[Col + C,Row + i] := D
                  else
                    XLS[SheetIndex].AsString[Col + C,Row + i] := Token;
                end;
              end;
            end;
          end;
        end;
        Inc(C);
      end;
    end;
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
  XLS[SheetIndex].CalcDimensions;
end;

function ImportCSV(XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; List: TStrings; SepChar: AxUCChar; HasQuoteChar: boolean): boolean; overload;
begin
  Result := ImportCSVInt(List,XLS,SheetIndex,Col,Row,SepChar,HasQuoteChar) > 0;
end;

function ImportCSV(XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; Filename: AxUCString; SepChar: AxUCChar; HasQuoteChar: boolean): boolean;
var
  Lines: TStringList;
begin
  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(Filename);
    Result := ImportCSVInt(Lines,XLS,SheetIndex,Col,Row,SepChar,HasQuoteChar) > 0;
  finally
    Lines.Free;
  end;
end;

function ImportCSVAuto(XLS: TXLSReadWriteII5; SheetIndex,Col,Row: integer; Filename: AxUCString): boolean;
var
  Lines: TStringList;
  SepCharCount: array[0..High(SepChars)] of integer;
  SepCharIndex: integer;

function FindSepChar(Index: integer): boolean;
var
  S: AXUCString;
  i,j,MaxCnt: integer;
  InQuote: boolean;
begin
  S := Lines[Index];
  InQuote := False;
  for i := 0 to High(SepChars) do
    SepCharCount[i] := 0;
  for i := 1 to Length(S) do begin
    if CharInSet(S[i],['"','''']) then
      InQuote := not InQuote;
    if not InQuote then begin
      for j := 0 to High(SepChars) do begin
        if S[i] = SepChars[j] then
          Inc(SepCharCount[j]);
      end;
    end;
  end;
  SepCharIndex := -1;
  MaxCnt := 0;
  for i := 0 to High(SepCharCount) do begin
    if (SepCharCount[i] > 2) and (SepCharCount[i] > MaxCnt) then begin
      MaxCnt := SepCharCount[i];
      SepCharIndex := i;
    end;
  end;
  Result := SepCharIndex >= 0;
end;

begin
//  Result := False;
  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(Filename);
//    if Lines.Count < 3 then
//      Exit;
{
    if DecimalSeparator = ',' then
      SepChars[0] := #$FF;
}
    if Lines.Count > 1 then
      Result := FindSepChar(1)
    else
      Result := FindSepChar(0);
    if not Result then
      Exit;
    Result := ImportCSVInt(Lines,XLS,SheetIndex,Col,Row,SepChars[SepCharIndex],True) > 0;
  finally
    Lines.Free;
  end;
end;

end.
