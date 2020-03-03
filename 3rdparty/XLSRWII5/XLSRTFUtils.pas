unit XLSRTFUtils;

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
     Xc12Utils5, Xc12Common5, Xc12DataStyleSheet5, Xc12DataWorksheet5, Xc12Manager5,
     XLSUtils5, XLSCellMMU5, XLSCmdFormat5, XLSColumn5, XLSRow5, XLSAXWEditor;

type TXLSDocConverter = class(TObject)
protected
     FManager  : TXc12Manager;
     FSheetObj : TIndexObject;
     FSheet    : TXc12DataWorksheet;
     FCols     : TXLSColumns;
     FRows     : TXLSRows;
     FDoc      : TAXWLogDocEditor;
     FCmdFormat: TXLSCmdFormat;

     procedure CHPXToFont(ACHPX: TAXWCHPX; AFont: TXc12Font);
     procedure ParaToDynFontRuns(APara: TAXWLogPara; out ARuns: TXc12DynFontRunArray);
     procedure ParaToWorksheet(APara: TAXWLogPara; var ACol,ARow: integer; AInTable: boolean);
     procedure TableToWorksheet(ATable: TAXWTable; var ACol,ARow: integer);
     procedure ParasToWorksheet(AParas: TAXWLogParas; var ACol,ARow: integer; AInTable: boolean);
public
     constructor Create(AManager: TXc12Manager; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; ADoc: TAXWLogDocEditor);
     destructor Destroy; override;

     procedure ToDynFontRuns(out ARuns: TXc12DynFontRunArray);
     procedure ToDoc(AText: AxUCString; ARuns: PXc12FontRunArray; ACount: integer);
     procedure ToWorksheet(var ACol,ARow: integer);
     end;

procedure AXWDocToDynFontRuns(AManager: TXc12Manager; ADoc: TAXWLogDocEditor; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; out ARuns: TXc12DynFontRunArray);
procedure FontRunsToAXWDoc(AManager: TXc12Manager; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; AText: AxUCString; ARuns: PXc12FontRunArray; ACount: integer; ADoc: TAXWLogDocEditor);
procedure AXWDocToWorksheet(AManager: TXc12Manager; ADoc: TAXWLogDocEditor; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; var ACol,ARow: integer);

implementation

procedure AXWDocToDynFontRuns(AManager: TXc12Manager; ADoc: TAXWLogDocEditor; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; out ARuns: TXc12DynFontRunArray);
var
  Cvt: TXLSDocConverter;
begin
  Cvt := TXLSDocConverter.Create(AManager,ASheet,ACols,ARows,ADoc);
  try
    Cvt.ToDynFontRuns(ARuns);
  finally
    Cvt.Free;
  end;
end;

procedure FontRunsToAXWDoc(AManager: TXc12Manager; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; AText: AxUCString; ARuns: PXc12FontRunArray; ACount: integer; ADoc: TAXWLogDocEditor);
var
  Cvt: TXLSDocConverter;
begin
  Cvt := TXLSDocConverter.Create(AManager,ASheet,ACols,ARows,ADoc);
  try
    Cvt.ToDoc(AText,ARuns,ACount);
  finally
    Cvt.Free;
  end;
end;

procedure AXWDocToWorksheet(AManager: TXc12Manager; ADoc: TAXWLogDocEditor; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; var ACol,ARow: integer);
var
  Cvt: TXLSDocConverter;
begin
  Cvt := TXLSDocConverter.Create(AManager,ASheet,ACols,ARows,ADoc);
  try
    Cvt.ToWorksheet(ACol,ARow);
  finally
    Cvt.Free;
  end;
end;

{ TXLSDocConverter }

procedure TXLSDocConverter.CHPXToFont(ACHPX: TAXWCHPX; AFont: TXc12Font);
begin
  AFont.Clear;

  AFont.Name := ACHPX.FontName;
  AFont.Size := ACHPX.Size;
  if ACHPX.Bold then
    AFont.Style := AFont.Style + [xfsBold];
  if ACHPX.Italic then
    AFont.Style := AFont.Style + [xfsItalic];
  if ACHPX.Underline <> axcuNone then begin
    if ACHPX.Underline = axcuDouble then
      AFont.Underline := xulDouble
    else
      AFont.Underline := xulSingle;
  end;

  AFont.Color := RGBColorToXc12(ACHPX.Color);
end;

constructor TXLSDocConverter.Create(AManager: TXc12Manager; ASheet: TIndexObject; ACols: TXLSColumns; ARows: TXLSRows; ADoc: TAXWLogDocEditor);
begin
  FDoc := ADoc;
  FManager := AManager;

  FSheetObj := ASheet;
  FSheet := TXc12DataWorksheet(FSheetObj.RequestObject(TXc12DataWorksheet));
  FCols := ACols;
  FRows := ARows;

  FCmdFormat := TXLSCmdFormat.Create(FManager);
  FCmdFormat.SetRowHeight := True;
end;

destructor TXLSDocConverter.Destroy;
begin
  FCmdFormat.Free;

  inherited;
end;

procedure TXLSDocConverter.ParasToWorksheet(AParas: TAXWLogParas; var ACol, ARow: integer; AInTable: boolean);
var
  i : integer;
begin
  for i := 0 to AParas.Count - 1 do begin
    case AParas.Items[i].Type_ of
      alptPara : ParaToWorksheet(AParas[i],ACol,ARow,AInTable);
      alptTable: begin
        // Ignore nested tables.
        if not AInTable then
          TableToWorksheet(AParas.Tables[i],ACol,ARow);
      end;
    end;
  end;
  if AInTable then
    Inc(ARow);
end;

procedure TXLSDocConverter.ParaToDynFontRuns(APara: TAXWLogPara; out ARuns: TXc12DynFontRunArray);
var
  i  : integer;
  CP : integer;
  CR : TAXWCharRun;
  FR : TXc12FontRun;
  Fnt: TXc12Font;
begin
  CP := 0;

  Fnt := TXc12Font.Create(Nil);
  try
    for i := 0 to APara.Runs.Count - 1 do begin
      Fnt.Clear;

      CR := APara.Runs[i];
      if CR.CHPX <> Nil then begin
        SetLength(ARuns,Length(ARuns) + 1);
        FR.Index := CP;

        CHPXToFont(CR.CHPX,Fnt);

        FR.Font := FManager.StyleSheet.Fonts.Find(Fnt);
        if FR.Font = Nil then begin
          FR.Font := FManager.StyleSheet.Fonts.Add;
          FR.Font.Assign(Fnt);
        end;

        FManager.StyleSheet.XFEditor.UseFont(FR.Font);

        ARuns[High(ARuns)] := FR;
      end;
      Inc(CP,Length(CR.Text));
    end;
  finally
    Fnt.Free;
  end;
end;

procedure TXLSDocConverter.ParaToWorksheet(APara: TAXWLogPara; var ACol, ARow: integer; AInTable: boolean);
var
  i   : integer;
  S   : AxUCString;
  Fnt : TXc12Font;
  FR  : TXc12DynFontRunArray;
  Cell: TXLSCellItem;
  vD  : double;
  vDT : TDateTime;
begin
  if FSheet.Cells.FindCell(ACol,ARow,Cell) then begin
    S := FSheet.Cells.GetString(@Cell) + #13#10;
    FSheet.Cells.StoreString(ACol,ARow,0,S + APara.PlainText);
    FCmdFormat.BeginEdit(FSheetObj);
    FCmdFormat.Alignment.WrapText := True;
    FCmdFormat.Apply(ACol,ARow);
  end
  else begin
    Fnt := TXc12Font.Create(Nil);
    try
      if APara.Runs.Count = 1 then begin
        S := Uppercase(APara.PlainText);

        if TryStrToFloat(S,vD) then
          FSheet.Cells.StoreFloat(ACol,ARow,0,vD)
        else if TryStrToDateTime(S,vDT) then begin
          if TryStrToTime(S,vDT) then
            i := XLS_NUMFMT_STD_TIME
          else if TryStrToDate(S,vDT) then
            i := XLS_NUMFMT_STD_DATE
          else
            i := XLS_NUMFMT_STD_DATETIME;

          FSheet.Cells.StoreFloat(ACol,ARow,0,vDT);

          FCmdFormat.BeginEdit(FSheetObj);
          FCmdFormat.Number.Format := ExcelStandardNumFormats[i];
          FCmdFormat.Apply(ACol,ARow);
        end
        else if (S = 'TRUE') or (S = 'FALSE') then
          FSheet.Cells.StoreBoolean(ACol,ARow,0,S = 'TRUE')
        else
          FSheet.Cells.StoreString(ACol,ARow,0,APara.PlainText);

        if APara.Runs[0].CHPX <> Nil then begin
          CHPXToFont(APara.Runs[0].CHPX,Fnt);

          FCmdFormat.BeginEdit(FSheetObj);
          FCmdFormat.Font.Assign(Fnt);
          FCmdFormat.Apply(ACol,ARow);
        end;
      end
      else if APara.Runs.Count > 1 then begin
        ParaToDynFontRuns(APara,FR);

        i := FManager.SST.AddFormattedString(APara.PlainText,FR);
        FSheet.Cells.StoreString(ACol,ARow,0,i);
      end;

      if not AInTable then
        Inc(ARow);
    finally
      Fnt.Free;
    end;
  end;
end;

procedure TXLSDocConverter.TableToWorksheet(ATable: TAXWTable; var ACol, ARow: integer);
var
  r,c  : integer;
  CC,RR: integer;
  Row  : TAXWTableRow;
  Cell : TAXWTableCell;
begin
  for r := 0 to ATable.Count - 1 do begin
    Row := ATable[r];

    for c := 0 to Row.Count - 1 do begin
      Cell := Row[c];

      CC := ACol + c;
      RR := ARow;
      if Cell.Width > FCols[CC].WidthPt then
        FCols[CC].WidthPt := Cell.Width;

      ParasToWorksheet(Cell.Paras, CC, RR,True);
    end;

    Inc(ARow);
  end;
end;

procedure TXLSDocConverter.ToDoc(AText: AxUCString; ARuns: PXc12FontRunArray; ACount: integer);
var
  i   : integer;
  n   : integer;
  S   : AxUCString;
  CR  : TAXWCharRun;
  Para: TAXWLogPara;
  Font: TXc12Font;
begin
  Para := FDoc.Paras._Add;

  for i := 1 to ACount - 1 do begin
    Font := ARuns[i - 1].Font;

    CR := Para.Runs.Add;

    CR.AddCHPX;

    CR.CHPX.FontName := Font.Name;
    CR.CHPX.Size := Font.Size;
    CR.CHPX.Bold := xfsBold in Font.Style;
    CR.CHPX.Italic := xfsItalic in Font.Style;
    CR.CHPX.Color := Xc12ColorToRGB(Font.Color);

    if Font.Underline in [xulDouble,xulDoubleAccount] then
      CR.CHPX.Underline := axcuDouble
    else if Font.Underline <> xulNone then
      CR.CHPX.Underline := axcuSingle;

    n := ARuns[i].Index - ARuns[i - 1].Index;
    S := Copy(AText,ARuns[i - 1].Index + 1,n);

    CR.Text := S;
  end;
end;

procedure TXLSDocConverter.ToDynFontRuns(out ARuns: TXc12DynFontRunArray);
var
  i  : integer;
  Fnt: TXc12Font;
begin
  if FDoc.Paras.Count < 1 then
    Exit;

  Fnt := TXc12Font.Create(Nil);
  try
    for i := 0 to FDoc.Paras[0].Runs.Count - 1 do begin
      Fnt.Clear;

      ParaToDynFontRuns(FDoc.Paras[0],ARuns);
    end;
  finally
    Fnt.Free;
  end;
end;

procedure TXLSDocConverter.ToWorksheet(var ACol, ARow: integer);
begin
  ParasToWorksheet(FDoc.Paras,ACol,ARow,False);
end;

end.

