unit XLSReadXLSX5;

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

uses Classes, SysUtils, Math, Contnrs,
{$ifdef MSWINDOWS}
     Windows,
{$endif}
     Xc12Utils5, Xc12Common5, Xc12DataStyleSheet5, Xc12DataWorkbook5,
     Xc12DataWorksheet5, Xc12DataComments5, Xc12DataXLinks5, Xc12DataAutofilter5,
     Xc12DataTable5, Xc12Manager5, Xc12FileData5,
     xpgParserXLSX, xpgParseXLinks, xpgParseDrawing, xpgParserPivot,
     xpgPXML, xpgPUtils, xpgPXMLUtils, xpgPSimpleDOM,
     XLSUtils5, XLSReadWriteOPC5, XLSMMU5, XLSCellMMU5, XLSFormulaTypes5,
     XLSFormula5, XLSCellAreas5, XLSPivotTables5;

type TXLSSharedFormula = class(TObject)
protected
     FRef: TXLSCellRef;
     FApplyArea: TXLSCellArea;
     FPtgs: PXLSPtgs;
     FPtgsSize: integer;
public
     destructor Destroy; override;

     procedure AllocPtgs(const ASize: integer);
     function  CopyPtgs(AIndex: integer): PXLSPtgs;

     property Ref: TXLSCellRef read FRef write FRef;
     property ApplyArea: TXLSCellArea read FApplyArea write FApplyArea;
     property Ptgs: PXLSPtgs read FPtgs write FPtgs;
     property PtgsSize: integer read FPtgsSize write FPtgsSize;
     end;

type TXLSSharedFormulaList = class(TObjectList)
private
     function GetItems(Index: integer): TXLSSharedFormula;
protected
public
     constructor Create;

     procedure Insert(AIndex: integer; ARef: PXLSCellRef; AApplyArea: PXLSCellArea; APtgs: PXLSPtgs; APtgsSize: integer);

     property Items[Index: integer]: TXLSSharedFormula read GetItems; default;
     end;

type TReadSheetThread = class(TThread)
protected
     FStream: TStream;
     FSheet: TXc12DataWorksheet;

     procedure Execute; override;
     procedure ReadSheet;
     procedure OnCell(ASender: TObject);
     procedure OnRow(ASender: TObject);
public
     constructor Create(AStream: TStream; ASheet: TXc12DataWorksheet);
     destructor Destroy; override;
     end;

type TXLSReadXLSX = class(TObject)
protected
     FManager       : TXc12Manager;
     FFormulas      : TXLSFormulaHandler;
     FSharedFormulas: TXLSSharedFormulaList;
     FArrayFormulas : TCellAreas;

     FCurrSheet     : TXc12DataWorksheet;
     // Statistics
     FTotCells      : integer;
     FTotRows       : integer;

     function  GetEnum(AValue, ADefault: integer): integer;

     function  GetColor(AColor: xpgParserXLSX.TCT_Color): TXc12Color;
     procedure GetBorder(ABorder: TCT_BorderPr; ADest: TXc12BorderPr);
     function  GetXfApply(AXf: TCT_Xf): TXc12ApplyFormats;
     function  GetProtections(AProtection: TCT_CellProtection): TXc12CellProtections;

     procedure GetFont(ASrcFont: TCT_Font; ADestFont: TXc12Font);
     function  GetPrFont(ASrcFont: TCT_RPrElt): TXc12Font;
     procedure GetFill(ASrcFill: TCT_Fill; ADestFill: TXc12Fill);
     procedure GetBorders(ASrcBorder: TCT_Border; ADestBorder: TXc12Border);
     procedure GetAlignment(ASrcAlign: TCT_CellAlignment; ADestAlign: TXc12CellAlignment);
     procedure GetRst(ARst: TCT_Rst; out AText: AxUCString; out ARuns: TXc12DynFontRunArray; out APhonetics: TXc12DynPhoneticRunArray);

     procedure ReadStyles;
     procedure ReadStyleNumFmts(AList: TCT_NumFmtXpgList);
     procedure ReadStyleFonts(AList: TCT_FontXpgList);
     procedure ReadStyleFills(AList: TCT_FillXpgList);
     procedure ReadStyleBorders(AList: TCT_BorderXpgList);
     procedure ReadStyleXFs(AList: TCT_XfXpgList; ADest: TXc12XFs);
     procedure ReadStyleCellStyles(AList: TCT_CellStyleXpgList);
     procedure ReadStyleDxfs(AList: TCT_DxfXpgList);
     procedure ReadStyleTableStyles(AList: TCT_TableStyles);
     procedure ReadStyleColors(AColors: TCT_Colors);

     procedure ReadMedia(ADrawinRId: AxUCString);
     procedure ReadSST;
     procedure ReadConnections;
     procedure ReadWorkbook;
     procedure ReadSheets;
     procedure ReadSheetsThread;
     procedure ReadSheet(const AId: AxUCString);
     procedure ReadChartSheet(const AId: AxUCString);
     procedure ReadSheetView(ASheetViews: TCT_SheetViews);
     procedure ReadSheetAutoFilter(ASrc: TCT_AutoFilter; ADest: TXc12AutoFilter);
     procedure ReadSheetSortState(ASrc: TCT_SortState; ADest: TXc12SortState);
     procedure ReadSheetCustomSheetViews(ACSViews: TCT_CustomSheetViews);
     procedure ReadSheetPane(ASrc: TCT_Pane; ADest: TXc12Pane);
     procedure ReadSheetSelection(ASrc: TCT_Selection; ADest: TXc12Selection);
     procedure ReadSheetPageBreak(ASrc: TCT_PageBreak; ADest: TXc12PageBreaks);
     procedure ReadSheetPageMargins(ASrc: TCT_PageMargins; ADest: TXc12PageMargins);
     procedure ReadSheetPrintOptions(ASrc: TCT_PrintOptions; ADest: TXc12PrintOptions);
     procedure ReadSheetPageSetup(ASrc: TCT_PageSetup; ADest: TXc12PageSetup);
     procedure ReadSheetHeaderFooter(ASrc: TCT_HeaderFooter; ADest: TXc12HeaderFooter);
     procedure ReadSheetCondFmt(ASrc: TCT_ConditionalFormatting; ADest: TXc12ConditionalFormatting);
     procedure ReadSheetCfvo(ASrc: TCT_Cfvo; ADest: TXc12Cfvo);
     procedure ReadSheetColors(ASrc: TCT_ColorXpgList; ADest: TXc12Colors);
     procedure ReadSheetSmartTags(ASrc: TCT_CellSmartTagsXpgList; ADest: TXc12SmartTags);
     procedure ReadTable(const AId: AxUCString);
     procedure ReadComments(const ASheetId: AxUCString);
     procedure ReadXLink(const AId: AxUCString);
     procedure ReadDrawing(const AId: AxUCString);
     procedure ReadVmlDrawing(const AId: AxUCString);

     procedure OnSSTItem(ASender: TObject);
     procedure OnCell(ASender: TObject);
     procedure OnCellDirect(ASender: TObject);
     procedure OnRow(ASender: TObject);
     procedure OnCol(ASender: TObject);
     procedure OnMergeCell(ASender: TObject);
     procedure OnHyperlink(ASender: TObject);
     procedure OnDefinedName(ASender: TObject);
     procedure OnComment(ASender: TObject);
     procedure OnDummy(ASender: TObject);

     procedure ClearArrayFormulas;

     procedure HandleParseError(ASender: TObject);
public
     constructor Create(AManager: TXc12Manager; AFormulas: TXLSFormulaHandler);
     destructor Destroy; override;

     procedure LoadFromStream(AZIPStream: TStream);
     procedure LoadSheetNamesFromStream(AZIPStream: TStream; AList: TStrings);
     end;

implementation

{ TXLSReadXLSX }

procedure TXLSReadXLSX.GetAlignment(ASrcAlign: TCT_CellAlignment; ADestAlign: TXc12CellAlignment);
begin
  ADestAlign.HorizAlignment := TXc12HorizAlignment(GetEnum(Integer(ASrcAlign.Horizontal),Integer(chaGeneral)));
  ADestAlign.VertAlignment := TXc12VertAlignment(GetEnum(Integer(ASrcAlign.Vertical),Integer(cvaBottom)));
  ADestAlign.Rotation := ASrcAlign.TextRotation;
  ADestAlign.Options := [];
  if ASrcAlign.WrapText then
    ADestAlign.Options := ADestAlign.Options + [foWrapText];
  if ASrcAlign.ShrinkToFit then
    ADestAlign.Options := ADestAlign.Options + [foShrinkToFit];
  if ASrcAlign.JustifyLastLine then
    ADestAlign.Options := ADestAlign.Options + [foJustifyLastLine];
  ADestAlign.Indent := ASrcAlign.Indent;
  ADestAlign.RelativeIndent := ASrcAlign.RelativeIndent;
  ADestAlign.ReadingOrder := TXc12ReadOrder(ASrcAlign.ReadingOrder);
end;

procedure TXLSReadXLSX.GetBorders(ASrcBorder: TCT_Border; ADestBorder: TXc12Border);
begin
  ADestBorder.Options := [];
  if ASrcBorder.DiagonalUp then
    ADestBorder.Options := ADestBorder.Options + [ecboDiagonalUp];
  if ASrcBorder.DiagonalDown then
    ADestBorder.Options := ADestBorder.Options + [ecboDiagonalDown];
  if ASrcBorder.Outline then
    ADestBorder.Options := ADestBorder.Options + [ecboOutline];
  GetBorder(ASrcBorder.Left,ADestBorder.Left);
  GetBorder(ASrcBorder.Right,ADestBorder.Right);
  GetBorder(ASrcBorder.Top,ADestBorder.Top);
  GetBorder(ASrcBorder.Bottom,ADestBorder.Bottom);
  GetBorder(ASrcBorder.Diagonal,ADestBorder.Diagonal);
  GetBorder(ASrcBorder.Vertical,ADestBorder.Vertical);
  GetBorder(ASrcBorder.Horizontal,ADestBorder.Horizontal);
end;

procedure TXLSReadXLSX.GetFill(ASrcFill: TCT_Fill; ADestFill: TXc12Fill);
var
  i: integer;
begin
  if ASrcFill.PatternFill.Available then begin
    ADestFill.PatternType := TXc12FillPattern(GetEnum(Integer(ASrcFill.PatternFill.PatternType),Integer(efpNone)));
    ADestFill.FgColor := GetColor(ASrcFill.PatternFill.FgColor);
    ADestFill.BgColor := GetColor(ASrcFill.PatternFill.BgColor);
  end
  else if ASrcFill.GradientFill.Available then begin
    ADestFill.IsGradientFill := True;
    for i := 0 to ASrcFill.GradientFill.StopXpgList.Count - 1 do
      ADestFill.GradientFill.Stops.Add(GetColor(ASrcFill.GradientFill.StopXpgList[i].Color),ASrcFill.GradientFill.StopXpgList[i].Position);
    ADestFill.GradientFill.GradientType := TXc12GradientFillType(ASrcFill.GradientFill.Type_);
    ADestFill.GradientFill.Degree := ASrcFill.GradientFill.Degree;
    ADestFill.GradientFill.Left := ASrcFill.GradientFill.Left;
    ADestFill.GradientFill.Right := ASrcFill.GradientFill.Right;
    ADestFill.GradientFill.Top := ASrcFill.GradientFill.Top;
    ADestFill.GradientFill.Bottom := ASrcFill.GradientFill.Bottom;
  end;
end;

procedure TXLSReadXLSX.GetFont(ASrcFont: TCT_Font; ADestFont: TXc12Font);
begin
  ADestFont.Charset := ASrcFont.Charset.Val;

  if ASrcFont.Color.Available then
    ADestFont.Color := GetColor(ASrcFont.Color)
  else
    ADestFont.Color := MakeXc12ColorAuto;

  if ASrcFont.Condense.Available then
    ADestFont.Style := ADestFont.Style + [xfsCondense];
  if ASrcFont.Extend.Available then
    ADestFont.Style := ADestFont.Style + [xfsExtend];
  if ASrcFont.Family.Available then
    ADestFont.Family := ASrcFont.Family.Val;
  ADestFont.Name := ASrcFont.Name.Val;
  if ASrcFont.Outline.Available then
    ADestFont.Style := ADestFont.Style + [xfsOutline];
  if ASrcFont.Shadow.Available then
    ADestFont.Style := ADestFont.Style + [xfsShadow];
  if ASrcFont.Scheme.Available then
    ADestFont.Scheme := TXc12FontScheme(ASrcFont.Scheme.Val);

  if ASrcFont.U.Val = stuvNone then
    ADestFont.Underline := xulNone
  else
    ADestFont.Underline := TXc12Underline(Integer(ASrcFont.U.Val) + 1);

  if ASrcFont.VertAlign.Available then
    ADestFont.SubSuperScript := TXc12SubSuperscript(ASrcFont.VertAlign.Val);
  ADestFont.Size := ASrcFont.Sz.Val;
  ADestFont.Style := [];
  if ASrcFont.B.Available and ASrcFont.B.Val then
    ADestFont.Style := ADestFont.Style + [xfsBold];
  if ASrcFont.I.Available and ASrcFont.I.Val then
    ADestFont.Style := ADestFont.Style + [xfsItalic];
  if ASrcFont.Strike.Available and ASrcFont.Strike.Val then
    ADestFont.Style := ADestFont.Style + [xfsStrikeOut];
end;

procedure TXLSReadXLSX.ClearArrayFormulas;
var
  i: integer;
begin
  for i := 0 to FArrayFormulas.Count - 1 do
    FArrayFormulas[i].Obj.Free;
  FArrayFormulas.Clear;
end;

constructor TXLSReadXLSX.Create(AManager: TXc12Manager; AFormulas: TXLSFormulaHandler);
begin
  FManager := AManager;
  FFormulas := AFormulas;
  FSharedFormulas := TXLSSharedFormulaList.Create;
  FArrayFormulas := TCellAreas.Create;
end;

destructor TXLSReadXLSX.Destroy;
begin
  FSharedFormulas.Free;
  ClearArrayFormulas;
  FArrayFormulas.Free;
  inherited;
end;

procedure TXLSReadXLSX.GetBorder(ABorder: TCT_BorderPr; ADest: TXc12BorderPr);
begin
  ADest.Color := GetColor(ABorder.Color);
  ADest.Style := TXc12CellBorderStyle(ABorder.Style);
end;

function TXLSReadXLSX.GetColor(AColor: xpgParserXLSX.TCT_Color): TXc12Color;
begin
  if AColor.Auto then
    Result.ColorType := exctAuto
  else if AColor.Theme >= 0 then begin
    Result.ColorType := exctTheme;
    Result.Theme := TXc12ClrSchemeColor(AColor.Theme);
  end
  else if AColor.Indexed >= 0 then begin
    Result.ColorType := exctIndexed;
    Result.Indexed := TXc12IndexColor(AColor.Indexed);
  end
  else begin
    Result.ColorType := exctRGB;
    Result.OrigRGB := AColor.RGB;
  end;
  Result.Tint := AColor.Tint;
  Result.ARGB := Xc12ColorToRGB(Result);
end;

function TXLSReadXLSX.GetEnum(AValue, ADefault: integer): integer;
begin
  if AValue <> XPG_UNKNOWN_ENUM then
    Result := AValue
  else
    Result := ADefault;
end;

function TXLSReadXLSX.GetPrFont(ASrcFont: TCT_RPrElt): TXc12Font;
var
  F: TXc12Font;
begin
  F := TXc12Font.Create(Nil);
  try
    F.Charset := ASrcFont.Charset.Val;
    if ASrcFont.Color.Available then
      F.Color := GetColor(ASrcFont.Color)
    else
      F.Color := MakeXc12ColorAuto;

    if ASrcFont.Family.Available then
      F.Family := ASrcFont.Family.Val;
    F.Name := ASrcFont.rFont.Val;
    if ASrcFont.Scheme.Available then
      F.Scheme := TXc12FontScheme(ASrcFont.Scheme.Val);

    if ASrcFont.U.Val = stuvNone then
      F.Underline := xulNone
    else
      F.Underline := TXc12Underline(Integer(ASrcFont.U.Val) + 1);

    if ASrcFont.VertAlign.Available then
      F.SubSuperScript := TXc12SubSuperscript(ASrcFont.VertAlign.Val);
    F.Size := ASrcFont.Sz.Val;

    F.Style := [];
    if ASrcFont.Outline.Available and ASrcFont.Outline.Val then
      F.Style := F.Style + [xfsOutline];
    if ASrcFont.Shadow.Available and ASrcFont.Shadow.Val then
      F.Style := F.Style + [xfsShadow];
    if ASrcFont.Extend.Available and ASrcFont.Extend.Val then
      F.Style := F.Style + [xfsExtend];
    if ASrcFont.Condense.Available and ASrcFont.Condense.Val then
      F.Style := F.Style + [xfsCondense];
    if ASrcFont.B.Available {and ASrcFont.B.Val} then
      F.Style := F.Style + [xfsBold];
    if ASrcFont.I.Available {and ASrcFont.I.Val} then
      F.Style := F.Style + [xfsItalic];
    if ASrcFont.Strike.Available and ASrcFont.Strike.Val then
      F.Style := F.Style + [xfsStrikeOut];

    Result := FManager.StyleSheet.XFEditor.UseRequiredFont(F);
  finally
    F.Free;
  end;
end;

function TXLSReadXLSX.GetProtections(AProtection: TCT_CellProtection): TXc12CellProtections;
begin
  Result := [];
  if AProtection.Locked then
    Result := Result + [cpLocked];
  if AProtection.Hidden then
    Result := Result + [cpHidden];
end;

procedure TXLSReadXLSX.GetRst(ARst: TCT_Rst; out AText: AxUCString; out ARuns: TXc12DynFontRunArray; out APhonetics: TXc12DynPhoneticRunArray);
var
  i,j: integer;
  Font: TXc12Font;
begin
  j := MAXINT;

  if ARst.RXpgList.Count > 0 then begin
    if ARst.RXpgList[0].RPr.Available then begin
      SetLength(ARuns,ARst.RXpgList.Count);
      j := 0;
      AText := '';
    end
    else begin
      SetLength(ARuns,ARst.RXpgList.Count - 1);
      j := 1;
      AText := ARst.RXpgList[0].T;
    end;
  end
  else begin
    SetLength(ARuns,1);
    ARuns[0].Index := 1;
    ARuns[0].Font := FManager.StyleSheet.Fonts.DefaultFont;
    AText := ARst.T;
  end;

  for i := j to ARst.RXpgList.Count - 1 do begin
    if ARst.RXpgList[i].RPr.Available then begin
      Font := GetPrFont(ARst.RXpgList[i].RPr);
      ARuns[i - j].Font := Font;
    end
    else
      ARuns[i - j].Font := Nil;
    ARuns[i - j].Index := Length(AText);
    AText := AText + ARst.RXpgList[i].T;
  end;

  SetLength(APhonetics,ARst.RPhXpgList.Count);
  for i := 0 to ARst.RPhXpgList.Count - 1 do begin
    APhonetics[i].Sb := ARst.RPhXpgList[i].Sb;
    APhonetics[i].Eb := ARst.RPhXpgList[i].Eb;
    APhonetics[i].Text := ARst.RPhXpgList[i].T;
  end;
end;

function TXLSReadXLSX.GetXfApply(AXf: TCT_Xf): TXc12ApplyFormats;
begin
  Result := [];
  if AXf.ApplyNumberFormat then Result := Result + [eafNumberFormat];
  if AXf.ApplyFont         then Result := Result + [eafFont];
  if AXf.ApplyFill         then Result := Result + [eafFill];
  if AXf.ApplyBorder       then Result := Result + [eafBorder];
  if AXf.ApplyAlignment    then Result := Result + [eafAlignment];
  if AXf.ApplyProtection   then Result := Result + [eafProtection];
end;

procedure TXLSReadXLSX.HandleParseError(ASender: TObject);
begin
  FManager.Errors.Warning(TXpgPErrors(ASender).LastErrorText,XLSWARN_FILEREAD_XMLPARSE);
end;

procedure TXLSReadXLSX.LoadFromStream(AZIPStream: TStream);
begin
  FTotCells := 0;
  FTotRows := 0;

  FManager.FileData.LoadFromStream(AZipStream);
  try
    ReadStyles;

    ReadSST;

    ReadConnections;

    ReadWorkbook;

    ReadSheets;

    FManager.FileData.ReadUnusedData;

  finally
    FManager.FileData.OPC.Close;
  end;
end;

procedure TXLSReadXLSX.LoadSheetNamesFromStream(AZIPStream: TStream; AList: TStrings);
var
  i: integer;
begin
  FManager.FileData.LoadFromStream(AZipStream);

  ReadWorkbook;

  for i := 0 to FManager.Workbook.Sheets.Count - 1 do
    AList.Add(FManager.Workbook.Sheets[i].Name);

  FManager.FileData.OPC.Close;
end;

procedure TXLSReadXLSX.OnCell(ASender: TObject);
var
  i: integer;
  C,R,FI: integer;
  CellType: TST_CellType;
  FormulaOptions: byte;
  TableOptions: byte;
  Cell: TCT_Cell;
  Ptgs: PXLSPtgs;
  Ref: TXLSCellRef;
  Area: TXLSCellArea;
  ApplyArea: TCellArea;
  Shared: TXLSSharedFormula;
begin
  Inc(FTotCells);

  Ptgs := Nil;
  FormulaOptions := 0;
  TableOptions := 0;

  Cell := TCT_Cell(ASender);

  CellType := Cell.T;
  FI := Cell.S;
  if xaRead in Cell.F.Assigneds then begin
    FCurrSheet.Cells.FormulaHelper.Clear;

    FCurrSheet.Cells.FormulaHelper.Ref := Cell.R;
    FCurrSheet.Cells.FormulaHelper.Style := FI;

    case Cell.F.T of
      stcftArray    : begin
        ApplyArea := TCellAreaImpl.Create;
        ApplyArea.Obj := TObject(1);
        ApplyArea.SetSize(FCurrSheet.Cells.FormulaHelper.Col,FCurrSheet.Cells.FormulaHelper.Row);
        FArrayFormulas.Add(Cell.F.Ref);
        FArrayFormulas.Last.Obj := ApplyArea;
        FCurrSheet.Cells.FormulaHelper.StrTargetRef := Cell.F.Ref;
        FCurrSheet.Cells.FormulaHelper.Formula := Cell.F.Content;

        if Cell.F.Aca   then FormulaOptions := FormulaOptions + Xc12FormulaOpt_ACA;
      end;
      stcftDataTable: begin
        FCurrSheet.Cells.FormulaHelper.IsTABLE := True;
        ApplyArea := TCellAreaImpl.Create;
        ApplyArea.Obj := TObject(0);
        ApplyArea.SetSize(FCurrSheet.Cells.FormulaHelper.Col,FCurrSheet.Cells.FormulaHelper.Row);
        FArrayFormulas.Add(Cell.F.Ref);
        FArrayFormulas.Last.Obj := ApplyArea;
        FCurrSheet.Cells.FormulaHelper.StrTargetRef := Cell.F.Ref;
        FCurrSheet.Cells.FormulaHelper.Formula := Cell.F.Content;

        if Cell.F.Dt2D  then TableOptions := TableOptions + Xc12FormulaTableOpt_DT2D;
        if Cell.F.Dtr   then TableOptions := TableOptions + Xc12FormulaTableOpt_DTR;
        if Cell.F.Del1  then TableOptions := TableOptions + Xc12FormulaTableOpt_DEL1;
        if Cell.F.Del2  then TableOptions := TableOptions + Xc12FormulaTableOpt_DEL2;

        if Cell.F.R1 <> '' then begin
          FCurrSheet.Cells.FormulaHelper.StrR1 := Cell.F.R1;
          if Cell.F.R2 <> '' then
            FCurrSheet.Cells.FormulaHelper.StrR2 := Cell.F.R2;
        end;
      end;
      stcftShared   : begin
        RefStrToColRow(Cell.R,Ref.Col,Ref.Row);
        if Cell.F.Content <> '' then begin
          FCurrSheet.Cells.FormulaHelper.PtgsSize := FFormulas.EncodeFormula(Cell.F.Content,Ptgs,FCurrSheet.Index);
          FCurrSheet.Cells.FormulaHelper.Ptgs := Ptgs;
          AreaStrToColRow(Cell.F.Ref,Area.Col1,Area.Row1,Area.Col2,Area.Row2);
          FSharedFormulas.Insert(Cell.F.Si,@Ref,@Area,Ptgs,FCurrSheet.Cells.FormulaHelper.PtgsSize);
        end
        else begin
          Shared := FSharedFormulas[Cell.F.Si];
          if (Ref.Col >= Shared.ApplyArea.Col1) and (Ref.Col <= Shared.ApplyArea.Col2) and (Ref.Row >= Shared.ApplyArea.Row1) and (Ref.Row <= Shared.ApplyArea.Row2) then begin
            Ptgs := Shared.CopyPtgs(Cell.F.Si);
            FFormulas.AdjustCell(Ptgs,Shared.PtgsSize,Ref.Col - Shared.Ref.Col,Ref.Row - Shared.Ref.Row);
            FCurrSheet.Cells.FormulaHelper.PtgsSize := Shared.PtgsSize;
            FCurrSheet.Cells.FormulaHelper.Ptgs := Ptgs;
          end;
        end;
      end
      else begin
        FCurrSheet.Cells.FormulaHelper.Formula := Cell.F.Content;
      end;
    end;

    FManager.Errors.LastError := 0;

    case CellType of
      stctB  : FCurrSheet.Cells.FormulaHelper.AsBoolean := XmlStrToBoolDef(Cell.V,False);
      stctN  : FCurrSheet.Cells.FormulaHelper.AsFloat := XmlStrToFloatDef(Cell.V,0);
      stctD  : raise XLSRWException.Create('Date cells not supported');
      stctE  : FCurrSheet.Cells.FormulaHelper.AsErrorStr := Cell.V;
      stctStr: FCurrSheet.Cells.FormulaHelper.AsString := Cell.V;
      else     raise XLSRWException.Create('Unsupported type of formula cell: ' + IntToStr(Integer(Cell.T)));
    end;

    case Cell.F.T of
      stcftNormal   : FCurrSheet.Cells.FormulaHelper.FormulaType := xcftNormal;
      stcftArray    : FCurrSheet.Cells.FormulaHelper.FormulaType := xcftArray;
      stcftDataTable: FCurrSheet.Cells.FormulaHelper.FormulaType := xcftDataTable;
      stcftShared   : FCurrSheet.Cells.FormulaHelper.FormulaType := xcftNormal;
    end;

    if Cell.F.Ca    then FormulaOptions := FormulaOptions + Xc12FormulaOpt_CA;
    if Cell.F.Bx    then FormulaOptions := FormulaOptions + Xc12FormulaOpt_BX;
    FCurrSheet.Cells.FormulaHelper.Options := FormulaOptions;
    FCurrSheet.Cells.FormulaHelper.TableOptions := TableOptions;

    FCurrSheet.Cells.StoreFormula;
    if Ptgs <> Nil then
      FreeMem(Ptgs);
  end
  else begin
    RefStrToColRow(Cell.R,C,R);

    if FArrayFormulas.Count > 0 then
      i := FArrayFormulas.CellInAreas(C,R)
    else
      i := -1;
    if i >= 0 then begin
      ApplyArea := TCellArea(FArrayFormulas[i].Obj);

      FCurrSheet.Cells.FormulaHelper.Clear;

      FCurrSheet.Cells.FormulaHelper.Ref := Cell.R;
      FCurrSheet.Cells.FormulaHelper.ParentCol := ApplyArea.Col1;
      FCurrSheet.Cells.FormulaHelper.ParentRow := ApplyArea.Row1;
      FCurrSheet.Cells.FormulaHelper.Style := FI;
      if Integer(ApplyArea.Obj) = 1 then
        FCurrSheet.Cells.FormulaHelper.FormulaType := xcftArrayChild
      else
        FCurrSheet.Cells.FormulaHelper.FormulaType := xcftDataTableChild;
      case CellType of
        stctB  : FCurrSheet.Cells.FormulaHelper.AsBoolean := XmlStrToBoolDef(Cell.V,False);
        stctN  : FCurrSheet.Cells.FormulaHelper.AsFloat := XmlStrToFloatDef(Cell.V,0);
        stctD  : raise XLSRWException.Create('Date cells not supported');
        stctE  : FCurrSheet.Cells.FormulaHelper.AsErrorStr := Cell.V;
        stctStr: FCurrSheet.Cells.FormulaHelper.AsString := Cell.V;
        else     raise XLSRWException.Create('Unsupported type of formula cell: ' + IntToStr(Integer(Cell.T)));
      end;
      FCurrSheet.Cells.StoreFormula;
    end
    else begin
      if Cell.V = '' then begin
        if CellType = stctInlineStr then
          FCurrSheet.Cells.StoreString(C,R,FI,Cell.Is_.T)
        else
          FCurrSheet.Cells.StoreBlank(C,R,FI);
      end
      else begin
        case CellType of
          stctB: FCurrSheet.Cells.StoreBoolean(C,R,FI,XmlStrToBoolDef(Cell.V,False));
          stctN: FCurrSheet.Cells.StoreFloat(C,R,FI,XmlStrToFloatDef(Cell.V,0));
          stctE: FCurrSheet.Cells.StoreError(C,R,FI,ErrorTextToCellError(Cell.V));
          stctS: FCurrSheet.Cells.StoreString(C,R,FI,XmlStrToIntDef(Cell.V,0));
          stctStr: FCurrSheet.Cells.StoreString(C,R,FI,Cell.V);
//          stctStr: raise XLSRWException.Create('Illegal use of formula string');
        end;
      end;
    end;
  end;
end;

procedure TXLSReadXLSX.OnCellDirect(ASender: TObject);
var
  i: integer;
  Sz: integer;
  C,R: integer;
  CellType: TST_CellType;
  Cell: TCT_Cell;
  Ptgs: PXLSPtgs;
  Ref: TXLSCellRef;
  Area: TXLSCellArea;
  Shared: TXLSSharedFormula;
begin
  Inc(FTotCells);

  Ptgs := Nil;

  FManager.EventCell.Clear;

  Cell := TCT_Cell(ASender);

  CellType := Cell.T;

  FManager.EventCell.SheetIndex := FCurrSheet.Index;
  FManager.EventCell.Ref := Cell.R;

  if xaRead in Cell.F.Assigneds then begin
    case Cell.F.T of
      stcftArray    : begin
        FArrayFormulas.Add(Cell.F.Ref);
        Exit;
      end;
      stcftDataTable: begin
        FArrayFormulas.Add(Cell.F.Ref);
        Exit;
      end;
      stcftShared   : begin
        RefStrToColRow(Cell.R,Ref.Col,Ref.Row);
        if Cell.F.Content <> '' then begin
          Sz := FFormulas.EncodeFormula(Cell.F.Content,Ptgs,FCurrSheet.Index);
          AreaStrToColRow(Cell.F.Ref,Area.Col1,Area.Row1,Area.Col2,Area.Row2);
          FSharedFormulas.Insert(Cell.F.Si,@Ref,@Area,Ptgs,Sz);
        end
        else begin
          Shared := FSharedFormulas[Cell.F.Si];
          if (Ref.Col >= Shared.ApplyArea.Col1) and (Ref.Col <= Shared.ApplyArea.Col2) and (Ref.Row >= Shared.ApplyArea.Row1) and (Ref.Row <= Shared.ApplyArea.Row2) then begin
            Ptgs := Shared.CopyPtgs(Cell.F.Si);
            FFormulas.AdjustCell(Ptgs,Shared.PtgsSize,Ref.Col - Shared.Ref.Col,Ref.Row - Shared.Ref.Row);

            FManager.EventCell.AsFormula := FFormulas.DecodeFormula(FCurrSheet.Cells,FCurrSheet.Index,Ptgs,Shared.PtgsSize,False)
          end;
        end;
      end
      else begin
//        FCurrSheet.Cells.FormulaHelper.PtgsSize := FFormulas.EncodeFormula(Cell.F.Content,Ptgs);
//        FCurrSheet.Cells.FormulaHelper.Ptgs := Ptgs;
        FManager.EventCell.AsFormula := Cell.F.Content;
      end;
    end;

    FManager.Errors.LastError := 0;

    case CellType of
      stctB  : FManager.EventCell.AsBoolean := XmlStrToBoolDef(Cell.V,False);
      stctN  : FManager.EventCell.AsFloat := XmlStrToFloatDef(Cell.V,0);
      stctD  : raise XLSRWException.Create('Date cells not supported');
      stctE  : FManager.EventCell.AsError := ErrorTextToCellError(Cell.V);
      stctStr: FManager.EventCell.AsString := Cell.V;
      else     raise XLSRWException.Create('Unsupported type of formula cell: ' + IntToStr(Integer(Cell.T)));
    end;

    if Ptgs <> Nil then
      FreeMem(Ptgs);
  end
  else begin
    if (FArrayFormulas.Count > 0) and (FArrayFormulas.CellInAreas(C,R) >= 0) then
      Exit;

    RefStrToColRow(Cell.R,C,R);

    if Cell.V = '' then begin
      if (CellType = stctInlineStr) and (Cell.Is_ <> Nil) then
        FManager.EventCell.AsString := Cell.Is_.T
      else
        Exit  // Blank cell
    end
    else begin
      case CellType of
        stctB: FManager.EventCell.AsBoolean := XmlStrToBoolDef(Cell.V,False);
        stctN: FManager.EventCell.AsFloat := XmlStrToFloatDef(Cell.V,0);
        stctE: FManager.EventCell.AsError := ErrorTextToCellError(Cell.V);
        stctS: begin
          i := XmlStrToIntDef(Cell.V,0);
          FManager.EventCell.AsString := FManager.SST.ItemText[i];
        end;
        stctStr: FManager.EventCell.AsString := Cell.V;
        stctInlineStr: raise XLSRWException.Create('Unhandled string type in file'); // FCurrSheet.IntWriteSSTString(C,R,FI,'Inline string');
      end;
    end;
  end;
  FManager.FireReadCellEvent;
end;

procedure TXLSReadXLSX.OnCol(ASender: TObject);
var
  C: TCT_Col;
  Col: TXc12Column;
begin
  C := TCT_Col(ASender);
  Col := FCurrSheet.Columns.Add(FManager.StyleSheet.XFs[C.Style]);
  FManager.StyleSheet.XFEditor.UseStyle(Col.Style);
  Col.Min := C.Min - 1;
  Col.Max := C.Max - 1;
  Col.Width := C.Width;
  Col.OutlineLevel := C.OutlineLevel;
  Col.Options := [];
  if C.Collapsed then
    Col.Options := Col.Options + [xcoCollapsed];
  if C.Hidden then
    Col.Options := Col.Options + [xcoHidden];
  if C.BestFit then
    Col.Options := Col.Options + [xcoBestFit];
  if C.CustomWidth then
    Col.Options := Col.Options + [xcoCustomWidth];
  if C.Phonetic then
    Col.Options := Col.Options + [xcoPhonetic];
end;

procedure TXLSReadXLSX.OnComment(ASender: TObject);
var
  S: AxUCString;
  Runs: TXc12DynFontRunArray;
  Phonetic: TXc12DynPhoneticRunArray;
  C: TCT_Comment;
  Comment: TXc12Comment;
  Col,Row: integer;
begin
  C := TCT_Comment(ASender);
  Comment := FCurrSheet.Comments.Add;
  Runs := Comment.FontRuns;
  Phonetic := Comment.PhoneticRuns;
  GetRst(C.Text,S,Runs,Phonetic);
  Comment.FontRuns := Runs;
  Comment.Text := S;
  RefStrToColRow(C.Ref,Col,Row);
  Comment.Col := Col;
  Comment.Row := Row;
  Comment.GUID := C.Guid;

  Comment.TempAuthorId := C.AuthorId;
end;

procedure TXLSReadXLSX.OnDefinedName(ASender: TObject);
var
  N: TCT_DefinedName;
  Name: TXc12DefinedName;
begin
  N := TCT_DefinedName(ASender);

  // Skip this.
  if N.Name = 'LOCAL_MYSQL_DATE_FORMAT' then
    Exit;

  Name := FManager.Workbook.DefinedNames.Add(N.Name,N.LocalSheetId);

  Name.Content := N.Content;

//  Name.Name := N.Name;
  Name.Comment := N.Comment;
  Name.CustomMenu := N.CustomMenu;
  Name.Description := N.Description;
  Name.Help := N.Help;
  Name.StatusBar := N.StatusBar;
  Name.LocalSheetId := N.LocalSheetId;
  Name.FunctionGroupId := N.FunctionGroupId;
  Name.ShortcutKey := N.ShortcutKey;

  Name.Hidden := N.Hidden;
  Name.Function_ := N.Function_;
  Name.VbProcedure := N.VbProcedure;
  Name.Xlm := N.Xlm;
  Name.PublishToServer := N.PublishToServer;
  Name.WorkbookParameter := N.WorkbookParameter;
end;

procedure TXLSReadXLSX.OnDummy(ASender: TObject);
begin

end;

procedure TXLSReadXLSX.OnHyperlink(ASender: TObject);
var
  H: TCT_Hyperlink;
  HLink: TXc12Hyperlink;
begin
  H := TCT_Hyperlink(ASender);
  HLink := FCurrSheet.Hyperlinks.Add;
  HLink.Ref := AreaStrToCellArea(TCT_Hyperlink(ASender).Ref);

  if H.R_Id <> '' then
    HLink.RawAddress := FManager.FileData.OPC.FindSheetItemTarget(FCurrSheet.RId,H.R_Id)
  else
    HLink.RawAddress := H.Location;
  HLink.Location := H.Location;

  HLink.Display := H.Display;
  HLink.ToolTip := H.Tooltip;
end;

procedure TXLSReadXLSX.OnMergeCell(ASender: TObject);
begin
  FCurrSheet.MergedCells.Add(AreaStrToCellArea(TCT_MergeCell(ASender).Ref));
end;

procedure TXLSReadXLSX.OnRow(ASender: TObject);
var
  R: TCT_Row;
  Row: PXLSMMURowHeader;
begin
  Inc(FTotRows);

  R := TCT_Row(ASender);
  Row := FCurrSheet.Cells.AddRow(R.R - 1,R.S);
  if R.Ht <> 0 then
    Row.Height := Round(R.Ht * 20)
  else
    Row.Height := XLS_DEFAULT_ROWHEIGHT_FLAG;
  Row.Options := [];
  if R.Hidden then
    Row.Options := Row.Options + [xroHidden];
  if R.Collapsed then
    Row.Options := Row.Options + [xroCollapsed];
  if R.CustomHeight then
    Row.Options := Row.Options + [xroCustomHeight];
  if R.Ph then
    Row.Options := Row.Options + [xroPhonetic];
  if R.ThickTop then
    Row.Options := Row.Options + [xroThickTop];
  if R.ThickBot then
    Row.Options := Row.Options + [xroThickBottom];
  Row.OutlineLevel := R.OutlineLevel;

  R.Clear;
end;

procedure TXLSReadXLSX.OnSSTItem(ASender: TObject);
var
  S: AxUCString;
  si: TCT_Rst;
  Runs: TXc12DynFontRunArray;
  Phonetics: TXc12DynPhoneticRunArray;
begin
  si := TCT_Rst(ASender);
  if si.RXpgList.Count > 0 then begin
    // TODO Phonetic
    GetRst(si,S,Runs,Phonetics);
    FManager.SST.AddFormattedString(S,Runs);
  end
  else
    FManager.SST.FileAddString(si.T);

  si.Clear;
end;

procedure TXLSReadXLSX.ReadSheetAutoFilter(ASrc: TCT_AutoFilter; ADest: TXc12AutoFilter);
var
  i,j: integer;
  FilterColumn: TXc12FilterColumn;
  FColumn: TCT_FilterColumn;
  DateGroupItem: TXc12DateGroupItem;
begin
  ADest.Ref := AreaStrToCellArea(ASrc.Ref);

  for i := 0 to ASrc.FilterColumnXpgList.Count - 1 do begin
    FColumn := ASrc.FilterColumnXpgList[i];
    FilterColumn := ADest.FilterColumns.Add;

    FilterColumn.ColId := FColumn.ColId;
    FilterColumn.HiddenButton := FColumn.HiddenButton;
    FilterColumn.ShowButton := FColumn.ShowButton;

    FilterColumn.Filters.Blank := FColumn.Filters.Blank;
    FilterColumn.Filters.CalendarType := TXc12CalendarType(FColumn.Filters.CalendarType);

    for j := 0 to FColumn.Filters.FilterXpgList.Count - 1 do
      FilterColumn.Filters.Filter.Add(FColumn.Filters.FilterXpgList[j].Val);

    for j := 0 to FColumn.Filters.DateGroupItemXpgList.Count - 1 do begin
      DateGroupItem := FilterColumn.Filters.DateGroupItems.Add;
      DateGroupItem.Year := FColumn.Filters.DateGroupItemXpgList[j].Year;
      DateGroupItem.Month := FColumn.Filters.DateGroupItemXpgList[j].Month;
      DateGroupItem.Day := FColumn.Filters.DateGroupItemXpgList[j].Day;
      DateGroupItem.Hour := FColumn.Filters.DateGroupItemXpgList[j].Hour;
      DateGroupItem.Minute := FColumn.Filters.DateGroupItemXpgList[j].Minute;
      DateGroupItem.Second := FColumn.Filters.DateGroupItemXpgList[j].Second;
      DateGroupItem.DateTimeGrouping := TXc12DateTimeGrouping(FColumn.Filters.DateGroupItemXpgList[j].DateTimeGrouping);
    end;

    if FColumn.Available then begin
      FilterColumn.Top10.Assigned := True;
      FilterColumn.Top10.Top := FColumn.Top10.Top;
      FilterColumn.Top10.Percent := FColumn.Top10.Percent;
      FilterColumn.Top10.Val := FColumn.Top10.Val;
      FilterColumn.Top10.FilterVal := FColumn.Top10.FilterVal;
    end;

    if FColumn.CustomFilters.CustomFilterXpgList.Count > 0 then begin
      FilterColumn.CustomFilters.Assigned := True;

      FilterColumn.CustomFilters.And_ := FColumn.CustomFilters.And_;

      FilterColumn.CustomFilters.Filter1.Operator_ := TXc12FilterOperator(FColumn.CustomFilters.CustomFilterXpgList[0].Operator_);
      FilterColumn.CustomFilters.Filter1.Val := FColumn.CustomFilters.CustomFilterXpgList[0].Val;
      if FColumn.CustomFilters.CustomFilterXpgList.Count > 1 then begin
        FilterColumn.CustomFilters.Filter2.Operator_ := TXc12FilterOperator(FColumn.CustomFilters.CustomFilterXpgList[1].Operator_);
        FilterColumn.CustomFilters.Filter2.Val := FColumn.CustomFilters.CustomFilterXpgList[1].Val;
      end;
    end;

    if FColumn.DynamicFilter.Available then begin
      FilterColumn.DynamicFilter.Type_ := TXc12DynamicFilterType(FColumn.DynamicFilter.Type_);
      FilterColumn.DynamicFilter.Val := FColumn.DynamicFilter.Val;
      FilterColumn.DynamicFilter.MaxVal := FColumn.DynamicFilter.MaxVal;
    end
    else
      FilterColumn.DynamicFilter.Type_ := TXc12DynamicFilterType(XPG_UNKNOWN_ENUM);

    if FColumn.ColorFilter.Available then begin
      FilterColumn.ColorFilter.DXF := FManager.StyleSheet.DXFs[FColumn.ColorFilter.DxfId];
      FilterColumn.ColorFilter.CellColor := FColumn.ColorFilter.CellColor;
    end;

    if FColumn.IconFilter.Available then begin
      FilterColumn.IconFilter.IconSet := TXc12IconSetType(FColumn.IconFilter.IconSet);
      FilterColumn.IconFilter.IconId := FColumn.IconFilter.IconId;
    end;

    if ASrc.SortState.Available then
      ReadSheetSortState(ASrc.SortState,ADest.SortState);
  end;
end;

procedure TXLSReadXLSX.ReadSheetCfvo(ASrc: TCT_Cfvo; ADest: TXc12Cfvo);
begin
  ADest.Type_ := TXc12CfvoType(ASrc.Type_);
  ADest.Val := ASrc.Val;
  ADest.Gte := ASrc.Gte;
end;

procedure TXLSReadXLSX.ReadSheetColors(ASrc: TCT_ColorXpgList; ADest: TXc12Colors);
var
  i: integer;
begin
  for i := 0 to ASrc.Count - 1 do
    ADest.Add(GetColor(ASrc[i]));
end;

procedure TXLSReadXLSX.ReadSheetCondFmt(ASrc: TCT_ConditionalFormatting; ADest: TXc12ConditionalFormatting);
var
  i,j: integer;
  CfRule: TXc12CfRule;
  CfR: TCt_CfRule;
begin
  ADest.Pivot := ASrc.Pivot;
  ADest.SQRef.DelimitedText := ASrc.SqrefXpgList.DelimitedText;

  for i := 0 to ASrc.CfRuleXpgList.Count - 1 do begin
    CfR := ASrc.CfRuleXpgList[i];
    CfRule := ADest.CfRules.Add;

    if Integer(CfR.Type_) <> XPG_UNKNOWN_ENUM then
      CfRule.Type_ := TXc12CfType(CfR.Type_);
    if CfR.DxfId >= 0 then
      CfRule.DXF := FManager.StyleSheet.DXFs[CfR.DxfId]
    else
      CfRule.DXF := Nil;
    CfRule.Priority := CfR.Priority;
    CfRule.StopIfTrue := CfR.StopIfTrue;
    CfRule.AboveAverage := CfR.AboveAverage;
    CfRule.Percent := CfR.Percent;
    CfRule.Bottom := CfR.Bottom;
    if Integer(CfR.Operator_) <> XPG_UNKNOWN_ENUM then
      CfRule.Operator_ := TXc12ConditionalFormattingOperator(CfR.Operator_);
    CfRule.Text := CfR.Text;
    if Integer(CfR.TimePeriod) <> XPG_UNKNOWN_ENUM then
      CfRule.TimePeriod := TXc12TimePeriod(CfR.TimePeriod);
    CfRule.Rank := CfR.Rank;
    CfRule.StdDev := CfR.StdDev;
    CfRule.EqualAverage := CfR.EqualAverage;

    for j := 0 to CfR.FormulaXpgList.Count - 1 do begin
      case j of
        0: CfRule.Formulas[j] := CfR.FormulaXpgList[j];
        1: CfRule.Formulas[j] := CfR.FormulaXpgList[j];
        2: CfRule.Formulas[j] := CfR.FormulaXpgList[j];
      end;
    end;

    if CfR.ColorScale.Available then begin
      for j := 0 to CfR.ColorScale.CfvoXpgList.Count - 1 do
        ReadSheetCfvo(CfR.ColorScale.CfvoXpgList[j],CFRule.ColorScale.Cfvos.Add);
      ReadSheetColors(CfR.ColorScale.ColorXpgList,CFRule.ColorScale.Colors);
    end;

    if CfR.DataBar.Available then begin
      CFRule.DataBar.MinLength := CfR.DataBar.MinLength;
      CFRule.DataBar.MaxLength := CfR.DataBar.MaxLength;
      CFRule.DataBar.ShowValue := CfR.DataBar.ShowValue;

      if CfR.DataBar.CfvoXpgList.Count > 0 then
        ReadSheetCfvo(CfR.DataBar.CfvoXpgList[0],CFRule.DataBar.Cfvo1);
      if CfR.DataBar.CfvoXpgList.Count > 1 then
        ReadSheetCfvo(CfR.DataBar.CfvoXpgList[1],CFRule.DataBar.Cfvo2);

      if CfR.DataBar.Color.Available then
        CFRule.DataBar.Color := GetColor(CfR.DataBar.Color);
    end;

    if CfR.IconSet.Available then begin
      CFRule.IconSet.IconSet := TXc12IconSetType(CfR.IconSet.IconSet);
      CFRule.IconSet.ShowValue := CfR.IconSet.ShowValue;
      CFRule.IconSet.Percent := CfR.IconSet.Percent;
      CFRule.IconSet.Reverse := CfR.IconSet.Reverse;

      for j := 0 to CfR.IconSet.CfvoXpgList.Count - 1 do
        ReadSheetCfvo(CfR.IconSet.CfvoXpgList[j],CFRule.IconSet.Cfvos.Add);
    end;
  end;
end;

procedure TXLSReadXLSX.ReadSheetCustomSheetViews(ACSViews: TCT_CustomSheetViews);
var
  i: integer;
  CSheetView: TCT_CustomSheetView;
  CustomSheetView: TXc12CustomSheetView;
begin
  for i := 0 to ACSViews.CustomSheetViewXpgList.Count - 1 do begin
    CSheetView := ACSViews.CustomSheetViewXpgList[i];
    CustomSheetView := FCurrSheet.CustomSheetViews.Add;

    CustomSheetView.GUID := CSheetView.GUID;
    CustomSheetView.Scale := CSheetView.Scale;
    CustomSheetView.ColorId := CSheetView.ColorId;
    CustomSheetView.ShowPageBreaks := CSheetView.ShowPageBreaks;
    CustomSheetView.ShowFormulas := CSheetView.ShowFormulas;
    CustomSheetView.ShowGridLines := CSheetView.ShowGridLines;
    CustomSheetView.ShowRowCol := CSheetView.ShowRowCol;
    CustomSheetView.OutlineSymbols := CSheetView.OutlineSymbols;
    CustomSheetView.ZeroValues := CSheetView.ZeroValues;
    CustomSheetView.FitToPage := CSheetView.FitToPage;
    CustomSheetView.PrintArea := CSheetView.PrintArea;
    CustomSheetView.Filter := CSheetView.Filter;
    CustomSheetView.ShowAutoFilter := CSheetView.ShowAutoFilter;
    CustomSheetView.HiddenRows := CSheetView.HiddenRows;
    CustomSheetView.HiddenColumns := CSheetView.HiddenColumns;
    CustomSheetView.State := TXc12SheetState(CSheetView.State);
    CustomSheetView.FilterUnique := CSheetView.FilterUnique;
    CustomSheetView.View := TXc12SheetViewType(CSheetView.View);
    CustomSheetView.ShowRuler := CSheetView.ShowRuler;
    CustomSheetView.TopLeftCell := AreaStrToCellArea(CSheetView.TopLeftCell);

    if CSheetView.Pane.Available then
      ReadSheetPane(CSheetView.Pane,CustomSheetView.Pane);

    if CSheetView.Selection.Available then
      ReadSheetSelection(CSheetView.Selection,CustomSheetView.Selection);

    if CSheetView.RowBreaks.BrkXpgList.Count > 0 then
      ReadSheetPageBreak(CSheetView.RowBreaks,CustomSheetView.RowBreaks);

    if CSheetView.ColBreaks.BrkXpgList.Count > 0 then
      ReadSheetPageBreak(CSheetView.ColBreaks,CustomSheetView.ColBreaks);

    if CSheetView.PageMargins.Available then
      ReadSheetPageMargins(CSheetView.PageMargins,CustomSheetView.PageMargins);

    if CSheetView.PrintOptions.Available then
      ReadSheetPrintOptions(CSheetView.PrintOptions,CustomSheetView.PrintOptions);

    if CSheetView.PageSetup.Available then
      ReadSheetPageSetup(CSheetView.PageSetup,CustomSheetView.PageSetup);

    if CSheetView.HeaderFooter.Available then
      ReadSheetHeaderFooter(CSheetView.HeaderFooter,CustomSheetView.HeaderFooter);

    if CSheetView.AutoFilter.Available then
      ReadSheetAutofilter(CSheetView.AutoFilter,CustomSheetView.AutoFilter);
  end;
end;

procedure TXLSReadXLSX.ReadSheetHeaderFooter(ASrc: TCT_HeaderFooter; ADest: TXc12HeaderFooter);
begin
  ADest.DifferentOddEven := ASrc.DifferentOddEven;
  ADest.DifferentFirst := ASrc.DifferentFirst;
  ADest.ScaleWithDoc := ASrc.ScaleWithDoc;
  ADest.AlignWithMargins := ASrc.AlignWithMargins;
  ADest.OddHeader := ASrc.OddHeader;
  ADest.OddFooter := ASrc.OddFooter;
  ADest.EvenHeader := ASrc.EvenHeader;
  ADest.EvenFooter := ASrc.EvenFooter;
  ADest.FirstHeader := ASrc.FirstHeader;
  ADest.FirstFooter := ASrc.FirstFooter;
end;

procedure TXLSReadXLSX.ReadSheetPageBreak(ASrc: TCT_PageBreak; ADest: TXc12PageBreaks);
var
  i: integer;
  Brk: TXc12Break;
begin
  for i := 0 to ASrc.BrkXpgList.Count - 1 do begin
    Brk := ADest.Add;
    Brk.Id := ASrc.BrkXpgList[i].Id;
    Brk.Min := ASrc.BrkXpgList[i].Min;
    Brk.Max := ASrc.BrkXpgList[i].Max;
    Brk.Man := ASrc.BrkXpgList[i].Man;
    Brk.Pt := ASrc.BrkXpgList[i].Pt;
  end;
end;

procedure TXLSReadXLSX.ReadSheetPageMargins(ASrc: TCT_PageMargins; ADest: TXc12PageMargins);
begin
  ADest.Left := ASrc.Left;
  ADest.Right := ASrc.Right;
  ADest.Top := ASrc.Top;
  ADest.Bottom := ASrc.Bottom;
  ADest.Header := ASrc.Header;
  ADest.Footer := ASrc.Footer;
end;

procedure TXLSReadXLSX.ReadSheetPageSetup(ASrc: TCT_PageSetup; ADest: TXc12PageSetup);
begin
  ADest.PaperSize := ASrc.PaperSize;
  ADest.Scale := ASrc.Scale;
  ADest.FirstPageNumber := ASrc.FirstPageNumber;
  ADest.FitToWidth := ASrc.FitToWidth;
  ADest.FitToHeight := ASrc.FitToHeight;
  ADest.PageOrder := TXc12PageOrder(ASrc.PageOrder);
  ADest.Orientation := TXc12Orientation(ASrc.Orientation);
  ADest.UsePrinterDefaults := ASrc.UsePrinterDefaults;
  ADest.BlackAndWhite := ASrc.BlackAndWhite;
  ADest.Draft := ASrc.Draft;
  ADest.CellComments := TXc12CellComments(ASrc.CellComments);
  ADest.UseFirstPageNumber := ASrc.UseFirstPageNumber;
  ADest.Errors := TXc12PrintError(ASrc.Errors);
  ADest.HorizontalDpi := ASrc.HorizontalDpi;
  ADest.VerticalDpi := ASrc.VerticalDpi;
  ADest.Copies := ASrc.Copies;
  if ASrc.R_Id <> '' then
    ADest.PrinterSettings := FManager.FileData.OPC.ReadPrinterSettings(ASrc.R_Id);
end;

procedure TXLSReadXLSX.ReadSheetPane(ASrc: TCT_Pane; ADest: TXc12Pane);
begin
  ADest.XSplit := ASrc.XSplit;
  ADest.YSplit := ASrc.YSplit;
  ADest.TopLeftCell := AreaStrToCellArea(ASrc.TopLeftCell);
  ADest.ActivePane := TXc12PaneEnum(ASrc.ActivePane);
  ADest.State := TXc12PaneState(ASrc.State);
end;

procedure TXLSReadXLSX.ReadSheetPrintOptions(ASrc: TCT_PrintOptions; ADest: TXc12PrintOptions);
begin
  ADest.HorizontalCentered := ASrc.HorizontalCentered;
  ADest.VerticalCentered := ASrc.VerticalCentered;
  ADest.Headings := ASrc.Headings;
  ADest.GridLines := ASrc.GridLines;
  ADest.GridLinesSet := ASrc.GridLinesSet;
end;

procedure TXLSReadXLSX.ReadSheetSelection(ASrc: TCT_Selection; ADest: TXc12Selection);
begin
  ADest.Pane := TXc12PaneEnum(ASrc.Pane);
  ADest.ActiveCell := AreaStrToCellArea(ASrc.ActiveCell);
  ADest.ActiveCellId := ASrc.ActiveCellId;
  if ASrc.SqrefXpgList.Count > 0 then
    ADest.SQRef.DelimitedText := ASrc.SqrefXpgList.DelimitedText;
end;

procedure TXLSReadXLSX.ReadSheetSmartTags(ASrc: TCT_CellSmartTagsXpgList; ADest: TXc12SmartTags);
var
  i,j,k: integer;
  CellSmartTags: TXc12CellSmartTags;
  CellSmartTag: TXc12CellSmartTag;
  CellSmartTagPr: TXc12CellSmartTagPr;
  CSTagPr: TCT_CellSmartTagPr;
begin
  for i := 0 to ASrc.Count - 1 do begin
    CellSmartTags := ADest.Add;

    CellSmartTags.Ref.AsString := ASrc[i].R;

    for j := 0 to ASrc[i].CellSmartTagXpgList.Count - 1 do begin
      CellSmartTag := CellSmartTags.Add;

      CellSmartTag.Type_ := ASrc[i].CellSmartTagXpgList[j].Type_;
      CellSmartTag.Deleted := ASrc[i].CellSmartTagXpgList[j].Deleted;
      CellSmartTag.XmlBased := ASrc[i].CellSmartTagXpgList[j].XmlBased;

      for k := 0 to ASrc[i].CellSmartTagXpgList[j].CellSmartTagPrXpgList.Count - 1 do begin
        CSTagPr := ASrc[i].CellSmartTagXpgList[j].CellSmartTagPrXpgList[k];
        CellSmartTagPr := CellSmartTag.Add;

        CellSmartTagPr.Key := CSTagPr.Key;
        CellSmartTagPr.Val := CSTagPr.Val;
      end;
    end;
  end;
end;

procedure TXLSReadXLSX.ReadSheetSortState(ASrc: TCT_SortState; ADest: TXc12SortState);
var
  i: integer;
  SortCond: TXc12SortCondition;
begin
  ADest.ColumnSort := ASrc.ColumnSort;
  ADest.CaseSensitive := ASrc.CaseSensitive;
  ADest.SortMethod := TXc12SortMethod(ASrc.SortMethod);
  ADest.Ref := AreaStrToCellArea(ASrc.Ref);

  for i := 0 to ASrc.SortConditionXpgList.Count - 1 do begin
    SortCond := ADest.SortConditions.Add;

    SortCond.Descending := ASrc.SortConditionXpgList[i].Descending;
    SortCond.SortBy := TXc12SortBy(ASrc.SortConditionXpgList[i].SortBy);
    SortCond.Ref := AreaStrToCellArea(ASrc.SortConditionXpgList[i].Ref);
    SortCond.CustomList := ASrc.SortConditionXpgList[i].CustomList;
    SortCond.DxfId := ASrc.SortConditionXpgList[i].DxfId;
    SortCond.IconSet := TXc12IconSetType(ASrc.SortConditionXpgList[i].IconSet);
    SortCond.IconId := ASrc.SortConditionXpgList[i].IconId;
  end;
end;

procedure TXLSReadXLSX.ReadSheetsThread;
{$ifdef DELPHI_XE_OR_LATER}
{$ifndef BABOON}
var
  i: integer;
  Sheet: TXc12DataWorksheet;
  Stream: TStream;
  Threads: array of TReadSheetThread;
  Handles: array [0..255] of THandle;
{$endif}
{$endif}
begin
{$ifdef DELPHI_XE_OR_LATER}
{$ifndef BABOON}
  SetLength(Threads,FManager.Workbook.Sheets.Count);
  for i := 0 to FManager.Workbook.Sheets.Count - 1 do begin
    Sheet := FManager.Worksheets.Add(FManager.Workbook.Sheets[i].SheetId - 1);

    Sheet.Name := FManager.Workbook.Sheets[i].Name;
    Sheet.State := FManager.Workbook.Sheets[i].State;
    Sheet.RId := FManager.Workbook.Sheets[i].RId;

    FManager.FileData.AddSaveSheet(Sheet.RId);

    Stream := FManager.FileData.OPC.OpenAndReadSheet(FManager.Workbook.Sheets[i].RId);

    Threads[i] := TReadSheetThread.Create(Stream,Sheet);
    if not Threads[i].IsSingleProcessor then
      Threads[i].Priority := tpHigher;
    Handles[i] := Threads[i].Handle;
  end;

  for i := 0 to FManager.Workbook.Sheets.Count - 1 do
    Threads[i].Start;

  WaitForMultipleObjects(FManager.Workbook.Sheets.Count,@Handles,True,INFINITE);

  for i := 0 to FManager.Workbook.Sheets.Count - 1 do
    Threads[i].Free;

  FManager.Workbook.Sheets.Clear;
{$endif}
{$endif}
end;

procedure TXLSReadXLSX.ReadChartSheet(const AId: AxUCString);
var
  i: integer;
  XML: TXPGDocXLSX;
  Stream: TStream;
  WebPublishItem: TXc12WebPublishItem;
  WPItem: TCT_WebPublishItem;
  CV: TCT_ChartsheetView;
  CView: TXc12SheetView;
begin
  Stream := FManager.FileData.OPC.OpenAndReadChartSheet(AId);

  try
    XML := TXPGDocXLSX.Create;
    try
      XML.Errors.OnError := HandleParseError;

      XML.LoadFromStream(Stream);

      if XML.Chartsheet.SheetPr.Available then begin
        FCurrSheet.SheetPr.CodeName := XML.Chartsheet.SheetPr.CodeName;
        FCurrSheet.SheetPr.TabColor := GetColor(XML.Chartsheet.SheetPr.TabColor);
        FCurrSheet.SheetPr.Published_ := XML.Chartsheet.SheetPr.Published_;
      end;

      if XML.Chartsheet.SheetViews.Available then begin
        FCurrSheet.SheetViews.Clear;
        for i := 0 to XML.Chartsheet.SheetViews.SheetViewXpgList.Count - 1 do begin
          CV := XML.Chartsheet.SheetViews.SheetViewXpgList[i];
          CView := FCurrSheet.SheetViews.Add;
          CView.TabSelected := CV.TabSelected;
          CView.ZoomScale := CV.ZoomScale;
          CView.WorkbookViewId := CV.WorkbookViewId;
          CView.ZoomToFit := CV.ZoomToFit;
        end;
      end;

      FCurrSheet.SheetProtection.Password := XML.Chartsheet.SheetProtection.Password;
      FCurrSheet.SheetProtection.Objects := XML.Chartsheet.SheetProtection.Objects;

//      if XML.Chartsheet.CustomSheetViews.CustomSheetViewXpgList.Count > 0 then
//        ReadSheetCustomSheetViews(XML.Chartsheet.CustomSheetViews);

      if XML.Chartsheet.PageMargins.Available then
        ReadSheetPageMargins(XML.Chartsheet.PageMargins,FCurrSheet.PageMargins);

//      if XML.Chartsheet.PageSetup.Available then
//        ReadSheetPageSetup(XML.Chartsheet.PageSetup,FCurrSheet.PageSetup);

      if XML.Chartsheet.HeaderFooter.Available then
        ReadSheetHeaderFooter(XML.Chartsheet.HeaderFooter,FCurrSheet.HeaderFooter);

      for i := 0 to XML.Chartsheet.WebPublishItems.WebPublishItemXpgList.Count - 1 do begin
        WPItem := XML.Chartsheet.WebPublishItems.WebPublishItemXpgList[i];
        WebPublishItem := FCurrSheet.WebPublishItems.Add;

        WebPublishItem.Id := WPItem.Id;
        WebPublishItem.DivId := WPItem.DivId;
        WebPublishItem.SourceType := TXc12WebSourceType(WPItem.SourceType);
        WebPublishItem.SourceRef.AsString := WPItem.SourceRef;
        WebPublishItem.SourceObject := WPItem.SourceObject;
        WebPublishItem.DestinationFile := WPItem.DestinationFile;
        WebPublishItem.Title := WPItem.Title;
        WebPublishItem.AutoRepublish := WPItem.AutoRepublish;
      end;

      if XML.Chartsheet.Drawing.Available then
        ReadDrawing(XML.Chartsheet.Drawing.R_Id);

      if XML.Chartsheet.LegacyDrawing.Available then
        ReadVmlDrawing(XML.Chartsheet.LegacyDrawing.R_Id);
    finally
      XML.Free;
    end;

  finally
    Stream.Free;
    FManager.FileData.OPC.CloseSheet;
  end;
end;

procedure TXLSReadXLSX.ReadComments(const ASheetId: AxUCString);
var
  i: integer;
  S: AxUCString;
  Anchor: AxUCString;
  Stream: TStream;
  XML: TXPGDocXLSX;
  Node,N,N2: TXpgDOMNode;
  Col,Row: integer;
  Comment: TXc12Comment;
  PPIX,PPIY: integer;
begin
  Stream := FManager.FileData.OPC.ReadComments(ASheetId);
  if Stream <> Nil then begin
    try
      XML := TXPGDocXLSX.Create;
      try
        XML.Errors.OnError := HandleParseError;
        XML.Root.Comments.CommentList.OnReadComment := OnComment;

        XML.LoadFromStream(Stream);

        for i := 0 to XML.Root.Comments.Authors.AuthorXpgList.Count - 1 do
          FCurrSheet.Comments.Authors.Add(XML.Root.Comments.Authors.AuthorXpgList[i]);

         for i := 0 to FCurrSheet.Comments.Count - 1 do
           FCurrSheet.Comments[i].Author := FCurrSheet.Comments.Authors[Integer(FCurrSheet.Comments[i].TempAuthorId)];
      finally
        XML.Free;
      end;
    finally
      Stream.Free;
    end;
  end;

  if FCurrSheet.VmlDrawing.Root.Count > 0 then begin
    FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);
    // xml
    Node := FCurrSheet.VmlDrawing.Root[0];
    i := 0;
    while i < Node.Count do begin
      Comment := Nil;

      N := Node[i];
      if N.QName = 'v:shape' then begin
        Col := -1;
        Row := -1;
        N2 := N.Find('x:ClientData/x:Column',False);
        if N2 <> Nil then
          Col := StrToIntDef(N2.Content,-1);
        N2 := N.Find('x:ClientData/x:Row',False);
        if N2 <> Nil then
          Row := StrToIntDef(N2.Content,-1);
        N2 := N.Find('x:ClientData/x:Anchor',False);
        if N2 <> Nil then
          Anchor := N2.Content;

        if (Col >= 0) and (Row >= 0) then begin
          Comment := FCurrSheet.Comments.Find(Col,Row);
          if Comment <> Nil then begin
            S := N.Attributes.FindValue('fillcolor');
            if S <> '' then
              Comment.Color := XmlStrHexToIntDef(S,XLS_COLOR_DEFAULT_COMMENT);


            Comment.VML := Node.Detach(i);

            Comment.Col1     := StrToIntDef(Trim(SplitAtChar(',',Anchor)),Col + 1);
            Comment.Col1Offs := StrToIntDef(Trim(SplitAtChar(',',Anchor)),15);
            Comment.Row1     := StrToIntDef(Trim(SplitAtChar(',',Anchor)),Max(Row - 1,0));
            Comment.Row1Offs := StrToIntDef(Trim(SplitAtChar(',',Anchor)),10);
            Comment.Col2     := StrToIntDef(Trim(SplitAtChar(',',Anchor)),Col + 3);
            Comment.Col2Offs := StrToIntDef(Trim(SplitAtChar(',',Anchor)),31);
            Comment.Row2     := StrToIntDef(Trim(SplitAtChar(',',Anchor)),Row + 3);
            Comment.Row2Offs := StrToIntDef(Trim(SplitAtChar(',',Anchor)),9);
          end;
        end;

//        // If Anchor element is deleted, excel will calculate the correct position for the comment.
//        N := N.Find('x:ClientData',False);
//        if N <> Nil then
//          N.DeleteChildFirst('x:Anchor');
      end;
      if Comment = Nil then
        Inc(i);
    end;
  end;
end;

procedure TXLSReadXLSX.ReadConnections;
var
  Stream: TStream;
begin
  Stream := FManager.FileData.OPC.ReadConnections;
  if Stream <> Nil then begin
    FManager.Workbook.Connections.LoadFromStream(Stream);
    Stream.Free;
  end;
end;

procedure TXLSReadXLSX.ReadDrawing(const AId: AxUCString);
var
  Stream: TStream;
begin
  Stream := FManager.FileData.OPC.OpenAndReadDrawing(AId);
  if Stream <> Nil then begin
    FManager.GrManager.XLSOPC := FManager.FileData.OPC;
    try
      ReadMedia(FManager.FileData.OPC.CurrSheet.Id);

      FCurrSheet.Drawing._LoadFromStream(Stream);
    finally
      Stream.Free;
      FManager.FileData.OPC.CloseDrawing;
    end;
  end;
end;

procedure TXLSReadXLSX.ReadMedia(ADrawinRId: AxUCString);
var
  i: integer;
  List: TStringList;
  Stream: TStream;
  Target: AxUCString;
begin
  List := TStringList.Create;
  try
    if FManager.FileData.OPC.GetDrawingImages(List) then begin
      for i := 0 to List.Count - 1 do begin
        Stream := FManager.FileData.OPC.ReadDrawingMedia(List[i],Target);
        if Stream <> Nil then begin
          try
            FManager.GrManager.Images.Xc12_Add(Stream,Copy(ExtractFileExt(Target),2,MAXINT),Target,ADrawinRId,List[i]);
          finally
            Stream.Free;
          end;
        end;
      end;
    end;
  finally
    List.Free;
  end;
end;

procedure TXLSReadXLSX.ReadSheet(const AId: AxUCString);
var
  i,j           : integer;
  XML           : TXPGDocXLSX;
  Stream        : TStream;
  Strm          : TStream;
  ProtectedRange: TXc12ProtectedRange;
  Scenario      : TXc12Scenario;
  Scen          : TCT_Scenario;
  InputCell     : TXc12InputCell;
  DataRef       : TXc12DataRef;
  DValidation   : TCT_DataValidation;
  DataValidation: TXc12DataValidation;
  Prop          : TXc12CustomProperty;
  IgnoredError  : TXc12IgnoredError;
  IError        : TCT_IgnoredError;
  OleObject     : TXc12OleObject;
  Control       : TXc12Control;
  WebPublishItem: TXc12WebPublishItem;
  WPItem        : TCT_WebPublishItem;
  PivotTable    : TCT_pivotTableDefinition;
  List          : TStringList;
begin
  FSharedFormulas.Clear;
  ClearArrayFormulas;

  Stream := FManager.FileData.OPC.OpenAndReadSheet(AId);
  try
    XML := TXPGDocXLSX.Create;
    try
      XML.Errors.OnError := HandleParseError;

      if FManager.DirectRead then begin
        XML.OnReadWorksheetColsCol := OnDummy;
        XML.Worksheet.SheetData.OnReadRow := OnDummy;
        XML.Worksheet.SheetData.Row.OnReadC := OnCellDirect;
        FManager.EventCell.SheetName := FCurrSheet.Name;
      end
      else begin
        XML.OnReadWorksheetColsCol := OnCol;
        XML.Worksheet.SheetData.OnReadRow := OnRow;
        XML.Worksheet.SheetData.Row.OnReadC := OnCell;
      end;

      XML.Worksheet.MergeCells.OnReadMergeCell := OnMergeCell;
      XML.Worksheet.Hyperlinks.OnReadHyperlink := OnHyperlink;
      XML.LoadFromStream(Stream);

      if XML.Worksheet.SheetPr.Available then begin
        FCurrSheet.SheetPr.SyncHorizontal := XML.Worksheet.SheetPr.SyncHorizontal;
        FCurrSheet.SheetPr.SyncVertical := XML.Worksheet.SheetPr.SyncVertical;
        FCurrSheet.SheetPr.SyncRef := AreaStrToCellArea(XML.Worksheet.SheetPr.SyncRef);
        FCurrSheet.SheetPr.TransitionEvaluation := XML.Worksheet.SheetPr.TransitionEvaluation;
        FCurrSheet.SheetPr.TransitionEntry := XML.Worksheet.SheetPr.TransitionEntry;
        FCurrSheet.SheetPr.Published_ := XML.Worksheet.SheetPr.Published_;
        FCurrSheet.SheetPr.CodeName := XML.Worksheet.SheetPr.CodeName;
        FCurrSheet.SheetPr.FilterMode := XML.Worksheet.SheetPr.FilterMode;
        FCurrSheet.SheetPr.EnableFormatConditionsCalculation := XML.Worksheet.SheetPr.EnableFormatConditionsCalculation;
        FCurrSheet.SheetPr.TabColor := GetColor(XML.Worksheet.SheetPr.TabColor);
        if XML.Worksheet.SheetPr.OutlinePr.Available then begin
          FCurrSheet.SheetPr.OutlinePr.ApplyStyles := XML.Worksheet.SheetPr.OutlinePr.ApplyStyles;
          FCurrSheet.SheetPr.OutlinePr.SummaryBelow := XML.Worksheet.SheetPr.OutlinePr.SummaryBelow;
          FCurrSheet.SheetPr.OutlinePr.SummaryRight := XML.Worksheet.SheetPr.OutlinePr.SummaryRight;
          FCurrSheet.SheetPr.OutlinePr.ShowOutlineSymbols := XML.Worksheet.SheetPr.OutlinePr.ShowOutlineSymbols;
        end;
        if XML.Worksheet.SheetPr.PageSetUpPr.Available then begin
          FCurrSheet.SheetPr.PageSetUpPr.AutoPageBreaks := XML.Worksheet.SheetPr.PageSetUpPr.AutoPageBreaks;
          FCurrSheet.SheetPr.PageSetUpPr.FitToPage := XML.Worksheet.SheetPr.PageSetUpPr.FitToPage;
        end;
      end;

      FCurrSheet.Dimension := AreaStrToCellArea(XML.Worksheet.Dimension.Ref);

      if XML.Worksheet.SheetViews.Available then
        ReadSheetView(XML.Worksheet.SheetViews);

      FCurrSheet.SheetFormatPr.BaseColWidth := XML.Worksheet.SheetFormatPr.BaseColWidth;
      FCurrSheet.SheetFormatPr.DefaultColWidth := XML.Worksheet.SheetFormatPr.DefaultColWidth;
      FCurrSheet.Columns.DefColWidth := XML.Worksheet.SheetFormatPr.DefaultColWidth;
      FCurrSheet.SheetFormatPr.DefaultRowHeight := XML.Worksheet.SheetFormatPr.DefaultRowHeight * 20;
      FCurrSheet.SheetFormatPr.CustomHeight := XML.Worksheet.SheetFormatPr.CustomHeight;
      FCurrSheet.SheetFormatPr.ZeroHeight := XML.Worksheet.SheetFormatPr.ZeroHeight;
      FCurrSheet.SheetFormatPr.ThickTop := XML.Worksheet.SheetFormatPr.ThickTop;
      FCurrSheet.SheetFormatPr.ThickBottom := XML.Worksheet.SheetFormatPr.ThickBottom;
      FCurrSheet.SheetFormatPr.OutlineLevelRow := XML.Worksheet.SheetFormatPr.OutlineLevelRow;
      FCurrSheet.SheetFormatPr.OutlineLevelCol := XML.Worksheet.SheetFormatPr.OutlineLevelCol;

      FCurrSheet.SheetCalcPr.FullCalcOnLoad := XML.Worksheet.SheetCalcPr.FullCalcOnLoad;

      FCurrSheet.SheetProtection.Password := XML.Worksheet.SheetProtection.Password;
      FCurrSheet.SheetProtection.Sheet := XML.Worksheet.SheetProtection.Sheet;
      FCurrSheet.SheetProtection.Objects := XML.Worksheet.SheetProtection.Objects;
      FCurrSheet.SheetProtection.Scenarios := XML.Worksheet.SheetProtection.Scenarios;
      FCurrSheet.SheetProtection.FormatCells := XML.Worksheet.SheetProtection.FormatCells;
      FCurrSheet.SheetProtection.FormatColumns := XML.Worksheet.SheetProtection.FormatColumns;
      FCurrSheet.SheetProtection.FormatRows := XML.Worksheet.SheetProtection.FormatRows;
      FCurrSheet.SheetProtection.InsertColumns := XML.Worksheet.SheetProtection.InsertColumns;
      FCurrSheet.SheetProtection.InsertRows := XML.Worksheet.SheetProtection.InsertRows;
      FCurrSheet.SheetProtection.InsertHyperlinks := XML.Worksheet.SheetProtection.InsertHyperlinks;
      FCurrSheet.SheetProtection.DeleteColumns := XML.Worksheet.SheetProtection.DeleteColumns;
      FCurrSheet.SheetProtection.DeleteRows := XML.Worksheet.SheetProtection.DeleteRows;
      FCurrSheet.SheetProtection.SelectLockedCells := XML.Worksheet.SheetProtection.SelectLockedCells;
      FCurrSheet.SheetProtection.Sort := XML.Worksheet.SheetProtection.Sort;
      FCurrSheet.SheetProtection.AutoFilter := XML.Worksheet.SheetProtection.AutoFilter;
      FCurrSheet.SheetProtection.PivotTables := XML.Worksheet.SheetProtection.PivotTables;
      FCurrSheet.SheetProtection.SelectUnlockedCells := XML.Worksheet.SheetProtection.SelectUnlockedCells;

      for i := 0 to XML.Worksheet.ProtectedRanges.ProtectedRangeXpgList.Count - 1 do begin
        ProtectedRange := FCurrSheet.ProtectedRanges.Add;

        ProtectedRange.Password := XML.Worksheet.ProtectedRanges.ProtectedRangeXpgList[i].Password;
        ProtectedRange.Sqref.DelimitedText := XML.Worksheet.ProtectedRanges.ProtectedRangeXpgList[i].SqrefXpgList.DelimitedText;
        ProtectedRange.Name := XML.Worksheet.ProtectedRanges.ProtectedRangeXpgList[i].Name;
        ProtectedRange.SecurityDescriptor := XML.Worksheet.ProtectedRanges.ProtectedRangeXpgList[i].SecurityDescriptor;
      end;

      if XML.Worksheet.Scenarios.ScenarioXpgList.Count > 0 then begin
        FCurrSheet.Scenarios.Current := XML.Worksheet.Scenarios.Current;
        FCurrSheet.Scenarios.Show := XML.Worksheet.Scenarios.Show;
        FCurrSheet.Scenarios._Sqref.DelimitedText := XML.Worksheet.Scenarios.SqrefXpgList.DelimitedText;

        for i := 0 to XML.Worksheet.Scenarios.ScenarioXpgList.Count - 1 do begin
          Scen := XML.Worksheet.Scenarios.ScenarioXpgList[i];
          Scenario := FCurrSheet.Scenarios.Add;

          Scenario.Name := Scen.Name;
          Scenario.Locked := Scen.Locked;
          Scenario.Hidden := Scen.Hidden;
          Scenario.User := Scen.User;
          Scenario.Comment := Scen.Comment;

          for j := 0 to Scen.InputCellsXpgList.Count - 1 do begin
            InputCell := Scenario.InputCells.Add(Scen.InputCellsXpgList[j].R);
            InputCell.Deleted := Scen.InputCellsXpgList[j].Deleted;
            InputCell.Undone := Scen.InputCellsXpgList[j].Undone;
            InputCell.Val := Scen.InputCellsXpgList[j].Val;
            InputCell.NumFmtId := Scen.InputCellsXpgList[j].NumFmtId;
          end;
        end;
      end;

      if XML.Worksheet.AutoFilter.Available then
        ReadSheetAutoFilter(XML.Worksheet.AutoFilter,FCurrSheet.AutoFilter);

      if XML.Worksheet.SortState.Available then
        ReadSheetSortState(XML.Worksheet.SortState,FCurrSheet.SortState);

      if XML.Worksheet.DataConsolidate.Available then begin
        FCurrSheet.DataConsolidate.Function_ := TXc12DataConsolidateFunction(XML.Worksheet.DataConsolidate.Function_);
        FCurrSheet.DataConsolidate.LeftLabels := XML.Worksheet.DataConsolidate.LeftLabels;
        FCurrSheet.DataConsolidate.TopLabels := XML.Worksheet.DataConsolidate.TopLabels;
        FCurrSheet.DataConsolidate.Link := XML.Worksheet.DataConsolidate.Link;

        for i := 0 to XML.Worksheet.DataConsolidate.DataRefs.DataRefXpgList.Count - 1 do begin
          DataRef := FCurrSheet.DataConsolidate.DataRefs.Add;

          DataRef.Ref := AreaStrToCellArea(XML.Worksheet.DataConsolidate.DataRefs.DataRefXpgList[i].Ref);
          DataRef.Name := XML.Worksheet.DataConsolidate.DataRefs.DataRefXpgList[i].Name;
          DataRef.Sheet := XML.Worksheet.DataConsolidate.DataRefs.DataRefXpgList[i].Sheet;
          DataRef.RId := XML.Worksheet.DataConsolidate.DataRefs.DataRefXpgList[i].R_Id;
        end;
      end;

      if XML.Worksheet.CustomSheetViews.CustomSheetViewXpgList.Count > 0 then
        ReadSheetCustomSheetViews(XML.Worksheet.CustomSheetViews);

      if XML.Worksheet.PhoneticPr.Available then begin
        FCurrSheet.PhoneticPr.FontId := XML.Worksheet.PhoneticPr.FontId;
        FCurrSheet.PhoneticPr.Type_ := TXc12PhoneticType(XML.Worksheet.PhoneticPr.Type_);
        FCurrSheet.PhoneticPr.Alignment := TXc12PhoneticAlignment(XML.Worksheet.PhoneticPr.Alignment);
      end;

      for i := 0 to XML.Worksheet.ConditionalFormattingXpgList.Count - 1 do
        ReadSheetCondFmt(XML.Worksheet.ConditionalFormattingXpgList[i],FCurrSheet.ConditionalFormatting.AddCF);

      if XML.Worksheet.DataValidations.Available then begin
        FCurrSheet.DataValidations.DisablePrompts := XML.Worksheet.DataValidations.DisablePrompts;
        FCurrSheet.DataValidations.XWindow := XML.Worksheet.DataValidations.XWindow;
        FCurrSheet.DataValidations.YWindow := XML.Worksheet.DataValidations.YWindow;

        for i := 0 to XML.Worksheet.DataValidations.DataValidationXpgList.Count - 1 do begin
          DValidation := XML.Worksheet.DataValidations.DataValidationXpgList[i];
          DataValidation := FCurrSheet.DataValidations.AddDV;

          DataValidation.Type_ := TXc12DataValidationType(DValidation.Type_);
          DataValidation.ErrorStyle := TXc12DataValidationErrorStyle(DValidation.ErrorStyle);
          DataValidation.ImeMode := TXc12DataValidationImeMode(DValidation.ImeMode);
          DataValidation.Operator_ := TXc12DataValidationOperator(DValidation.Operator_);
          DataValidation.AllowBlank := DValidation.AllowBlank;
          DataValidation.ShowDropDown := DValidation.ShowDropDown;
          DataValidation.ShowInputMessage := DValidation.ShowInputMessage;
          DataValidation.ShowErrorMessage := DValidation.ShowErrorMessage;
          DataValidation.ErrorTitle := DValidation.ErrorTitle;
          DataValidation.Error := DValidation.Error;
          DataValidation.PromptTitle := DValidation.PromptTitle;
          DataValidation.Prompt := DValidation.Prompt;
          DataValidation.Sqref.DelimitedText := DValidation.SqrefXpgList.DelimitedText;
          DataValidation.Formula1 := DValidation.Formula1;
          DataValidation.Formula2 := DValidation.Formula2;
        end;
      end;

      if XML.Worksheet.PrintOptions.Available then
        ReadSheetPrintOptions(XML.Worksheet.PrintOptions,FCurrSheet.PrintOptions);

      if XML.Worksheet.PageMargins.Available then
        ReadSheetPageMargins(XML.Worksheet.PageMargins,FCurrSheet.PageMargins);

      if XML.Worksheet.PageSetup.Available then
        ReadSheetPageSetup(XML.Worksheet.PageSetup,FCurrSheet.PageSetup);

      if XML.Worksheet.HeaderFooter.Available then
        ReadSheetHeaderFooter(XML.Worksheet.HeaderFooter,FCurrSheet.HeaderFooter);

      if XML.Worksheet.RowBreaks.Available then
        ReadSheetPageBreak(XML.Worksheet.RowBreaks,FCurrSheet.RowBreaks);

      if XML.Worksheet.ColBreaks.Available then
        ReadSheetPageBreak(XML.Worksheet.ColBreaks,FCurrSheet.ColBreaks);

      for i := 0 to XML.Worksheet.CustomProperties.CustomPrXpgList.Count - 1 do begin
        Prop := FCurrSheet.CustomProperties.Add;

        Prop.Name := XML.Worksheet.CustomProperties.CustomPrXpgList[i].Name;
        Prop.RId := XML.Worksheet.CustomProperties.CustomPrXpgList[i].R_Id;
      end;

      for i := 0 to XML.Worksheet.CellWatches.CellWatchXpgList.Count - 1 do
        FCurrSheet.CellWatches.Add(XML.Worksheet.CellWatches.CellWatchXpgList[i].R);

      for i := 0 to XML.Worksheet.IgnoredErrors.IgnoredErrorXpgList.Count - 1 do begin
        IError := XML.Worksheet.IgnoredErrors.IgnoredErrorXpgList[i];
        IgnoredError := FCurrSheet.IgnoredErrors.Add;

        IgnoredError.Sqref.DelimitedText := IError.SqrefXpgList.DelimitedText;
        IgnoredError.EvalError := IError.EvalError;
        IgnoredError.TwoDigitTextYear := IError.TwoDigitTextYear;
        IgnoredError.NumberStoredAsText := IError.NumberStoredAsText;
        IgnoredError.Formula := IError.Formula;
        IgnoredError.FormulaRange := IError.FormulaRange;
        IgnoredError.UnlockedFormula := IError.UnlockedFormula;
        IgnoredError.EmptyCellReference := IError.EmptyCellReference;
        IgnoredError.ListDataValidation := IError.ListDataValidation;
        IgnoredError.CalculatedColumn := IError.CalculatedColumn;
      end;

      if XML.Worksheet.SmartTags.CellSmartTagsXpgList.Count > 0 then
        ReadSheetSmartTags(XML.Worksheet.SmartTags.CellSmartTagsXpgList,FCurrSheet.SmartTags);

      for i := 0 to XML.Worksheet.OleObjects.OleObjectXpgList.Count - 1 do begin
        OleObject := FCurrSheet.OleObjects.Add;

        OleObject.ProgId := XML.Worksheet.OleObjects.OleObjectXpgList[i].ProgId;
        OleObject.DvAspect := TXc12DvAspect(XML.Worksheet.OleObjects.OleObjectXpgList[i].DvAspect);
        OleObject.Link := XML.Worksheet.OleObjects.OleObjectXpgList[i].Link;
        OleObject.OleUpdate := TXc12OleUpdate(XML.Worksheet.OleObjects.OleObjectXpgList[i].OleUpdate);
        OleObject.AutoLoad := XML.Worksheet.OleObjects.OleObjectXpgList[i].AutoLoad;
        OleObject.ShapeId := XML.Worksheet.OleObjects.OleObjectXpgList[i].ShapeId;
        OleObject.RId := XML.Worksheet.OleObjects.OleObjectXpgList[i].R_Id;
      end;

      for i := 0 to XML.Worksheet.Controls.ControlXpgList.Count - 1 do begin
        Control := FCurrSheet.Controls.Add;

        Control.ShapeId := XML.Worksheet.Controls.ControlXpgList[i].ShapeId;
        Control.RId := XML.Worksheet.Controls.ControlXpgList[i].R_Id;
        Control.Name := XML.Worksheet.Controls.ControlXpgList[i].Name;
      end;

      for i := 0 to XML.Worksheet.WebPublishItems.WebPublishItemXpgList.Count - 1 do begin
        WPItem := XML.Worksheet.WebPublishItems.WebPublishItemXpgList[i];
        WebPublishItem := FCurrSheet.WebPublishItems.Add;

        WebPublishItem.Id := WPItem.Id;
        WebPublishItem.DivId := WPItem.DivId;
        WebPublishItem.SourceType := TXc12WebSourceType(WPItem.SourceType);
        WebPublishItem.SourceRef.AsString := WPItem.SourceRef;
        WebPublishItem.SourceObject := WPItem.SourceObject;
        WebPublishItem.DestinationFile := WPItem.DestinationFile;
        WebPublishItem.Title := WPItem.Title;
        WebPublishItem.AutoRepublish := WPItem.AutoRepublish;
      end;

      for i := 0 to XML.Worksheet.TableParts.TablePartXpgList.Count - 1 do
        ReadTable(XML.Worksheet.TableParts.TablePartXpgList[i].R_Id);

      if XML.Worksheet.Drawing.Available then
        ReadDrawing(XML.Worksheet.Drawing.R_Id);

      if XML.Worksheet.LegacyDrawing.Available then
        ReadVmlDrawing(XML.Worksheet.LegacyDrawing.R_Id);
    finally
      XML.Free;
    end;

    List := TStringList.Create;
    try
      FManager.FileData.OPC.GetSheetPivotTables(List);
      for i := 0 to List.Count - 1 do begin
        Strm := FManager.FileData.OPC.ReadPivotTable(List[i]);
        try
          PivotTable := FCurrSheet.PivotTables.LoadFromStream(Strm);
          if PivotTable <> Nil then begin
            PivotTable.Cache := FManager.Workbook.PivotCaches.Find(PivotTable.CacheId);
            if PivotTable.Cache <> Nil then
              PivotTable.Cache.Use;
          end;
        finally
          Strm.Free;
        end;
      end;
    finally
      List.Free;
    end;

  finally
    Stream.Free;
    FManager.FileData.OPC.CloseSheet;
  end;

  if not FCurrSheet.IsChartSheet then
    ReadComments(AId);
end;

procedure TXLSReadXLSX.ReadSheetView(ASheetViews: TCT_SheetViews);
var
  i,j,k,l: integer;
  SView: TCT_SheetView;
  SheetView: TXc12SheetView;
  Selection: TXc12Selection;
  PSelection: TCT_PivotSelection;
  PivotSelection: TXc12PivotSelection;
  Reference: TXc12PivotAreaReference;
begin
  FCurrSheet.SheetViews.Clear;
  for i := 0 to ASheetViews.SheetViewXpgList.Count - 1 do begin
    SView := ASheetViews.SheetViewXpgList[i];
    SheetView := FCurrSheet.SheetViews.Add;

    SheetView.WindowProtection := SView.WindowProtection;
    SheetView.ShowFormulas := SView.ShowFormulas;
    SheetView.ShowGridLines := SView.ShowGridLines;
    SheetView.ShowRowColHeaders := SView.ShowRowColHeaders;
    SheetView.ShowZeros := SView.ShowZeros;
    SheetView.RightToLeft := SView.RightToLeft;
    SheetView.TabSelected := SView.TabSelected;
    SheetView.ShowRuler := SView.ShowRuler;
    SheetView.ShowOutlineSymbols := SView.ShowOutlineSymbols;
    SheetView.DefaultGridColor := SView.DefaultGridColor;
    SheetView.ShowWhiteSpace := SView.ShowWhiteSpace;
    SheetView.View := TXc12SheetViewType(SView.View);
    SheetView.TopLeftCell := AreaStrToCellArea(SView.TopLeftCell);
    SheetView.ColorId := SView.ColorId;
    SheetView.ZoomScale := SView.ZoomScale;
    SheetView.ZoomScaleNormal := SView.ZoomScaleNormal;
    SheetView.ZoomScaleSheetLayoutView := SView.ZoomScaleSheetLayoutView;
    SheetView.ZoomScalePageLayoutView := SView.ZoomScalePageLayoutView;
    SheetView.WorkbookViewId := SView.WorkbookViewId;

    if SView.Pane.Available then
      ReadSheetPane(SView.Pane,SheetView.Pane);

    for j := 0 to SView.SelectionXpgList.Count - 1 do begin
      Selection := SheetView.Selection.Add;
      ReadSheetSelection(SView.SelectionXpgList[j],Selection);
    end;

    for j := 0 to SView.PivotSelectionXpgList.Count - 1 do begin
      PSelection := SView.PivotSelectionXpgList[j];
      PivotSelection := SheetView.PivotSelection.Add;

      PivotSelection.Pane := TXc12PaneEnum(PSelection.Pane);
      PivotSelection.ShowHeader := PSelection.ShowHeader;
      PivotSelection.Label_ := PSelection.Label_;
      PivotSelection.Data := PSelection.Data;
      PivotSelection.Extendable := PSelection.Extendable;
      PivotSelection.Axis := TXc12Axis(PSelection.Axis);
      PivotSelection.Dimension := PSelection.Dimension;
      PivotSelection.Start := PSelection.Start;
      PivotSelection.Min := PSelection.Min;
      PivotSelection.Max := PSelection.Max;
      PivotSelection.ActiveRow := PSelection.ActiveRow;
      PivotSelection.ActiveCol := PSelection.ActiveCol;
      PivotSelection.PreviousRow := PSelection.PreviousRow;
      PivotSelection.PreviousCol := PSelection.PreviousCol;
      PivotSelection.Click := PSelection.Click;
      PivotSelection.RId := PSelection.R_Id;

      PivotSelection.PivotArea.Field := PSelection.PivotArea.Field;
      PivotSelection.PivotArea.Type_ := TXc12PivotAreaType(PSelection.PivotArea.Type_);
      PivotSelection.PivotArea.DataOnly := PSelection.PivotArea.DataOnly;
      PivotSelection.PivotArea.LabelOnly := PSelection.PivotArea.LabelOnly;
      PivotSelection.PivotArea.GrandRow := PSelection.PivotArea.GrandRow;
      PivotSelection.PivotArea.GrandCol := PSelection.PivotArea.GrandCol;
      PivotSelection.PivotArea.CacheIndex := PSelection.PivotArea.CacheIndex;
      PivotSelection.PivotArea.Outline := PSelection.PivotArea.Outline;
      PivotSelection.PivotArea.Offset := AreaStrToCellArea(PSelection.PivotArea.Offset);
      PivotSelection.PivotArea.CollapsedLevelsAreSubtotals := PSelection.PivotArea.CollapsedLevelsAreSubtotals;
      PivotSelection.PivotArea.Axis := TXc12Axis(PSelection.PivotArea.Axis);
      PivotSelection.PivotArea.FieldPosition := PSelection.PivotArea.FieldPosition;

      for k := 0 to PSelection.PivotArea.References.ReferenceXpgList.Count -1 do begin
        Reference := PivotSelection.PivotArea.References.Add;

        Reference.Field := PSelection.PivotArea.References.ReferenceXpgList[k].Field;
        Reference.Selected := PSelection.PivotArea.References.ReferenceXpgList[k].Selected;
        Reference.ByPosition := PSelection.PivotArea.References.ReferenceXpgList[k].ByPosition;
        Reference.Relative := PSelection.PivotArea.References.ReferenceXpgList[k].Relative;
        Reference.DefaultSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].DefaultSubtotal;
        Reference.SumSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].SumSubtotal;
        Reference.CountASubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].CountASubtotal;
        Reference.AvgSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].AvgSubtotal;
        Reference.MaxSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].MaxSubtotal;
        Reference.MinSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].MinSubtotal;
        Reference.ProductSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].ProductSubtotal;
        Reference.CountSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].CountSubtotal;
        Reference.StdDevSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].StdDevSubtotal;
        Reference.StdDevPSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].StdDevPSubtotal;
        Reference.VarSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].VarSubtotal;
        Reference.VarPSubtotal := PSelection.PivotArea.References.ReferenceXpgList[k].VarPSubtotal;

        for l := 0 to PSelection.PivotArea.References.ReferenceXpgList[k].XXpgList.Count - 1 do
          Reference.Values.Add(PSelection.PivotArea.References.ReferenceXpgList[k].XXpgList[l].V);
      end;
    end;
  end;
end;

procedure TXLSReadXLSX.ReadSheets;
var
  i: integer;
begin
  for i := 0 to FManager.Workbook.Sheets.Count - 1 do begin
    FCurrSheet := FManager.Worksheets.Add(FManager.Workbook.Sheets[i].SheetId - 1);

    FCurrSheet.Name := FManager.Workbook.Sheets[i].Name;
    FCurrSheet.State := FManager.Workbook.Sheets[i].State;
    FCurrSheet.RId := FManager.Workbook.Sheets[i].RId;
  end;

  for i := 0 to FManager.Workbook.Sheets.Count - 1 do begin
    FCurrSheet := FManager.Worksheets[i];
    FManager.FileData.AddSaveSheet(FCurrSheet.RId);
    FCurrSheet.IsChartSheet := FManager.FileData.OPC.SheetIsChartSheet(FManager.Workbook.Sheets[i].RId);
    if FCurrSheet.IsChartSheet then
      ReadChartSheet(FManager.Workbook.Sheets[i].RId)
    else
      ReadSheet(FManager.Workbook.Sheets[i].RId);
  end;
  FManager.Workbook.Sheets.Clear;
end;

procedure TXLSReadXLSX.ReadSST;
var
  XML: TXPGDocXLSX;
  Stream: TStream;
begin
  Stream := FManager.FileData.OPC.ReadSST;
  if Stream = Nil then
    Exit;
  try
    XML := TXPGDocXLSX.Create;
    try
      XML.Errors.OnError := HandleParseError;
      XML.Root.Sst.OnReadSi := OnSSTItem;
      XML.LoadFromStream(Stream);
    finally
      XML.Free;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSReadXLSX.ReadStyleBorders(AList: TCT_BorderXpgList);
var
  i,n: integer;
  Border: TXc12Border;
begin
  n := FManager.StyleSheet.Borders.DefaultCount;
  for i := 0 to n - 1 do begin
    Border := FManager.StyleSheet.Borders[i];
    Border.Clear;
    GetBorders(AList[i],Border);
  end;

  for i := n to AList.Count - 1 do begin
    Border := FManager.StyleSheet.Borders.Add;
    GetBorders(AList[i],Border);
  end;
end;

procedure TXLSReadXLSX.ReadStyleCellStyles(AList: TCT_CellStyleXpgList);
var
  i: integer;
  S: TCT_CellStyle;
  Style: TXc12CellStyle;
begin
   for i := 0 to AList.Count - 1 do begin
     S := AList[i];
//   Don't use this. There can be duplicate style names.
//     Style := FManager.StyleSheet.Styles.AddOrGet(S.Name);

     Style := FManager.StyleSheet.Styles.Add;
     Style.Name := S.Name;

     Style.BuiltInId := S.BuiltinId;
     Style.CustomBuiltIn := S.CustomBuiltin;
     Style.Hidden := S.Hidden;
     Style.Level := S.ILevel;
     Style.XF := FManager.StyleSheet.StyleXFs[S.XfId];

     FManager.StyleSheet.XFEditor.UseStyleStyle(Style.XF.Index);
   end;
end;

procedure TXLSReadXLSX.ReadStyleColors(AColors: TCT_Colors);
var
  i: integer;
begin
  for i := 0 to AColors.IndexedColors.RgbColorXpgList.Count - 1 do begin
    FManager.StyleSheet.InxColors.Add.RGB := AColors.IndexedColors.RgbColorXpgList[i].Rgb;
    Xc12IndexColorPalette[i] := AColors.IndexedColors.RgbColorXpgList[i].Rgb;
  end;
  for i := 0 to AColors.MruColors.ColorXpgList.Count - 1 do
    FManager.StyleSheet.MruColors.Add(GetColor(AColors.MruColors.ColorXpgList[i]));
end;

procedure TXLSReadXLSX.ReadStyleDxfs(AList: TCT_DxfXpgList);
var
  i: integer;
  D: TCT_Dxf;
  DXF: TXc12DXF;
  NumFmt: TXc12NumberFormat;
begin
  for i := 0 to AList.Count - 1 do begin
    D := AList[i];
    DXF := FManager.StyleSheet.DXFs.Add;

    if D.Font.Available then
      GetFont(D.Font,DXF.AddFont);
    if D.NumFmt.Available then begin
      NumFmt := DXF.AddNumFmt;
      NumFmt.Value := D.NumFmt.FormatCode;
    end;
    if D.Fill.Available then
      GetFill(D.Fill,DXF.AddFill);
    if D.Alignment.Available then
      GetAlignment(D.Alignment,DXF.AddAlignment);
    if D.Border.Available then
      GetBorders(D.Border,DXF.AddBorder);
    if D.Protection.Available then
      DXF.Protection := GetProtections(D.Protection);
  end;
end;

procedure TXLSReadXLSX.ReadStyleFills(AList: TCT_FillXpgList);
var
  i,n: integer;
  Fill: TXc12Fill;
begin
  n := FManager.StyleSheet.Fills.DefaultCount;
  for i := 0 to n - 1 do begin
    Fill := FManager.StyleSheet.Fills[i];
    Fill.Clear;
    GetFill(AList[i],Fill);
  end;

  for i := n to AList.Count - 1 do begin
    Fill := FManager.StyleSheet.Fills.Add;
    GetFill(AList[i],Fill);
  end;
end;

procedure TXLSReadXLSX.ReadStyleFonts(AList: TCT_FontXpgList);
var
  i,n: integer;
  Font: TXc12Font;
begin
  n := FManager.StyleSheet.Fonts.DefaultCount;
  for i := 0 to n - 1 do begin
    Font := FManager.StyleSheet.Fonts[i];
    Font.Clear;
    GetFont(AList[i],Font);
  end;

  for i := n to AList.Count - 1 do begin
    Font := FManager.StyleSheet.Fonts.Add;
    GetFont(AList[i],Font);
  end;
end;

procedure TXLSReadXLSX.ReadStyleNumFmts(AList: TCT_NumFmtXpgList);
var
  i: integer;
begin
  for i := 0 to AList.Count - 1 do
    FManager.StyleSheet.NumFmts.Add(AList[i].FormatCode,AList[i].NumFmtId);
end;

procedure TXLSReadXLSX.ReadStyles;
var
  XML: TXPGDocXLSX;
  Stream: TStream;
begin
  Stream := FManager.FileData.OPC.ReadStyles;
  if Stream = Nil then
    Exit;
  try
    XML := TXPGDocXLSX.Create;
    try
      XML.Errors.OnError := HandleParseError;
      XML.LoadFromStream(Stream);

      ReadStyleNumFmts(XML.StyleSheet.NumFmts.NumFmtXpgList);
      ReadStyleFonts(XML.StyleSheet.Fonts.FontXpgList);
      ReadStyleFills(XML.StyleSheet.Fills.FillXpgList);
      ReadStyleBorders(XML.StyleSheet.Borders.BorderXpgList);
      ReadStyleXFs(XML.StyleSheet.CellStyleXfs.XfXpgList,FManager.StyleSheet.StyleXFs);
      ReadStyleXFs(XML.StyleSheet.CellXfs.XfXpgList,FManager.StyleSheet.XFs);
      ReadStyleCellStyles(XML.StyleSheet.CellStyles.CellStyleXpgList);
      ReadStyleDxfs(XML.StyleSheet.Dxfs.DxfXpgList);
      ReadStyleTableStyles(XML.StyleSheet.TableStyles);
      ReadStyleColors(XML.StyleSheet.Colors);
    finally
      XML.Free;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSReadXLSX.ReadStyleTableStyles(AList: TCT_TableStyles);
var
  i,j: integer;
  T: TCT_TableStyle;
  Style: TXc12TableStyle;
  StyleElement: TXc12TableStyleElement;
begin
  FManager.StyleSheet.TableStyles.DefaultTableStyle := AList.DefaultTableStyle;
  FManager.StyleSheet.TableStyles.DefaultPivotStyle := AList.DefaultPivotStyle;
  for i := 0 to AList.TableStyleXpgList.Count - 1 do begin
    T := AList.TableStyleXpgList[i];
    Style := FManager.StyleSheet.TableStyles.Add;
    Style.Name := T.Name;
    Style.Pivot := T.Pivot;
    Style.Table := T.Table;
    for j := 0 to T.TableStyleElementXpgList.Count - 1 do begin
      StyleElement := Style.TableStyleElements.Add;
      StyleElement.Type_ := TXc12TableStyleType(T.TableStyleElementXpgList[j].Type_);
      StyleElement.Size :=  T.TableStyleElementXpgList[j].Size;
      StyleElement.DXF := FManager.StyleSheet.DXFs[T.TableStyleElementXpgList[j].DxfId];
    end;
  end;
end;

procedure TXLSReadXLSX.ReadStyleXFs(AList: TCT_XfXpgList; ADest: TXc12XFs);
var
  i,n: integer;
  X: TCT_Xf;
  XF: TXc12XF;
begin
  n := ADest.DefaultCount;
  for i := 0 to AList.Count - 1 do begin
    X := AList[i];
    if i < n then
      XF := ADest[i]
    else
      XF := ADest.Add;

    XF.Apply := GetXfApply(X);
    XF.Fill := FManager.StyleSheet.Fills[X.FillId];
    XF.Border := FManager.StyleSheet.Borders[X.BorderId];
    XF.Font := FManager.StyleSheet.Fonts[X.FontId];
    XF.NumFmt := FManager.StyleSheet.NumFmts[X.NumFmtId];
    XF.XF := FManager.StyleSheet.StyleXFs[X.XfId];
    if X.Alignment.Available then
      GetAlignment(X.Alignment,XF.Alignment);
    if X.Protection.Available then
      XF.Protection := GetProtections(X.Protection);
    XF.QuotePrefix := X.QuotePrefix;
    XF.PivotButton := X.PivotButton;
  end;
end;

procedure TXLSReadXLSX.ReadTable(const AId: AxUCString);
var
  i: integer;
  T: TXc12Table;
  TC: TXc12TableColumn;
  Stream: TStream;
  XML: TXPGDocXLSX;
begin
  Stream := FManager.FileData.OPC.ReadTable(AId);
  if Stream <> Nil then begin
    try
      XML := TXPGDocXLSX.Create;
      try
        XML.Errors.OnError := HandleParseError;
        XML.LoadFromStream(Stream);
        T := FCurrSheet.Tables.Add;

        T.Id := XML.Table.Id;
        T.Name := XML.Table.Name;
        T.DisplayName := XML.Table.DisplayName;
        T.Comment := XML.Table.Comment;
        T.Ref := AreaStrToCellArea(XML.Table.Ref);
        T.TableType := TXc12TableType(XML.Table.TableType);
        T.HeaderRowCount := XML.Table.HeaderRowCount;
        T.InsertRow := XML.Table.InsertRow;
        T.InsertRowShift := XML.Table.InsertRowShift;
        T.TotalsRowCount := XML.Table.TotalsRowCount;
        T.TotalsRowShown := XML.Table.TotalsRowShown;
        T.Published_ := XML.Table.Published_;
        T.HeaderRowDxfId := XML.Table.HeaderRowDxfId;
        T.DataDxfId := XML.Table.DataDxfId;
        T.TotalsRowDxfId := XML.Table.TotalsRowDxfId;
        T.HeaderRowBorderDxfId := XML.Table.HeaderRowBorderDxfId;
        T.TableBorderDxfId := XML.Table.TableBorderDxfId;
        T.TotalsRowBorderDxfId := XML.Table.TotalsRowBorderDxfId;
        T.HeaderRowCellStyle := XML.Table.HeaderRowCellStyle;
        T.DataCellStyle := XML.Table.DataCellStyle;
        T.TotalsRowCellStyle := XML.Table.TotalsRowCellStyle;
        T.ConnectionId := XML.Table.ConnectionId;

        if XML.Table.AutoFilter.Available then
          ReadSheetAutoFilter(XML.Table.AutoFilter,T.AutoFilter);

        if XML.Table.SortState.Available then
          ReadSheetSortState(XML.Table.SortState,T.SortState);

        for i := 0 to XML.Table.TableColumns.TableColumnXpgList.Count - 1 do begin
          TC := T.TableColumns.Add;
          TC.Id := XML.Table.TableColumns.TableColumnXpgList[i].Id;
          TC.UniqueName := XML.Table.TableColumns.TableColumnXpgList[i].UniqueName;
          TC.Name := XML.Table.TableColumns.TableColumnXpgList[i].Name;
          TC.TotalsRowFunction := TXc12TotalsRowFunction(XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowFunction);
          TC.TotalsRowLabel := XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowLabel;
          TC.QueryTableFieldId := XML.Table.TableColumns.TableColumnXpgList[i].QueryTableFieldId;
          TC.HeaderRowDxfId := XML.Table.TableColumns.TableColumnXpgList[i].HeaderRowDxfId;
          TC.DataDxfId := XML.Table.TableColumns.TableColumnXpgList[i].DataDxfId;
          TC.TotalsRowDxfId := XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowDxfId;
          TC.HeaderRowCellStyle := XML.Table.TableColumns.TableColumnXpgList[i].HeaderRowCellStyle;
          TC.DataCellStyle := XML.Table.TableColumns.TableColumnXpgList[i].DataCellStyle;
          TC.TotalsRowCellStyle := XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowCellStyle;

          if XML.Table.TableColumns.TableColumnXpgList[i].CalculatedColumnFormula.Available then begin
            TC.CalculatedColumnFormula.Array_ := XML.Table.TableColumns.TableColumnXpgList[i].CalculatedColumnFormula.Array_;
            TC.CalculatedColumnFormula.Content := XML.Table.TableColumns.TableColumnXpgList[i].CalculatedColumnFormula.Content;
          end;
          if XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowFormula.Available then begin
            TC.TotalsRowFormula.Array_ := XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowFormula.Array_;
            TC.TotalsRowFormula.Content := XML.Table.TableColumns.TableColumnXpgList[i].TotalsRowFormula.Content;
          end;
          if XML.Table.TableColumns.TableColumnXpgList[i].XmlColumnPr.Available then begin
            TC.XmlColumnPr.MapId := XML.Table.TableColumns.TableColumnXpgList[i].XmlColumnPr.MapId;
            TC.XmlColumnPr.Xpath := XML.Table.TableColumns.TableColumnXpgList[i].XmlColumnPr.Xpath;
            TC.XmlColumnPr.Denormalized := XML.Table.TableColumns.TableColumnXpgList[i].XmlColumnPr.Denormalized;
            TC.XmlColumnPr.XmlDataType := TXc12XmlDataType(XML.Table.TableColumns.TableColumnXpgList[i].XmlColumnPr.XmlDataType);
          end;
        end;

        if XML.Table.TableStyleInfo.Available then begin
          T.TableStyleInfo.Name := XML.Table.TableStyleInfo.Name;
          T.TableStyleInfo.ShowFirstColumn := XML.Table.TableStyleInfo.ShowFirstColumn;
          T.TableStyleInfo.ShowLastColumn := XML.Table.TableStyleInfo.ShowLastColumn;
          T.TableStyleInfo.ShowRowStripes := XML.Table.TableStyleInfo.ShowRowStripes;
          T.TableStyleInfo.ShowColumnStripes := XML.Table.TableStyleInfo.ShowColumnStripes;
        end;
      finally
        XML.Free;
      end;
    finally
      Stream.Free;
    end;

    Stream := FManager.FileData.OPC.ReadQueryTable;
    if Stream <> Nil then begin
      try
        T.QueryTable.LoadFromStream(Stream);
      finally
        Stream.Free;
      end;
    end;

    FManager.FileData.OPC.CloseTable;
  end;
end;

procedure TXLSReadXLSX.ReadVmlDrawing(const AId: AxUCString);
var
  Stream: TStream;
begin
  Stream := FManager.FileData.OPC.ReadVmlDrawing(AId,FCurrSheet.VmlDrawingRels);
  if Stream <> Nil then begin
    FCurrSheet.VmlDrawing.LoadFromStream(Stream);

    Stream.Free;
  end;
end;

procedure TXLSReadXLSX.ReadWorkbook;
var
  i: integer;
  S: AxUCString;
  Id: AxUCString;
  XML: TXPGDocXLSX;
  Stream: TStream;
  Strm1: TStream;
  Strm2: TStream;
  Sht: TCT_Sheet;
  Sheet: TXc12Sheet;
  BookView: TXc12BookView;
  BView: TCT_BookView;
  CWView: TXc12CustomWorkbookView;
  CWV: TCT_CustomWorkbookView;
  PivotCache: TCT_PivotCacheDefinition;
  STagType: TXc12SmartTagType;
  WebPubObj: TXc12WebPublishObject;
  WPO: TCT_WebPublishObject;
  FileRecoveryPr: TXc12FileRecoveryPr;
begin
  Stream := FManager.FileData.OPC.ReadWorkbook;
  if Stream = Nil then
    Exit;
  try
    XML := TXPGDocXLSX.Create;
    try
      XML.Errors.OnError := HandleParseError;
      XML.Workbook.DefinedNames.OnReadDefinedName := OnDefinedName;

      XML.LoadFromStream(Stream);

      FManager.Workbook.FileVersion.AppName := XML.Workbook.FileVersion.AppName;
      FManager.Workbook.FileVersion.LastEdited := XML.Workbook.FileVersion.LastEdited;
      FManager.Workbook.FileVersion.LowestEdited := XML.Workbook.FileVersion.LowestEdited;
      FManager.Workbook.FileVersion.RupBuild := XML.Workbook.FileVersion.RupBuild;
      FManager.Workbook.FileVersion.CodeName := XML.Workbook.FileVersion.CodeName;

      FManager.Workbook.FileSharing.ReadOnlyReccomended := XML.Workbook.FileSharing.ReadOnlyRecommended;
      FManager.Workbook.FileSharing.Username := XML.Workbook.FileSharing.UserName;
      FManager.Workbook.FileSharing.ReservationPassword := XML.Workbook.FileSharing.ReservationPassword;

      FManager.Workbook.WorkbookPr.Date1904 := XML.Workbook.WorkbookPr.Date1904;
      FManager.Workbook.WorkbookPr.ShowObjects := TXc12Objects(XML.Workbook.WorkbookPr.ShowObjects);
      FManager.Workbook.WorkbookPr.ShowBorderUnselectedTables := XML.Workbook.WorkbookPr.ShowBorderUnselectedTables;
      FManager.Workbook.WorkbookPr.FilterPrivacy := XML.Workbook.WorkbookPr.FilterPrivacy;
      FManager.Workbook.WorkbookPr.PromptedSolutions := XML.Workbook.WorkbookPr.PromptedSolutions;
      FManager.Workbook.WorkbookPr.ShowInkAnnotation := XML.Workbook.WorkbookPr.ShowInkAnnotation;
      FManager.Workbook.WorkbookPr.BackupFile := XML.Workbook.WorkbookPr.BackupFile;
      FManager.Workbook.WorkbookPr.SaveExternalLinkValues := XML.Workbook.WorkbookPr.SaveExternalLinkValues;
      FManager.Workbook.WorkbookPr.UpdateLinks := TXc12UpdateLinks(XML.Workbook.WorkbookPr.UpdateLinks);
      FManager.Workbook.WorkbookPr.CodeName := XML.Workbook.WorkbookPr.CodeName;
      FManager.Workbook.WorkbookPr.HidePivotFieldList := XML.Workbook.WorkbookPr.HidePivotFieldList;
      FManager.Workbook.WorkbookPr.ShowPivotChartFilter := XML.Workbook.WorkbookPr.ShowPivotChartFilter;
      FManager.Workbook.WorkbookPr.AllowRefreshQuery := XML.Workbook.WorkbookPr.AllowRefreshQuery;
      FManager.Workbook.WorkbookPr.PublishItems := XML.Workbook.WorkbookPr.PublishItems;
      FManager.Workbook.WorkbookPr.CheckCompatibility := XML.Workbook.WorkbookPr.CheckCompatibility;
      FManager.Workbook.WorkbookPr.AutoCompressPictures := XML.Workbook.WorkbookPr.AutoCompressPictures;
      FManager.Workbook.WorkbookPr.RefreshAllConnections := XML.Workbook.WorkbookPr.RefreshAllConnections;
      FManager.Workbook.WorkbookPr.DefaultThemeVersion := XML.Workbook.WorkbookPr.DefaultThemeVersion;

      FManager.Workbook.WorkbookProtection.WorkbookPassword := XML.Workbook.WorkbookProtection.WorkbookPassword;
      FManager.Workbook.WorkbookProtection.RevisionPassword := XML.Workbook.WorkbookProtection.RevisionsPassword;
      FManager.Workbook.WorkbookProtection.LockStructure := XML.Workbook.WorkbookProtection.LockStructure;
      FManager.Workbook.WorkbookProtection.LockWindows := XML.Workbook.WorkbookProtection.LockWindows;
      FManager.Workbook.WorkbookProtection.LockRevision := XML.Workbook.WorkbookProtection.LockRevision;

      if XML.Workbook.BookViews.WorkbookViewXpgList.Count > 0 then
        FManager.Workbook.BookViews.Clear(False);
      for i := 0 to XML.Workbook.BookViews.WorkbookViewXpgList.Count - 1 do begin
        BookView := FManager.Workbook.BookViews.Add;
        BView := XML.Workbook.BookViews.WorkbookViewXpgList[i];
        BookView.Visibility := TXc12Visibility(BView.Visibility);
        BookView.Minimized := BView.Minimized;
        BookView.ShowHorizontalScroll := BView.ShowHorizontalScroll;
        BookView.ShowVerticalScroll := BView.ShowVerticalScroll;
        BookView.ShowSheetTabs := BView.ShowSheetTabs;
        BookView.XWindow := BView.XWindow;
        BookView.YWindow := BView.YWindow;
        BookView.WindowWidth := BView.WindowWidth;
        BookView.WindowHeight := BView.WindowHeight;
        BookView.TabRatio := BView.TabRatio;
        BookView.FirstSheet := BView.FirstSheet;
        BookView.ActiveTab := BView.ActiveTab;
        BookView.AutoFilterDateGrouping := BView.AutoFilterDateGrouping;
      end;

      FManager.Workbook.FunctionGroups.BuiltInGroupCount := XML.Workbook.FunctionGroups.BuiltInGroupCount;
      for i := 0 to XML.Workbook.FunctionGroups.FunctionGroupXpgList.Count - 1 do
        FManager.Workbook.FunctionGroups.Add(XML.Workbook.FunctionGroups.FunctionGroupXpgList[i].Name);

      for i := 0 to XML.Workbook.ExternalReferences.ExternalReferenceXpgList.Count - 1 do
        ReadXLink(XML.Workbook.ExternalReferences.ExternalReferenceXpgList[i].R_Id);

      FManager.Workbook.CalcPr.CalcId := XML.Workbook.CalcPr.CalcId;
      FManager.Workbook.CalcPr.CalcMode := TXc12CalcMode(XML.Workbook.CalcPr.CalcMode);
      FManager.Workbook.CalcPr.FullCalcOnLoad := XML.Workbook.CalcPr.FullCalcOnLoad;
      FManager.Workbook.CalcPr.RefMode := TXc12RefMode(XML.Workbook.CalcPr.RefMode);
      FManager.Workbook.CalcPr.Iterate := XML.Workbook.CalcPr.Iterate;
      FManager.Workbook.CalcPr.IterateCount := XML.Workbook.CalcPr.IterateCount;
      FManager.Workbook.CalcPr.IterateDelta := XML.Workbook.CalcPr.IterateDelta;
      FManager.Workbook.CalcPr.FullPrecision := XML.Workbook.CalcPr.FullPrecision;
      FManager.Workbook.CalcPr.CalcCompleted := XML.Workbook.CalcPr.CalcCompleted;
      FManager.Workbook.CalcPr.CalcOnSave := XML.Workbook.CalcPr.CalcOnSave;
      FManager.Workbook.CalcPr.ConcurrentCalc := XML.Workbook.CalcPr.ConcurrentCalc;
      FManager.Workbook.CalcPr.ConcurrentManualCount := XML.Workbook.CalcPr.ConcurrentManualCount;
      FManager.Workbook.CalcPr.ForceFullCalc := XML.Workbook.CalcPr.ForceFullCalc;

      FManager.Workbook.OleSize := AreaStrToCellArea(XML.Workbook.OleSize.Ref);

      for i := 0 to XML.Workbook.CustomWorkbookViews.CustomWorkbookViewXpgList.Count - 1 do begin
        CWV := XML.Workbook.CustomWorkbookViews.CustomWorkbookViewXpgList[i];
        CWView := FManager.Workbook.CustomWorkbookViews.Add;

        CWView.Name := CWV.Name;
        CWView.Guid := CWV.Guid;
        CWView.AutoUpdate := CWV.AutoUpdate;
        CWView.MergeInterval := CWV.MergeInterval;
        CWView.ChangesSavedWin := CWV.ChangesSavedWin;
        CWView.OnlySync := CWV.OnlySync;
        CWView.PersonalView := CWV.PersonalView;
        CWView.IncludePrintSettings := CWV.IncludePrintSettings;
        CWView.IncludeHiddenRowCol := CWV.IncludeHiddenRowCol;
        CWView.Maximized := CWV.Maximized;
        CWView.Minimized := CWV.Minimized;
        CWView.ShowHorizontalScroll := CWV.ShowHorizontalScroll;
        CWView.ShowVerticalScroll := CWV.ShowVerticalScroll;
        CWView.ShowSheetTabs := CWV.ShowSheetTabs;
        CWView.XWindow := CWV.XWindow;
        CWView.YWindow := CWV.YWindow;
        CWView.WindowWidth := CWV.WindowWidth;
        CWView.WindowHeight := CWV.WindowHeight;
        CWView.TabRatio := CWV.TabRatio;
        CWView.ActiveSheetId := CWV.ActiveSheetId;
        CWView.ShowFormulaBar := CWV.ShowFormulaBar;
        CWView.ShowStatusbar := CWV.ShowStatusbar;
        CWView.ShowComments := TXc12Comments_(CWV.ShowComments);
        CWView.ShowObjects := TXc12Objects(CWV.ShowObjects);
      end;

      for i := 0 to XML.Workbook.PivotCaches.PivotCacheXpgList.Count - 1 do begin
        Id := XML.Workbook.PivotCaches.PivotCacheXpgList[i].R_Id;

        Strm1 := FManager.FileData.OPC.ReadPivotCacheDefinition(Id);
        try
          PivotCache := FManager.Workbook.PivotCaches.LoadFromStream(Strm1);
          PivotCache.CacheId := XML.Workbook.PivotCaches.PivotCacheXpgList[i].CacheId;

          S := PivotCache.R_Id;
          if S <> '' then begin
            Strm2 := FManager.FileData.OPC.ReadPivotCacheRecords(Id,S);
            try
              PivotCache.Records.LoadFromStream(Strm2);
            finally
              Strm2.Free;
            end;
          end;

        finally
          Strm1.Free;
        end;
      end;

      FManager.Workbook.SmartTagPr.Embed := XML.Workbook.SmartTagPr.Embed;
      FManager.Workbook.SmartTagPr.Show := TXc12SmartTagShow(XML.Workbook.SmartTagPr.Show);

      for i := 0 to XML.Workbook.SmartTagTypes.SmartTagTypeXpgList.Count - 1 do begin
        STagType := FManager.Workbook.SmartTagTypes.Add;

        STagType.NamespaceUri := XML.Workbook.SmartTagTypes.SmartTagTypeXpgList[i].NamespaceUri;
        STagType.Name := XML.Workbook.SmartTagTypes.SmartTagTypeXpgList[i].Name;
        STagType.Url := XML.Workbook.SmartTagTypes.SmartTagTypeXpgList[i].Url;
      end;

      FManager.Workbook.WebPublishing.Css := XML.Workbook.WebPublishing.Css;
      FManager.Workbook.WebPublishing.Thicket := XML.Workbook.WebPublishing.Thicket;
      FManager.Workbook.WebPublishing.LongFileNames := XML.Workbook.WebPublishing.LongFileNames;
      FManager.Workbook.WebPublishing.Vml := XML.Workbook.WebPublishing.Vml;
      FManager.Workbook.WebPublishing.AllowPng := XML.Workbook.WebPublishing.AllowPng;
      FManager.Workbook.WebPublishing.TargetScreenSize := TXc12TargetScreenSize(XML.Workbook.WebPublishing.TargetScreenSize);
      FManager.Workbook.WebPublishing.Dpi := XML.Workbook.WebPublishing.Dpi;
      FManager.Workbook.WebPublishing.CodePage := XML.Workbook.WebPublishing.CodePage;

      for i := 0 to XML.Workbook.FileRecoveryPrXpgList.Count - 1 do begin
        FileRecoveryPr := FManager.Workbook.FileRecoveryPrs.Add;
        FileRecoveryPr.AutoRecover := XML.Workbook.FileRecoveryPrXpgList[0].AutoRecover;
        FileRecoveryPr.CrashSave := XML.Workbook.FileRecoveryPrXpgList[0].CrashSave;
        FileRecoveryPr.DataExtractLoad := XML.Workbook.FileRecoveryPrXpgList[0].DataExtractLoad;
        FileRecoveryPr.RepairLoad := XML.Workbook.FileRecoveryPrXpgList[0].RepairLoad;
      end;

      for i := 0 to XML.Workbook.WebPublishObjects.WebPublishObjectXpgList.Count - 1 do begin
        WebPubObj := FManager.Workbook.WebPublishObjects.Add;
        WPO := XML.Workbook.WebPublishObjects.WebPublishObjectXpgList[i];

        WebPubObj.Id := WPO.Id;
        WebPubObj.DivId := WPO.DivId;
        WebPubObj.SourceObject := WPO.SourceObject;
        WebPubObj.DestinationFile := WPO.DestinationFile;
        WebPubObj.Title := WPO.Title;
        WebPubObj.AutoRepublish := WPO.AutoRepublish;
      end;

      for i := 0 to XML.Workbook.Sheets.SheetXpgList.Count - 1 do begin
        Sht := XML.Workbook.Sheets.SheetXpgList[i];
        if Sht.R_Id <> '' then begin
          Sheet := FManager.Workbook.Sheets.Add;
          Sheet.Name := Sht.Name;
          Sheet.SheetId := Sht.SheetId;
          Sheet.State := TXc12Visibility(Sht.State);
          Sheet.RId := Sht.R_Id;
        end;
      end;
    finally
      XML.Free;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSReadXLSX.ReadXLink(const AId: AxUCString);
var
  i,j,k: integer;
  OPC: TOPCItem;
  XML: TXPGDocXLink;
  Stream: TStream;
  XLink: TXc12ExternalLink;
  DName: TCT_ExternalDefinedName;
  DefName: TXc12XDefinedName;
  SheetData: TXc12XSheetData;
  SData: TCT_ExternalSheetData;
  Row: TXc12XRow;
  Cell: TXc12XCell;
  C: TCT_ExternalCell;
  DdeLink: TXc12DdeLink;
  Dde: TCT_DdeItem;
  DdeItem: TXc12DdeItem;
  DVal: TCT_DdeValue;
  DdeVal: TXc12DdeValue;
  OleLink: TXc12OleLink;
  OleItem: TXc12OleItem;
  OItem: TCT_OleItem;
begin
  Stream := FManager.FileData.OPC.ReadXLink(AId);
  try
    XML := TXPGDocXLink.Create;
    try
//      XML.Errors.OnError := HandleParseError;

      XML.LoadFromStream(Stream);

      XLink := FManager.XLinks.Add;

      if XML.ExternalLink.ExternalBook.Available then begin
        OPC := FManager.FileData.OPC.Workbook.Find(AId);
        if OPC <> Nil then begin
          OPC := OPC.Find(XML.ExternalLink.ExternalBook.R_Id);
          if OPC <> Nil then
            XLink.ExternalBook.Filename := OPC.Target;
        end;
      end;

      for i := 0 to XML.ExternalLink.ExternalBook.SheetNames.SheetNameXpgList.Count - 1 do
        XLink.ExternalBook.SheetNames.Add(XML.ExternalLink.ExternalBook.SheetNames.SheetNameXpgList[i].Val);

      for i := 0 to XML.ExternalLink.ExternalBook.DefinedNames.DefinedNameXpgList.Count - 1 do begin
        DName := XML.ExternalLink.ExternalBook.DefinedNames.DefinedNameXpgList[i];
        DefName := XLink.ExternalBook.DefinedNames.Add;
        DefName.Name := DName.Name;
        DefName.RefersTo := DName.RefersTo;
        DefName.SheetId := DName.SheetId;
      end;

      for i := 0 to XML.ExternalLink.ExternalBook.SheetDataSet.SheetDataXpgList.Count - 1 do begin
        SData := XML.ExternalLink.ExternalBook.SheetDataSet.SheetDataXpgList[i];
        SheetData := XLink.ExternalBook.SheetDataSet.Add;
        SheetData.SheetId := SData.SheetId;
        SheetData.RefreshError := SData.RefreshError;
        for j := 0 to SData.RowXpgList.Count - 1 do begin
          Row := SheetData.Add;
          Row.Row := SData.RowXpgList[j].R - 1;
          for k := 0 to SData.RowXpgList[j].CellXpgList.Count - 1 do begin
            C := SData.RowXpgList[j].CellXpgList[k];
            Cell := Row.Add;
            Cell.CellRef := C.R;
            Cell.ValueMetadata := C.Vm;
            case C.T of
              stectB  : Cell.AsBoolean := XmlStrToBool(C.V);
              stectN  : Cell.AsFloat := XmlStrToFloat(C.V);
              stectE  : Cell.AsError := ErrorTextToCellError(C.V);
              stectStr: Cell.AsString := C.V;
            end;
          end;
        end;
      end;

      if XML.ExternalLink.DdeLink.Available then begin
        DdeLink := XLink.DdeLink;
        DdeLink.DdeService := XML.ExternalLink.DdeLink.DdeService;
        DdeLink.DdeTopic := XML.ExternalLink.DdeLink.DdeTopic;
        for i := 0 to XML.ExternalLink.DdeLink.DdeItems.DdeItemXpgList.Count - 1 do begin
          Dde := XML.ExternalLink.DdeLink.DdeItems.DdeItemXpgList[i];
          DdeItem := DdeLink.Add;
          DdeItem.Name := Dde.Name;
          DdeItem.Ole := Dde.Ole;
          DdeItem.Advise := Dde.Advise;
          DdeItem.PreferPic := Dde.PreferPic;

          DdeItem.Values.Cols := Dde.Values.Cols;
          DdeItem.Values.Rows := Dde.Values.Rows;
          for j := 0 to Dde.Values.ValueXpgList.Count - 1 do begin
            DVal := Dde.Values.ValueXpgList[i];
            DdeVal := DdeItem.Values.Add;
            case DVal.T of
              stdvtNil: DdeVal.AsNil := True;
              stdvtB  : DdeVal.AsBoolean := XmlStrToBool(DVal.Val);
              stdvtN  : DdeVal.AsFloat := XmlStrToFloat(DVal.Val);
              stdvtE  : DdeVal.AsError := ErrorTextToCellError(DVal.Val);
              stdvtStr: DdeVal.AsString := DVal.Val;
            end;
          end;
        end;

        if XML.ExternalLink.OleLink.Available then begin
          OleLink := XLink.OleLink;

          OPC := FManager.FileData.OPC.Workbook.Find(AId);
          if OPC <> Nil then begin
            OPC := OPC.Find(XML.ExternalLink.OleLink.R_Id);
            if OPC <> Nil then
              OleLink.Filename := OPC.Target;
          end;

          OleLink.ProgId := XML.ExternalLink.OleLink.ProgId;
          for i := 0 to XML.ExternalLink.OleLink.OleItems.OleItemXpgList.Count - 1 do begin
            OItem := XML.ExternalLink.OleLink.OleItems.OleItemXpgList[i];
            OleItem := OleLink.Add;
            OleItem.Name := OItem.Name;
            OleItem.Icon := OItem.Icon;
            OleItem.Advice := OItem.Advise;
            OleItem.PreferPic := OItem.PreferPic;
          end;
        end;
      end;
    finally
      XML.Free;
    end;
  finally
    Stream.Free;
  end;
end;

{ TReadSheetThread }

constructor TReadSheetThread.Create(AStream: TStream; ASheet: TXc12DataWorksheet);
begin
  FStream := AStream;
  FSheet := ASheet;
  inherited Create(True);
end;

destructor TReadSheetThread.Destroy;
begin
  FStream.Free;
  inherited;
end;

procedure TReadSheetThread.Execute;
begin
  ReadSheet;
end;

procedure TReadSheetThread.OnCell(ASender: TObject);
begin
  // If needed, copy from normal OnCell.
end;

procedure TReadSheetThread.OnRow(ASender: TObject);
var
  R: TCT_Row;
  Row: PXLSMMURowHeader;
begin
  R := TCT_Row(ASender);
  if R.CustomFormat then
    Row := FSheet.Cells.AddRow(R.R - 1,R.S)
  else
    Row := FSheet.Cells.AddRow(R.R - 1,0);
  if R.Ht <> 0 then
    Row.Height := Round(R.Ht * 20)
  else
    Row.Height := XLS_DEFAULT_ROWHEIGHT_FLAG;
  Row.Options := [];
  if R.Hidden then
    Row.Options := Row.Options + [xroHidden];
  if R.Collapsed then
    Row.Options := Row.Options + [xroCollapsed];
  if R.CustomHeight then
    Row.Options := Row.Options + [xroCustomHeight];
  if R.Ph then
    Row.Options := Row.Options + [xroPhonetic];
  if R.ThickTop then
    Row.Options := Row.Options + [xroThickTop];
  if R.ThickBot then
    Row.Options := Row.Options + [xroThickBottom];
  Row.OutlineLevel := R.OutlineLevel;

  R.Clear;
end;

procedure TReadSheetThread.ReadSheet;
var
  XML: TXPGDocXLSX;
begin
  XML := TXPGDocXLSX.Create;
  try

    XML.Worksheet.SheetData.Row.OnReadC := OnCell;
    XML.Worksheet.SheetData.OnReadRow := OnRow;
    XML.LoadFromStream(FStream);

  finally
    XML.Free;
  end;
end;

{ TXLSSharedFormulaList }

constructor TXLSSharedFormulaList.Create;
begin
  inherited Create;
end;

function TXLSSharedFormulaList.GetItems(Index: integer): TXLSSharedFormula;
begin
  Result := TXLSSharedFormula(inherited Items[Index]);
end;

procedure TXLSSharedFormulaList.Insert(AIndex: integer; ARef: PXLSCellRef; AApplyArea: PXLSCellArea; APtgs: PXLSPtgs; APtgsSize: integer);
var
  Item: TXLSSharedFormula;
begin
  while AIndex >= Count do
    Add(Nil);

  Item := TXLSSharedFormula.Create;

  Item.Ref := ARef^;
  Item.ApplyArea := AApplyArea^;
  Item.AllocPtgs(APtgsSize);
  System.Move(APtgs^,Item.Ptgs^,APtgsSize);

  inherited Insert(AIndex,Item);
end;

{ TXLSSharedFormula }

procedure TXLSSharedFormula.AllocPtgs(const ASize: integer);
begin
  FPtgsSize := ASize;
  GetMem(FPtgs,FPtgsSize);
end;

function TXLSSharedFormula.CopyPtgs(AIndex: integer): PXLSPtgs;
begin
  GetMem(Result,FPtgsSize);
  System.Move(FPtgs^,Result^,FPtgsSize);
end;

destructor TXLSSharedFormula.Destroy;
begin
  if FPtgs <> Nil then
    FreeMem(FPtgs);
  inherited;
end;

end.
