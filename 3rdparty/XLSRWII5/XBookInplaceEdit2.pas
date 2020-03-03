unit XBookInplaceEdit2;

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

interface

uses Classes, SysUtils, vcl.Controls, Windows, Messages, vcl.Graphics, Math,

     Xc12Utils5, Xc12DataStylesheet5, Xc12DataSST5, Xc12Manager5, Xc12Common5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSTools5, XLSFormulaTypes5, XLSEncodeFmla5,
     XLSDecodeFmla5,

     XBookUtils2, XBook_System_2, XBookSkin2, XBookWindows2,

     XSSIEDefs, XSSIEUtils, XSSIEDocProps, XSSIELogParas, XSSIEEditor, XSSIEKeys, XSSIECharRun;

const EDITREF_TAG_START = '<!';
const EDITREF_TAG_END   = '!>';

type TXBookEditorCloseEvent = procedure (ASender: TObject; AKey: word) of object;

type TXBookIECellAreaHelper = class(TObject)
protected
     FEditorArea  : TAXWEditorArea;

     FDispCol     : integer;
     FDispRow     : integer;
     FDispCols    : TDynIntegerArray;
     FDispRows    : TDynIntegerArray;

     FAlignment   : TAXWPapTextAlign;

     FOrigLeft    : integer;
     FOrigTop     : integer;
public
     constructor Create(AEditorArea: TAXWEditorArea);

     procedure CalcExpand(const ATextWidth, ATextHeight: integer);

     property Alignment : TAXWPapTextAlign read FAlignment write FAlignment;

     property OrigLeft  : integer read FOrigLeft write FOrigLeft;
     property OrigTop   : integer read FOrigTop write FOrigTop;
     end;

type TXBookInplaceEditor = class(TAXWWinEditor)
private
     function  GetCell: PXLSCellItem;
     function  GetTextChangedEvent: TAXWTextChangedEvent;
     procedure SetTextChangedEvent(const Value: TAXWTextChangedEvent);
     function  GetOnCharFmtChanged: TNotifyEvent;
     procedure SetOnCharFmtChanged(const Value: TNotifyEvent);
     function  GetCHP: TAXWCHP;
     function  GetPAP: TAXWPAP;
protected
     FIsVisible : boolean;
     FManager   : TXc12Manager;
     FCloseEvent: TXBookEditorCloseEvent;
     FSheetIndex: integer;
     FCol       : integer;
     FRow       : integer;
     FEditCell  : TXLSCellItem;

     FCRLFMode  : boolean;

     FEditorArea: TAXWEditorArea;
     FAreaHelper: TXBookIECellAreaHelper;

     function  GetText: AxUCString; override;

     procedure CalcEditorArea(const AX,AY: integer; const AWrapText: boolean);
     procedure Xc12FontToCHPX(AFont: TXc12Font; ACHPX: TAXWCHPX);
     procedure CHPToXc12Font(ACHP: TAXWCHP; AFont: TXc12Font);
     procedure CHPXToXc12Font(ACHPX: TAXWCHPX; AFont: TXc12Font);
     procedure AddFormatted(const ASSTIndex: integer; ADefaultFont: TXc12Font);
     procedure DoShow;
     procedure EditorBeforePaint(ASender: TObject);
     procedure EditorAfterPaint(ASender: TObject);
     function  PrepareTextBefore(const AText: AxUCString): AxUCString;
     function  PrepareTextAfter(const AText: AxUCString): AxUCString;
public
     constructor Create(AParent: TXSSWindow; AManager: TXc12Manager);
     destructor Destroy; override;

     procedure SetColsRows(const ADispCols, ADispRows: TDynIntegerArray; const ADispCol, ADispColCount, ADispRow, ADispRowCount: integer);
     procedure SetCellPos(const ASheetIndex, ACol, ARow: integer);

     procedure BeginAddRefs;
     procedure AddRefText(const AText: AxUCString); overload;
     function  AddRefText(const AText: AxUCString; const AColor: longword): integer; overload;
     procedure EndAddRefs;
     procedure UpdateRefText(const AIndex: integer; const AText: AxUCString);

     function  LastCharIsOperator: boolean;
     // True if there is more than one char run.
     function  IsFormatted: boolean;
     procedure GetFormatted(var AFontRuns: TXc12DynFontRunArray);

     procedure BeginUpdate;
     procedure EndUpdate;
     procedure Command(ACmd: TAXWCommand);
     procedure ClearCharFormatting;

     function  Show(const AX, AY, AWidth,AHeight: integer; AssignCellValue: boolean): AxUCString;
     procedure Hide;

     function  IsFormula: boolean;

     property Cell: PXLSCellItem read GetCell;
     property SheetIndex: integer read FSheetIndex;
     property Col: integer read FCol;
     property Row: integer read FRow;
     property PAP: TAXWPAP read GetPAP;
     property CHP: TAXWCHP read GetCHP;

     property OnClose: TXBookEditorCloseEvent read FCloseEvent write FCloseEvent;
     property OnTextChanged: TAXWTextChangedEvent read GetTextChangedEvent write SetTextChangedEvent;
     property OnCharFmtChanged: TNotifyEvent read GetOnCharFmtChanged write SetOnCharFmtChanged;
     end;

implementation

{ TXBookInplaceEditor }

{
	<xsd:simpleType name="ST_TblStyleOverrideType">
		<xsd:restriction base="xsd:string">
			<xsd:enumeration value="wholeTable"/>
			<xsd:enumeration value="firstRow"/>
			<xsd:enumeration value="lastRow"/>
			<xsd:enumeration value="firstCol"/>
			<xsd:enumeration value="lastCol"/>
			<xsd:enumeration value="band1Vert"/>
			<xsd:enumeration value="band2Vert"/>
			<xsd:enumeration value="band1Horz"/>
			<xsd:enumeration value="band2Horz"/>
			<xsd:enumeration value="neCell"/>
			<xsd:enumeration value="nwCell"/>
			<xsd:enumeration value="seCell"/>
			<xsd:enumeration value="swCell"/>
		</xsd:restriction>
	</xsd:simpleType>

}
procedure TXBookInplaceEditor.AddFormatted(const ASSTIndex: integer; ADefaultFont: TXc12Font);
var
  i: integer;
  Txt: AxUCString;
  Cnt: integer;
  PStr: PXLSString;
  FontRuns: PXc12FontRunArray;
  FmtText: TXLSFormattedText;
  CHPX: TAXWCHPX;
begin
  Txt := FManager.SST.ItemText[ASSTIndex];
  PStr := FManager.SST[ASSTIndex];
  Cnt := FManager.SST.GetFontRunsCount(PStr);
  FontRuns := FManager.SST.GetFontRuns(PStr);

  FmtText := TXLSFormattedText.Create;
  try
    FmtText.SplitAtCR := True;
    FmtText.Add(Txt,ADefaultFont,FontRuns,Cnt);

    CHPX := TAXWCHPX.Create(FEditor.MasterCHP);
    try
      for i := 0 to FmtText.Count - 1 do begin
        if FmtText[i].NewLine then
          FmtText[i].AddSpecialCR(AXW_CHAR_SPECIAL_HARDLINEBREAK);

        CHPX.Clear;
        Xc12FontToCHPX(FmtText[i].Font,CHPX);
        FEditor.AppendText(FmtText[i].Text,CHPX);
      end;
    finally
      CHPX.Free;
    end;
  finally
    FmtText.Free;
  end;

//  FEditor.CommitText;
end;

procedure TXBookInplaceEditor.AddRefText(const AText: AxUCString);
begin
  FEditor.Paras.Last.AppendTextUncond(Nil,AText);
end;

function TXBookInplaceEditor.AddRefText(const AText: AxUCString; const AColor: longword): integer;
var
  CHPX: TAXWCHPX;
  Para: TAXWLogPara;
begin
  CHPX := TAXWCHPX.Create(FEditor.MasterCHP);
  try
    CHPX.Color := AColor;
    Para := FEditor.Paras.Last;
    Para.AppendTextUncond(CHPX,AText);
  finally
    CHPX.Free;
  end;

  Result := Para.Count - 1;
end;

procedure TXBookInplaceEditor.BeginAddRefs;
begin
  FEditor.Clear;
  FEditor.Paras.Add;
end;

procedure TXBookInplaceEditor.BeginUpdate;
begin
  FEditor.BeginUpdate;
end;

procedure TXBookInplaceEditor.CalcEditorArea(const AX,AY: integer; const AWrapText: boolean);
var
  W: integer;
  H: integer;
begin
  FEditorArea.X1 := AX;
  FEditorArea.Y1 := AY;

  FEditorArea.X2 := AX + FEditorArea.CellWidth;
  FEditorArea.Y2 := AY + FEditorArea.CellHeight;
  if AWrapText then begin
    FEditorArea.MinX := FEditorArea.X1;
    FEditorArea.MaxX := FEditorArea.X2;
  end
  else begin
    case FEditor.CaretPAP.Alignment of
      axptaLeft   : begin
        FEditorArea.MinX := AX;
        FEditorArea.MinY := AY;
      end;
      axptaCenter : begin
        W := Min(FEditorArea.X1 - FEditorArea.MinX,FEditorArea.MaxX - FEditorArea.X2);
        H := Min(FEditorArea.Y1 - FEditorArea.MinY,FEditorArea.MaxY - FEditorArea.Y2);

        FEditorArea.MinX := FEditorArea.X1 - W;
        FEditorArea.MinY := FEditorArea.Y1 - H;
        FEditorArea.MaxX := FEditorArea.X2 + W;
        FEditorArea.MaXY := FEditorArea.Y2 + H;
      end;
      axptaRight  : begin
        FEditorArea.MaxX := AX + FEditorArea.CellWidth;
        FEditorArea.MaxY := AY + FEditorArea.CellHeight;
      end;
      axptaJustify: ;
    end;
  end;
end;

procedure TXBookInplaceEditor.CHPToXc12Font(ACHP: TAXWCHP; AFont: TXc12Font);
begin
  AFont.Name := ACHP.FontName;
  AFont.Size := ACHP.Size;
  AFont.Color := RGBColorToXc12(ACHP.Color);
  AFont.Style := [];
  if ACHP.Bold         then AFont.Style := AFont.Style + [xfsBold];
  if ACHP.Italic       then AFont.Style := AFont.Style + [xfsItalic];
  if ACHP.StrikeTrough then AFont.Style := AFont.Style + [xfsStrikeOut];
  case ACHP.Underline of
    axcuNone  : AFont.Underline := xulNone;
    axcuSingle: AFont.Underline := xulSingle;
    axcuDouble: AFont.Underline := xulDouble;
  end;
  case ACHP.SubSuperscript of
    axcssNone       : AFont.SubSuperscript := xssNone;
    axcssSubscript  : AFont.SubSuperscript := xssSuperscript;
    axcssSuperscript: AFont.SubSuperscript := xssSubscript;
  end;
end;

procedure TXBookInplaceEditor.CHPXToXc12Font(ACHPX: TAXWCHPX; AFont: TXc12Font);
begin
  AFont.Name := ACHPX.FontName;
  AFont.Size := ACHPX.Size;
  AFont.Color := RGBColorToXc12(ACHPX.Color);
  AFont.Style := [];
  if ACHPX.Bold         then AFont.Style := AFont.Style + [xfsBold];
  if ACHPX.Italic       then AFont.Style := AFont.Style + [xfsItalic];
  if ACHPX.StrikeTrough then AFont.Style := AFont.Style + [xfsStrikeOut];
  case ACHPX.Underline of
    axcuNone  : AFont.Underline := xulNone;
    axcuSingle: AFont.Underline := xulSingle;
    axcuDouble: AFont.Underline := xulDouble;
  end;
  case ACHPX.SubSuperscript of
    axcssNone       : AFont.SubSuperscript := xssNone;
    axcssSubscript  : AFont.SubSuperscript := xssSuperscript;
    axcssSuperscript: AFont.SubSuperscript := xssSubscript;
  end;
end;

procedure TXBookInplaceEditor.ClearCharFormatting;
begin
  FEditor.CmdFormat(axcClearAllCharFmt);
end;

procedure TXBookInplaceEditor.Command(ACmd: TAXWCommand);
begin
  FEditor.Command(ACmd);
  FSystem.ProcessRequests;
end;

constructor TXBookInplaceEditor.Create(AParent: TXSSWindow; AManager: TXc12Manager);
begin
  FEditorArea := TAXWEditorArea.Create;
  inherited Create(AParent,FEditorArea);

  FManager := AManager;

  FEditor.OnBeforePaint := EditorBeforePaint;
  FEditor.OnAfterPaint := EditorAfterPaint;

  Visible := False;
  FCRLFMode := True;

  FAreaHelper := TXBookIECellAreaHelper.Create(FEditorArea);

  FEditor.SetNoFormattingOnFirstEqual := True;
end;

destructor TXBookInplaceEditor.Destroy;
begin
  FAreaHelper.Free;
  FEditorArea.Free;
  FSystem.HideCaret;
  inherited;
end;

procedure TXBookInplaceEditor.Hide;
begin
  Visible := False;
end;

function TXBookInplaceEditor.IsFormatted: boolean;
begin
  Result := FEditor.Paras.CountCR > 1;
end;

function TXBookInplaceEditor.IsFormula: boolean;
begin
  Result := FManager.Worksheets[FSheetIndex].Cells.CellType(@FEditCell) in XLSCellTypeFormulas;
end;

function TXBookInplaceEditor.LastCharIsOperator: boolean;
var
  S: AxUCString;
  Para: TAXWLogPara;
begin
  Result := False;
  Para := FEditor.Paras.Last;
  if (Para <> Nil) and (Para.CRLast <> Nil) then begin
    S := Trim(Para.CRLast.Text);
    Result := S <> '';
    if Result then begin
      Result := CharInSet(S[Length(S)],['^','*','/','+','-','&','=','%','(']);
    end;
  end;
end;

function TXBookInplaceEditor.PrepareTextAfter(const AText: AxUCString): AxUCString;
begin
  if FCRLFMode then
    Result := StringReplace(AText,AXW_CHAR_SPECIAL_HARDLINEBREAK,#$0A + #$0D,[rfReplaceAll])
  else
    Result := StringReplace(AText,AXW_CHAR_SPECIAL_HARDLINEBREAK,#$0A,[rfReplaceAll]);
end;

function TXBookInplaceEditor.PrepareTextBefore(const AText: AxUCString): AxUCString;
var
  S,S2: AxUCString;
  V: double;
begin
  S := AText;

  Result := '';

  while S <> '' do begin
    S2 := SplitAtChar(#$0A,S);
    Result := Result + Trim(S2) + AXW_CHAR_SPECIAL_HARDLINEBREAK;

    if S2 <> '' then
      FCRLFMode := S2[Length(S2)] = #$0D;
  end;
  Result := Trim(Result);

  if TryStrToFloat(Result,V) then
    Result := '''' + Result;
end;

procedure TXBookInplaceEditor.DoShow;
begin
  Enabled := True;
  Visible := True;
  FEditor.CommitText;
  Focus;
end;

procedure TXBookInplaceEditor.EditorAfterPaint(ASender: TObject);
begin
  EndPaint;
end;

procedure TXBookInplaceEditor.EditorBeforePaint(ASender: TObject);
begin
  FAreaHelper.CalcExpand(Round(FEditor.TextWidth),Round(FEditor.TextHeight));

  BeginPaint;
end;

procedure TXBookInplaceEditor.EndAddRefs;
begin
  FEditor.Paras.Sync;
  FEditor.CommitText;
end;

procedure TXBookInplaceEditor.EndUpdate;
begin
  FEditor.EndUpdate;
end;

function TXBookInplaceEditor.GetCell: PXLSCellItem;
begin
  Result := @FEditCell;
end;

function TXBookInplaceEditor.GetCHP: TAXWCHP;
begin
  Result := FEditor.CaretCHP;
end;

procedure TXBookInplaceEditor.GetFormatted(var AFontRuns: TXc12DynFontRunArray);
var
  i,j,k,l: integer;
  Para: TAXWLogPara;
  Font: TXc12Font;
begin
  SetLength(AFontRuns,FEditor.Paras.CountCR);

  k := 0;
  l := 0;

  for i := 0 to FEditor.Paras.Count - 1 do begin
    Para := FEditor.Paras[i];
    for j := 0 to Para.Count - 1 do begin
      Font := FManager.StyleSheet.Fonts.Add;
      if Para[j].CHPX <> Nil then
        CHPXToXc12Font(Para[j].CHPX,Font)
      else
        CHPToXc12Font(FEditor.MasterCHP,Font);
      AFontRuns[k].Index := l;
      AFontRuns[k].Font := FManager.StyleSheet.Fonts.Find(Font);
      FManager.StyleSheet.XFEditor.UseFont(AFontRuns[k].Font);
      Inc(l,Para[j].Len);
      Inc(k);
    end;
  end;
end;

function TXBookInplaceEditor.GetOnCharFmtChanged: TNotifyEvent;
begin
  Result := FEditor.OnCharFmtChanged;
end;

function TXBookInplaceEditor.GetPAP: TAXWPAP;
begin
  Result := FEditor.CaretPAP;
end;

function TXBookInplaceEditor.GetText: AxUCString;
begin
  Result := PrepareTextAfter(inherited GetText);
end;

function TXBookInplaceEditor.GetTextChangedEvent: TAXWTextChangedEvent;
begin
  Result := FEditor.OnTextChanged;
end;

procedure TXBookInplaceEditor.SetCellPos(const ASheetIndex, ACol, ARow: integer);
begin
  FSheetIndex := ASheetIndex;
  FCol := ACol;
  FRow := ARow;
end;

procedure TXBookInplaceEditor.SetColsRows(const ADispCols, ADispRows: TDynIntegerArray; const ADispCol, ADispColCount, ADispRow, ADispRowCount: integer);
var
  i: integer;
  W: integer;
  H: integer;
begin
  W := ADispCols[ADispCol];
  for i := 1 to ADispColCount - 1 do
    Inc(W,ADispCols[ADispCol + i]);
  FEditorArea.CellWidth := W;

  H := ADispRows[ADispRow];
  for i := 1 to ADispRowCount - 1 do
    Inc(H,ADispRows[ADispRow + i]);
  FEditorArea.CellHeight := H;

  FAreaHelper.FDispCol := ADispCol;
  FAreaHelper.FDispRow := ADispRow;
  FAreaHelper.FDispCols := ADispCols;
  FAreaHelper.FDispRows := ADispRows;
end;

procedure TXBookInplaceEditor.SetOnCharFmtChanged(const Value: TNotifyEvent);
begin
  FEditor.OnCharFmtChanged := Value;
end;

procedure TXBookInplaceEditor.SetTextChangedEvent(const Value: TAXWTextChangedEvent);
begin
  FEditor.OnTextChanged := Value;
end;

function TXBookInplaceEditor.Show(const AX, AY, AWidth, AHeight: integer; AssignCellValue: boolean): AxUCString;
var
  i: integer;
//  H: integer;
  Cells: TXLSCellMMU;
  XF: TXc12XF;
  CHPX: TAXWCHPX;
  Ptgs: PXLSPtgs;
  PtgsSz: integer;
begin
  FEditorArea.MinX := TXSSClientWindow(FParent).CX1;
  FEditorArea.MinY := TXSSClientWindow(FParent).CY1;
  FEditorArea.MaxX := TXSSClientWindow(FParent).CX2;
  FEditorArea.MaxY := TXSSClientWindow(FParent).CY2;

  FEditorArea.CellWidth := AWidth;
  FEditorArea.CellHeight := AHeight;

  FCX1 := TXSSClientWindow(FParent).CX1;
  FCY1 := TXSSClientWindow(FParent).CY1;
  FCX2 := TXSSClientWindow(FParent).CX2;
  FCY2 := TXSSClientWindow(FParent).CY2;

  FAreaHelper.OrigLeft := AX;
  FAreaHelper.OrigTop := AY;

  Cells := FManager.Worksheets[FSheetIndex].Cells;
  FEditCell := Cells.FindCell(FCol,FRow);

  FEditor.CaretPAP.Alignment := axptaLeft;
  if FEditCell.Data <> Nil then begin
    XF := FManager.StyleSheet.XFs[Cells.GetStyle(@FEditCell)];
    case Cells.CellType(@FEditCell) of
      xctBoolean,
      xctError  : FEditor.CaretPAP.Alignment := axptaCenter;
      xctFloat  : FEditor.CaretPAP.Alignment := axptaRight;
    end;
  end
  else
    XF := FManager.StyleSheet.XFs.DefaultXF;

  FEditor.BgColor := Xc12ColorToRGB(XF.Fill.FgColor,XSS_SYS_COLOR_WHITE);

  case XF.Alignment.HorizAlignment of
    chaGeneral         : ;
    chaLeft            : FEditor.CaretPAP.Alignment := axptaLeft;
    chaCenter          : FEditor.CaretPAP.Alignment := axptaCenter;
    chaRight           : FEditor.CaretPAP.Alignment := axptaRight;
    chaFill            : ;
    chaJustify         : FEditor.CaretPAP.Alignment := axptaJustify;
    chaCenterContinuous: ;
    chaDistributed     : ;
  end;

  CHPX := TAXWCHPX.Create(FEditor.MasterCHP);
  try
    Xc12FontToCHPX(XF.Font,CHPX);
    FEditor.MasterCHP.Assign(CHPX);
  finally
    CHPX.Free;
  end;

  FAreaHelper.Alignment := FEditor.CaretPAP.Alignment;

  CalcEditorArea(AX,AY,XF.Alignment.IsWrapText);

  if AssignCellValue  and (FEditCell.Data <> Nil) then begin
    case Cells.CellType(@FEditCell) of
      xctBlank  : Text := '';
      xctBoolean: if Cells.GetBoolean(@FEditCell) then Text := G_StrTRUE else Text := G_StrFALSE;
      xctError  : Text := Xc12CellErrorNames[Cells.GetError(@FEditCell)];
      xctString : begin
        i := Cells.GetStringSST(@FEditCell);
        if FManager.SST.IsFormatted[i] then
          AddFormatted(i,XF.Font)
        else
          Result := PrepareTextBefore(FManager.SST.ItemText[i]);
      end;
      xctFloat  : begin
        if XF.NumFmt.IsDate then
          Text := DateToStr(Cells.GetFloat(@FEditCell))
        else if XF.NumFmt.IsTime then
          Text := FloatToStr(Cells.GetFloat(@FEditCell))
        else if XF.NumFmt.IsDateTime then
          Text := DateTimeToStr(Cells.GetFloat(@FEditCell))
        else
          Text := FloatToStr(Cells.GetFloat(@FEditCell));
      end;
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula: begin
        if Cells.IsFormulaCompiled(@FEditCell) then begin
          PtgsSz := Cells.GetFormulaPtgs(@FEditCell,Ptgs);
          Result := '=' + XLSDecodeFormula(FManager,Cells,FSheetIndex,Ptgs,PtgsSz);
        end
        else
          Result := '=' + Cells.GetFormula(@FEditCell);

        FEditor.NoFormatting := True;
      end;
    end;
  end
  else
    Result := '';

  // If formula, Text is set by TXBookEditRefList
  if (FEditCell.Data = Nil) or not (Cells.CellType(@FEditCell) in XLSCellTypeFormulas) then
    Text := Result
  else
    Text := '';

  case CellBorderWidthts[XF.Border.Left.Style] of
    0,
    1,
    2: FX1 := AX;
    3: FX1 := AX + 1;
  end;

//  H := 0;
  case CellBorderWidthts[XF.Border.Top.Style] of
    0,
    1,
    2: begin
      FY1 := AY;
//      Inc(H);
    end;
    3: begin
      FY1 := AY + 1;
//      Inc(H,2);
    end;
  end;

//  case CellBorderWidthts[XF.Border.Bottom.Style] of
//    0,
//    1: Dec(H);
//    2: ;
//    3: ;
//  end;

  FX2 := FX1 + FEditorArea.CellWidth;
  FY2 := FY1 + FEditorArea.CellHeight;

  FSkin.SetCellFont(XF.Font);

  case XF.Alignment.VertAlignment of
    cvaTop        : FEditor.VertAlign := avaTop;
    cvaCenter     : FEditor.VertAlign := avaCenter;
    cvaBottom     : FEditor.VertAlign := avaBottom;
    cvaJustify    : FEditor.VertAlign := avaTop;
    cvaDistributed: FEditor.VertAlign := avaDistributed;
  end;

  FEditorArea.MargLeft := FSkin.CellLeftMarg;
  FEditorArea.MargTop := FSkin.CellTopMarg;
  FEditorArea.MargRight := FSkin.CellRightMarg;
  FEditorArea.MargBottom := FSkin.CellBottomMarg;

  FEditorArea.PaddingLeft := 0;
  FEditorArea.PaddingTop := 0;
  FEditorArea.PaddingRight := 0;
  FEditorArea.PaddingBottom := 0;

  DoShow;
end;

procedure TXBookInplaceEditor.UpdateRefText(const AIndex: integer; const AText: AxUCString);
var
  Para: TAXWLogPara;
begin
  Para := FEditor.Paras.Last;
  Para[AIndex].Text := AText;
  FEditor.Paras.Sync;
end;

procedure TXBookInplaceEditor.Xc12FontToCHPX(AFont: TXc12Font; ACHPX: TAXWCHPX);
begin
  ACHPX.FontName := AFont.Name;
  ACHPX.Size := AFont.Size;
  ACHPX.Color := AFont.Color.ARGB;
  ACHPX.Bold := xfsBold in AFont.Style;
  ACHPX.Italic := xfsItalic in AFont.Style;
  ACHPX.StrikeTrough := xfsStrikeOut in AFont.Style;
  case AFont.Underline of
    xulNone         : ;
    xulSingle       : ACHPX.Underline := axcuSingle;
    xulDouble       : ACHPX.Underline := axcuDouble;
    xulSingleAccount: ACHPX.Underline := axcuSingle;
    xulDoubleAccount: ACHPX.Underline := axcuDouble;
  end;
  case AFont.SubSuperscript of
    xssNone       : ;
    xssSuperscript: ACHPX.SubSuperscript := axcssSuperscript;
    xssSubscript  : ACHPX.SubSuperscript := axcssSubscript;
  end;
end;

{ TXBookIECellAreaHelper }

procedure TXBookIECellAreaHelper.CalcExpand(const ATextWidth, ATextHeight: integer);
var
  W: integer;
  WL,WR: integer;
  H: integer;
  Col: integer;
  Row: integer;
begin
  Col := FDispCol;
  W := FEditorArea.CellWidth;
  case FAlignment of
    axptaJustify,
    axptaLeft   : begin
      while (ATextWidth > W) and (ATextWidth <= FEditorArea.PaginateWidth) and (Col < High(FDispCols)) do begin
        Inc(Col);
        W := W + FDispCols[Col];
      end;
      FEditorArea.X2 := FOrigLeft + W;
    end;
    axptaCenter : begin
      WL := 0;
      WR := 0;
      Col := 1;
      while (ATextWidth > (W + Min(WL,WR) * 2)) and ((FDispCol - Col) >= 0) and ((FDispCol + Col) < High(FDispCols)) do begin
        Inc(WL,FDispCols[FDispCol - Col]);
        Inc(WR,FDispCols[FDispCol + Col]);
        Inc(Col);
      end;
      FEditorArea.X1 := FOrigLeft - Min(WL,WR);
      FEditorArea.X2 := FOrigLeft + W + Min(WL,WR);
    end;
    axptaRight  : begin
      while (ATextWidth > W) and (Col > 0) do begin
        Dec(Col);
        W := W + FDispCols[Col];
      end;
      FEditorArea.X1 := FEditorArea.X2 - W;
    end;
  end;

  Row := FDispRow + 1;
  H := FEditorArea.CellHeight;

  while (ATextHeight > H) and (Row <= High(FDispRows)) do begin
    H := H + FDispRows[Row];
    Inc(Row);
  end;
  FEditorArea.Y2 := FEditorArea.Y1 + H;

  if W > FEditorArea.CellWidth then
    FEditorArea.CellWidth := W;

  FEditorArea.ClipToMax;
end;

constructor TXBookIECellAreaHelper.Create(AEditorArea: TAXWEditorArea);
begin
  FEditorArea := AEditorArea;
end;

end.
