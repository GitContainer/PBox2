unit XLSWriteXLSX5;

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
     xpgPUtils, xpgParseDocPropsApp, xpgParserXLSX, xpgParseXLinks, xpgParseDrawing,
     xpgPSimpleDOM, xpgParserPivot,
     Xc12Utils5, Xc12DefaultData5, Xc12DataStylesheet5, Xc12DataWorkbook5,
     Xc12DataWorksheet5, Xc12DataSST5, Xc12DataComments5, Xc12DataXLinks5,
     Xc12DataAutofilter5, Xc12DataTable5, Xc12FileData5, Xc12Manager5,
     XLSUtils5, XLSReadWriteOPC5, XLSCellMMU5, XLSMMU5, XLSFormula5;

type TXLSWriteXLSX = class;

     TXLSWriteSheetData = class(TObject)
protected
     FOwner    : TXLSWriteXLSX;
     FManager  : TXc12Manager;
     FCurrSheet: TXc12DataWorksheet;
     FSheetData: TCT_Sheetdata;
     FComments : TCT_Comments;
     FCurrCol  : integer;
     FCurrIndex: integer;

     procedure OnWriteRow(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteRowDirect(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteCell(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteCellDirect(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteCol(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteComment(ASender: TObject; var AWriteElement: boolean);
public
     constructor Create(AOwner: TXLSWriteXLSX; ASheet: TXc12DataWorksheet; ASheetData: TCT_Sheetdata; AComments : TCT_Comments);

     procedure Initiate;
     procedure WriteComments(AXML: TXPGDocXLSX; AOwner: TOPCItem);
     end;

     TXLSWriteXLSX = class(TObject)
protected
     FManager      : TXc12Manager;
     FFormulas     : TXLSFormulaHandler;
     FCurrSheet    : TXc12DataWorksheet;
     FOPCSheet     : TOPCItem;
     FCurrCount    : integer;
     FCurrIndex    : integer;
     FCurrMergeCell: integer;
     FCurrHyperlink: integer;

     FXML: TXPGDocXLSX;

     function  GetEnum(AValue: integer; ADefault: integer): integer;

     procedure GetRst(AText: AxUCString; ARuns: TXc12DynFontRunArray; APhonetics: TXc12DynPhoneticRunArray; ADestRst: TCT_Rst);

     procedure GetColor(ASrcColor: TXc12Color; ADestColor: TCT_Color);
     procedure GetBorder(ASrcBorder: TXc12BorderPr; ADestBorder: TCT_BorderPr);
     procedure GetBorders(ASrcBorder: TXc12Border; ADestBorder: TCT_Border);
     procedure GetFont(ASrcFont: TXc12Font; ADestFont: TCT_Font);
     procedure GetPrFont(ASrcFont: TXc12Font; ADestFont: TCT_RPrElt);
     procedure GetFill(ASrcFill: TXc12Fill; ADestFill: TCT_Fill);
     procedure GetAlignment(ASrcAlign: TXc12CellAlignment; ADestAlign: TCT_CellAlignment);
     procedure GetProtection(ASrcProtection: TXc12CellProtections; ADestProtection: TCT_CellProtection);

     procedure WriteDocPropsApp;
     procedure WriteMedia;
     procedure WriteSST;
     procedure WriteConnections;
     procedure WriteBorders;
     procedure WriteFills;
     procedure WriteFonts;
     procedure WriteNumberFormats;
     procedure WriteXFs(ASrc: TXc12XFs; ADest: TCT_XfXpgList);
     procedure WriteCellStyles;
     procedure WriteDXFs;
     procedure WriteTableStyles;
     procedure WriteColors;
     procedure WriteStyles;
     function  WriteDrawing(AOwner: TOPCItem): AxUCString;
     function  WriteVmlDrawing(AOwner: TOPCItem): AxUCString;
     procedure WriteSheet(ASheet: TXc12DataWorksheet);
     procedure WriteChartSheet(ASheet: TXc12DataWorksheet);
     procedure WriteSheetViews(ASheet: TXc12DataWorksheet);
     procedure WriteSheetCustomSheetViews(ACSViews: TXc12CustomSheetViews);
     procedure WriteSheetAutoFilter(ASrc: TXc12AutoFilter; ADest: TCT_AutoFilter);
     procedure WriteSheetSortState(ASrc: TXc12SortState; ADest: TCT_SortState);
     procedure WriteSheetPane(ASrc: TXc12Pane; ADest: TCT_Pane);
     procedure WriteSheetSelection(ASrc: TXc12Selection; ADest: TCT_Selection);
     procedure WriteSheetPageBreak(ASrc: TXc12PageBreaks; ADest: TCT_PageBreak);
     procedure WriteSheetPageMargins(ASrc: TXc12PageMargins; ADest: TCT_PageMargins);
     procedure WriteSheetPrintOptions(ASrc: TXc12PrintOptions; ADest: TCT_PrintOptions);
     procedure WriteSheetPageSetup(ASrc: TXc12PageSetup; ADest: TCT_PageSetup);
     procedure WriteSheetHeaderFooter(ASrc: TXc12HeaderFooter; ADest: TCT_HeaderFooter);
     function  WriteSheetTable(ASrc: TXc12Table; AOwner: TOPCItem): AxUCString;
     procedure WriteSheetCondFmt(ASrc: TXc12ConditionalFormatting; ADest: TCT_ConditionalFormatting);
     procedure WriteSheetCfvo(ASrc: TXc12Cfvo; ADest: TCT_Cfvo);
     procedure WriteSheetColors(ASrc: TXc12Colors; ADest: TCT_ColorXpgList);
     procedure WriteSheetSmartTags(ASrc: TXc12SmartTags; ADest: TCT_CellSmartTagsXpgList);
     function  WriteXLink(AIndex: integer): AxUCString;
     procedure BeginWriteWorkbook;
     procedure EndWriteWorkbook;

     procedure OnWriteSSTItem(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteDefinedName(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteDefinedNames(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteMergeCell(ASender: TObject; var AWriteElement: boolean);
     procedure OnWriteHyperlink(ASender: TObject; var AWriteElement: boolean);
public
     constructor Create(AManager: TXc12Manager; AFormulas: TXLSFormulaHandler);
     destructor Destroy; override;

     procedure SaveToStream(AStream: TStream);
     end;


implementation

{ TXLSReadXLSX }

constructor TXLSWriteXLSX.Create(AManager: TXc12Manager; AFormulas: TXLSFormulaHandler);
begin
  FManager := AManager;
  FFormulas := AFormulas;
  FXML := TXPGDocXLSX.Create;
  FXML.Worksheet.MergeCells.OnWriteMergeCell := OnWriteMergeCell;
  FXML.Worksheet.Hyperlinks.OnWriteHyperlink := OnWriteHyperlink;
end;

destructor TXLSWriteXLSX.Destroy;
begin
  FXML.Free;
  inherited;
end;

procedure TXLSWriteXLSX.EndWriteWorkbook;
var
  i: integer;
  S: AxUCString;
  OPC: TOPCItem;
  OPC2: TOPCItem;
  OPCSheets: array of TOPCItem;
  Stream: TMemoryStream;
  Strm: TStream;
  Sheet: TCT_Sheet;
  BView: TXc12BookView;
  BookView: TCT_BookView;
  CWV: TXc12CustomWorkbookView;
  CWView: TCT_CustomWorkbookView;
  PivotCache: TCT_PivotCache;
  STagType: TCT_SmartTagType;
  WPO: TXc12WebPublishObject;
  WebPubObj: TCT_WebPublishObject;
  FileRecoveryPr: TCT_FileRecoveryPr;
  PivTbl: TCT_pivotTableDefinition;
begin
  SetLength(OPCSheets,FManager.Worksheets.Count);
  for i := 0 to FManager.Worksheets.Count - 1 do begin
    if FManager.Worksheets[i].IsChartSheet then
      OPCSheets[i] := FManager.FileData.OPC.CreateChartSheet(i + 1)
    else
      OPCSheets[i] := FManager.FileData.OPC.CreateSheet(i + 1);
    Sheet := FXML.Workbook.Sheets.SheetXpgList.Add;
    Sheet.Name := FManager.Worksheets[i].Name;
    Sheet.SheetId := FManager.Worksheets[i].Index + 1;
    Sheet.State := TST_SheetState(FManager.Worksheets[i].State);
    Sheet.R_Id := OPCSheets[i].Id;
  end;

  for i := 0 to FManager.Workbook.PivotCaches.Count - 1 do begin
    PivTbl := FManager.FindPivotTable(FManager.Workbook.PivotCaches[i]);

    if PivTbl <> Nil then begin
      FManager.FileData.OPC.CreatePivotCacheDefinition(TOPCItem(PivTbl._OPC),i + 1);

      OPC := FManager.FileData.OPC.CreatePivotCacheDefinition(Nil,i + 1);
      OPC2 := FManager.FileData.OPC.CreatePivotCacheRecords(OPC,i + 1);

      Strm := FManager.FileData.OPC.ItemCreateStream(OPC2);
      try
        FManager.Workbook.PivotCaches.RecordsSaveToStream(Strm,i);
      finally
        FManager.FileData.OPC.ItemCloseStream(OPC2,Strm);
      end;

      FManager.Workbook.PivotCaches[i].R_Id := OPC2.Id;

      Strm := FManager.FileData.OPC.ItemCreateStream(OPC);
      try
        FManager.Workbook.PivotCaches.SaveToStream(Strm,i);
      finally
        FManager.FileData.OPC.ItemCloseStream(OPC,Strm);
      end;

      PivotCache := FXML.Workbook.PivotCaches.PivotCacheXpgList.Add;
      PivotCache.CacheId := FManager.Workbook.PivotCaches[i].CacheId;
      PivotCache.R_Id := OPC.Id;
    end;
  end;

  FXML.Workbook.FileVersion.AppName := FManager.Workbook.FileVersion.AppName;
  FXML.Workbook.FileVersion.LastEdited := FManager.Workbook.FileVersion.LastEdited;
  FXML.Workbook.FileVersion.LowestEdited := FManager.Workbook.FileVersion.LowestEdited;
  FXML.Workbook.FileVersion.RupBuild := FManager.Workbook.FileVersion.RupBuild;
  FXML.Workbook.FileVersion.CodeName := FManager.Workbook.FileVersion.CodeName;

  FXML.Workbook.FileSharing.ReadOnlyRecommended := FManager.Workbook.FileSharing.ReadOnlyReccomended;
  FXML.Workbook.FileSharing.UserName := FManager.Workbook.FileSharing.Username;
  FXML.Workbook.FileSharing.ReservationPassword := FManager.Workbook.FileSharing.ReservationPassword;

  FXML.Workbook.WorkbookPr.Date1904 := FManager.Workbook.WorkbookPr.Date1904;
  FXML.Workbook.WorkbookPr.ShowObjects := TST_Objects(FManager.Workbook.WorkbookPr.ShowObjects);

  FXML.Workbook.WorkbookPr.ShowBorderUnselectedTables := FManager.Workbook.WorkbookPr.ShowBorderUnselectedTables;
  FXML.Workbook.WorkbookPr.FilterPrivacy := FManager.Workbook.WorkbookPr.FilterPrivacy;
  FXML.Workbook.WorkbookPr.PromptedSolutions := FManager.Workbook.WorkbookPr.PromptedSolutions;
  FXML.Workbook.WorkbookPr.ShowInkAnnotation := FManager.Workbook.WorkbookPr.ShowInkAnnotation;
  FXML.Workbook.WorkbookPr.BackupFile := FManager.Workbook.WorkbookPr.BackupFile;
  FXML.Workbook.WorkbookPr.SaveExternalLinkValues := FManager.Workbook.WorkbookPr.SaveExternalLinkValues;
  FXML.Workbook.WorkbookPr.UpdateLinks := TST_UpdateLinks(FManager.Workbook.WorkbookPr.UpdateLinks);
  FXML.Workbook.WorkbookPr.CodeName := FManager.Workbook.WorkbookPr.CodeName;
  FXML.Workbook.WorkbookPr.HidePivotFieldList := FManager.Workbook.WorkbookPr.HidePivotFieldList;
  FXML.Workbook.WorkbookPr.ShowPivotChartFilter := FManager.Workbook.WorkbookPr.ShowPivotChartFilter;
  FXML.Workbook.WorkbookPr.AllowRefreshQuery := FManager.Workbook.WorkbookPr.AllowRefreshQuery;
  FXML.Workbook.WorkbookPr.PublishItems := FManager.Workbook.WorkbookPr.PublishItems;
  FXML.Workbook.WorkbookPr.CheckCompatibility := FManager.Workbook.WorkbookPr.CheckCompatibility;
  FXML.Workbook.WorkbookPr.AutoCompressPictures := FManager.Workbook.WorkbookPr.AutoCompressPictures;
  FXML.Workbook.WorkbookPr.RefreshAllConnections := FManager.Workbook.WorkbookPr.RefreshAllConnections;
  FXML.Workbook.WorkbookPr.DefaultThemeVersion := FManager.Workbook.WorkbookPr.DefaultThemeVersion;

  FXML.Workbook.WorkbookProtection.WorkbookPassword := FManager.Workbook.WorkbookProtection.WorkbookPassword;
  FXML.Workbook.WorkbookProtection.RevisionsPassword := FManager.Workbook.WorkbookProtection.RevisionPassword;
  FXML.Workbook.WorkbookProtection.LockStructure := FManager.Workbook.WorkbookProtection.LockStructure;
  FXML.Workbook.WorkbookProtection.LockWindows := FManager.Workbook.WorkbookProtection.LockWindows;
  FXML.Workbook.WorkbookProtection.LockRevision := FManager.Workbook.WorkbookProtection.LockRevision;

  for i := 0 to FManager.Workbook.BookViews.Count - 1 do begin
    BView := FManager.Workbook.BookViews[i];
    BookView := FXML.Workbook.BookViews.WorkbookViewXpgList.Add;
    BookView.Visibility := TST_Visibility(BView.Visibility);
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

  FXML.Workbook.FunctionGroups.BuiltInGroupCount := FManager.Workbook.FunctionGroups.BuiltInGroupCount;
  for i := 0 to FManager.Workbook.FunctionGroups.Count - 1 do
    FXML.Workbook.FunctionGroups.FunctionGroupXpgList.Add.Name := FManager.Workbook.FunctionGroups[i];

  for i := 0 to FManager.XLinks.Count - 1 do begin
    S := WriteXLink(i + 1);
    FXML.Workbook.ExternalReferences.ExternalReferenceXpgList.Add.R_Id := S;
  end;

  // Shall be zero, otherwhise calculations may not work correct on some excel versions.
  FXML.Workbook.CalcPr.CalcId := 0; //FManager.Workbook.CalcPr.CalcId;
  FXML.Workbook.CalcPr.CalcMode := TST_CalcMode(FManager.Workbook.CalcPr.CalcMode);
  FXML.Workbook.CalcPr.FullCalcOnLoad := FManager.Workbook.CalcPr.FullCalcOnLoad;
  FXML.Workbook.CalcPr.RefMode := TST_RefMode(FManager.Workbook.CalcPr.RefMode);
  FXML.Workbook.CalcPr.Iterate := FManager.Workbook.CalcPr.Iterate;
  FXML.Workbook.CalcPr.IterateCount := FManager.Workbook.CalcPr.IterateCount;
  FXML.Workbook.CalcPr.IterateDelta := FManager.Workbook.CalcPr.IterateDelta;
  FXML.Workbook.CalcPr.FullPrecision := FManager.Workbook.CalcPr.FullPrecision;
  FXML.Workbook.CalcPr.CalcCompleted := FManager.Workbook.CalcPr.CalcCompleted;
  FXML.Workbook.CalcPr.CalcOnSave := FManager.Workbook.CalcPr.CalcOnSave;
  FXML.Workbook.CalcPr.ConcurrentCalc := FManager.Workbook.CalcPr.ConcurrentCalc;
  FXML.Workbook.CalcPr.ConcurrentManualCount := FManager.Workbook.CalcPr.ConcurrentManualCount;
  FXML.Workbook.CalcPr.ForceFullCalc := FManager.Workbook.CalcPr.ForceFullCalc;

  // Misstake. Shall not be event.
  if FManager.Workbook.DefinedNames.Count > 0 then
    FXML.Workbook.OnWriteDefinedNames := OnWriteDefinedNames;
  // Be aware. FCurrCount/FCurrIndex can only be used once per root element.
  FCurrCount := FManager.Workbook.DefinedNames.Count;
  FCurrIndex := 0;
  FXML.Workbook.DefinedNames.OnWriteDefinedName := OnWriteDefinedName;

  FXML.Workbook.OleSize.Ref := CellAreaToAreaStr(FManager.Workbook.OleSize);

  for i := 0 to FManager.Workbook.CustomWorkbookViews.Count - 1 do begin
    CWView := FXML.Workbook.CustomWorkbookViews.CustomWorkbookViewXpgList.Add;
    CWV := FManager.Workbook.CustomWorkbookViews[i];

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
    CWView.ShowComments := TST_Comments(CWV.ShowComments);
    CWView.ShowObjects := TST_Objects(CWV.ShowObjects);
  end;

  FXML.Workbook.SmartTagPr.Embed := FManager.Workbook.SmartTagPr.Embed;
  FXML.Workbook.SmartTagPr.Show := TST_SmartTagShow(FManager.Workbook.SmartTagPr.Show);

  for i := 0 to FManager.Workbook.SmartTagTypes.Count - 1 do begin
    STagType := FXML.Workbook.SmartTagTypes.SmartTagTypeXpgList.Add;

    STagType.NamespaceUri := FManager.Workbook.SmartTagTypes[i].NamespaceUri;
    STagType.Name := FManager.Workbook.SmartTagTypes[i].Name;
    STagType.Url := FManager.Workbook.SmartTagTypes[i].Url;
  end;

  FXML.Workbook.WebPublishing.Css := FManager.Workbook.WebPublishing.Css;
  FXML.Workbook.WebPublishing.Thicket := FManager.Workbook.WebPublishing.Thicket;
  FXML.Workbook.WebPublishing.LongFileNames := FManager.Workbook.WebPublishing.LongFileNames;
  FXML.Workbook.WebPublishing.Vml := FManager.Workbook.WebPublishing.Vml;
  FXML.Workbook.WebPublishing.AllowPng := FManager.Workbook.WebPublishing.AllowPng;
  FXML.Workbook.WebPublishing.TargetScreenSize := TST_TargetScreenSize(FManager.Workbook.WebPublishing.TargetScreenSize);
  FXML.Workbook.WebPublishing.Dpi := FManager.Workbook.WebPublishing.Dpi;
  FXML.Workbook.WebPublishing.CodePage := FManager.Workbook.WebPublishing.CodePage;

  for i := 0 to FManager.Workbook.FileRecoveryPrs.Count - 1 do begin
    FileRecoveryPr := FXML.Workbook.FileRecoveryPrXpgList.Add;
    FileRecoveryPr.AutoRecover := FManager.Workbook.FileRecoveryPrs[i].AutoRecover;
    FileRecoveryPr.CrashSave := FManager.Workbook.FileRecoveryPrs[i].CrashSave;
    FileRecoveryPr.DataExtractLoad := FManager.Workbook.FileRecoveryPrs[i].DataExtractLoad;
    FileRecoveryPr.RepairLoad := FManager.Workbook.FileRecoveryPrs[i].RepairLoad;
  end;

  for i := 0 to FManager.Workbook.WebPublishObjects.Count - 1 do begin
    WebPubObj := FXML.Workbook.WebPublishObjects.WebPublishObjectXpgList.Add;
    WPO := FManager.Workbook.WebPublishObjects[i];

    WebPubObj.Id := WPO.Id;
    WebPubObj.DivId := WPO.DivId;
    WebPubObj.SourceObject := WPO.SourceObject;
    WebPubObj.DestinationFile := WPO.DestinationFile;
    WebPubObj.Title := WPO.Title;
    WebPubObj.AutoRepublish := WPO.AutoRepublish;
  end;

  FXML.Root.RootAttributes.Clear;
  FXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);
  FXML.Root.RootAttributes.AddNameValue('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);

  Stream := TMemoryStream.Create;
  try
    FXML.SaveToStream(Stream,TCT_Workbook);
    FManager.FileData.OPC.ItemWrite(FManager.FileData.OPC.Workbook,Stream);
//    FXLS.FileDataXLSX.OPC.EndAddWorkbook(Stream);
  finally
    Stream.Free;
  end;

  FXML.Workbook.Clear;
end;

procedure TXLSWriteXLSX.GetAlignment(ASrcAlign: TXc12CellAlignment; ADestAlign: TCT_CellAlignment);
begin
  if ASrcAlign.HorizAlignment <> chaGeneral then
    ADestAlign.Horizontal := TST_HorizontalAlignment(GetEnum(Integer(ASrcAlign.HorizAlignment),Integer(chaGeneral)));
  if ASrcAlign.VertAlignment <> cvaBottom then
    ADestAlign.Vertical := TST_VerticalAlignment(GetEnum(Integer(ASrcAlign.VertAlignment),Integer(cvaBottom)));
  ADestAlign.TextRotation := ASrcAlign.Rotation;
  ADestAlign.WrapText := foWrapText in ASrcAlign.Options;
  ADestAlign.ShrinkToFit := foShrinkToFit in ASrcAlign.Options;
  ADestAlign.JustifyLastLine := foJustifyLastLine in ASrcAlign.Options;
  ADestAlign.Indent := ASrcAlign.Indent;
  ADestAlign.RelativeIndent := ASrcAlign.RelativeIndent;
  ADestAlign.ReadingOrder := Integer(ASrcAlign.ReadingOrder);
end;

procedure TXLSWriteXLSX.GetBorders(ASrcBorder: TXc12Border; ADestBorder: TCT_Border);
begin
  ADestBorder.DiagonalUp := ecboDiagonalUp in ASrcBorder.Options;
  ADestBorder.DiagonalDown := ecboDiagonalDown in ASrcBorder.Options;
  ADestBorder.Outline := ecboOutline in ASrcBorder.Options;
  GetBorder(ASrcBorder.Left,ADestBorder.Left);
  GetBorder(ASrcBorder.Right,ADestBorder.Right);
  GetBorder(ASrcBorder.Top,ADestBorder.Top);
  GetBorder(ASrcBorder.Bottom,ADestBorder.Bottom);
  GetBorder(ASrcBorder.Diagonal,ADestBorder.Diagonal);
  GetBorder(ASrcBorder.Vertical,ADestBorder.Vertical);
  GetBorder(ASrcBorder.Horizontal,ADestBorder.Horizontal);
end;

procedure TXLSWriteXLSX.GetBorder(ASrcBorder: TXc12BorderPr; ADestBorder: TCT_BorderPr);
begin
  if ASrcBorder.Style <> cbsNone then
    ADestBorder.Style := TST_BorderStyle(ASrcBorder.Style);
  GetColor(ASrcBorder.Color,ADestBorder.Color);
end;

procedure TXLSWriteXLSX.GetColor(ASrcColor: TXc12Color; ADestColor: TCT_Color);
begin
  case ASrcColor.ColorType of
    exctAuto      : ADestColor.Auto := True;
    exctIndexed   : ADestColor.Indexed := Integer(ASrcColor.Indexed);
    exctRgb       : ADestColor.Rgb := ASrcColor.OrigRGB;
    exctTheme     : ADestColor.Theme := Integer(ASrcColor.Theme);
    exctUnassigned: ;
  end;
  ADestColor.Tint := ASrcColor.Tint;
end;

function TXLSWriteXLSX.GetEnum(AValue, ADefault: integer): integer;
begin
  if AValue <> ADefault then
    Result := AValue
  else
    Result := XPG_UNKNOWN_ENUM;
end;

procedure TXLSWriteXLSX.GetFill(ASrcFill: TXc12Fill; ADestFill: TCT_Fill);
var
  i: integer;
  Stop: xpgParserXLSX.TCT_GradientStop;
begin
  if ASrcFill.IsGradientFill then begin
    for i := 0 to ASrcFill.GradientFill.Stops.Count - 1 do begin
      Stop := ADestFill.GradientFill.StopXpgList.Add;
      Stop.Position := ASrcFill.GradientFill.Stops[i].Position;
      GetColor(ASrcFill.GradientFill.Stops[i].Color,Stop.Color);
    end;
    if ASrcFill.GradientFill.GradientType <> gftLinear then
      ADestFill.GradientFill.Type_ := TST_GradientType(GetEnum(Integer(ASrcFill.GradientFill.GradientType),Integer(gftLinear)));
    ADestFill.GradientFill.Degree := ASrcFill.GradientFill.Degree;
    ADestFill.GradientFill.Left := ASrcFill.GradientFill.Left;
    ADestFill.GradientFill.Right := ASrcFill.GradientFill.Right;
    ADestFill.GradientFill.Top := ASrcFill.GradientFill.Top;
    ADestFill.GradientFill.Bottom := ASrcFill.GradientFill.Bottom;
  end
  else begin
    ADestFill.PatternFill.PatternType := TST_PatternType(GetEnum(Integer(ASrcFill.PatternType),Integer(efpNone)));
    GetColor(ASrcFill.FgColor,ADestFill.PatternFill.FgColor);
    GetColor(ASrcFill.BgColor,ADestFill.PatternFill.BgColor);
  end;
end;

procedure TXLSWriteXLSX.GetFont(ASrcFont: TXc12Font; ADestFont: TCT_Font);
begin
  ADestFont.Charset.Val := ASrcFont.Charset;
  GetColor(ASrcFont.Color,ADestFont.Color);
  ADestFont.Condense.Val := xfsCondense in ASrcFont.Style;
  ADestFont.Extend.Val := xfsExtend in ASrcFont.Style;
  ADestFont.Family.Val := ASrcFont.Family;
  ADestFont.Name.Val := ASrcFont.Name;
  ADestFont.Outline.Val := xfsOutline in ASrcFont.Style;
  ADestFont.Shadow.Val := xfsShadow in ASrcFont.Style;

//  if ASrcFont.Scheme <> efsNone then
  ADestFont.Scheme.Val := TST_FontScheme(ASrcFont.Scheme);

  case ASrcFont.Underline of
    xulNone         : ADestFont.U.Val := stuvNone;
    xulSingle       : ADestFont.U.Val := stuvSingle;
    xulDouble       : ADestFont.U.Val := stuvDouble;
    xulSingleAccount: ADestFont.U.Val := stuvSingleAccounting;
    xulDoubleAccount: ADestFont.U.Val := stuvDoubleAccounting;
  end;
  if ASrcFont.SubSuperScript <> xssNone then
    ADestFont.VertAlign.Val := TST_VerticalAlignRun(ASrcFont.SubSuperScript);
  ADestFont.Sz.Val := ASrcFont.Size;
  ADestFont.B.Val := xfsBold in ASrcFont.Style;
  ADestFont.I.Val := xfsItalic in ASrcFont.Style;
  ADestFont.Strike.Val := xfsStrikeOut in ASrcFont.Style;
end;

procedure TXLSWriteXLSX.GetPrFont(ASrcFont: TXc12Font; ADestFont: TCT_RPrElt);
begin
  ADestFont.Charset.Val := ASrcFont.Charset;
  GetColor(ASrcFont.Color,ADestFont.Color);
  ADestFont.Condense.Val := xfsCondense in ASrcFont.Style;
  ADestFont.Extend.Val := xfsExtend in ASrcFont.Style;
  ADestFont.Family.Val := ASrcFont.Family;
  ADestFont.rFont.Val := ASrcFont.Name;
  ADestFont.Outline.Val := xfsOutline in ASrcFont.Style;
  ADestFont.Shadow.Val := xfsShadow in ASrcFont.Style;
  if ASrcFont.Scheme <> efsNone then
    ADestFont.Scheme.Val := TST_FontScheme(ASrcFont.Scheme);
  if ASrcFont.Underline = xulNone then
    ADestFont.U.Val := stuvNone
  else
    ADestFont.U.Val := TST_UnderlineValues(Integer(ASrcFont.Underline) - 1);
  if ASrcFont.SubSuperScript <> xssNone then
    ADestFont.VertAlign.Val := TST_VerticalAlignRun(ASrcFont.SubSuperScript);
  ADestFont.Sz.Val := ASrcFont.Size;
  ADestFont.B.Val := xfsBold in ASrcFont.Style;
  ADestFont.I.Val := xfsItalic in ASrcFont.Style;
  ADestFont.Strike.Val := xfsStrikeOut in ASrcFont.Style;
end;

procedure TXLSWriteXLSX.GetProtection(ASrcProtection: TXc12CellProtections; ADestProtection: TCT_CellProtection);
begin
  ADestProtection.Locked := cpLocked in ASrcProtection;
  ADestProtection.Hidden := cpHidden in ASrcProtection;
end;

procedure TXLSWriteXLSX.GetRst(AText: AxUCString; ARuns: TXc12DynFontRunArray; APhonetics: TXc12DynPhoneticRunArray; ADestRst: TCT_Rst);

procedure GetRichStr(F: TXc12Font; S: AxUCString);
var
  RElt: TCT_RElt;
begin
  RElt := ADestRst.RXpgList.Add;
  if F <> Nil then
    GetPrFont(F,RElt.RPr);
  RElt.T := S;
end;

procedure GetRichXML(S: AxUCString);
var
  i: integer;
  F1,F2: TXc12FontRun;
begin
  F1 := ARuns[0];
  for i := 1 to High(ARuns) do begin
    F2 := ARuns[i];
    GetRichStr(F1.Font,Copy(S,F1.Index + 1,F2.Index - F1.Index));
    F1 := F2;
  end;
  GetRichStr(F1.Font,Copy(S,F1.Index + 1,Length(S) - F1.Index));
end;

procedure WritePhonetic;
var
  i: integer;
  rPh: TCT_PhoneticRun;
begin
  for i := 0 to High(APhonetics) do begin
    rPh := ADestRst.RPhXpgList.Add;
    rPh.Sb := APhonetics[i].Sb;
    rPh.Eb := APhonetics[i].Eb;
    rPh.T := APhonetics[i].Text;
  end;
end;

begin
  if (Length(ARuns) <= 0) then
    ADestRst.T := AText
  else
    GetRichXML(AText);
end;

procedure TXLSWriteXLSX.OnWriteDefinedName(ASender: TObject; var AWriteElement: boolean);
var
  N   : TXc12DefinedName;
  Name: TCT_DefinedName;
  CS  : AxUCChar;
begin
  AWriteElement := FCurrIndex < FCurrCount;
  if AWriteElement then begin
    N := FManager.Workbook.DefinedNames[FCurrIndex];
    if not N.Deleted then begin
      Name := TCT_DefinedName(ASender);
      Name.Clear;

      case N.BuiltIn of
        bnConsolidateArea: Name.Name := '_xlnm.Consolidate_Area';
        bnExtract        : Name.Name := '_xlnm.Extract';
        bnDatabase       : Name.Name := '_xlnm.Database';
        bnCriteria       : Name.Name := '_xlnm.Criteria';
        bnPrintArea      : Name.Name := '_xlnm.Print_Area';
        bnPrintTitles    : Name.Name := '_xlnm.Print_Titles';
        bnSheetTitle     : Name.Name := '_xlnm.Sheet_Title';
        bnFilterDatabase : Name.Name := '_xlnm._FilterDatabase';
        bnNone           : Name.Name := N.Name;
      end;

      Name.Hidden := N.Hidden;
      Name.Function_ := N.Function_;
      Name.VBProcedure := N.VBProcedure;
      Name.Xlm := N.Xlm;
      Name.PublishToServer := N.PublishToServer;
      Name.WorkbookParameter := N.WorkbookParameter;

      Name.Comment := N.Comment;
      Name.CustomMenu := N.CustomMenu;
      Name.Description := N.Description;
      Name.Help := N.Help;
      Name.StatusBar := N.StatusBar;
      if (N.LocalSheetId >= 0) and (N.LocalSheetId < $FFFF) then
        Name.LocalSheetId := N.LocalSheetId;

      CS := FManager.ColSeparator;
      FManager.ColSeparator := ',';
      Name.Content := N.Content;
      FManager.ColSeparator := CS;
    end;
    Inc(FCurrIndex);
  end;
end;

procedure TXLSWriteXLSX.OnWriteDefinedNames(ASender: TObject; var AWriteElement: boolean);
begin
  AWriteElement := FCurrIndex = 0;
end;

procedure TXLSWriteXLSX.OnWriteHyperlink(ASender: TObject; var AWriteElement: boolean);
var
  H: TXc12Hyperlink;
  HLink: TCT_Hyperlink;
  OPC: TOPCItem;
begin
  if ASender = Nil then
    AWriteElement := FCurrSheet.Hyperlinks.Count > 0
  else begin
    AWriteElement := FCurrHyperlink < FCurrSheet.Hyperlinks.Count;
    if AWriteElement then begin
      H := FCurrSheet.Hyperlinks[FCurrHyperlink];
      HLink := TCT_Hyperlink(ASender);
      HLink.Ref := CellAreaToAreaStr(H.Ref);
      HLink.Tooltip := H.ToolTip;
      HLink.Display := H.Display;

      if H.HyperlinkType = xhltWorkbook then begin
        HLink.Location := H.RawAddress;
        HLink.R_Id := '';
      end
      else begin
        HLink.Location := H.Location;
        OPC := FManager.FileData.OPC.CreateHyperlink(FOPCSheet,H.RawAddress);
        HLink.R_Id := OPC.Id;
        OPC.Close;
      end;

      Inc(FCurrHyperlink);
    end;
  end;
end;

procedure TXLSWriteXLSX.OnWriteMergeCell(ASender: TObject; var AWriteElement: boolean);
begin
  if ASender = Nil then
    AWriteElement := FCurrSheet.MergedCells.Count > 0
  else begin
    AWriteElement := FCurrMergeCell < FCurrSheet.MergedCells.Count;
    if AWriteElement then begin
      TCT_MergeCell(ASender).Ref := CellAreaToAreaStr(FCurrSheet.MergedCells[FCurrMergeCell].Ref);
      Inc(FCurrMergeCell);
    end;
  end;
end;

procedure TXLSWriteXLSX.OnWriteSSTItem(ASender: TObject; var AWriteElement: boolean);
var
  S: AxUCString;
  S8: XLS8String;
  si: TCT_Rst;
  XStr: PXLSString;
  FontRuns: TXc12DynFontRunArray;
  PhoneticRuns: TXc12DynPhoneticRunArray;

function CvtFontRuns(ACount: integer; ARuns: PXc12FontRunArray): TXc12DynFontRunArray;
var
  i,j: integer;
begin
  if ARuns[0].Index = 0 then begin
    SetLength(Result,ACount);
    j := 0;
  end
  else begin
    SetLength(Result,ACount + 1);
    j := 1;
    Result[0].Index := 0;
    Result[0].Font := Nil;
  end;

  for i := 0 to ACount - 1 do begin
    Result[i + j].Index := ARuns[i].Index;
    Result[i + j].Font := ARuns[i].Font;
  end;
end;

begin
  AWriteElement := FCurrIndex < FCurrCount;
  if AWriteElement then begin
    si := TCT_Rst(ASender);

    XStr := FManager.SST[FCurrIndex];
    if XStr = Nil then
      si.T := ''
    else begin
      case XStr.Options of
        STRID_COMPRESSED: begin
          SetLength(S8,XStr.Len);
          Move(PByteArray(@XStr.Data)^,Pointer(S8)^,XStr.Len);
          S := AxUCString(S8);
//          si.T := StripCRLF(S);
          si.T := S;
        end;
        STRID_UNICODE: begin
          SetLength(S,XStr.Len);
          Move(PByteArray(@XStr.Data)^,Pointer(S)^,XStr.Len * 2);
          si.T := S;
//          si.T := StripCRLF(S);
        end;
        STRID_RICH: begin
          SetLength(S8,XStr.Len);
          with PXLSStringRich(XStr)^ do
            Move(PByteArray(@Data)^,Pointer(S8)^,XStr.Len);
          S := AxUCString(S8);
          // TODO
          SetLength(PhoneticRuns,0);

          FontRuns := CvtFontRuns(PXLSStringRichUC(XStr).FormatCount,FManager.SST.GetFontRuns(XStr));
          GetRst(S,FontRuns,PhoneticRuns,si);
        end;
        STRID_RICH_UNICODE: begin
          SetLength(S,XStr.Len);
          with PXLSStringRichUC(XStr)^ do
            Move(PByteArray(@Data)^,Pointer(S)^,XStr.Len * 2);
          // TODO
          SetLength(PhoneticRuns,0);

          FontRuns := CvtFontRuns(PXLSStringRichUC(XStr).FormatCount,FManager.SST.GetFontRuns(XStr));
          GetRst(S,FontRuns,PhoneticRuns,si);
        end;
        STRID_FAREAST: begin
          SetLength(S8,XStr.Len);
          with PXLSStringFarEast(XStr)^ do
            Move(PByteArray(@Data)^,Pointer(S8)^,XStr.Len);
          si.T := AxUCString(S8);
        end;
        STRID_FAREAST_RICH: begin
          SetLength(S8,XStr.Len);
          with PXLSStringFarEastRich(XStr)^ do
            Move(PByteArray(@Data)^,Pointer(S8)^,XStr.Len);
          // TODO
          SetLength(PhoneticRuns,0);

          FontRuns := CvtFontRuns(PXLSStringRichUC(XStr).FormatCount,FManager.SST.GetFontRuns(XStr));
          GetRst(S,FontRuns,PhoneticRuns,si);
        end;
        STRID_FAREAST_UC: begin
          SetLength(S,XStr.Len);
          with PXLSStringFarEastUC(XStr)^ do
            Move(PByteArray(@Data)^,Pointer(S)^,XStr.Len * 2);
          si.T := S;
        end;
        STRID_FAREAST_RICH_UC: begin
          SetLength(S,XStr.Len);
          with PXLSStringFarEastRichUC(XStr)^ do
            Move(PByteArray(@Data)^,Pointer(S)^,XStr.Len * 2);
          // TODO
          SetLength(PhoneticRuns,0);

          FontRuns := CvtFontRuns(PXLSStringRichUC(XStr).FormatCount,FManager.SST.GetFontRuns(XStr));
          GetRst(S,FontRuns,PhoneticRuns,si);
        end;
      end;
    end;
  end;
  Inc(FCurrIndex);
end;

procedure TXLSWriteXLSX.SaveToStream(AStream: TStream);
var
  i: integer;
  TempDS: AxUCChar;
  TempLS: AxUCChar;
begin
  TempLS := FormatSettings.ListSeparator;
  FormatSettings.ListSeparator := ',';
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    FManager.FileData.BeginSaveToStream(AStream);

    FManager.Workbook.PivotCaches.Enumerate;

    BeginWriteWorkbook;

    WriteDocPropsApp;

    WriteStyles;

  //  FManager.Worksheets.Enumerate;
    for i := 0 to FManager.Worksheets.Count - 1 do begin
      if FManager.Worksheets[i].IsChartSheet then
        WriteChartSheet(FManager.Worksheets[i])
      else
        WriteSheet(FManager.Worksheets[i]);
    end;

    WriteSST;

    WriteConnections;

    EndWriteWorkbook;

    FManager.FileData.CommitSaveToStream;
  finally
    FormatSettings.DecimalSeparator := TempDS;
    FormatSettings.ListSeparator := TempLS;
  end;
end;

procedure TXLSWriteXLSX.WriteDocPropsApp;
var
  i: integer;
  Props: TCT_Properties;
  V: TCT_Variant;
  Stream: TMemoryStream;
begin
  Props := FManager.FileData.DocPropsApp.Properties;

  Props.Application := 'Microsoft Excel';
  Props.AppVersion := '12.000';

  Props.HeadingPairs.Vt_Vector.Size := 2;
  Props.HeadingPairs.Vt_Vector.BaseType := stvbtVariant;
  Props.HeadingPairs.Vt_Vector.CreateTCT_VariantXpgList;
  V := Props.HeadingPairs.Vt_Vector.Vt_VariantXpgList.Add;
  V.Vt_Lpstr := 'Worksheets';
  V := Props.HeadingPairs.Vt_Vector.Vt_VariantXpgList.Add;
  V.Vt_i4 := FManager.Worksheets.Count;

  Props.TitlesOfParts.Vt_Vector.Size := FManager.Worksheets.Count;
  Props.TitlesOfParts.Vt_Vector.BaseType := stvbtLpstr;
  Props.TitlesOfParts.Vt_Vector.Vt_LpstrXpgList.Clear;
  for i := 0 to FManager.Worksheets.Count - 1 do
    Props.TitlesOfParts.Vt_Vector.Vt_LpstrXpgList.Add(FManager.Worksheets[i].Name);

  Stream := TMemoryStream.Create;
  try
    FManager.FileData.DocPropsApp.SaveToStream(Stream);
    FManager.FileData.OPC.AddDocPropsApp(Stream);
  finally
    Stream.Free;
  end;
end;

function TXLSWriteXLSX.WriteDrawing(AOwner: TOPCItem): AxUCString;
var
  OPC: TOPCItem;
  Stream: TMemoryStream;
begin
  OPC := FManager.FileData.OPC.CreateDrawing(AOwner,FCurrSheet.Index + 1);
  Result := OPC.Id;
  FManager.GrManager.XLSOPC := FManager.FileData.OPC;

  FCurrSheet.Drawing.Root.RootAttributes.Clear;
  FCurrSheet.Drawing.Root.RootAttributes.AddNameValue('xmlns:xdr','http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing');
  FCurrSheet.Drawing.Root.RootAttributes.AddNameValue('xmlns:a','http://schemas.openxmlformats.org/drawingml/2006/main');

  Stream := TMemoryStream.Create;
  try

    FCurrSheet.Drawing.SaveToStream(Stream);

    FManager.FileData.OPC.ItemWrite(OPC,Stream);

    FManager.FileData.OPC.ItemClose(OPC);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWriteXLSX.WriteDXFs;
var
  i: integer;
  D: TXc12DXF;
  DXF: TCT_DXF;
begin
  FXML.StyleSheet.Dxfs.Count_ := FManager.StyleSheet.DXFs.Count;
  for i := 0 to FManager.StyleSheet.DXFs.Count - 1 do begin
    D := FManager.StyleSheet.DXFs[i];
    DXF := FXML.StyleSheet.Dxfs.DxfXpgList.Add;
    if D.Font <> Nil then
      GetFont(D.Font as TXc12Font,DXF.Font);
    if (D.NumFmt <> Nil) and (D.NumFmt.Value <> '') then begin
      DXF.NumFmt.NumFmtId := D.NumFmt.Index;
      DXF.NumFmt.FormatCode := D.NumFmt.Value;
    end;
    if D.Fill <> Nil then
      GetFill(D.Fill as TXc12Fill,DXF.Fill);
    if D.Alignment <> Nil then
      GetAlignment(D.Alignment as TXc12CellAlignment,DXF.Alignment);
    if D.Border <> Nil then
      GetBorders(D.Border,DXF.Border);
    GetProtection(D.Protection,DXF.Protection);
  end;
end;

procedure TXLSWriteXLSX.WriteSheet(ASheet: TXc12DataWorksheet);
var
  i,j           : integer;
  S             : AxUCString;
  Stream        : TStream;
  SheetDataWrite: TXLSWriteSheetData;
  ProtectedRange: TCT_ProtectedRange;
  Scenario      : TCT_Scenario;
  Scen          : TXc12Scenario;
  InputCell     : TCT_InputCells;
  DataRef       : TCT_DataRef;
  DValidation   : TXc12DataValidation;
  DataValidation: TCT_DataValidation;
  Prop          : TCT_CustomProperty;
  IgnoredError  : TCT_IgnoredError;
  IError        : TXc12IgnoredError;
  OleObject     : TCT_OleObject;
  Control       : TCT_Control;
  WebPublishItem: TCT_WebPublishItem;
  WPItem        : TXc12WebPublishItem;
  TablePart     : TCT_TablePart;
  OPC           : TOPCItem;
begin
  FCurrSheet := ASheet;
  FCurrMergeCell := 0;
  FCurrHyperlink := 0;

  FOPCSheet := FManager.FileData.OPC.CreateSheet(ASheet.Index + 1);

  FManager.FileData.WriteUnusedDataSheet(ASheet.Index,FOPCSheet);

  // Write substreams before sheet stream
  if FCurrSheet.Drawing.WsDr.HasData then
    FXML.Worksheet.Drawing.R_Id := WriteDrawing(FOPCSheet);

  if not FCurrSheet.VmlDrawing.Empty or (FCurrSheet.Comments.Count > 0) then
    FXML.Worksheet.LegacyDrawing.R_Id := WriteVmlDrawing(FOPCSheet);

  if ASheet.PageSetup.PrinterSettings <> Nil then
    FXML.Worksheet.PageSetup.R_Id := FManager.FileData.OPC.WritePrinterSettings(FOPCSheet,ASheet.PageSetup.PrinterSettings);

  if FCurrSheet.Comments.Count > 0 then begin
    SheetDataWrite := TXLSWriteSheetData.Create(Self,ASheet,FXML.Worksheet.SheetData,FXML.Comments);
    try
      SheetDataWrite.WriteComments(FXML,FOPCSheet);
    finally
      SheetDataWrite.Free;
    end;
  end;

  for i := 0 to ASheet.Tables.Count - 1 do begin
    S := WriteSheetTable(ASheet.Tables[i],FOPCSheet);
    TablePart := FXML.Worksheet.TableParts.TablePartXpgList.Add;
    TablePart.R_Id := S;
  end;

  for i := 0 to ASheet.PivotTables.Count - 1 do begin
    OPC := FManager.FileData.OPC.CreatePivotTable;

    ASheet.PivotTables[i].CacheId := ASheet.PivotTables[i].Cache.CacheId;

    Stream := FManager.FileData.OPC.ItemCreateStream(OPC);
    try
      ASheet.PivotTables.SaveToStream(Stream,i,OPC);
    finally
      FManager.FileData.OPC.ItemCloseStream(OPC,Stream);
    end;
  end;


  // End Write substreams

  FXML.Root.RootAttributes.Clear;
  FXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);
  FXML.Root.RootAttributes.AddNameValue('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);

  FXML.Worksheet.Dimension.Ref := CellAreaToAreaStr(ASheet.Dimension);

  Stream := FManager.FileData.OPC.ItemCreateStream(FOPCSheet);

  FXML.Worksheet.SheetPr.SyncHorizontal := ASheet.SheetPr.SyncHorizontal;
  FXML.Worksheet.SheetPr.SyncVertical := ASheet.SheetPr.SyncVertical;
  FXML.Worksheet.SheetPr.SyncRef := CellAreaToAreaStr(ASheet.SheetPr.SyncRef);
  FXML.Worksheet.SheetPr.TransitionEvaluation := ASheet.SheetPr.TransitionEvaluation;
  FXML.Worksheet.SheetPr.TransitionEntry := ASheet.SheetPr.TransitionEntry;
  FXML.Worksheet.SheetPr.Published_ := ASheet.SheetPr.Published_;
  FXML.Worksheet.SheetPr.CodeName := ASheet.SheetPr.CodeName;
  FXML.Worksheet.SheetPr.FilterMode := ASheet.SheetPr.FilterMode;
  FXML.Worksheet.SheetPr.EnableFormatConditionsCalculation := ASheet.SheetPr.EnableFormatConditionsCalculation;
  if ASheet.SheetPr.TabColor.ColorType <> exctUnassigned then
    GetColor(ASheet.SheetPr.TabColor,FXML.Worksheet.SheetPr.TabColor);

  FXML.Worksheet.SheetPr.OutlinePr.ApplyStyles := ASheet.SheetPr.OutlinePr.ApplyStyles;
  FXML.Worksheet.SheetPr.OutlinePr.SummaryBelow := ASheet.SheetPr.OutlinePr.SummaryBelow;
  FXML.Worksheet.SheetPr.OutlinePr.SummaryRight := ASheet.SheetPr.OutlinePr.SummaryRight;
  FXML.Worksheet.SheetPr.OutlinePr.ShowOutlineSymbols := ASheet.SheetPr.OutlinePr.ShowOutlineSymbols;

  FXML.Worksheet.SheetPr.PageSetUpPr.AutoPageBreaks := ASheet.SheetPr.PageSetUpPr.AutoPageBreaks;
  FXML.Worksheet.SheetPr.PageSetUpPr.FitToPage := ASheet.SheetPr.PageSetUpPr.FitToPage;

  if ASheet.SheetViews.Count > 0 then
    WriteSheetViews(ASheet);

  FXML.Worksheet.SheetFormatPr.BaseColWidth := ASheet.SheetFormatPr.BaseColWidth;
  if ASheet.Columns.DefColWidth > 0 then
    FXML.Worksheet.SheetFormatPr.DefaultColWidth := ASheet.Columns.DefColWidth
  else
    FXML.Worksheet.SheetFormatPr.DefaultColWidth := ASheet.SheetFormatPr.DefaultColWidth;
  FXML.Worksheet.SheetFormatPr.DefaultRowHeight := ASheet.SheetFormatPr.DefaultRowHeight / 20;
  FXML.Worksheet.SheetFormatPr.CustomHeight := ASheet.SheetFormatPr.CustomHeight;
  FXML.Worksheet.SheetFormatPr.ZeroHeight := ASheet.SheetFormatPr.ZeroHeight;
  FXML.Worksheet.SheetFormatPr.ThickTop := ASheet.SheetFormatPr.ThickTop;
  FXML.Worksheet.SheetFormatPr.ThickBottom := ASheet.SheetFormatPr.ThickBottom;
  FXML.Worksheet.SheetFormatPr.OutlineLevelRow := ASheet.SheetFormatPr.OutlineLevelRow;
  FXML.Worksheet.SheetFormatPr.OutlineLevelCol := ASheet.SheetFormatPr.OutlineLevelCol;

  FXML.Worksheet.SheetCalcPr.FullCalcOnLoad := ASheet.SheetCalcPr.FullCalcOnLoad;

  FXML.Worksheet.SheetProtection.Password := ASheet.SheetProtection.Password;
  FXML.Worksheet.SheetProtection.Sheet := ASheet.SheetProtection.Sheet;
  FXML.Worksheet.SheetProtection.Objects := ASheet.SheetProtection.Objects;
  FXML.Worksheet.SheetProtection.Scenarios := ASheet.SheetProtection.Scenarios;
  FXML.Worksheet.SheetProtection.FormatCells := ASheet.SheetProtection.FormatCells;
  FXML.Worksheet.SheetProtection.FormatColumns := ASheet.SheetProtection.FormatColumns;
  FXML.Worksheet.SheetProtection.FormatRows := ASheet.SheetProtection.FormatRows;
  FXML.Worksheet.SheetProtection.InsertColumns := ASheet.SheetProtection.InsertColumns;
  FXML.Worksheet.SheetProtection.InsertRows := ASheet.SheetProtection.InsertRows;
  FXML.Worksheet.SheetProtection.InsertHyperlinks := ASheet.SheetProtection.InsertHyperlinks;
  FXML.Worksheet.SheetProtection.DeleteColumns := ASheet.SheetProtection.DeleteColumns;
  FXML.Worksheet.SheetProtection.DeleteRows := ASheet.SheetProtection.DeleteRows;
  FXML.Worksheet.SheetProtection.SelectLockedCells := ASheet.SheetProtection.SelectLockedCells;
  FXML.Worksheet.SheetProtection.Sort := ASheet.SheetProtection.Sort;
  FXML.Worksheet.SheetProtection.AutoFilter := ASheet.SheetProtection.AutoFilter;
  FXML.Worksheet.SheetProtection.PivotTables := ASheet.SheetProtection.PivotTables;
  FXML.Worksheet.SheetProtection.SelectUnlockedCells := ASheet.SheetProtection.SelectUnlockedCells;

  for i := 0 to ASheet.ProtectedRanges.Count - 1 do begin
    ProtectedRange := FXML.Worksheet.ProtectedRanges.ProtectedRangeXpgList.Add;

    ProtectedRange.Password := ASheet.ProtectedRanges[i].Password;
    ProtectedRange.SqrefXpgList.DelimitedText := ASheet.ProtectedRanges[i].Sqref.DelimitedText;
    ProtectedRange.Name := ASheet.ProtectedRanges[i].Name;
    ProtectedRange.SecurityDescriptor := ASheet.ProtectedRanges[i].SecurityDescriptor;
  end;

  if ASheet.Scenarios.Count > 0 then begin
    FXML.Worksheet.Scenarios.Current := ASheet.Scenarios.Current;
    FXML.Worksheet.Scenarios.Show := ASheet.Scenarios.Show;
    FXML.Worksheet.Scenarios.SqrefXpgList.DelimitedText := ASheet.Scenarios._Sqref.DelimitedText;

    for i := 0 to ASheet.Scenarios.Count - 1 do begin
      Scen := ASheet.Scenarios[i];
      Scenario := FXML.Worksheet.Scenarios.ScenarioXpgList.Add;

      Scenario.Name := Scen.Name;
      Scenario.Locked := Scen.Locked;
      Scenario.Hidden := Scen.Hidden;
      Scenario.User := Scen.User;
      Scenario.Comment := Scen.Comment;

      for j := 0 to Scen.InputCells.Count - 1 do begin
        InputCell := Scenario.InputCellsXpgList.Add;
        InputCell.R := ColRowToRefStr(Scen.InputCells[j].Col,Scen.InputCells[j].Row);
        InputCell.Deleted := Scen.InputCells[j].Deleted;
        InputCell.Undone := Scen.InputCells[j].Undone;
        InputCell.Val := Scen.InputCells[j].Val;
        InputCell.NumFmtId := Scen.InputCells[j].NumFmtId;
      end;
    end;
  end;

  if (ASheet.AutoFilter.FilterColumns.Count > 0) or CellAreaAssigned(ASheet.AutoFilter.Ref) then
    WriteSheetAutoFilter(ASheet.AutoFilter,FXML.Worksheet.AutoFilter);

  if ASheet.SortState.SortConditions.Count > 0 then
    WriteSheetSortState(ASheet.SortState,FXML.Worksheet.SortState);

  if ASheet.DataConsolidate.DataRefs.Count > 0 then begin
    FXML.Worksheet.DataConsolidate.Function_ := TST_DataConsolidateFunction(ASheet.DataConsolidate.Function_);
    FXML.Worksheet.DataConsolidate.LeftLabels := ASheet.DataConsolidate.LeftLabels;
    FXML.Worksheet.DataConsolidate.TopLabels := ASheet.DataConsolidate.TopLabels;
    FXML.Worksheet.DataConsolidate.Link := ASheet.DataConsolidate.Link;

    for i := 0 to ASheet.DataConsolidate.DataRefs.Count - 1 do begin
      DataRef := FXML.Worksheet.DataConsolidate.DataRefs.DataRefXpgList.Add;

      DataRef.Ref := CellAreaToAreaStr(ASheet.DataConsolidate.DataRefs[i].Ref);
      DataRef.Name := ASheet.DataConsolidate.DataRefs[i].Name;
      DataRef.Sheet := ASheet.DataConsolidate.DataRefs[i].Sheet;
      DataRef.R_Id := ASheet.DataConsolidate.DataRefs[i].RId;
    end;
  end;

  if ASheet.CustomSheetViews.Count > 0 then
    WriteSheetCustomSheetViews(ASheet.CustomSheetViews);

  if ASheet.PhoneticPr.FontId <> 0 then begin
    FXML.Worksheet.PhoneticPr.FontId := ASheet.PhoneticPr.FontId;
    FXML.Worksheet.PhoneticPr.Type_ := TST_PhoneticType(ASheet.PhoneticPr.Type_);
    FXML.Worksheet.PhoneticPr.Alignment := TST_PhoneticAlignment(ASheet.PhoneticPr.Alignment);
  end;

  for i := 0 to ASheet.ConditionalFormatting.Count - 1 do
    WriteSheetCondFmt(ASheet.ConditionalFormatting[i],FXML.Worksheet.ConditionalFormattingXpgList.Add);

  if ASheet.DataValidations.Count > 0 then begin
    ASheet.DataValidations.DisablePrompts := FXML.Worksheet.DataValidations.DisablePrompts;
    ASheet.DataValidations.XWindow := FXML.Worksheet.DataValidations.XWindow;
    ASheet.DataValidations.YWindow := FXML.Worksheet.DataValidations.YWindow;

    for i := 0 to ASheet.DataValidations.Count - 1 do begin
      DValidation := FCurrSheet.DataValidations[i];
      DataValidation := FXML.Worksheet.DataValidations.DataValidationXpgList.Add;

      DataValidation.Type_ := TST_DataValidationType(DValidation.Type_);
      DataValidation.ErrorStyle := TST_DataValidationErrorStyle(DValidation.ErrorStyle);
      DataValidation.ImeMode := TST_DataValidationImeMode(DValidation.ImeMode);
      DataValidation.Operator_ := TST_DataValidationOperator(DValidation.Operator_);
      DataValidation.AllowBlank := DValidation.AllowBlank;
      DataValidation.ShowDropDown := DValidation.ShowDropDown;
      DataValidation.ShowInputMessage := DValidation.ShowInputMessage;
      DataValidation.ShowErrorMessage := DValidation.ShowErrorMessage;
      DataValidation.ErrorTitle := DValidation.ErrorTitle;
      DataValidation.Error := DValidation.Error;
      DataValidation.PromptTitle := DValidation.PromptTitle;
      DataValidation.Prompt := DValidation.Prompt;
      DataValidation.SqrefXpgList.DelimitedText := DValidation.Sqref.DelimitedText;
      DataValidation.Formula1 := DValidation.Formula1;
      DataValidation.Formula2 := DValidation.Formula2;
    end;
  end;

  WriteSheetPrintOptions(ASheet.PrintOptions,FXML.Worksheet.PrintOptions);

  WriteSheetPageMargins(ASheet.PageMargins,FXML.Worksheet.PageMargins);

  WriteSheetPageSetup(ASheet.PageSetup,FXML.Worksheet.PageSetup);

  WriteSheetHeaderFooter(ASheet.HeaderFooter,FXML.Worksheet.HeaderFooter);

  if ASheet.RowBreaks.Count > 0 then
    WriteSheetPageBreak(ASheet.RowBreaks,FXML.Worksheet.RowBreaks);

  if ASheet.ColBreaks.Count > 0 then
    WriteSheetPageBreak(ASheet.ColBreaks,FXML.Worksheet.ColBreaks);

  for i := 0 to ASheet.CustomProperties.Count - 1 do begin
    Prop := FXML.Worksheet.CustomProperties.CustomPrXpgList.Add;

    Prop.Name := ASheet.CustomProperties[i].Name;
    Prop.R_Id := ASheet.CustomProperties[i].RId;
  end;

  for i := 0 to ASheet.IgnoredErrors.Count - 1 do begin
    IError := ASheet.IgnoredErrors[i];
    IgnoredError := FXML.Worksheet.IgnoredErrors.IgnoredErrorXpgList.Add;

    IgnoredError.SqrefXpgList.DelimitedText := IError.Sqref.DelimitedText;
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

  if ASheet.SmartTags.Count > 0 then
    WriteSheetSmartTags(ASheet.SmartTags,FXML.Worksheet.SmartTags.CellSmartTagsXpgList);

//  FXML.Worksheet.LegacyDrawing.R_Id := FManager.FileData.OPC.FindVmlDrawingId(FOPCSheet);
//  FXML.Worksheet.Drawing.R_Id := FManager.FileData.OPC.FindDrawingId(FOPCSheet);

  // Not supported. Need to save the picture.
//    FXML.Worksheet.Picture.R_Id := FXLS.FileDataXLSX.OPC.FindPictureId(FOPCSheet,ASheet.Index + 1);

  for i := 0 to ASheet.OleObjects.Count - 1 do begin
    OleObject := FXML.Worksheet.OleObjects.OleObjectXpgList.Add;

    OleObject.ProgId := ASheet.OleObjects[i].ProgId;
    OleObject.DvAspect := TST_DvAspect(ASheet.OleObjects[i].DvAspect);
    OleObject.Link := ASheet.OleObjects[i].Link;
    OleObject.OleUpdate := TST_OleUpdate(ASheet.OleObjects[i].OleUpdate);
    OleObject.AutoLoad := ASheet.OleObjects[i].AutoLoad;
    OleObject.ShapeId := ASheet.OleObjects[i].ShapeId;
    OleObject.R_Id := ASheet.OleObjects[i].RId;
  end;

  for i := 0 to ASheet.Controls.Count - 1 do begin
    Control := FXML.Worksheet.Controls.ControlXpgList.Add;

    Control.ShapeId := ASheet.Controls[i].ShapeId;
    Control.R_Id := ASheet.Controls[i].RId;
    Control.Name := ASheet.Controls[i].Name;
  end;

  for i := 0 to ASheet.WebPublishItems.Count - 1 do begin
    WebPublishItem := FXML.Worksheet.WebPublishItems.WebPublishItemXpgList.Add;
    WPItem := ASheet.WebPublishItems[i];

    WebPublishItem.Id := WPItem.Id;
    WebPublishItem.DivId := WPItem.DivId;
    WebPublishItem.SourceType := TST_WebSourceType(WPItem.SourceType);
    WebPublishItem.SourceRef := WPItem.SourceRef.AsString;
    WebPublishItem.SourceObject := WPItem.SourceObject;
    WebPublishItem.DestinationFile := WPItem.DestinationFile;
    WebPublishItem.Title := WPItem.Title;
    WebPublishItem.AutoRepublish := WPItem.AutoRepublish;
  end;

  SheetDataWrite := TXLSWriteSheetData.Create(Self,ASheet,FXML.Worksheet.SheetData,FXML.Comments);
  try
    FXML.OnWriteWorksheetColsCol := SheetDataWrite.OnWriteCol;

    if ASheet.Columns.Count > 0 then
      FXML.Worksheet.ColsXpgList.Add;

    SheetDataWrite.Initiate;

    FXML.SaveToStream(Stream,TCT_Worksheet);
  finally
    SheetDataWrite.Free;
  end;

  FManager.FileData.OPC.ItemCloseStream(FOPCSheet,Stream);
  FManager.FileData.OPC.ItemClose(FOPCSheet);
  FXML.Worksheet.Clear;
end;

procedure TXLSWriteXLSX.WriteSheetAutoFilter(ASrc: TXc12AutoFilter; ADest: TCT_AutoFilter);
var
  i,j: integer;
  FilterColumn: TCT_FilterColumn;
  FColumn: TXc12FilterColumn;
  DateGroupItem: TCT_DateGroupItem;
  CustomFilter: TCT_CustomFilter;
begin
  ADest.Ref := CellAreaToAreaStr(ASrc.Ref);

  for i := 0 to ASrc.FilterColumns.Count - 1 do begin
    FColumn := ASrc.FilterColumns[i];
    FilterColumn := ADest.FilterColumnXpgList.Add;

    FilterColumn.ColId := FColumn.ColId;
    FilterColumn.HiddenButton := FColumn.HiddenButton;
    FilterColumn.ShowButton := FColumn.ShowButton;

    FilterColumn.Filters.Blank := FColumn.Filters.Blank;
    FilterColumn.Filters.CalendarType := TST_CalendarType(FColumn.Filters.CalendarType);

    for j := 0 to FColumn.Filters.Filter.Count - 1 do
      FilterColumn.Filters.FilterXpgList.Add.Val := FColumn.Filters.Filter[j];

    for j := 0 to FColumn.Filters.DateGroupItems.Count - 1 do begin
      DateGroupItem := FilterColumn.Filters.DateGroupItemXpgList.Add;
      DateGroupItem.Year := FColumn.Filters.DateGroupItems[j].Year;
      DateGroupItem.Month := FColumn.Filters.DateGroupItems[j].Month;
      DateGroupItem.Day := FColumn.Filters.DateGroupItems[j].Day;
      DateGroupItem.Hour := FColumn.Filters.DateGroupItems[j].Hour;
      DateGroupItem.Minute := FColumn.Filters.DateGroupItems[j].Minute;
      DateGroupItem.Second := FColumn.Filters.DateGroupItems[j].Second;
      DateGroupItem.DateTimeGrouping := TST_DateTimeGrouping(FColumn.Filters.DateGroupItems[j].DateTimeGrouping);
    end;

    FilterColumn.Top10.Top := FColumn.Top10.Top;
    FilterColumn.Top10.Percent := FColumn.Top10.Percent;
    FilterColumn.Top10.Val := FColumn.Top10.Val;
    FilterColumn.Top10.FilterVal := FColumn.Top10.FilterVal;

    if FColumn.CustomFilters.Filter1.Val <> '' then begin
      FilterColumn.CustomFilters.And_ := FColumn.CustomFilters.And_;

      CustomFilter := FilterColumn.CustomFilters.CustomFilterXpgList.Add;

      CustomFilter.Operator_ := TST_FilterOperator(FColumn.CustomFilters.Filter1.Operator_);
      CustomFilter.Val := FColumn.CustomFilters.Filter1.Val;
      if FColumn.CustomFilters.Filter2.Val <> '' then begin
        CustomFilter := FilterColumn.CustomFilters.CustomFilterXpgList.Add;

        CustomFilter.Operator_ := TST_FilterOperator(FColumn.CustomFilters.Filter2.Operator_);
        CustomFilter.Val := FColumn.CustomFilters.Filter2.Val;
      end;
    end;

    if Integer(FColumn.DynamicFilter.Type_) <> XPG_UNKNOWN_ENUM then begin
      FilterColumn.DynamicFilter.Type_ := TST_DynamicFilterType(FColumn.DynamicFilter.Type_);
      FilterColumn.DynamicFilter.Val := FColumn.DynamicFilter.Val;
      FilterColumn.DynamicFilter.MaxVal := FColumn.DynamicFilter.MaxVal;
    end;

    if FColumn.ColorFilter.Dxf <> Nil then begin
      FilterColumn.ColorFilter.DxfId := FColumn.ColorFilter.Dxf.Index;
      FilterColumn.ColorFilter.CellColor := FColumn.ColorFilter.CellColor;
    end;

    if FColumn.IconFilter.IconId >= 0 then begin
      FilterColumn.IconFilter.IconSet := TST_IconSetType(FColumn.IconFilter.IconSet);
      FilterColumn.IconFilter.IconId := FColumn.IconFilter.IconId;
    end;
  end;
  WriteSheetSortState(ASrc.SortState,ADest.SortState);
end;

procedure TXLSWriteXLSX.WriteSheetCfvo(ASrc: TXc12Cfvo; ADest: TCT_Cfvo);
begin
  ADest.Type_ := TST_CfvoType(ASrc.Type_);
  ADest.Val := ASrc.Val;
  ADest.Gte := ASrc.Gte;
end;

procedure TXLSWriteXLSX.WriteSheetColors(ASrc: TXc12Colors; ADest: TCT_ColorXpgList);
var
  i: integer;
begin
  for i := 0 to ASrc.Count - 1 do
    GetColor(ASrc[i],ADest.Add);
end;

procedure TXLSWriteXLSX.WriteSheetCondFmt(ASrc: TXc12ConditionalFormatting; ADest: TCT_ConditionalFormatting);
var
  i,j: integer;
  CfRule: TCT_CfRule;
  CfR: TXc12CfRule;
begin
  ADest.Pivot := ASrc.Pivot;
  ADest.SQRefXpgList.DelimitedText := ASrc.SQRef.DelimitedText;

  for i := 0 to ASrc.CfRules.Count - 1 do begin
    CfR := ASrc.CfRules[i];
    CfRule := ADest.CfRuleXpgList.Add;

    if Integer(CfR.Type_) <> XPG_UNKNOWN_ENUM then
      CfRule.Type_ := TST_CfType(CfR.Type_);
    if CfR.DXF <> Nil then
      CfRule.DxfId := CfR.DXF.Index;
    CfRule.Priority := CfR.Priority;
    CfRule.StopIfTrue := CfR.StopIfTrue;
    CfRule.AboveAverage := CfR.AboveAverage;
    CfRule.Percent := CfR.Percent;
    CfRule.Bottom := CfR.Bottom;
    if Integer(CfR.Operator_) <> XPG_UNKNOWN_ENUM then
      CfRule.Operator_ := TST_ConditionalFormattingOperator(CfR.Operator_);
    CfRule.Text := CfR.Text;
    if Integer(CfR.TimePeriod) <> XPG_UNKNOWN_ENUM then
      CfRule.TimePeriod := TST_TimePeriod(CfR.TimePeriod);
    CfRule.Rank := CfR.Rank;
    CfRule.StdDev := CfR.StdDev;
    CfRule.EqualAverage := CfR.EqualAverage;

    if CfR.Formulas[0] <> '' then
      CfRule.FormulaXpgList.Add(CfR.Formulas[0]);
    if CfR.Formulas[1] <> '' then
      CfRule.FormulaXpgList.Add(CfR.Formulas[1]);
    if CfR.Formulas[2] <> '' then
      CfRule.FormulaXpgList.Add(CfR.Formulas[2]);

    for j := 0 to CfR.ColorScale.Cfvos.Count - 1 do
      WriteSheetCfvo(CfR.ColorScale.Cfvos[j],CFRule.ColorScale.CfvoXpgList.Add);
    WriteSheetColors(CfR.ColorScale.Colors,CFRule.ColorScale.ColorXpgList);

    CFRule.DataBar.MinLength := CfR.DataBar.MinLength;
    CFRule.DataBar.MaxLength := CfR.DataBar.MaxLength;
    CFRule.DataBar.ShowValue := CfR.DataBar.ShowValue;

    if CFR.DataBar.Cfvo1.Val <> '' then
      WriteSheetCfvo(CfR.DataBar.Cfvo1,CFRule.DataBar.CfvoXpgList.Add);
    if CFR.DataBar.Cfvo2.Val <> '' then
      WriteSheetCfvo(CfR.DataBar.Cfvo2,CFRule.DataBar.CfvoXpgList.Add);

    GetColor(CfR.DataBar.Color,CFRule.DataBar.Color);

    CFRule.IconSet.IconSet := TST_IconSetType(CfR.IconSet.IconSet);
    CFRule.IconSet.ShowValue := CfR.IconSet.ShowValue;
    CFRule.IconSet.Percent := CfR.IconSet.Percent;
    CFRule.IconSet.Reverse := CfR.IconSet.Reverse;

    for j := 0 to CfR.IconSet.Cfvos.Count - 1 do
      WriteSheetCfvo(CfR.IconSet.Cfvos[j],CFRule.IconSet.CfvoXpgList.Add);
  end;
end;

procedure TXLSWriteXLSX.WriteSheetCustomSheetViews(ACSViews: TXc12CustomSheetViews);
var
  i: integer;
  CSheetView: TXc12CustomSheetView;
  CustomSheetView: TCT_CustomSheetView;
begin
  for i := 0 to ACSViews.Count - 1 do begin
    CSheetView := ACSViews[i];
    CustomSheetView := FXML.Worksheet.CustomSheetViews.CustomSheetViewXpgList.Add;

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
    CustomSheetView.State := TST_SheetState(CSheetView.State);
    CustomSheetView.FilterUnique := CSheetView.FilterUnique;
    CustomSheetView.View := TST_SheetViewType(CSheetView.View);
    CustomSheetView.ShowRuler := CSheetView.ShowRuler;
    CustomSheetView.TopLeftCell := CellAreaToAreaStr(CSheetView.TopLeftCell);

    WriteSheetPane(CSheetView.Pane,CustomSheetView.Pane);

    WriteSheetSelection(CSheetView.Selection,CustomSheetView.Selection);

    if CSheetView.RowBreaks.Count > 0 then
      WriteSheetPageBreak(CSheetView.RowBreaks,CustomSheetView.RowBreaks);

    if CSheetView.ColBreaks.Count > 0 then
      WriteSheetPageBreak(CSheetView.ColBreaks,CustomSheetView.ColBreaks);

    WriteSheetPageMargins(CSheetView.PageMargins,CustomSheetView.PageMargins);

    WriteSheetPrintOptions(CSheetView.PrintOptions,CustomSheetView.PrintOptions);

    WriteSheetPageSetup(CSheetView.PageSetup,CustomSheetView.PageSetup);

    WriteSheetHeaderFooter(CSheetView.HeaderFooter,CustomSheetView.HeaderFooter);

    WriteSheetAutofilter(CSheetView.AutoFilter,CustomSheetView.AutoFilter);
  end;
end;

procedure TXLSWriteXLSX.WriteSheetHeaderFooter(ASrc: TXc12HeaderFooter; ADest: TCT_HeaderFooter);
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

procedure TXLSWriteXLSX.WriteSheetPageBreak(ASrc: TXc12PageBreaks; ADest: TCT_PageBreak);
var
  i: integer;
  Brk: TCT_Break;
begin
  ADest.ManualBreakCount := 0;
  for i := 0 to ASrc.Count - 1 do begin
    Brk := ADest.BrkXpgList.Add;
    Brk.Id := ASrc[i].Id;
    Brk.Min := ASrc[i].Min;
    Brk.Max := ASrc[i].Max;
    Brk.Man := ASrc[i].Man;
    if Brk.Man then
      ADest.ManualBreakCount := ADest.ManualBreakCount + 1;
    Brk.Pt := ASrc[i].Pt;
  end;
end;

procedure TXLSWriteXLSX.WriteSheetPageMargins(ASrc: TXc12PageMargins; ADest: TCT_PageMargins);
begin
  ADest.Left := ASrc.Left;
  ADest.Right := ASrc.Right;
  ADest.Top := ASrc.Top;
  ADest.Bottom := ASrc.Bottom;
  ADest.Header := ASrc.Header;
  ADest.Footer := ASrc.Footer;
end;

procedure TXLSWriteXLSX.WriteSheetPageSetup(ASrc: TXc12PageSetup; ADest: TCT_PageSetup);
begin
  ADest.PaperSize := ASrc.PaperSize;
  ADest.Scale := ASrc.Scale;
  ADest.FirstPageNumber := ASrc.FirstPageNumber;
  ADest.FitToWidth := ASrc.FitToWidth;
  ADest.FitToHeight := ASrc.FitToHeight;
  ADest.PageOrder := TST_PageOrder(ASrc.PageOrder);
  ADest.Orientation := TST_Orientation(ASrc.Orientation);
  ADest.UsePrinterDefaults := ASrc.UsePrinterDefaults;
  ADest.BlackAndWhite := ASrc.BlackAndWhite;
  ADest.Draft := ASrc.Draft;
  ADest.CellComments := TST_CellComments(ASrc.CellComments);
  ADest.UseFirstPageNumber := ASrc.UseFirstPageNumber;
  ADest.Errors := TST_PrintError(ASrc.Errors);
  ADest.HorizontalDpi := ASrc.HorizontalDpi;
  ADest.VerticalDpi := ASrc.VerticalDpi;
  ADest.Copies := ASrc.Copies;
end;

procedure TXLSWriteXLSX.WriteSheetPane(ASrc: TXc12Pane; ADest: TCT_Pane);
begin
  ADest.XSplit := ASrc.XSplit;
  ADest.YSplit := ASrc.YSplit;
  ADest.TopLeftCell := CellAreaToAreaStr(ASrc.TopLeftCell);
  ADest.ActivePane := TST_Pane(ASrc.ActivePane);
  ADest.State := TST_PaneState(ASrc.State);
end;

procedure TXLSWriteXLSX.WriteSheetPrintOptions(ASrc: TXc12PrintOptions; ADest: TCT_PrintOptions);
begin
  ADest.HorizontalCentered := ASrc.HorizontalCentered;
  ADest.VerticalCentered := ASrc.VerticalCentered;
  ADest.Headings := ASrc.Headings;
  ADest.GridLines := ASrc.GridLines;
  ADest.GridLinesSet := ASrc.GridLinesSet;
end;

procedure TXLSWriteXLSX.WriteSheetSelection(ASrc: TXc12Selection; ADest: TCT_Selection);
begin
  ADest.Pane := TST_Pane(ASrc.Pane);
  ADest.ActiveCell := CellAreaToAreaStr(ASrc.ActiveCell);
  ADest.ActiveCellId := ASrc.ActiveCellId;
  if ASrc.SQRef.Count > 0 then
    ADest.SqrefXpgList.DelimitedText := ASrc.SQRef.DelimitedText;
end;

procedure TXLSWriteXLSX.WriteSheetSmartTags(ASrc: TXc12SmartTags; ADest: TCT_CellSmartTagsXpgList);
var
  i,j,k: integer;
  CellSmartTags: TCT_CellSmartTags;
  CellSmartTag: TCT_CellSmartTag;
  CellSmartTagPr: TCT_CellSmartTagPr;
  CSTagPr: TXc12CellSmartTagPr;
begin
  for i := 0 to ASrc.Count - 1 do begin
    CellSmartTags := ADest.Add;

    CellSmartTags.R := ASrc[i].Ref.AsString;

    for j := 0 to ASrc[i].Count - 1 do begin
      CellSmartTag := CellSmartTags.CellSmartTagXpgList.Add;

      CellSmartTag.Type_ := ASrc[i].Items[j].Type_;
      CellSmartTag.Deleted := ASrc[i].Items[j].Deleted;
      CellSmartTag.XmlBased := ASrc[i].Items[j].XmlBased;

      for k := 0 to ASrc[i].Items[j].Count - 1 do begin
        CSTagPr := ASrc[i].Items[j].Items[k];
        CellSmartTagPr := CellSmartTag.CellSmartTagPrXpgList.Add;

        CellSmartTagPr.Key := CSTagPr.Key;
        CellSmartTagPr.Val := CSTagPr.Val;
      end;
    end;
  end;
end;

procedure TXLSWriteXLSX.WriteSheetSortState(ASrc: TXc12SortState; ADest: TCT_SortState);
var
  i: integer;
  SortCond: TCT_SortCondition;
begin
  ADest.ColumnSort := ASrc.ColumnSort;
  ADest.CaseSensitive := ASrc.CaseSensitive;
  ADest.SortMethod := TST_SortMethod(ASrc.SortMethod);
  ADest.Ref := CellAreaToAreaStr(ASrc.Ref);

  for i := 0 to ASrc.SortConditions.Count - 1 do begin
    SortCond := ADest.SortConditionXpgList.Add;

    SortCond.Descending := ASrc.SortConditions[i].Descending;
    SortCond.SortBy := TST_SortBy(ASrc.SortConditions[i].SortBy);
    SortCond.Ref := CellAreaToAreaStr(ASrc.SortConditions[i].Ref);
    SortCond.CustomList := ASrc.SortConditions[i].CustomList;
    SortCond.DxfId := ASrc.SortConditions[i].DxfId;
    SortCond.IconSet := TST_IconSetType(ASrc.SortConditions[i].IconSet);
    SortCond.IconId := ASrc.SortConditions[i].IconId;
  end;
end;

function TXLSWriteXLSX.WriteSheetTable(ASrc: TXc12Table; AOwner: TOPCItem): AxUCString;
var
  i: integer;
  Table: TCT_Table;
  TC: TCT_TableColumn;
  Stream: TMemoryStream;
  OPC: TOPCItem;
begin
  OPC := FManager.FileData.OPC.CreateTable(AOwner);
  Result := OPC.Id;

  Stream := TMemoryStream.Create;
  try
    Table := FXML.Table;
    Table.Clear;

    Table.Id := ASrc.Id;
    Table.Name := ASrc.Name;
    Table.DisplayName := ASrc.DisplayName;
    Table.Comment := ASrc.Comment;
    Table.Ref := CellAreaToAreaStr(ASrc.Ref);
    Table.TableType := TST_TableType(ASrc.TableType);
    Table.HeaderRowCount := ASrc.HeaderRowCount;
    Table.InsertRow := ASrc.InsertRow;
    Table.InsertRowShift := ASrc.InsertRowShift;
    Table.TotalsRowCount := ASrc.TotalsRowCount;
    Table.TotalsRowShown := ASrc.TotalsRowShown;
    Table.Published_ := ASrc.Published_;
    Table.HeaderRowDxfId := ASrc.HeaderRowDxfId;
    Table.DataDxfId := ASrc.DataDxfId;
    Table.TotalsRowDxfId := ASrc.TotalsRowDxfId;
    Table.HeaderRowBorderDxfId := ASrc.HeaderRowBorderDxfId;
    Table.TableBorderDxfId := ASrc.TableBorderDxfId;
    Table.TotalsRowBorderDxfId := ASrc.TotalsRowBorderDxfId;
    Table.HeaderRowCellStyle := ASrc.HeaderRowCellStyle;
    Table.DataCellStyle := ASrc.DataCellStyle;
    Table.TotalsRowCellStyle := ASrc.TotalsRowCellStyle;
    Table.ConnectionId := ASrc.ConnectionId;

    if CellAreaAssigned(ASrc.AutoFilter.Ref) then
      WriteSheetAutoFilter(ASrc.AutoFilter,Table.AutoFilter);

    if CellAreaAssigned(ASrc.SortState.Ref) then
      WriteSheetSortState(ASrc.SortState,Table.SortState);

    Table.TableColumns.Count := ASrc.TableColumns.Count;
    for i := 0 to ASrc.TableColumns.Count - 1 do begin
      TC := Table.TableColumns.TableColumnXpgList.Add;
      TC.Id := ASrc.TableColumns[i].Id;
      TC.UniqueName := ASrc.TableColumns[i].UniqueName;
      TC.Name := ASrc.TableColumns[i].Name;
      TC.TotalsRowFunction := TST_TotalsRowFunction(ASrc.TableColumns[i].TotalsRowFunction);
      TC.TotalsRowLabel := ASrc.TableColumns[i].TotalsRowLabel;
      TC.QueryTableFieldId := ASrc.TableColumns[i].QueryTableFieldId;
      TC.HeaderRowDxfId := ASrc.TableColumns[i].HeaderRowDxfId;
      TC.DataDxfId := ASrc.TableColumns[i].DataDxfId;
      TC.TotalsRowDxfId := ASrc.TableColumns[i].TotalsRowDxfId;
      TC.HeaderRowCellStyle := ASrc.TableColumns[i].HeaderRowCellStyle;
      TC.DataCellStyle := ASrc.TableColumns[i].DataCellStyle;
      TC.TotalsRowCellStyle := ASrc.TableColumns[i].TotalsRowCellStyle;

      if ASrc.TableColumns[i].CalculatedColumnFormula.Content <> '' then begin
        TC.CalculatedColumnFormula.Array_ := ASrc.TableColumns[i].CalculatedColumnFormula.Array_;
        TC.CalculatedColumnFormula.Content := ASrc.TableColumns[i].CalculatedColumnFormula.Content;
      end;
      if ASrc.TableColumns[i].TotalsRowFormula.Content <> '' then begin
        TC.TotalsRowFormula.Array_ := ASrc.TableColumns[i].TotalsRowFormula.Array_;
        TC.TotalsRowFormula.Content := ASrc.TableColumns[i].TotalsRowFormula.Content;
      end;
      if ASrc.TableColumns[i].XmlColumnPr.Xpath <> '' then begin
        TC.XmlColumnPr.MapId := ASrc.TableColumns[i].XmlColumnPr.MapId;
        TC.XmlColumnPr.Xpath := ASrc.TableColumns[i].XmlColumnPr.Xpath;
        TC.XmlColumnPr.Denormalized := ASrc.TableColumns[i].XmlColumnPr.Denormalized;
        TC.XmlColumnPr.XmlDataType := TST_XmlDataType(ASrc.TableColumns[i].XmlColumnPr.XmlDataType);
      end;
    end;

    Table.TableStyleInfo.Name := ASrc.TableStyleInfo.Name;
    Table.TableStyleInfo.ShowFirstColumn := ASrc.TableStyleInfo.ShowFirstColumn;
    Table.TableStyleInfo.ShowLastColumn := ASrc.TableStyleInfo.ShowLastColumn;
    Table.TableStyleInfo.ShowRowStripes := ASrc.TableStyleInfo.ShowRowStripes;
    Table.TableStyleInfo.ShowColumnStripes := ASrc.TableStyleInfo.ShowColumnStripes;


    FXML.SaveToStream(Stream,TCT_Table);

    FManager.FileData.OPC.ItemWrite(OPC,Stream);
    FManager.FileData.OPC.ItemClose(OPC);
  finally
    Stream.Free;
  end;

  if not ASrc.QueryTable.Empty then begin
    OPC := FManager.FileData.OPC.CreateQueryTable(OPC);

    Stream := TMemoryStream.Create;
    try
      ASrc.QueryTable.SaveToStream(Stream);
      FManager.FileData.OPC.ItemWrite(OPC,Stream);
      FManager.FileData.OPC.ItemClose(OPC);
    finally
      Stream.Free;
    end;
  end;
end;

procedure TXLSWriteXLSX.WriteSheetViews(ASheet: TXc12DataWorksheet);
var
  i,j,k,l: integer;
  SView: TXc12SheetView;
  SheetView: TCT_SheetView;
  Selection: TCT_Selection;
  PSelection: TXc12PivotSelection;
  PivotSelection: TCT_PivotSelection;
  Reference: xpgParserXLSX.TCT_PivotAreaReference;
begin
  for i := 0 to ASheet.SheetViews.Count - 1 do begin
    SView := ASheet.SheetViews[i];
    SheetView := FXML.Worksheet.SheetViews.SheetViewXpgList.Add;

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
    SheetView.View := TST_SheetViewType(SView.View);
    SheetView.TopLeftCell := CellAreaToAreaStr(SView.TopLeftCell);
    SheetView.ColorId := SView.ColorId;
    SheetView.ZoomScale := SView.ZoomScale;
    SheetView.ZoomScaleNormal := SView.ZoomScaleNormal;
    SheetView.ZoomScaleSheetLayoutView := SView.ZoomScaleSheetLayoutView;
    SheetView.ZoomScalePageLayoutView := SView.ZoomScalePageLayoutView;
    SheetView.WorkbookViewId := SView.WorkbookViewId;

    SheetView.Pane.XSplit := SView.Pane.XSplit;
    SheetView.Pane.YSplit := SView.Pane.YSplit;
    SheetView.Pane.TopLeftCell := CellAreaToAreaStr(SView.Pane.TopLeftCell);
    SheetView.Pane.ActivePane := TST_Pane(SView.Pane.ActivePane);
    SheetView.Pane.State := TST_PaneState(SView.Pane.State);

    for j := 0 to SView.Selection.Count - 1 do begin
      Selection := SheetView.SelectionXpgList.Add;

      Selection.Pane := TST_Pane(SView.Selection[j].Pane);
      Selection.ActiveCell := CellAreaToAreaStr(SView.Selection[j].ActiveCell);
      Selection.ActiveCellId := SView.Selection[j].ActiveCellId;
      if SView.Selection[j].Sqref.Count > 0 then begin
        Selection.SqrefXpgList.Clear;
        Selection.SqrefXpgList.DelimitedText := SView.Selection[j].Sqref.DelimitedText;
      end;
    end;

    for j := 0 to SView.PivotSelection.Count - 1 do begin
      PSelection := SView.PivotSelection[j];
      PivotSelection := SheetView.PivotSelectionXpgList.Add;

      PivotSelection.Pane := TST_Pane(PSelection.Pane);
      PivotSelection.ShowHeader := PSelection.ShowHeader;
      PivotSelection.Label_ := PSelection.Label_;
      PivotSelection.Data := PSelection.Data;
      PivotSelection.Extendable := PSelection.Extendable;
      PivotSelection.Axis := xpgParserXLSX.TST_Axis(PSelection.Axis);
      PivotSelection.Dimension := PSelection.Dimension;
      PivotSelection.Start := PSelection.Start;
      PivotSelection.Min := PSelection.Min;
      PivotSelection.Max := PSelection.Max;
      PivotSelection.ActiveRow := PSelection.ActiveRow;
      PivotSelection.ActiveCol := PSelection.ActiveCol;
      PivotSelection.PreviousRow := PSelection.PreviousRow;
      PivotSelection.PreviousCol := PSelection.PreviousCol;
      PivotSelection.Click := PSelection.Click;
      PivotSelection.R_Id := PSelection.RId;

      PivotSelection.PivotArea.Field := PSelection.PivotArea.Field;
      PivotSelection.PivotArea.Type_ := xpgParserXLSX.TST_PivotAreaType(PSelection.PivotArea.Type_);
      PivotSelection.PivotArea.DataOnly := PSelection.PivotArea.DataOnly;
      PivotSelection.PivotArea.LabelOnly := PSelection.PivotArea.LabelOnly;
      PivotSelection.PivotArea.GrandRow := PSelection.PivotArea.GrandRow;
      PivotSelection.PivotArea.GrandCol := PSelection.PivotArea.GrandCol;
      PivotSelection.PivotArea.CacheIndex := PSelection.PivotArea.CacheIndex;
      PivotSelection.PivotArea.Outline := PSelection.PivotArea.Outline;
      PivotSelection.PivotArea.Offset := CellAreaToAreaStr(PSelection.PivotArea.Offset);
      PivotSelection.PivotArea.CollapsedLevelsAreSubtotals := PSelection.PivotArea.CollapsedLevelsAreSubtotals;
      PivotSelection.PivotArea.Axis := xpgParserXLSX.TST_Axis(PSelection.PivotArea.Axis);
      PivotSelection.PivotArea.FieldPosition := PSelection.PivotArea.FieldPosition;

      for k := 0 to PSelection.PivotArea.References.Count -1 do begin
        Reference := PivotSelection.PivotArea.References.ReferenceXpgList.Add;

        Reference.Field := PSelection.PivotArea.References[k].Field;
        Reference.Selected := PSelection.PivotArea.References[k].Selected;
        Reference.ByPosition := PSelection.PivotArea.References[k].ByPosition;
        Reference.Relative := PSelection.PivotArea.References[k].Relative;
        Reference.DefaultSubtotal := PSelection.PivotArea.References[k].DefaultSubtotal;
        Reference.SumSubtotal := PSelection.PivotArea.References[k].SumSubtotal;
        Reference.CountASubtotal := PSelection.PivotArea.References[k].CountASubtotal;
        Reference.AvgSubtotal := PSelection.PivotArea.References[k].AvgSubtotal;
        Reference.MaxSubtotal := PSelection.PivotArea.References[k].MaxSubtotal;
        Reference.MinSubtotal := PSelection.PivotArea.References[k].MinSubtotal;
        Reference.ProductSubtotal := PSelection.PivotArea.References[k].ProductSubtotal;
        Reference.CountSubtotal := PSelection.PivotArea.References[k].CountSubtotal;
        Reference.StdDevSubtotal := PSelection.PivotArea.References[k].StdDevSubtotal;
        Reference.StdDevPSubtotal := PSelection.PivotArea.References[k].StdDevPSubtotal;
        Reference.VarSubtotal := PSelection.PivotArea.References[k].VarSubtotal;
        Reference.VarPSubtotal := PSelection.PivotArea.References[k].VarPSubtotal;

        for l := 0 to PSelection.PivotArea.References[k].Values.Count - 1 do
          Reference.XXpgList.Add.V := PSelection.PivotArea.References[k].Values[l];
      end;
    end;
  end;
end;

//procedure TXLSWriteXLSX.WriteSST;
//var
//  Stream: TMemoryStream;
//  OPC: TOPCItem;
//begin
//  if FManager.SST.TotalCount < 1 then
//    Exit;
//
//  OPC := FManager.FileData.OPC.CreateSST;
//
//  FXML.Root.RootAttributes.Clear;
//  FXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);
//
//  FXML.Root.Sst.Count_ := FManager.SST.TotalCount;
//
//  FCurrCount := FManager.SST.TotalCount;
//  FCurrIndex := 0;
//  FXML.Root.Sst.OnWriteSi := OnWriteSSTItem;
//
//  Stream := TMemoryStream.Create;
//  try
//    FXML.SaveToStream(Stream,TCT_Sst);
//    FManager.FileData.OPC.ItemWrite(OPC,Stream);
//    FManager.FileData.OPC.ItemClose(OPC);
//  finally
//    Stream.Free;
//  end;
//
//  FXML.SST.Clear;
//end;

procedure TXLSWriteXLSX.WriteSST;
var
  Stream: TStream;
  OPC: TOPCItem;
begin
  if FManager.SST.TotalCount < 1 then
    Exit;

  OPC := FManager.FileData.OPC.CreateSST;

  FXML.Root.RootAttributes.Clear;
  FXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);

  FXML.Root.Sst.Count_ := FManager.SST.TotalCount;

  FCurrCount := FManager.SST.TotalCount;
  FCurrIndex := 0;
  FXML.Root.Sst.OnWriteSi := OnWriteSSTItem;

  Stream := FManager.FileData.OPC.ItemCreateStream(OPC);

  FXML.SaveToStream(Stream,TCT_Sst);

  FManager.FileData.OPC.ItemCloseStream(OPC,Stream);
  FManager.FileData.OPC.ItemClose(OPC);

  FXML.SST.Clear;
end;


procedure TXLSWriteXLSX.WriteNumberFormats;
var
  i,Cnt: integer;
  NumFmt: TCT_NumFmt;
begin
  Cnt := 0;
  for i := 0 to FManager.StyleSheet.NumFmts.Count - 1 do begin
    if FManager.StyleSheet.NumFmts[i].Assigned then begin
      NumFmt := FXML.StyleSheet.NumFmts.NumFmtXpgList.Add;
      NumFmt.NumFmtId := FManager.StyleSheet.NumFmts[i].Index;
      NumFmt.FormatCode := FManager.StyleSheet.NumFmts[i].Value;
      Inc(Cnt);
    end;
  end;
  FXML.StyleSheet.NumFmts.Count_ := Cnt;
end;

procedure TXLSWriteXLSX.WriteFonts;
var
  i: integer;
  F: TXc12Font;
  Font: TCT_Font;
begin
  FXML.StyleSheet.Fonts.Count_ := FManager.StyleSheet.Fonts.Count;
  for i := 0 to FManager.StyleSheet.Fonts.Count - 1 do begin
    Font := FXML.StyleSheet.Fonts.FontXpgList.Add;
    F := FManager.StyleSheet.Fonts[i];

    GetFont(F,Font);
  end;
end;

procedure TXLSWriteXLSX.WriteMedia;
begin

end;

procedure TXLSWriteXLSX.WriteFills;
var
  i: integer;
  F: TXc12Fill;
  Fill: TCT_Fill;
begin
  for i := 0 to FManager.StyleSheet.Fills.Count - 1 do begin
    Fill := FXML.StyleSheet.Fills.FillXpgList.Add;
    F := FManager.StyleSheet.Fills[i];

    GetFill(F,Fill);
  end;
  FXML.StyleSheet.Fills.Count_ := FManager.StyleSheet.Fills.Count;
end;

procedure TXLSWriteXLSX.WriteBorders;
var
  i: integer;
  B: TXc12Border;
  Border: TCT_Border;
begin
  FXML.StyleSheet.Borders.Count_ := FManager.StyleSheet.Borders.Count;
  for i := 0 to FManager.StyleSheet.Borders.Count - 1 do begin
    Border := FXML.StyleSheet.Borders.BorderXpgList.Add;
    B := FManager.StyleSheet.Borders[i];

    GetBorders(B,Border);
  end;
end;

procedure TXLSWriteXLSX.WriteCellStyles;
var
  i: integer;
  S: TXc12CellStyle;
  Style: TCT_CellStyle;
begin
  FXML.StyleSheet.CellStyles.Count_ := FManager.StyleSheet.Styles.Count;
  for i := 0 to FManager.StyleSheet.Styles.Count - 1 do begin
    S := FManager.StyleSheet.Styles[i];
    Style := FXML.StyleSheet.CellStyles.CellStyleXpgList.Add;

    Style.BuiltinId := S.BuiltInId;
    Style.CustomBuiltin := S.CustomBuiltIn;
    Style.Hidden := S.Hidden;
    Style.ILevel := S.Level;
    if S.XF <> nil then
      Style.XfId := S.XF.Index
    else
      Style.XfId := 0;
    Style.Name := S.Name;
  end;
end;

procedure TXLSWriteXLSX.WriteChartSheet(ASheet: TXc12DataWorksheet);
var
  i             : integer;
  Stream        : TStream;
  WebPublishItem: TCT_WebPublishItem;
  WPItem        : TXc12WebPublishItem;
  CView         : TCT_ChartsheetView;
  CV            : TXc12SheetView;
begin
  FCurrSheet := ASheet;

  FOPCSheet := FManager.FileData.OPC.CreateChartSheet(ASheet.Index + 1);

  FManager.FileData.WriteUnusedDataSheet(ASheet.Index,FOPCSheet);

  // Write substreams before sheet stream
  if FCurrSheet.Drawing.WsDr.HasData then
    FXML.Chartsheet.Drawing.R_Id := WriteDrawing(FOPCSheet);

  if not FCurrSheet.VmlDrawing.Empty or (FCurrSheet.Comments.Count > 0) then
    FXML.Chartsheet.LegacyDrawing.R_Id := WriteVmlDrawing(FOPCSheet);

  if ASheet.PageSetup.PrinterSettings <> Nil then
    FXML.Chartsheet.PageSetup.R_Id := FManager.FileData.OPC.WritePrinterSettings(FOPCSheet,ASheet.PageSetup.PrinterSettings);

  // End Write substreams

  FXML.Root.RootAttributes.Clear;
  FXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);
  FXML.Root.RootAttributes.AddNameValue('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);

  Stream := FManager.FileData.OPC.ItemCreateStream(FOPCSheet);

  FXML.Chartsheet.SheetPr.Published_ := ASheet.SheetPr.Published_;
  FXML.Chartsheet.SheetPr.CodeName := ASheet.SheetPr.CodeName;
  if ASheet.SheetPr.TabColor.ColorType <> exctUnassigned then
    GetColor(ASheet.SheetPr.TabColor,FXML.Chartsheet.SheetPr.TabColor);

  for i := 0 to ASheet.SheetViews.Count - 1 do begin
    CV := ASheet.SheetViews[i];
    CView := FXML.Chartsheet.SheetViews.SheetViewXpgList.Add;
    CView.ZoomScale := CV.ZoomScale;
    CView.WorkbookViewId := CV.WorkbookViewId;
    CView.ZoomToFit := CV.ZoomToFit;
  end;

  FXML.Chartsheet.SheetProtection.Password := ASheet.SheetProtection.Password;
  FXML.Chartsheet.SheetProtection.Objects := ASheet.SheetProtection.Objects;

//  if ASheet.CustomSheetViews.Count > 0 then
//    WriteSheetCustomSheetViews(ASheet.CustomSheetViews);

  WriteSheetPageMargins(ASheet.PageMargins,FXML.Chartsheet.PageMargins);

//  WriteSheetPageSetup(ASheet.PageSetup,FXML.Chartsheet.PageSetup);

  WriteSheetHeaderFooter(ASheet.HeaderFooter,FXML.Chartsheet.HeaderFooter);

  for i := 0 to ASheet.WebPublishItems.Count - 1 do begin
    WebPublishItem := FXML.Chartsheet.WebPublishItems.WebPublishItemXpgList.Add;
    WPItem := ASheet.WebPublishItems[i];

    WebPublishItem.Id := WPItem.Id;
    WebPublishItem.DivId := WPItem.DivId;
    WebPublishItem.SourceType := TST_WebSourceType(WPItem.SourceType);
    WebPublishItem.SourceRef := WPItem.SourceRef.AsString;
    WebPublishItem.SourceObject := WPItem.SourceObject;
    WebPublishItem.DestinationFile := WPItem.DestinationFile;
    WebPublishItem.Title := WPItem.Title;
    WebPublishItem.AutoRepublish := WPItem.AutoRepublish;
  end;

  FXML.SaveToStream(Stream,TCT_Chartsheet);

  FManager.FileData.OPC.ItemCloseStream(FOPCSheet,Stream);
  FManager.FileData.OPC.ItemClose(FOPCSheet);
  FXML.Chartsheet.Clear;
end;

procedure TXLSWriteXLSX.WriteColors;
var
  i    : integer;
  RGB  : TCT_RgbColor;
  Color: TCT_Color;
begin
  if FManager.PaletteChanged then begin
    FXML.StyleSheet.Colors.IndexedColors.Clear;

    for i := 0 to High(Xc12IndexColorPalette) do
      FXML.StyleSheet.Colors.IndexedColors.RgbColorXpgList.Add.Rgb := Xc12IndexColorPalette[i];
  end;

  for i := 0 to FManager.StyleSheet.InxColors.Count - 1 do begin
    RGB := FXML.StyleSheet.Colors.IndexedColors.RgbColorXpgList.Add;
    RGB.Rgb := FManager.StyleSheet.InxColors[i].RGB;
  end;

  for i := 0 to FManager.StyleSheet.MruColors.Count - 1 do begin
    Color := FXML.StyleSheet.Colors.MruColors.ColorXpgList.Add;
    GetColor(FManager.StyleSheet.MruColors[i].Color,Color);
  end;
end;

procedure TXLSWriteXLSX.WriteConnections;
var
  Stream: TStream;
  OPC: TOPCItem;
begin
  if FManager.Workbook.Connections.Empty then
    Exit;

  OPC := FManager.FileData.OPC.CreateConnections;

  Stream := FManager.FileData.OPC.ItemCreateStream(OPC);

  FManager.Workbook.Connections.SaveToStream(Stream);

  FManager.FileData.OPC.ItemCloseStream(OPC,Stream);
  FManager.FileData.OPC.ItemClose(OPC);
end;

procedure TXLSWriteXLSX.WriteStyles;
var
  OPC: TOPCItem;
  Stream: TMemoryStream;
begin
  OPC := FManager.FileData.OPC.CreateStyles;

  FXML.Root.RootAttributes.Clear;
  FXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);

  WriteNumberFormats;
  WriteFonts;
  WriteFills;
  WriteBorders;

  FXML.Root.StyleSheet.CellStyleXfs.Count_ := FManager.StyleSheet.StyleXFs.Count;
  WriteXFs(FManager.StyleSheet.StyleXFs,FXML.Root.StyleSheet.CellStyleXfs.XfXpgList);

  FXML.Root.StyleSheet.CellXfs.Count_ := FManager.StyleSheet.XFs.Count;
  WriteXFs(FManager.StyleSheet.XFs,FXML.Root.StyleSheet.CellXfs.XfXpgList);

  WriteCellStyles;
  WriteDXFs;
  WriteTableStyles;
  WriteColors;

  Stream := TMemoryStream.Create;
  try
    FXML.SaveToStream(Stream,TCT_Stylesheet);
    FManager.FileData.OPC.ItemWrite(OPC,Stream);
    FManager.FileData.OPC.ItemClose(OPC);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWriteXLSX.WriteTableStyles;
var
  i,j: integer;
  T: TXc12TableStyle;
  TableStyle: TCT_TableStyle;
  StyleElement: TCT_TableStyleElement;
begin
  FXML.StyleSheet.TableStyles.DefaultTableStyle := FManager.StyleSheet.TableStyles.DefaultTableStyle;
  FXML.StyleSheet.TableStyles.DefaultPivotStyle := FManager.StyleSheet.TableStyles.DefaultPivotStyle;
  FXML.StyleSheet.TableStyles.Count_ := FManager.StyleSheet.TableStyles.Count;
  for i := 0 to FManager.StyleSheet.TableStyles.Count - 1 do begin
    T := FManager.StyleSheet.TableStyles[i];
    TableStyle := FXML.StyleSheet.TableStyles.TableStyleXpgList.Add;

    TableStyle.Name := T.Name;
    TableStyle.Pivot := T.Pivot;
    TableStyle.Table := T.Table;
    for j := 0 to T.TableStyleElements.Count - 1 do begin
      StyleElement := TableStyle.TableStyleElementXpgList.Add;
      StyleElement.Type_ := TST_TableStyleType(T.TableStyleElements[j].Type_);
      StyleElement.Size := T.TableStyleElements[j].Size;
      StyleElement.DxfId := T.TableStyleElements[j].Dxf.Index;
    end;
  end;
end;

function TXLSWriteXLSX.WriteVmlDrawing(AOwner: TOPCItem): AxUCString;
var
  i     : integer;
  S1,S2 : AxUCString;
  OPC   : TOPCItem;
  DOM   : TXpgSimpleDOM;
  Node  : TXpgDOMNode;
  Stream: TStream;
begin
  OPC := FManager.FileData.OPC.CreateVmlDrawing(AOwner,FCurrSheet.Index + 1);
  Result := OPC.Id;

  if FCurrSheet.VmlDrawing.Empty then
    FCurrSheet.VmlDrawing.LoadFromString(XLS_DEFAULT_VML);

  Stream := FManager.FileData.OPC.ItemCreateStream(OPC);

  DOM := TXpgSimpleDOM.Create;
  try
    if (FCurrSheet.VmlDrawing.Root.Count > 0) and (FCurrSheet.VmlDrawing.Root[0].QName = 'xml') then begin
      DOM.Root.Insert(FCurrSheet.VmlDrawing.Root[0].Clone);

      Node := DOM.Root[0];
      for i := 0 to FCurrSheet.Comments.Count - 1 do begin
        FCurrSheet.Comments[i].UpdateVML;
        Node.Insert(FCurrSheet.Comments[i].VML.Clone);
      end;
    end;
    DOM.SaveToStream(Stream);
  finally
    DOM.Free;
  end;

  FManager.FileData.OPC.ItemCloseStream(OPC,Stream);

  for i := 0 to FCurrSheet.VmlDrawingRels.Count - 1 do begin
    S1 := FCurrSheet.VmlDrawingRels[i];
    S2 := SplitAtChar('&',S1);

    OPC.AddNew(S2,S1);
  end;

  FManager.FileData.OPC.ItemClose(OPC);
end;

procedure TXLSWriteXLSX.BeginWriteWorkbook;
begin
  FManager.FileData.OPC.CreateWorkbook(FManager.FileData.HasMacros);
end;

procedure TXLSWriteXLSX.WriteXFs(ASrc: TXc12XFs; ADest: TCT_XfXpgList);
var
  i: integer;
  X: TXc12XF;
  XF: TCT_Xf;
begin
  for i := 0 to ASrc.Count - 1 do begin
    X := ASrc[i];
    XF := ADest.Add;
    if X <> Nil then begin
      XF.ApplyNumberFormat := eafNumberFormat in X.Apply;
      XF.ApplyFont := eafFont in X.Apply;
      XF.ApplyFill := eafFill in X.Apply;
      XF.ApplyBorder := eafBorder in X.Apply;
      XF.ApplyAlignment := eafAlignment in X.Apply;
      XF.ApplyProtection := eafProtection in X.Apply;

      XF.FillId := X.Fill.Index;
      XF.BorderId := X.Border.Index;
      XF.FontId := X.Font.Index;
      XF.NumFmtId := X.NumFmt.Index;

      if X.XF <> Nil then
        XF.XfId := X.XF.Index;

      GetAlignment(X.Alignment as TXc12CellAlignment,XF.Alignment);

      GetProtection(X.Protection,XF.Protection);

      XF.QuotePrefix := X.QuotePrefix;
      XF.PivotButton := X.PivotButton;
    end;
  end;
end;

function TXLSWriteXLSX.WriteXLink(AIndex: integer): AxUCString;
var
  OPC: TOPCItem;
  XML: TXPGDocXLink;
  Stream: TMemoryStream;
  XLink: TXc12ExternalLink;

procedure WriteExternalBook(AXBook: TXc12XExternalBook);
var
  i,j,k: integer;
  O: TOPCItem;
  DefName: TCT_ExternalDefinedName;
  SheetData: TCT_ExternalSheetData;
  Row: TCT_ExternalRow;
  Cell: TCT_ExternalCell;
  C: TXc12XCell;
begin
  O := FManager.FileData.OPC.CreateXLinkPath(OPC,AXBook.Filename);

  FXML.Root.RootAttributes.Clear;
  XML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);
  XML.Root.RootAttributes.AddNameValue('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);

  XML.ExternalLink.ExternalBook.R_Id := O.Id;
  for i := 0 to AXBook.SheetNames.Count - 1 do
    XML.ExternalLink.ExternalBook.SheetNames.SheetNameXpgList.Add.Val := AXBook.SheetNames[i];

  for i := 0 to AXBook.DefinedNames.Count - 1 do begin
    DefName := XML.ExternalLink.ExternalBook.DefinedNames.DefinedNameXpgList.Add;

    DefName.Name := AXBook.DefinedNames[i].Name;
    DefName.RefersTo := AXBook.DefinedNames[i].RefersTo;
    DefName.SheetId := AXBook.DefinedNames[i].SheetId;
  end;

  for i := 0 to AXBook.SheetDataSet.Count - 1 do begin
    SheetData := XML.ExternalLink.ExternalBook.SheetDataSet.SheetDataXpgList.Add;
    SheetData.SheetId := AXBook.SheetDataSet[i].SheetId;
    SheetData.RefreshError := AXBook.SheetDataSet[i].RefreshError;
    for j := 0 to AXBook.SheetDataSet[i].Count - 1 do begin
      Row := SheetData.RowXpgList.Add;
      Row.R := AXBook.SheetDataSet[i].Rows[j].Row + 1;
      for k := 0 to AXBook.SheetDataSet[i].Rows[j].Count - 1 do begin
        C := AXBook.SheetDataSet[i].Rows[j].Cells[k];
        Cell := Row.CellXpgList.Add;
        Cell.R := C.CellRef;
        case C.Type_ of
          x12ctBoolean: Cell.T := stectB;
          x12ctError  : Cell.T := stectE;
          x12ctFloat  : Cell.T := stectN;
          x12ctString : Cell.T := stectStr;
        end;
        Cell.Vm := C.ValueMetadata;
        Cell.V := C.AsString;
      end;
    end;
  end;
end;

procedure WriteDde(ADde: TXc12DdeLink);
var
  i,j: integer;
  DLink: TCT_DdeLink;
  DItem: TCT_DdeItem;
  DVal: TCT_DdeValue;
begin
  FXML.Root.RootAttributes.Clear;
  XML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);

  DLink := XML.ExternalLink.DdeLink;
  DLink.DdeService := ADde.DdeService;
  DLink.DdeTopic := ADde.DdeTopic;
  for i := 0 to ADde.Count - 1 do begin
    DItem := DLink.DdeItems.DdeItemXpgList.Add;
    DItem.Name := ADde[i].Name;
    DItem.Ole := ADde[i].Ole;
    DItem.Advise := ADde[i].Advise;
    DItem.PreferPic := ADde[i].PreferPic;
    for j := 0 to ADde[i].Values.Count - 1 do begin
      DVal := DItem.Values.ValueXpgList.Add;
      case ADde[i].Values[j].Type_ of
        x12ctNil    : DVal.T := stdvtNil;
        x12ctBoolean: DVal.T := stdvtB;
        x12ctError  : DVal.T := stdvtE;
        x12ctFloat  : DVal.T := stdvtN;
        x12ctString : DVal.T := stdvtStr;
      end;
      DVal.Val := ADde[i].Values[j].AsString;
    end;
  end;
end;

procedure WriteOle(AOle: TXc12OleLink);
var
  i: integer;
  O: TOPCItem;
  OLink: TCT_OleLink;
  OItem: TCT_OleItem;
begin
  O := FManager.FileData.OPC.CreateXLinkPath(OPC,AOle.Filename);

  FXML.Root.RootAttributes.Clear;
  XML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);
  XML.Root.RootAttributes.AddNameValue('xmlns:r',OOXML_URI_OFFICEDOC_RELATIONSHIPS);

  OLink := XML.ExternalLink.OleLink;
  OLink.R_Id := O.Id;
  OLink.ProgId := AOle.ProgId;
  for i := 0 to AOle.Count - 1 do begin
    OItem := OLink.OleItems.OleItemXpgList.Add;
    OItem.Name := AOle[i].Name;
    OItem.Icon := AOle[i].Icon;
    OItem.Advise := AOle[i].Advice;
    OItem.PreferPic := AOle[i].PreferPic;
  end;
end;

begin
  OPC := FManager.FileData.OPC.CreateXLink(AIndex);

  Stream := TMemoryStream.Create;
  try
    XML := TXPGDocXLink.Create;
    try
      XLink := FManager.XLinks[AIndex - 1];
      if XLink.ExternalBook.Filename <> '' then
        WriteExternalBook(XLink.ExternalBook);

      if XLink.DdeLink.Count > 0 then
        WriteDde(XLink.DdeLink);

      if XLink.OleLink.Count > 0 then
        WriteOle(XLink.OleLink);

      XML.SaveToStream(Stream);
    finally
      XML.Free;
    end;

    FManager.FileData.OPC.ItemWrite(OPC,Stream);
    FManager.FileData.OPC.ItemClose(OPC);
  finally
    Stream.Free;
  end;

  Result := OPC.Id;
end;

{ TXLSWriteSheetData }

constructor TXLSWriteSheetData.Create(AOwner: TXLSWriteXLSX; ASheet: TXc12DataWorksheet; ASheetData: TCT_Sheetdata; AComments: TCT_Comments);
begin
  FOwner := AOwner;
  FManager := FOwner.FManager;
  FCurrSheet := ASheet;
  FSheetData := ASheetData;
  FComments := AComments;
end;

procedure TXLSWriteSheetData.OnWriteCell(ASender: TObject; var AWriteElement: boolean);
var
//  C: PXLSCellItem;
  Cell: TCT_Cell;

procedure WriteFormula;
var
  FH: TXLSFormulaHelper;
begin
  FH := FCurrSheet.Cells.FormulaHelper;

  Cell.F.T := TST_CellFormulaType(FH.FormulaType);
  if FH.HasApply then
    Cell.F.Ref := FH.StrTargetRef;

  if FH.IsTABLE then begin
    Cell.F.R1 := FH.StrR1;
    if FH.StrR2 <> '' then
      Cell.F.R2 := FH.StrR2;
    Cell.F.DT2D := (FH.TableOptions and Xc12FormulaTableOpt_DT2D) <> 0;
    Cell.F.DTR  := (FH.TableOptions and Xc12FormulaTableOpt_DTR) <> 0;
    Cell.F.DEL1 := (FH.TableOptions and Xc12FormulaTableOpt_DEL1) <> 0;
    Cell.F.DEL2 := (FH.TableOptions and Xc12FormulaTableOpt_DEL2) <> 0;
  end
  else begin
    if FH.IsCompiled then begin
      Cell.F.Content := FOwner.FFormulas.DecodeFormula(FCurrSheet.Cells,FCurrSheet.Index,FH.Ptgs,FH.PtgsSize,True)
    end
    else
      Cell.F.Content := FH.Formula;
  end;

  Cell.F.Aca := (FH.Options and Xc12FormulaOpt_ACA) <> 0;
  Cell.F.CA := FOwner.FFormulas.FormulasUncalced;
//  Cell.F.CA  := (FH.Options and Xc12FormulaOpt_CA) <> 0;
  Cell.F.BX  := (FH.Options and Xc12FormulaOpt_BX) <> 0;
end;

begin
  AWriteElement := FCurrSheet.Cells.IterateNextCol;

  if AWriteElement then begin
    Cell := TCT_Cell(ASender);

    Cell.R := ColRowToRefStr(FCurrSheet.Cells.IterCellCol,FCurrSheet.Cells.IterCellRow);

    if FCurrSheet.Cells.IterGetStyleIndex <> XLS_STYLE_DEFAULT_XF then
      Cell.S := FCurrSheet.Cells.IterGetStyleIndex;

    case FCurrSheet.Cells.IterCellType of
      xctNone: ;
      xctBlank: ;
      xctBoolean: begin
        Cell.V := XMlBoolToStr(FCurrSheet.Cells.IterGetBoolean);
        Cell.T := stctB;
      end;
      xctError: begin
        Cell.V := CellErrorToErrorText(FCurrSheet.Cells.IterGetError);
        Cell.T := stctE;
      end;
      xctString: begin
        Cell.V := XMLIntToStr(FCurrSheet.Cells.IterGetStringIndex);
        Cell.T := stctS;
      end;
      xctFloat: TCT_Cell(ASender).V := XMLFloatToStr(FCurrSheet.Cells.IterGetFloat);
      xctFloatFormula : begin
        FCurrSheet.Cells.IterGetFormula;
        if not (FCurrSheet.Cells.FormulaHelper.FormulaType in [xcftArrayChild,xcftDataTableChild]) then
          WriteFormula;
        Cell.V := XMLFloatToStr(FCurrSheet.Cells.FormulaHelper.AsFloat);
      end;
      xctStringFormula : begin
        FCurrSheet.Cells.IterGetFormula;
        if not (FCurrSheet.Cells.FormulaHelper.FormulaType in [xcftArrayChild,xcftDataTableChild]) then
          WriteFormula;
        Cell.T := stctStr;
        Cell.V := FCurrSheet.Cells.FormulaHelper.AsString;
      end;
      xctBooleanFormula: begin
        FCurrSheet.Cells.IterGetFormula;
        if not (FCurrSheet.Cells.FormulaHelper.FormulaType in [xcftArrayChild,xcftDataTableChild]) then
          WriteFormula;
        Cell.T := stctB;
        Cell.V := XMLBoolToStr(FCurrSheet.Cells.FormulaHelper.AsBoolean);
      end;
      xctErrorFormula  : begin
        FCurrSheet.Cells.IterGetFormula;
        if not (FCurrSheet.Cells.FormulaHelper.FormulaType in [xcftArrayChild,xcftDataTableChild]) then
          WriteFormula;
        Cell.T := stctE;
        Cell.V := FCurrSheet.Cells.FormulaHelper.AsErrorStr;
      end;
    end;
  end;
end;

procedure TXLSWriteSheetData.OnWriteCellDirect(ASender: TObject; var AWriteElement: boolean);
var
  Cell: TCT_Cell;
begin
  FManager.FireWriteCellEvent(False);
  AWriteElement := FManager.EventCell.Active;
  if AWriteElement then begin
    if FManager.CheckDirectWrite then begin
      Cell := TCT_Cell(ASender);

      Cell.R := ColRowToRefStr(FManager.EventCell.Col,FManager.EventCell.Row);

//      if FCurrSheet.Cells.IterGetStyleIndex <> XLS_STYLE_DEFAULT_XF then
//        Cell.S := FCurrSheet.Cells.IterGetStyleIndex;

      case FManager.EventCell.CellType of
        xctNone: ;
        xctBlank: ;
        xctBoolean: begin
          Cell.V := XMlBoolToStr(FManager.EventCell.AsBoolean);
          Cell.T := stctB;
        end;
        xctError: begin
          Cell.V := CellErrorToErrorText(FManager.EventCell.AsError);
          Cell.T := stctE;
        end;
        xctString: begin
          Cell.V := FManager.EventCell.AsString;
          Cell.T := stctStr;
        end;
        xctFloat: Cell.V := XMLFloatToStr(FManager.EventCell.AsFloat);
        xctFloatFormula : begin
          Cell.F.Content := FManager.EventCell.AsFormula;
          Cell.V := XMLFloatToStr(FManager.EventCell.AsFloat);
        end;
        xctStringFormula : begin
          Cell.F.Content := FManager.EventCell.AsFormula;
          Cell.V := FManager.EventCell.AsString;
          Cell.T := stctStr;
        end;
        xctBooleanFormula: begin
          Cell.F.Content := FManager.EventCell.AsFormula;
          Cell.V := XMlBoolToStr(FManager.EventCell.AsBoolean);
          Cell.T := stctB;
        end;
        xctErrorFormula  : begin
          Cell.F.Content := FManager.EventCell.AsFormula;
          Cell.V := CellErrorToErrorText(FManager.EventCell.AsError);
          Cell.T := stctE;
        end;
      end;
      if FManager.EventCell.XF <> Nil then
        Cell.S := FManager.EventCell.XF.Index;
    end
    else
      FManager.Errors.Error('',XLSERR_DIRECTWR_CELL_NOT_INC);
  end;
  FManager.EventCell.ClearValues;
end;

procedure TXLSWriteSheetData.OnWriteRow(ASender: TObject; var AWriteElement: boolean);
var
  Row: TCT_Row;
  R: PXLSMMURowHeader;
begin
  AWriteElement := FCurrSheet.Cells.IterateNextRow;
  if AWriteElement then begin
    Row := TCT_Row(ASender);
    Row.Clear;

    R := FCurrSheet.Cells.IterRow;
    Row.R := FCurrSheet.Cells.IterCellRow + 1;

    if R.Height <> XLS_DEFAULT_ROWHEIGHT_FLAG then
      Row.Ht := R.Height / 20
    else
      Row.Ht := 0;
    Row.Hidden := xroHidden in R.Options;
    Row.CustomFormat := R.Style <> 0;
    if Row.CustomFormat then
      Row.S := R.Style;
    Row.Collapsed := xroCollapsed in R.Options;
    Row.OutlineLevel := R.OutlineLevel;
    Row.CustomHeight := xroCustomHeight in R.Options;
    Row.Ph := xroPhonetic in R.Options;
    Row.ThickTop := xroThickTop in R.Options;
    Row.ThickBot := xroThickBottom in R.Options;

//    if R.Spans <> '' then
//      Row.SpansXpgList.Add(R.Spans);

    FCurrSheet.Cells.BeginIterateCol;
  end;
end;

procedure TXLSWriteSheetData.WriteComments(AXML: TXPGDocXLSX; AOwner: TOPCItem);
var
  i: integer;
  S: AxUCString;
  Stream: TMemoryStream;
  OPC: TOPCItem;
begin
  Stream := TMemoryStream.Create;
  try
    OPC := FManager.FileData.OPC.CreateComment(AOwner,FCurrSheet.Index + 1);

    AXML.Root.RootAttributes.Clear;
    AXML.Root.RootAttributes.AddNameValue('xmlns',OOXML_URI_SPREADSHEETML_MAIN);

    FCurrIndex := 0;
    for i := 0 to FCurrSheet.Comments.Count - 1 do begin
      S := FCurrSheet.Comments[i].Author.Name;
      if FComments.Authors.AuthorXpgList.IndexOf(S) < 0 then begin
        FCurrSheet.Comments[i].AuthorId := FComments.Authors.AuthorXpgList.Count;
        FComments.Authors.AuthorXpgList.Add(S);
      end;
    end;

    FComments.CommentList.OnWriteComment := OnWriteComment;
    AXML.SaveToStream(Stream,TCT_Comments);

    FManager.FileData.OPC.ItemWrite(OPC,Stream);
    FManager.FileData.OPC.ItemClose(OPC);
  finally
    Stream.Free;
  end;
end;

procedure TXLSWriteSheetData.OnWriteCol(ASender: TObject; var AWriteElement: boolean);
var
  C: TXc12Column;
  Col: TCT_Col;
begin
  AWriteElement := FCurrCol < FCurrSheet.Columns.Count;
  if AWriteElement then begin
    C := FCurrSheet.Columns[FCurrCol];
    Col := TCT_Col(ASender);

    Col.Min := C.Min + 1;
    Col.Max := C.Max + 1;
    Col.Width := C.Width;
    Col.CustomWidth := xcoCustomWidth in C.Options;
    Col.Collapsed := xcoCollapsed in C.Options;
    Col.Hidden := xcoHidden in C.Options;
    Col.OutlineLevel := C.OutlineLevel;
    Col.BestFit := xcoBestFit in C.Options;
    Col.Phonetic := xcoPhonetic in C.Options;
    Col.Style := C.Style.Index;

    Inc(FCurrCol);
  end;
end;

procedure TXLSWriteSheetData.OnWriteComment(ASender: TObject; var AWriteElement: boolean);
var
  C: TXc12Comment;
  Comment: TCT_Comment;
begin
  AWriteElement := FCurrIndex < FCurrSheet.Comments.Count;

  if AWriteElement then begin
    Comment := TCT_Comment(ASender);

    C := FCurrSheet.Comments[FCurrIndex];
    Comment.Ref := CellAreaToAreaStr(C.Ref);

    Comment.AuthorId := C.AuthorId; // C.Author.Index;
    Comment.Guid := C.GUID;

    FOwner.GetRst(C.Text,C.FontRuns,C.PhoneticRuns,Comment.Text);
  end;

  Inc(FCurrIndex);
end;

procedure TXLSWriteSheetData.OnWriteRowDirect(ASender: TObject; var AWriteElement: boolean);
var
  Row: TCT_Row;
begin
  FManager.FireWriteCellEvent(True);
  AWriteElement := FManager.EventCell.Active;
  if AWriteElement then begin
    if FManager.CheckDirectWrite then begin
      Row := TCT_Row(ASender);
      Row.Clear;

      Row.R := FManager.EventCell.Row + 1;
    end
    else
      FManager.Errors.Error('',XLSERR_DIRECTWR_CELL_NOT_INC);
  end;
end;

procedure TXLSWriteSheetData.Initiate;
begin
  FCurrCol := 0;
  if FManager.DirectWrite and (FCurrSheet.Index = FManager.EventCell.SheetIndex) then begin
    FManager.EventCell.Clear;
    FSheetData.OnWriteRow := OnWriteRowDirect;
    FSheetData.Row.OnWriteC := OnWriteCellDirect;
  end
  else begin
    FSheetData.OnWriteRow := OnWriteRow;
    FSheetData.Row.OnWriteC := OnWriteCell;
    FCurrSheet.Cells.BeginIterateRow;
  end;
  FComments.CommentList.OnWriteComment := OnWriteComment;
end;

end.
