unit XLSExportHTML5;

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
     Xc12Utils5, Xc12DataStylesheet5, Xc12DataWorksheet5,
     XLSUtils5, XLSTools5, XLSExport5, XLSMMU5, XLSCellMMU5, XLSSheetData5,
     XLSCellAreas5, XLSMergedCells5, XLSHyperlinks5, XLSDrawing5, XLSComment5;


type TXLSExportOptionHTML = (xeohComments,xeohImages,xeohWriteImages,xeohHyperlinks);
     TXLSExportOptionsHTML = set of TXLSExportOptionHTML;

//* Settings for the tables created on the destionation file.
type TTABLEProperties = class(TPersistent)
private
     FBordeWidth: integer;
     FCellPadding: integer;
     FCellSpacing: integer;
protected
public
published
     //* Border width of the cells on the table.
     property BordeWidth: integer read FBordeWidth write FBordeWidth;
     //* Cell padding of the cells on the table.
     property CellPadding: integer read FCellPadding write FCellPadding;
     //* Cell spacing of the cells on the table.
     property CellSpacing: integer read FCellSpacing write FCellSpacing;
     end;

//* Exports a workbook to a HTML file. On the destionation file one table is created
//* for each worksheet of the source TXLSReadWriteII object.
type TXLSExportHTML5 = class(TXLSExport5)
private
    function  GetWriteOnlyTables: boolean;
    procedure SetWriteOnlyTables(const Value: boolean);
    function  GetTitle: AxUCString;
protected
    FStream      : TStream;

    FLinkTargets : array of TXLSCellRef;
    FTargetCount : integer;

    FImageCount  : integer;
    FCommentCount: integer;

    FTitle       : AxUCString;
    FTABLE       : TTABLEProperties;

    FSimpleExport: boolean;
    FHTMLOptions : TXLSExportOptionsHTML;

    procedure WString(S: AxUCString);
    procedure WStringCR(S: AxUCString);

    procedure AddLinkTarget(const ACol,ARow: integer);
    function  FindLinkTarget(const ACol,ARow: integer): boolean;

    procedure WriteCSS;
    function  GetCellHTML(Sheet: TXLSWorksheet; const ACell: TXLSCellItem): AxUCString;
    function  GetImageHTML(Sheet: TXLSWorksheet; AImage: TXLSDrawingImage; const AImageName: AxUCString): AxUCString;

    function  ExportImages: boolean; override;

    procedure OpenFile;        override;
    procedure WriteFilePrefix; override;
    procedure WritePagePrefix; override;
    procedure WriteRowPrefix(const ARow: integer);  override;
    procedure WriteCell(const Col,Row: integer; const IsFirstCol,IsFirstRow: boolean); override;
    procedure WriteCellSimple(const Col,Row: integer; const IsFirstCol,IsFirstRow: boolean);
    procedure WriteRowSuffix;  override;
    procedure WritePageSuffix; override;
    procedure WriteFileSuffix; override;
    procedure CloseFile;       override;
public
    //* ~exclude
    constructor Create(AOwner: TComponent); override;
    //* ~exclude
    destructor Destroy; override;
    //* Writes the data to a stream.
    //* ~param Stream The destination stream.
    procedure SaveToStream(Stream: TStream); override;
published
    //* Title of HTML document. Default is the XLS filename.
    property Title: AxUCString read GetTitle write FTitle;
    //* Obsolete. Can't be implemented when CSS is used.
    property WriteOnlyTables: boolean read GetWriteOnlyTables write SetWriteOnlyTables;
    //* When SimpleExport is True, simplest, and smallest, possible file will
    //* be written. There will not be any formatting, images, hyperlinks etc.
    property SimpleExport: boolean read FSimpleExport write FSimpleExport;
    //* Settings for the tables created on the destionation file.
    property HTMLOPtions: TXLSExportOptionsHTML read FHTMLOptions write FHTMLOptions;
    //* HTML export options.
    //* xeohComments = Include comments. The comment texts will be placed at the
    //* end of the worksheet.
    //* xeohImages = Include images.
    //* xeohWriteImages = write image files to disk.
    //* xeohHyperlinks = include hyperlinks.
    property TABLE: TTABLEProperties read FTABLE write FTABLE;
    end;

implementation

{ TXLSExport }

procedure TXLSExportHTML5.AddLinkTarget(const ACol, ARow: integer);
begin
  if Length(FLinkTargets) <=  FTargetCount then
    SetLength(FLinkTargets,Length(FLinkTargets) + 64);
  FLinkTargets[FTargetCount].Col := ACol;
  FLinkTargets[FTargetCount].Row := ARow;
  Inc(FTargetCount);
end;

procedure TXLSExportHTML5.CloseFile;
begin
  inherited;
  FStream.Free;
  FStream := Nil;
end;

constructor TXLSExportHTML5.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FFileExtension := 'htm';
  FTABLE := TTABLEProperties.Create;
  FHTMLOptions := [xeohComments,xeohImages,xeohWriteImages,xeohHyperlinks];
end;

destructor TXLSExportHTML5.Destroy;
begin
  FTABLE.Free;
  inherited;
end;

function TXLSExportHTML5.ExportImages: boolean;
begin
  Result := xeohImages in FHTMLOptions;
end;

function TXLSExportHTML5.FindLinkTarget(const ACol, ARow: integer): boolean;
var
  i: integer;
begin
  for i := 0 to FTargetCount - 1 do begin
    Result := (FLinkTargets[i].Col = ACol) and (FLinkTargets[i].Row = ARow);
    if Result then
      Exit;
  end;
  Result := False;
end;

function TXLSExportHTML5.GetCellHTML(Sheet: TXLSWorksheet; const ACell: TXLSCellItem): AxUCString;
var
  FT: TXLSFormattedText;
begin
  FT := Sheet.MakeFormattedText(ACell);
  if FT <> Nil then begin
    try
      Result := FT.AsHTML;
    finally
      FT.Free;
    end;
  end
  else
    Result := '';
end;

function TXLSExportHTML5.GetImageHTML(Sheet: TXLSWorksheet; AImage: TXLSDrawingImage; const AImageName: AxUCString): AxUCString;
var
  W,H: integer;
begin
  Sheet.GetShapePixelSize(AImage,W,H);
  Inc(W,Round(Sheet.Columns[AImage.Col1].PixelWidth * AImage.Col1Offs));
  Inc(H,Round(Sheet.Rows[AImage.Row1].PixelHeight * AImage.Row1Offs));
  Result := Format('<span style="position:absolute;z-index:1;"><img src="file:///%s" width="%d" height="%d"></span>',[AImageName,W,H]);
end;

function TXLSExportHTML5.GetTitle: AxUCString;
begin
  if (FTitle = '') and (FXLS <> Nil) then
    Result := FXLS.Filename
  else
    Result := FTitle;
end;

function TXLSExportHTML5.GetWriteOnlyTables: boolean;
begin
  Result := False;
end;

procedure TXLSExportHTML5.OpenFile;
begin
  FStream := TFileStream.Create(FCurrFilename,fmCreate);

  SetLength(FLinkTargets,0);
  FTargetCount := 0;
  FImageCount := 0;
  FCommentCount := 0;

  FXLS.Manager.GrManager.Images.SetWritten(True);

  inherited;
end;

procedure TXLSExportHTML5.SaveToStream(Stream: TStream);
begin
  FStream := Stream;

  inherited;
end;

procedure TXLSExportHTML5.SetWriteOnlyTables(const Value: boolean);
begin

end;

procedure TXLSExportHTML5.WriteCell(const Col, Row: integer; const IsFirstCol,IsFirstRow: boolean);
var
  i,j: integer;
  C,R: integer;
  v: integer;
  sStyle: AxUCString;
  sText: AxUCString;
  sImage: AxUCString;
  sHLink: AxUCString;
  sTDAttr: AxUCString;
  sComment: AxUCString;
  sHA: AxUCString;
  tagTD: AxUCString;
  Sheet: TXLSWorksheet;
  Cell: TXLSCellItem;
  XF: TXc12XF;
  HLink: TXLSHyperlink;
  Image: TXLSDrawingImage;
begin
  inherited;

  if FSimpleExport then begin
    WriteCellSimple(Col,Row,IsFirstCol,IsFirstRow);
    Exit;
  end;

  Sheet := FXLS[FCurrSheetIndex];

  if (Sheet.Columns[Col].Width <= 0) or Sheet.Columns[Col].Hidden then
    Exit;

  tagTD := '<td';

  sStyle := '';
  sText := '&nbsp';
  sImage := '';
  sHLink := '';
  sTDAttr := '';
  sComment := '';

  Image := Nil;

  i := Sheet.MergedCells.FindArea(Col,Row);
  if i >= 0 then begin
    if (Sheet.MergedCells[i].Col1 = Col) and (Sheet.MergedCells[i].Row1 = Row) then begin
      sTDAttr := sTDAttr + ' colspan=' + IntToStr(Sheet.MergedCells[i].Col2 - Sheet.MergedCells[i].Col1 + 1);
      sTDAttr := sTDAttr + ' rowspan=' + IntToStr(Sheet.MergedCells[i].Row2 - Sheet.MergedCells[i].Row1 + 1);

      sStyle := sStyle + 'white-space:normal;';

      if Sheet.MMUCells.FindCell(Sheet.MergedCells[i].Col2,Sheet.MergedCells[i].Row2,Cell) then begin
        XF := FXLS.Manager.StyleSheet.XFs[Sheet.MMUCells.GetStyle(@Cell)];
        sStyle := sStyle + XF.BordersAsCSS;
      end;
      // Todo Check if there is an image in the rest of the cells.
    end
    else
      Exit;
  end;

  if xeohImages in FHTMLOptions then begin
    Image := Sheet.Drawing.Images.FindTopLeft(Col,Row);

    if Image <> Nil then begin
      Inc(FImageCount);
      Image.TagForWriting;
    end;
  end;

  if IsFirstCol or IsFirstRow then begin
    if IsFirstCol then
      sStyle := sStyle + Format('height:%dpt;',[Round((Sheet.Rows[Row].Height / 20) * 0.75)]);
    if IsFirstRow then begin
      if i >= 0 then begin
        v := 0;
        for j := Sheet.MergedCells[i].Col1 to Sheet.MergedCells[i].Col2 do
          Inc(v,Sheet.Columns[j].PixelWidth);
        sStyle := sStyle + Format('width:%dpt;',[Round(v * 0.75)]);
      end
      else
        sStyle := sStyle + Format('width:%dpt;',[Round(Sheet.Columns[Col].PixelWidth * 0.75)]);
    end;
  end;

  if (xeohComments in FHTMLOptions) and (Sheet.Comments.Find(Col,Row) <> Nil) then begin
    Inc(FCommentCount);
    sComment := Format('[%d]',[FCommentCount]);
  end;

  if (xeohHyperlinks in FHTMLOptions) and FindLinkTarget(Col,Row) then
    sHLink := '<a name="' + AxUCString(XLSUTF8Encode(Sheet.Name)) + '_' + ColRowToRefStr(Col,Row) + '">';

  if not Sheet.MMUCells.FindCell(Col,Row,Cell) then
    XF := FXLS.Manager.StyleSheet.XFs[XLS_STYLE_DEFAULT_XF]
  else if Sheet.MMUCells.CellType(@Cell) = xctBlank then
    XF := FXLS.Manager.StyleSheet.XFs[Sheet.MMUCells.GetStyle(@Cell)]
  else begin
    XF := FXLS.Manager.StyleSheet.XFs[Sheet.MMUCells.GetStyle(@Cell)];

    sHA := '';
    case XF.Alignment.HorizAlignment of
      chaGeneral          : begin
        case Sheet.MMUCells.CellType(@Cell) of
          xctBlank         : ;
          xctBoolean       : sHA := 'center';
          xctError         : ;
          xctString        : ;
          xctFloat         : sHA := 'right';
          xctFloatFormula  : sHA := 'right';
          xctStringFormula : ;
          xctBooleanFormula: sHA := 'center';
          xctErrorFormula  : ;
        end;
      end;
      chaCenter           : sHA := 'center';
      chaRight            : sHA := 'right';
      chaFill             : ;
      chaJustify          : sHA := 'justify';
      chaCenterContinuous : sHA := 'center';
      chaDistributed      : ;
    end;
    if sHA <> '' then
      sTDAttr := sTDAttr + ' style="text-align:' + sHA + '";';

    sText := GetCellHTML(Sheet,Cell);
    if sText = '' then
      sText := '&nbsp;'
    else if xeohHyperlinks in FHTMLOptions then begin
      HLink := Sheet.Hyperlinks.Find(Col,Row);
      if HLink <> Nil then begin
        case HLink.HyperlinkType of
          xhltURL     : sHLink := '<a href="' + HLink.Address + '">';
          xhltFile    : sHLink := '<a href="' + HLink.Address + '">';
          xhltUNC     : sHLink := '<a href="' + HLink.Address + '">';
          xhltWorkbook: begin
            if HLink.GetWorkbookTarget(C,R) then
              sHLink := '<a href="#' + AxUCString(XLSUTF8Encode(Sheet.Name)) + '_' + ColRowToRefStr(C,R) + '">';
          end;
        end;
      end;
    end;
  end;

  if Image <> Nil then begin
    sImage := sImage + GetImageHTML(Sheet,Image,Directory + Image.UniqueName);
    if xeohWriteImages in FHTMLOptions then
      Image.SaveToFile(Directory + Image.UniqueName);
  end;

  tagTD := '<td';

  if sTDAttr <> '' then
    tagTD := tagTD + sTDAttr;

  if XF.Index <> XLS_STYLE_DEFAULT_XF then
    tagTD := tagTD + ' class="' + XF.CSSSelector + '"';

  if sStyle <> '' then
    tagTD := tagTD + ' style="' + sStyle + '"';

  tagTD := tagTD + '>';

  if sImage <> '' then
    tagTD := tagTD + sImage;

  if sHLink <> '' then
    sText := sHLink + sText + '</a>';

  if sComment <> '' then
    sText := sText + Format('<a href="#%s_Comment%d">%s</a>',[XLSUTF8Encode(Sheet.Name),FCommentCount,sComment]);

  WStringCR(tagTD + sText + '</td>');
end;

procedure TXLSExportHTML5.WriteCellSimple(const Col, Row: integer; const IsFirstCol,IsFirstRow: boolean);
var
  tagTD: AxUCString;
  sText: AxUCString;
  sStyle: AxUCString;
  sTDAttr: AxUCString;
  Cell: TXLSCellItem;
  Sheet: TXLSWorksheet;
begin
  Sheet := FXLS[FCurrSheetIndex];

  tagTD := '<td';

  sText := '&nbsp';
  sTDAttr := '';

  if IsFirstRow then
    sStyle := sStyle + Format('width:%dpt;',[Round(Sheet.Columns[Col].PixelWidth * 0.75)]);

  if Sheet.MMUCells.FindCell(Col,Row,Cell) and (Sheet.MMUCells.CellType(@Cell) <> xctBlank) then begin
    if not (Sheet.MMUCells.CellType(@Cell) in [xctString,xctBoolean,xctError,xctStringFormula]) then
      sTDAttr := sTDAttr + ' style="text-align:right";';

    sText := GetCellHTML(Sheet,Cell);
    if sText = '' then
      sText := '&nbsp;'
  end;

  tagTD := '<td';

  if sTDAttr <> '' then
    tagTD := tagTD + sTDAttr;

  tagTD := tagTD + '>';

  WStringCR(tagTD + sText + '</td>');
end;

procedure TXLSExportHTML5.WriteCSS;
var
  i: integer;
  S: AxUCString;
begin
  WStringCR('<style>');

  S := 'tr {' + #13;
  S := S + Format('height:%dpt;',[Round((XLS_DEFAULT_ROWHEIGHT / 20) * 0.75)]);
  S := S + '}' + #13;
  WStringCR(S);

  S := 'td {' + #13;
	S := S + 'padding-top:1px;';
	S := S + 'padding-right:1px;';
	S := S + 'padding-left:1px;';
  S := S + 'white-space:nowrap;';
  S := S + FXLS.Manager.StyleSheet.Fonts.DefaultFont.AsCSS(False);
  S := S + '}' + #13;
  WStringCR(S);

  for i := 0 to FXLS.Manager.StyleSheet.Fonts.Count - 1 do begin
    S := '.' + FXLS.Manager.StyleSheet.Fonts[i].AsCSS;
    WStringCR(S);
  end;
  for i := 0 to FXLS.Manager.StyleSheet.XFs.Count - 1 do begin
    if FXLS.Manager.StyleSheet.XFs[i] <> Nil then begin
      S := '.' + FXLS.Manager.StyleSheet.XFs[i].AsCSS;
      WStringCR(S);
    end;
  end;
  WStringCR('</style>');
end;

procedure TXLSExportHTML5.WriteFilePrefix;
begin
  inherited;
  WStringCR('<!DOCTYPE html>');
  WStringCR('<html><head>');
  WStringCR('<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">');
  WStringCR('<title>' + AxUCString(XLSUTF8Encode(FXLS.Filename)) + '</TITLE>');
  WriteCSS;
  WStringCR('</head>');
  WStringCR('<body bgcolor="#FFFFFF">');
end;

procedure TXLSExportHTML5.WriteFileSuffix;
begin
  inherited;
  WStringCR('</body>');
  WStringCR('</html>');
end;

procedure TXLSExportHTML5.WritePagePrefix;
var
  i: integer;
  S: AxUCString;
  sFont: AxUCString;
  C,R: integer;
  w: integer;
  Sheet: TXLSWorksheet;
begin
  inherited;

  Sheet := FXLS[FCurrSheetIndex];

  if xeohHyperlinks in FHTMLOptions then begin
    for i := 0 to Sheet.Hyperlinks.Count - 1 do begin
      if (Sheet.Hyperlinks[i].HyperlinkType = xhltWorkbook) and Sheet.Hyperlinks[i].GetWorkbookTarget(C,R) then begin
        if not FindLinkTarget(C,R) then
          AddLinkTarget(C,R);
      end;
    end;
  end;

  if xeohComments in FHTMLOptions then begin
    for i := 0 to Sheet.Comments.Count - 1 do begin
      C := Sheet.Comments[i].Col;
      R := Sheet.Comments[i].Row;
      if not FindLinkTarget(C,R) then
        AddLinkTarget(C,R);
    end;
  end;


  WStringCR('<hr>');
  sFont := FXLS.Manager.StyleSheet.Fonts.DefaultFont.CSSSelector;
  S := Format('<div class="%s" style="font-size:large;font-weight:bold;text-align:center;">%s</div>',[sFont,XLSUTF8Encode(Sheet.Name)]);
  WStringCR(S);
  WStringCR('<hr>');

  w := 0;
  for i := FFirstCol to FXLS[FCurrSheetIndex].LastCol do
    Inc(w,FXLS[FCurrSheetIndex].Columns[i].PixelWidth);

  WStringCR(Format('<table border=%d cellpadding=%d cellspacing=%d style="border-collapse:collapse;table-layout:fixed;width:%dpt">',
     [FTABLE.BordeWidth,FTABLE.CellPadding,FTABLE.CellSpacing,Round(w * 0.75)]));
end;

procedure TXLSExportHTML5.WritePageSuffix;
var
  i: integer;
  S: AxUCString;
  sFont: AxUCString;
  Sheet: TXLSWorksheet;
  Comment: TXLSComment;
begin
  inherited;

  WStringCR('</table>');

  Sheet := FXLS[FCurrSheetIndex];

  if (xeohComments in FHTMLOPtions) and (Sheet.Comments.Count > 0) then begin
    WStringCR('<hr>');

    sFont := FXLS.Manager.StyleSheet.Fonts.DefaultFont.CSSSelector;

    for i := 0 to Sheet.Comments.Count - 1 do begin
      Comment := Sheet.Comments[i];
      S := Format('<div class="%s"><a name="%s_Comment%d"><a href="#%s_%s">[%d]</a></div>',[sFont,XLSUTF8Encode(Sheet.Name),i + 1,XLSUTF8Encode(Sheet.Name),ColRowToRefStr(Comment.Col,Comment.Row),i + 1]);
      WStringCR(S);
      if Comment.Author <> '' then
        WStringCR('<div class="' + sFont + '"><b>' + XLSUTF8Encode(Comment.Author) + ':</b></div>');
      WStringCR('<div class="' + sFont + '">' + XLSUTF8Encode(Comment.PlainTextNoAuthor) + '</div>');
    end;
  end;
end;

procedure TXLSExportHTML5.WriteRowPrefix(const ARow: integer);
var
  Row: PXLSMMURowHeader;
begin
  inherited;

  Row := FXLS[FCurrSheetIndex].MMUCells.FindRow(ARow);
  if (Row <> Nil) and (Row.Height <> XLS_DEFAULT_ROWHEIGHT) and (Row.Height <= 8190) then
    WStringCR(Format('<tr style="height:%dpt">',[Round((Row.Height / 20) * 0.75)]))
  else
    WStringCR('<tr>');
end;

procedure TXLSExportHTML5.WriteRowSuffix;
begin
  inherited;
  WStringCR('</tr>');
end;

procedure TXLSExportHTML5.WString(S: AxUCString);
var
  AStr: AnsiString;
begin
  AStr := AnsiString(S);
  FStream.Write(Pointer(AStr)^,Length(AStr));
end;

procedure TXLSExportHTML5.WStringCR(S: AxUCString);
begin
  WString(S + #13#10);
end;

end.
