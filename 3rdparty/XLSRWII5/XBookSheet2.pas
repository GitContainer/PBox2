unit XBookSheet2;

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

uses { Delphi  } Classes, SysUtils, vcl.Controls, vcl.StdCtrls, vcl.Forms, Windows, Math,
                 Contnrs, StrUtils,
     { XLSRWII } Xc12Utils5, Xc12DataStyleSheet5, Xc12Manager5, Xc12DataWorkbook5,
                 XLSCondFormat5, XLSHyperlinks5,XLSUtils5, XLSCellAreas5,
                 Xc12DataWorksheet5, Xc12DataAutofilter5,
                 XLSMMU5, XLSCellMMU5, XLSSheetData5, XLSDrawing5,
                 XLSComment5, XLSMergedCells5, XLSPivotTables5, FormPivotTable,
     { XLSBook } XBookSkin2, XBookOptions2, XBookRows2, XBookColumns2, XBookSysVar2,
                 XBookTabSet2, XBookRects2, XBookPaintGDI2, XBookWindows2,
                 XBookPaint2, XBookUtils2, XBook_System_2,
                 XBookHeaders2, XBookButtons2, XBookComponent2, XBookEditSelArea2,
                 XBookRefEdit2, XBookTypes2,XBookPaintLayers2, XBookRichPainter2,
                 XBookInplaceEdit2, XBookControls2, XBookDrawing2;

type TXSSSelectOption = (xsoHyperlink,xsoShape,xsoPicture);
     TXSSSelectOptions = set of TXSSSelectOption;

type TCellsToPaint = (ctpAll,ctpSelected,ctpInvSelected);
type TPointerArray = array[0..$FFFF] of Pointer;

type TCellPaintBorder = (cpbLeft,cpbTop,cpbRight,cpbBottom);

type TCellCorners = (ccN,ccS,ccE,ccW);
type TCellCorner = array[ccN..ccW] of TXc12CellBorderStyle;

type TCellStyleFlag = (csfHorizOverride,csfTextExpandLeft,csfTextExpandRight,csfComment);
     TCellStyleFlags = set of TCellStyleFlag;

type TXSSChildWindows = (xcwUnknown,xcwSheet,xcwColumns,xcwRows,xcwTopLeft,xcwTabs,xcwVertScroll,xcwHorizScroll,xcwBottomRight,xcwDrawing);

const COND_FMT_VAL_SCALE = 1000000;

type TCellStyleRec = record
     Ref         : TXLSCellItem;
     DXF         : TXc12DXF;
     CondFmtRes  : TXSSCondFmtResultType;
     CondFmtVal  : integer;
     CondFmtColor: longword;
     ForeColor   : longword;
     BackColor   : longword;
     Pattern     : TXc12FillPattern;
     RightColor,
     BottomColor : longword;
     Flags       : TCellStyleFlags;
     InMerged    : integer;
     Hyperlink   : TXLSHyperlink;
     StyleNW     : TCellCorner;
     StyleNE     : TCellCorner;
     StyleSW     : TCellCorner;
     StyleSE     : TCellCorner;
     Diagonal    : TDiagLines;
     DiagColor   : longword;
     DiagStyle   : TXc12CellBorderStyle;
     end;

type TCellStyleRecArray = array[0..$FFFF] of TCellStyleRec;
type PCellStyleRecArray = ^TCellStyleRecArray;

type TCellLineCacheRec = record
     Pos       : TXYRect;
     PenColor  : longword;
     BackColor : longword;
     Style     : TPaintLineStyle;
     end;

type TCellPaintData = class(TObject)
private
     FCellStyle: TCellStyleRec;
     FCellRect: TXYRect;
     FPaintRect: TXYRect;
     FRightPos,FBottomPos: TXYRect;

     FIsPrinting    : boolean;
     FVertLCache    : array of TCellLineCacheRec;
     FHorizLCache   : array of TCellLineCacheRec;
     FLineCacheCount: integer;

     procedure SetIsPrinting(const Value: boolean);
public
     procedure Clear(X1,Y1,X2,Y2: integer; CellStyle: TCellStyleRec);
     procedure Calc(CSR: TCellStyleRec; Left,Top,Right,Bottom: TXc12CellBorderStyle);
     procedure CalcPaintRect(Left, Top, Right, Bottom: TXc12CellBorderStyle);
     procedure PaintBorders(Skin: TXLSBookSkin; Right,Bottom: longword);
     procedure PaintBorderLeft(Skin: TXLSBookSkin; AColor: longword; AStyle: TXc12CellBorderStyle);
     procedure PaintBorderTop(Skin: TXLSBookSkin; AColor: longword; AStyle: TXc12CellBorderStyle);
     procedure PaintBordersTopLeft(Skin: TXLSBookSkin; ATop,ALeft: TXc12BorderPr);
     procedure PaintCell(Skin: TXLSBookSkin);
     procedure PaintDataBar(Skin: TXLSBookSkin; const APercent: double; const AColor: longword);
     procedure PaintIcon(Skin: TXLSBookSkin; const AIconSet: TXc12IconSetType; const AIndex: integer; const ACentre: boolean);
     procedure SetCacheSize(ACols,ARows: integer);
     procedure PaintLineCache(ASkin: TXLSBookSkin);
     procedure PaintDiagonal(ASkin: TXLSBookSkin);

     property IsPrinting: boolean read FIsPrinting write SetIsPrinting;
     end;

type TCursorAreaHit = (cahNo,cahInside,cahEdge,cahSize);

type TXLSBookSheet = class(TXSSClientWindow)
private
     function  GetLeftCol: integer;
     procedure SetLeftCol(Value: integer);
     procedure PaintInvalidate;
     function  GetRightCol: integer;
     procedure SetRightCol(const Value: integer);
     procedure SetTopRow(Value: integer);
     function  GetBottomRow: integer;
     function  GetTopRow: integer;
     procedure SetBottomRow(Value: integer);
protected
     FOwner         : TCustomControl;
     FLayers        : TXPaintLayers;
     FManager       : TXc12Manager;
     FXLSSheet      : TXLSWorkSheet;
     FWorkbook      : TXLSWorkbook;
     FCells         : TXLSCellMMU;
     FOptions       : TXLSBookOptions;
     FSelectOptions : TXSSSelectOptions;
     FTopLeft       : TCornerRect;
     FTopTopLeft    : TCornerRect;
     FBottomRight   : TCornerRect;
     FRows          : TXLSBookRows;
     FColumns       : TXLSBookColumns;
     FTabs          : TXTabSet;
     FVScroll       : TXLSScrollBar;
     FHScroll       : TXLSScrollBar;
     FHSplit        : TSplitterHandle;
     FVSplit        : TSplitterHandle;
     FCSplit        : TSplitterHandle;
     FHideHSplitter,
     FHideVSplitter : boolean;
     FColGroupBtns  : TGroupButtons;
     FRowGroupBtns  : TGroupButtons;
     FLastX         : integer;
     FLastDragTab   : integer;
     FSheetArea     : TXLSCellArea;
     FCursorXY,
     FCursorXYMargO,
     FCursorXYMargI : TXYRect;
     FSizeCol,
     FSizeRow       : integer;
     FScrollTimerId : integer;
     FCommentTimerId: integer;
     FCellStyleCache: PCellStyleRecArray;
     FCellCacheCols,
     FCellCacheRows : integer;
     FLastMouseDown : TXYPoint;
     FMouseIsDown   : boolean;
     FSelectionEdit : TXLSBookEditSelArea;
     FEditRefs      : TXBookEditRefList;
     FCursorSizePos : TCursorSizePos;
     FCellPaint     : TCellPaintData;
     FInMergedCells : integer;
     FDrawing       : TXBookDrawing;
     FSheetChanged  : boolean;
     FCacheRequest  : boolean;
     FVisibleComment: TXLSComment;
     FInplaceEditor : TXBookInplaceEditor;
     FCopyAreas     : TCellAreas;
     FDeleteOptions : TXLSDeleteOptions;

     FIsPrinting           : boolean;
     FPrintStdFontWidth    : integer;
     FPrintMaxX            : integer;
     FPrintMaxY            : integer;

     FFormPivot            : TfrmPivotTable;

     FCellChangedEvent     : TXColRowEvent;
     FSelectionChangedEvent: TNotifyEvent;
     FHyperlinkClickEvent  : TXStringEvent;

     FIECharFmtEvent       : TNotifyEvent;
     FEditorOpenEvent      : TXBooleanEvent;
     FEditorCloseEvent     : TXBooleanEvent;
     FEditorDisableFmtEvent: TNotifyEvent;
     FNotificationEvent    : TXNotifyEvent;

     procedure SetTopRowPrint(Value: integer; MaxRowToCalc: integer);
     procedure SetVisibleRects;
     function  CalcVScroll: boolean;
     function  CalcHScroll: boolean;
     procedure TabSizeChanged(Sender: TObject; X,Y: integer);
     procedure RepaintRequest(Sender: TObject);
     procedure PaintLine(Sender: TObject; LineType: TPaintLineType; var X,Y: integer; Execute: boolean);
     procedure PaintCursor(C1,R1,C2,R2: integer);
     procedure PaintMarchingAnts(C1,R1,C2,R2: integer);
     procedure PaintCells(Col1,Row1,Col2,Row2: integer; CellsToPaint: TCellsToPaint);
     procedure PaintTheCell(Col,Row: integer; pX1,pY1,pX2,pY2: integer; Selected: boolean = False);
     procedure PaintCellText(Col,Row: integer; Ref: TXLSCellItem; Rect: TXYRect; ADXF: TXc12DXF);
     procedure ClearSelected(Col,Row: integer; DoPaint: boolean; CursorSizePos: TCursorSizePos = cspCell);
     procedure BeginSelection(Col,Row: integer);
     procedure AppendSelection(Col,Row: integer; DoPaint: boolean);
     function  Scroll(Col,Row: integer): boolean;
     procedure ClipToExtent(var Col,Row: integer);
     procedure ClipToSheet(var C1,R1,C2,R2: integer);
     procedure MarchingAntsTimer(Sender: TObject; TimerId,Data: integer);
     procedure CommentTimer(Sender: TObject; AData: Pointer);
     procedure CacheCells;
     procedure CacheMergedCells(const AMergedIndex: integer; const AFgColor,ABgColor: longword);
     procedure ApplyCondFmt;
     function  GetCursorAreaHit(X,Y: integer): TCursorAreaHit;
     procedure UpdateSheetArea;
     procedure PaintEditRefs(Sender: TObject; C1,R1,C2,R2: integer);
     procedure ScrollEditRefs(Sender: TObject; Direction: integer);
     procedure ColumnClick(Sender: TObject; Header: integer; SizeHit: TSizeHitType; Button: TXSSMouseButton; Shift: TXSSShiftState);
     procedure RowClick(Sender: TObject; Header: integer; SizeHit: TSizeHitType; Button: TXSSMouseButton; Shift: TXSSShiftState);
     procedure SelColChanged(Sender: TObject; Header,ScrollDirection: integer);
     procedure SelRowChanged(Sender: TObject; Header,ScrollDirection: integer);
     procedure SetSelectedHeaders;
     procedure DoSheetChanged;
     procedure InvalidateCell(Col,Row: integer);
     procedure HideComment(Col,Row: integer);
     function  FitNumAsString(Width: integer; Val: double): AxUCString;
     procedure IEPaintCells(Sender: TObject; const Col1,Row1,Col2,Row2: integer);
     procedure HandleScrollTimer(ASender: TObject; AData: Pointer);
     procedure IECharFormatChanged(ASender: TObject);
     procedure DoReloadSheet(ASender: TObject);
     procedure DoShowPivotForm(APivTbl: TXLSPivotTable);
     procedure ClosePivotForm(ASender: TObject; var Action: TCloseAction);
public
     constructor Create(AOwner: TCustomControl; AParent: TXSSWindow; Options: TXLSBookOptions; Layers: TXPaintLayers; Manager: TXc12Manager; AWorkbook: TXLSWorkbook; Sheet: TXLSWorkSheet);
     destructor Destroy; override;
     procedure SheetChanged(Sheet: TXLSWorkSheet);
     procedure Paint; override;
     procedure Invalidate; override;
     procedure InvalidateAndReload;

     procedure BeginPrint(const AVertAdjustment: double; const AStdFontWidth: integer);
     procedure PaintPrint(ASheet: TXLSWorkSheet; const ACol1,ARow1,ACol2,ARow2: integer);
     procedure EndPrint;

     procedure HideForms;
     procedure ShowPivotForm;

     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseWheel(const X,Y,Delta: integer); override;
     procedure MouseEnter(X, Y: Integer); override;
     procedure MouseLeave; override;
     procedure HandleKey(const Key: TXSSKey; const Shift: TXSSShiftKeys); override;
     procedure KeyPress(Key: AxUCChar); override;
     procedure DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
     procedure DragDrop(Source: TObject; X, Y: Integer);
     procedure OnVScroll(Sender: TObject; Value: TXScrollState);
     procedure OnHScroll(Sender: TObject; Value: TXScrollState);
     procedure InvalidateArea(C1,R1,C2,R2: integer);
     procedure HideCursor;
     procedure ShowCursor;
     procedure MoveCursor(const ACol,ARow: integer);

     procedure ShowInplaceEdit(AAssignCellValue: boolean; const AIsFormula: boolean = False);
     procedure HideInplaceEdit(const ASaveText: boolean = False; AKey: TXSSKey = kyDown);

     procedure ColGroupBtnsClick(Sender: TObject; Hdr, Level: integer);
     procedure RowGroupBtnsClick(Sender: TObject; Hdr, Level: integer);

     function  FocusedChild: TXSSChildWindows;

     property Columns: TXLSBookColumns read FColumns;
     property Rows: TXLSBookRows read FRows;
     property LeftCol: integer read GetLeftCol write SetLeftCol;
     property RightCol: integer read GetRightCol write SetRightCol;
     property TopRow: integer read GetTopRow write SetTopRow;
     property BottomRow: integer read GetBottomRow write SetBottomRow;
     property Tabs: TXTabSet read FTabs;
     property XLSSheet: TXLSWorkSheet read FXLSSheet;
     property Drawing: TXBookDrawing read FDrawing;
     property InplaceEditor: TXBookInplaceEditor read FInplaceEditor;

     property CursorArea: TXYRect read FCursorXY;

     property SelectOptions : TXSSSelectOptions read FSelectOptions write FSelectOptions;
     property DeleteOptions : TXLSDeleteOptions read FDeleteOptions write FDeleteOptions;

     property OnCellChanged: TXColRowEvent read FCellChangedEvent write FCellChangedEvent;
     property OnSelectionChanged: TNotifyEvent read FSelectionChangedEvent write FSelectionChangedEvent;
     property OnEditorOpen : TXBooleanEvent read FEditorOpenEvent write FEditorOpenEvent;
     property OnEditorClose: TXBooleanEvent read FEditorCloseEvent write FEditorCloseEvent;
     property OnEditorDisableFmt: TNotifyEvent read FEditorDisableFmtEvent write FEditorDisableFmtEvent;
     property OnHyperlinkClick: TXStringEvent read FHyperlinkClickEvent write FHyperlinkClickEvent;
     property OnIECharFmt     : TNotifyEvent read FIECharFmtEvent write FIECharFmtEvent;
     property OnNotification: TXNotifyEvent read FNotificationEvent write FNotificationEvent;
     end;

implementation

const TBorderLinePri: array[0..13] of integer = (0,6,10,4,5,13,11,1,9,2,8,3,7,12);

type TSetCellFormat = record
     ForeColor: longword;
     BackColor: longword;
     Pattern: TXc12FillPattern;
     LeftColor: longword;
     TopColor: longword;

     RightColor: longword;
     BottomColor: longword;
     LeftStyle: TXc12CellBorderStyle;
     TopStyle: TXc12CellBorderStyle;
     RightStyle: TXc12CellBorderStyle;
     BottomStyle: TXc12CellBorderStyle;
     DiagLines: TDiagLines;
     DiagColor: longword;
     DiagStyle: TXc12CellBorderStyle;
     end;

function StyleMax(S1,S2: TXc12CellBorderStyle): TXc12CellBorderStyle;
begin
  if S1 > S2 then
    Result := S1
  else
    Result := S2;
end;

{ TXLSBookSheet }

function TXLSBookSheet.CalcHScroll: boolean;
var
  D: integer;
begin
  D := FColumns.RightCol - FColumns.LeftCol;
  if D <= 0 then
    Result := False
  else begin
    FHScroll.LargeChange := D;
    FHScroll.MaxVal := Max(FColumns.HdrFromSizeRev(FXLSSheet.LastCol + 1,FCX2 - FCX1),Max(FColumns.LeftCol,1));
    FHScroll.TotalValues := Max(FXLSSheet.LastCol + 1,FColumns.RightCol);
    FHScroll.Position := FColumns.LeftCol;
    Result := True;
  end;
end;

function TXLSBookSheet.CalcVScroll: boolean;
var
  D: integer;
begin
  D := FRows.BottomRow - FRows.TopRow;
  if D <= 0 then
    Result := False
  else begin
    FVScroll.LargeChange := D;
    FVScroll.MaxVal := Max(FRows.HdrFromSizeRev(FXLSSheet.LastRow + 1,FCY2 - FCY1),Max(FRows.TopRow,1));
    FVScroll.TotalValues := Max(FXLSSheet.LastRow + 1,FRows.BottomRow);
    FVScroll.Position := FRows.TopRow;
    Result := True;
  end;
end;

procedure TXLSBookSheet.ColGroupBtnsClick(Sender: TObject; Hdr, Level: integer);
begin
  if Hdr >= 0 then
    FXLSSheet.Columns.ToggleGroupedAll(Level);
  FColumns.CalcHeaders;
  Invalidate;
end;

constructor TXLSBookSheet.Create(AOwner: TCustomControl; AParent: TXSSWindow; Options: TXLSBookOptions; Layers: TXPaintLayers; Manager: TXc12Manager; AWorkbook: TXLSWorkbook; Sheet: TXLSWorkSheet);
begin
  inherited Create(AParent);

  FOwner := AOwner;

  FOptions := Options;

  FSelectOptions := [xsoHyperlink,xsoShape,xsoPicture];

  FLayers := Layers;
  FManager := Manager;
  FWorkbook := AWorkbook;
  FXLSSheet := Sheet;
  FCells := FXLSSheet.MMUCells;

  FScrollTimerId := -1;
  FInMergedCells := -1;

  FHideHSplitter := True;
  FHideVSplitter := True;

  FDeleteOptions := [xdoCellValues,xdoComments,xdoDataValidations,xdoConditionalFormats,xdoAutofilters,xdoDrawing];

  FCellPaint := TCellPaintData.Create;

  Fid := Integer(xcwSheet);

  FTopLeft := TCornerRect.Create(Self);
  FTopLeft.Id := Integer(xcwTopLeft);
  FTopTopLeft := TCornerRect.Create(Self);
  FBottomRight := TCornerRect.Create(Self);

  FRows := TXLSBookRows.Create(Self,FXLSSheet.Rows,FXLSSheet.Autofilter);
  FRows.Id := Integer(xcwRows);
  FRows.DefaultHeight := FXLSSheet.DefaultRowHeight;
  FRows.OnChanged := RepaintRequest;
  FRows.OnPaintLine := PaintLine;
  FRows.OnHeaderClick := RowClick;
  FRows.OnSelHeaderChanged := SelRowChanged;

  FColumns := TXLSBookColumns.Create(Self,FXLSSheet.Columns);
  FColumns.Id := Integer(xcwColumns);
//  FColumns.DefaultWidthPix := FXLSSheet.DefaultColWidth;
  FColumns.DefaultHeight := FXLSSheet.DefaultRowHeight;
  FColumns.OnChanged := RepaintRequest;
  FColumns.OnPaintLine := PaintLine;
  FColumns.OnHeaderClick := ColumnClick;
  FColumns.OnSelHeaderChanged := SelColChanged;

{
  if Sheet.Pane.PaneType = ptSplit then begin
    FRows.FrozenRow := Sheet.Pane.SplitRowY;
    FColumns.FrozenCol := Sheet.Pane.SplitColX;
  end;
}

  FTabs := TXTabSet.Create(Self);
  FTabs.Id := Integer(xcwTabs);
  FTabs.OnSize := TabSizeChanged;
  FTabs.Visible := False;

  FVScroll := TXLSScrollBar.Create(Self,Nil);
  FVScroll.Id := Integer(xcwVertScroll);
  FVScroll.MinVal := 0;
  FVScroll.Visible := False;
  FVScroll.ExcelStyle := True;
  FVScroll.OnScroll := OnVScroll;

  FHScroll := TXLSScrollBar.Create(Self,Nil);
  FHScroll.Id := Integer(xcwHorizScroll);
  FHScroll.KindHoriz := True;
  FHScroll.MinVal := 0;
  FHScroll.Visible := False;
  FHScroll.ExcelStyle := True;
  FHScroll.OnScroll := OnHScroll;

  FHSplit := TSplitterHandle.Create(Self,shtHorizontal);
  FHSplit.OnPaintLine := PaintLine;
  FHSplit.Enabled := False;

  FVSplit := TSplitterHandle.Create(Self,shtVertical);
  FVSplit.OnPaintLine := PaintLine;
  FVSplit.Enabled := False;

  FCSplit := TSplitterHandle.Create(Self,shtCenter);
  FCSplit.OnPaintLine := PaintLine;

  FColGroupBtns := TGroupButtons.Create(Self,False);
  FColGroupBtns.Visible := False;
  FColGroupBtns.Buttons.OnButtonClick := ColGroupBtnsClick;

  FRowGroupBtns := TGroupButtons.Create(Self,True);
  FRowGroupBtns.Visible := False;
  FRowGroupBtns.Buttons.OnButtonClick := RowGroupBtnsClick;

  if FLayers <> Nil then
    FDrawing := TXBookDrawing.Create(Self,FManager,FLayers.DrawingLayer,FSkin.Style,FColumns,FRows)
  else
    FDrawing := TXBookDrawing.Create(Self,FManager,FSkin.GDI,FSkin.Style,FColumns,FRows);
  FDrawing.Id := Integer(xcwDrawing);

  FCopyAreas := TCellAreas.Create;

  Add(FTopLeft);
  Add(FTopTopLeft);
  Add(FRows);
  Add(FVSplit);
  Add(FBottomRight);
  Add(FHSplit);
  Add(FTabs);
  Add(FColumns);
  Add(FCSplit);
  Add(FColGroupBtns);
  Add(FRowGroupBtns);
  Add(FVScroll);
  Add(FHScroll);
  Add(FDrawing);

  FXLSSheet.XSSCellInvalidateEvent := InvalidateCell;

  ClearSelected(0,0,False);
//  SetVisibleRects;
end;

destructor TXLSBookSheet.Destroy;
begin
//  HideInplaceEdit;

  FreeMem(FCellStyleCache);
  if FSelectionEdit <> Nil then
    FSelectionEdit.Free;
  if FEditRefs <> Nil then
    FEditRefs.Free;
  FCellPaint.Free;
  FCopyAreas.Free;

  if FFormPivot <> Nil then
    FFormPivot.Free;

  inherited;
end;

procedure TXLSBookSheet.DragDrop(Source: TObject; X, Y: Integer);
var
  i: integer;
begin
  i := FTabs.SnapX(X);
  if i >= 0 then
    FTabs.Exchange(i,FTabs.SelectedTab);
end;

procedure TXLSBookSheet.DragOver(Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
  Accept := True;
  FLastDragTab := FTabs.SnapX(X);
  if (State = dsDragMove) and (X = FLastX) then
    Exit;
  case State of
    dsDragEnter:
      FLastX := X;
    dsDragLeave: FLastDragTab := -1;
    dsDragMove: begin
      FSkin.GDI.PaintTabMoveBmp(FLastX,FCY2);
      FLastX := X;
    end;
  end;
  FSkin.GDI.PaintTabMoveBmp(FLastX,FCY2);
end;

procedure TXLSBookSheet.EndPrint;
begin
  FIsPrinting := False;
end;

function TXLSBookSheet.FitNumAsString(Width: integer; Val: double): AxUCString;
var
  W: integer;
  S: AxUCString;
  Decimals: integer;
begin
  Result := FloatToStr(Val);
  if FSkin.GDI.TextWidth(Result) > Width then begin
    if (Abs(Val) < 1000) or (Frac(Val) > 0) then begin
      S := IntToStr(Round(Int(Val)));
      W := FSkin.GDI.TextWidth(S) + FSkin.GDI.CharWidth(WideChar(FormatSettings.DecimalSeparator));
      Decimals := Round(Int((Width - W) / FSkin.GDI.CharWidth('8')));
      Result := FloatToStr(RoundToDecimal(Val,Decimals));
    end
    else begin
      Result := FloatToStrF(Val,ffExponent,2,2);
    end;
  end;
end;

function TXLSBookSheet.FocusedChild: TXSSChildWindows;
begin
  Result := TXSSChildWindows(TXSSRootWindow(RootWindow).FocusedWin.Id);
end;

procedure TXLSBookSheet.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  Col,Row: integer;
  CacheIndex: integer;
begin
  inherited MouseDown(Button,Shift,X,Y);

  FMouseIsDown := True;

  if not (Button = xmbLeft) then
    Exit;

  Col := FColumns.PosToHeader(X);
  Row := FRows.PosToHeader(Y);

  if (xssDouble in Shift) and (FInplaceEditor = Nil) then
    ShowInplaceEdit(True)
  else if (FEditRefs <> Nil) and not FEditRefs.HasSelect and FInplaceEditor.LastCharIsOperator then
    FEditRefs.AddSelect(Col,Row)
  else if not ((FEditRefs <> Nil) and FEditRefs.IsFocused) then
    HideInplaceEdit(True);

//  if FChildRectHit then
//    Exit;

  FLastMouseDown.X := X;
  FLastMouseDown.Y := Y;

  if FEditRefs <> Nil then
    FEditRefs.MouseDown(Button,Shift,X,Y,Col,Row)
  else begin
    case GetCursorAreaHit(X,Y) of
      cahEdge: begin
        FSkin.GDI.SetCursor(xctArrow);
        FSelectionEdit := TXLSBookMoveSelArea.Create(FSkin,FColumns,FRows,Col,Row);
      end;
      cahSize: begin
        FSkin.GDI.SetCursor(xctSizeCells);
        FSelectionEdit := TXLSBookSizeSelArea.Create(FSkin,FColumns,FRows);
      end;
      else begin
        CacheIndex := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);
        if (FCellStyleCache[CacheIndex].Hyperlink <> Nil) and (xsoHyperlink in FSelectOptions) then begin
          FSkin.GDI.SetCursor(xctHandPoint);
          if Assigned(FHyperlinkClickEvent) then
            FHyperlinkClickEvent(Self,FCellStyleCache[CacheIndex].Hyperlink.Address);
        end
        else
          FSkin.GDI.SetCursor(xctCell);
        if xssCtrl in Shift then
          BeginSelection(Col,Row)
        else if xssShift in Shift then
          AppendSelection(Col,Row,True)
        else
          ClearSelected(Col,Row,True);
      end;
    end;
  end;
end;

procedure TXLSBookSheet.MouseEnter;
begin
  inherited;

end;

procedure TXLSBookSheet.MouseLeave;
begin
  inherited;
  if FVisibleComment <> Nil then
    HideComment(FVisibleComment.Col,FVisibleComment.Row);
end;

procedure TXLSBookSheet.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
var
  CacheIndex: integer;
  Col,Row: integer;

procedure SetSizeDirection;
var
  HX,HY: integer;
begin
  with FXLSSheet.SelectedAreas.First do begin
    HX := FColumns.HalfPos[Col - FSheetArea.Col1];
    HY := FRows.HalfPos[Row - FSheetArea.Row1];

    TXLSBookSizeSelArea(FSelectionEdit).IsDelete := PtInXYRect(X,Y,FCursorXYMargI);
    if TXLSBookSizeSelArea(FSelectionEdit).IsDelete then begin
      case FCursorSizePos of
        cspColumn: begin
          Row := Row1
        end;
        cspRow: begin
          Col := Col1;
        end;
        else begin
          if X > HX then
            Inc(Col);
          if Y > HY then
            Inc(Row);
          if Abs(X - FCursorXY.X2) > Abs(Y - FCursorXY.Y2) then
            Row := Row1
          else
            Col := Col1;
        end;
      end;
    end
    else begin
      if (Col > Col2) and (X < HX) then
        Dec(Col)
      else if (Col < Col1) and (X > HX) then
        Inc(Col);

      if (Row > Row2) and (Y < HY) then
        Dec(Row)
      else if (Row < Row1) and (Y > HY) then
        Inc(Row);

      if (Col > Col2) and (Row > Row2) then begin
        if Abs(X - FCursorXY.X2) > Abs(Y - FCursorXY.Y2) then
          Row := Row2
        else
          Col := Col2;
      end
      else if (Col > Col2) and (Row < Row1) then begin
        if Abs(X - FCursorXY.X2) > Abs(Y - FCursorXY.Y1) then
          Row := Row2
        else
          Col := Col1;
      end
      else if (Col < Col1) and (Row < Row1) then begin
        if Abs(X - FCursorXY.X1) > Abs(Y - FCursorXY.Y1) then
          Row := Row1
        else
          Col := Col1;
      end
      else if (Col < Col1) and (Row > Row2) then begin
        if Abs(X - FCursorXY.X1) > Abs(Y - FCursorXY.Y2) then
          Row := Row2
        else
          Col := Col1;
      end;
    end;
  end;
end;

procedure MouseMoveSelf;
begin
  Col := FColumns.PosToHeader(X);
  Row := FRows.PosToHeader(Y);
  CacheIndex := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);
  HideComment(Col + FColumns.LeftCol,Row + FRows.TopRow);
  if FSelectionEdit <> Nil then begin
    BeginPaint;
    case FSelectionEdit.Mode of
      semMove: begin
        if FSelectionEdit.Started then begin
          FSkin.GDI.SetCursor(xctArrow);
          FSelectionEdit.UpdateEdit(Col,Row,True);
        end
        else begin
          FSkin.GDI.SetCursor(xctArrow);
          if ((Abs(X - FLastMouseDown.X) > MARG_CELLCURSORSIZEHIT) or (Abs(Y - FLastMouseDown.Y) > MARG_CELLCURSORSIZEHIT)) then
            FSelectionEdit.BeginEdit(FXLSSheet.SelectedAreas.First,Col,Row);
        end;
      end;
      semSize: begin
        if FSelectionEdit.Started then begin
          FSkin.GDI.SetCursor(xctSizeCells);
          ClipToSheet(Col,Row,Col,Row);
          SetSizeDirection;
          if FCellStyleCache[CacheIndex].InMerged >= 0 then begin
            Row := FXLSSheet.MergedCells[FCellStyleCache[CacheIndex].InMerged].Row1;
            Col := FXLSSheet.MergedCells[FCellStyleCache[CacheIndex].InMerged].Col1;
          end;
          FSelectionEdit.UpdateEdit(Col,Row,FCursorSizePos = cspCell);
        end
        else begin
          FSkin.GDI.SetCursor(xctSizeCells);
          if (Abs(X - FLastMouseDown.X) > Abs(FLastMouseDown.X - FColumns.HalfPos[Col - FSheetArea.Col1])) or
             (Abs(Y - FLastMouseDown.Y) > Abs(FLastMouseDown.Y - FRows.HalfPos[Row - FSheetArea.Row1]))
             then begin
            SetSizeDirection;
            FSelectionEdit.BeginEdit(FXLSSheet.SelectedAreas.First,Col,Row);
          end;
        end;
      end;
    end;
    EndPaint;
  end
  else begin
    if True {not FMouseDown} {FLastClick = Nil} then begin
      case GetCursorAreaHit(X,Y) of
        cahEdge: FSkin.GDI.SetCursor(xctMoveCells);
        cahSize: FSkin.GDI.SetCursor(xctSizeCells);
        else
          FSkin.GDI.SetCursor(xctCell);
      end;
    end;
  end;
  if FEditRefs <> Nil then begin
    FEditRefs.MouseMove(Shift,X,Y,Col,Row);
  end
  else if not FOptions.ReadOnly and (xssLeft in Shift) then begin
    if FSelectionEdit = Nil then
      AppendSelection(Col,Row,True);

    if FScrollTimerId < 0 then begin
      if X > FCX2 then
        FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,100,Pointer(kyRight))
      else if (X < FCX1) and (FSheetArea.Col1 > 0) then
        FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,100,Pointer(kyLeft))
      else if Y > FCY2 then
        FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyDown))
      else if (Y < FCY1) and (FSheetArea.Row1 > 0) then
        FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyUp));
    end
    else if (FScrollTimerId >= 0) and Hit(X,Y) then begin
      FSystem.StopTimer(FScrollTimerId);
      FScrollTimerId := -1;
//        PaintMoveRect;
    end;
  end
  else
    inherited MouseMove(Shift,X,Y);;
end;

begin
  if FMouseIsDown then
    MouseMoveSelf;
  if (FEditRefs = Nil) and (FSelectionEdit = Nil) then begin
//    inherited MouseMove(Shift,X,Y);

//    MouseMoveSelf; // is commented out in order to solve problem that cell area was selected when
    // opening file by doubleclick on name in OpenDialog.
    // Replaced with SetCursor.
    case GetCursorAreaHit(X,Y) of
      cahEdge: FSkin.GDI.SetCursor(xctMoveCells);
      cahSize: FSkin.GDI.SetCursor(xctSizeCells);
      else begin
        Col := FColumns.PosToHeader(X);
        Row := FRows.PosToHeader(Y);
        CacheIndex := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);
        HideComment(Col,Row);
        if FVisibleComment = Nil then begin
          FSystem.StopTimer(FCommentTimerId);
          if (FCellStyleCache <> Nil) and (csfComment in FCellStyleCache[CacheIndex].Flags) then begin
            FVisibleComment := FXLSSheet.Comments.Find(Col,Row);
            if FVisibleComment.AutoVisible then
              FVisibleComment := Nil
            else
              FCommentTimerId := FSystem.StartTimer(Self,CommentTimer,300,Nil);
          end;
        end;

        if (FCellStyleCache <> Nil) and (FCellStyleCache[CacheIndex].Hyperlink <> Nil) and (xsoHyperlink in FSelectOptions) then
          FSkin.GDI.SetCursor(xctHandPoint)
        else
          FSkin.GDI.SetCursor(xctCell);
      end;
    end;
  end;
end;

procedure TXLSBookSheet.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  Col,Row: integer;
  dCol,dRow: integer;
  Area,A: TXLSCellArea;
begin
  FMouseIsDown := False;

  if FScrollTimerId >= 0 then begin
    FSystem.StopTimer(FScrollTimerId);
    FScrollTimerId := -1;
  end;
  inherited MouseUp(Button,Shift,X,Y);
  Col := FColumns.PosToHeader(X);
  Row := FRows.PosToHeader(Y);
  if FSelectionEdit <> Nil then begin
    if not FSelectionEdit.Started then
      FreeAndNil(FSelectionEdit)
    else begin
      BeginPaint;
      FSelectionEdit.EndEdit(Col,Row);
      case FSelectionEdit.Mode of
        semMove: begin
          with FXLSSheet.SelectedAreas.First do begin
            dCol := TXLSBookMoveSelArea(FSelectionEdit).DeltaCol;
            dRow := TXLSBookMoveSelArea(FSelectionEdit).DeltaRow;
            FreeAndNil(FSelectionEdit);
            PaintCursor(Col1,Row1,Col2,Row2);
            PaintCells(Col1,Row1,Col2,Row2,ctpInvSelected);
            A := FXLSSheet.SelectedAreas[0].AsRect;
            FXLSSheet.MoveCells(A.Col1,A.Row1,A.Col2,A.Row2,A.Col1 + dCol,A.Row1 + dRow,CopyAllCells);
            CacheCells;
            FXLSSheet.SelectedAreas.ActiveCol := FXLSSheet.SelectedAreas.ActiveCol + dCol;
            FXLSSheet.SelectedAreas.ActiveRow := FXLSSheet.SelectedAreas.ActiveRow + dRow;
            SetArea(Col1 + dCol,Row1 + dRow,Col2 + dCol,Row2 + dRow);
            PaintCells(A.Col1,A.Row1,A.Col2,A.Row2,ctpAll);
            PaintCells(Col1,Row1,Col2,Row2,ctpAll);
            PaintCursor(Col1,Row1,Col2,Row2);
          end;
        end;
        semSize: begin
          if TXLSBookSizeSelArea(FSelectionEdit).IsDelete then begin
            A := TXLSBookSizeSelArea(FSelectionEdit).DeleteArea;
            Area.Col1 := FXLSSheet.SelectedAreas.First.Col1;
            Area.Row1 := FXLSSheet.SelectedAreas.First.Row1;
            Area.Col2 := A.Col2;
            Area.Row2 := A.Row2;
            if (A.Col1 = Area.Col1) and (A.Row1 = Area.Row1) then begin
              Area.Col2 := Area.Col1;
              Area.Row2 := Area.Row1;
            end
            else if A.Col1 > Area.Col1 then
              Area.Col2 := A.Col1 - 1
            else if A.Row1 > Area.Row1 then
              Area.Row2 := A.Row1 - 1;
            if not CellInArea(FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow,Area) then begin
              FXLSSheet.SelectedAreas.ActiveCol := FXLSSheet.SelectedAreas.First.Col1;
              FXLSSheet.SelectedAreas.ActiveRow := FXLSSheet.SelectedAreas.First.Row1;
            end;
          end
          else begin
            Area := TXLSBookSizeSelArea(FSelectionEdit).SizeArea;
          end;
          FreeAndNil(FSelectionEdit);
          with FXLSSheet.SelectedAreas.First do begin
            PaintCursor(Col1,Row1,Col2,Row2);
            PaintCells(Col1,Row1,Col2,Row2,ctpInvSelected);
            SetArea(Area.Col1,Area.Row1,Area.Col2,Area.Row2);
            PaintCells(Col1,Row1,Col2,Row2,ctpAll);
            PaintCursor(Col1,Row1,Col2,Row2);
          end;
        end;
      end;
      EndPaint;
    end;
  end
  else if FEditRefs <> Nil then
    FEditRefs.MouseUp(Button,Shift,X,Y,Col,Row);

  if FInplaceEditor <> Nil then
    FInplaceEditor.Focus;
end;

procedure TXLSBookSheet.MouseWheel(const X,Y,Delta: integer);
begin
  if Delta <> 0 then begin
    if Delta < 0 then
      SetTopRow(FRows.TopRow + 4)
    else if Delta > 0 then
      SetTopRow(FRows.TopRow - 4);

    if CalcVScroll then begin
      Repaint;
      FRows.Paint;
    end;
  end;
end;

procedure TXLSBookSheet.MoveCursor(const ACol, ARow: integer);
var
  C,R: integer;
  Scrolled: boolean;
  Merged: TXc12MergedCell;
begin
  C := ACol;
  R := ARow;

  if FInMergedCells >= 0 then begin
    Merged := FXLSSheet.MergedCells[FInMergedCells];
    if Merged.Row2 >= FRows.BottomRow then
      R := Merged.Row2;
    if Integer(Merged.Col2) > FColumns.RightCol then
      C := Merged.Col2;
  end;

  Scrolled := False;
  if C >= FColumns.RightCol then begin
    FColumns.RightCol := C;
    Scrolled := True;
  end
  else if C < FColumns.LeftCol then begin
    FColumns.LeftCol := C;
    Scrolled := True;
  end;

  if R >= FRows.BottomRow then begin
    SetBottomRow(R);
    Scrolled := True;
  end
  else if R < FRows.TopRow then begin
    SetTopRow(R);
    Scrolled := True;
  end;

  if Scrolled then begin
    FCacheRequest := True;
    UpdateSheetArea;
    CalcVScroll;
    CalcHScroll;
    Repaint;
  end;
end;

procedure TXLSBookSheet.CommentTimer(Sender: TObject; AData: Pointer);
begin
  FSystem.StopTimer(FCommentTimerId);
  FDrawing.AddComment(FVisibleComment);
end;

procedure TXLSBookSheet.OnHScroll(Sender: TObject; Value: TXScrollState);
var
  D: integer;
begin
  D := Max(FColumns.RightCol - FColumns.LeftCol,0);
  case Value of
    xssUpClicked       : SetLeftCol(FColumns.NextVisible(Max(FColumns.LeftCol - 1,0)));
    xssDownClicked     : SetLeftCol(FColumns.NextVisible(Min(FColumns.LeftCol + 1,XLS_MAXCOL)));
    xssLargeUpClicked  : SetLeftCol(FColumns.PrevVisible(Max(FColumns.LeftCol - D,0)));
    xssLargeDownClicked: SetLeftCol(FColumns.PrevVisible(Min(FColumns.LeftCol + D,XLS_MAXCOL)));
    xssTrack           : SetLeftCol(FHScroll.Position);
    xssReleaseTrack    : SetLeftCol(FColumns.LeftCol);
  end;

  if CalcHScroll then begin
    Repaint;
    FColumns.Paint;
  end;
end;

procedure TXLSBookSheet.OnVScroll(Sender: TObject; Value: TXScrollState);
var
  D: integer;
begin
  D := Max(FRows.BottomRow - FRows.TopRow,0);
  case Value of
    xssUpClicked       : SetTopRow(FRows.NextVisible(Max(FRows.TopRow - 1,0)));
    xssDownClicked     : SetTopRow(FRows.NextVisible(Min(FRows.TopRow + 1,XLS_MAXROW)));
    xssLargeUpClicked  : SetTopRow(FRows.PrevVisible(Max(FRows.TopRow - D,0)));
    xssLargeDownClicked: SetTopRow(FRows.PrevVisible(Min(FRows.TopRow + D,XLS_MAXROW)));
    xssTrack           : SetTopRow(FVScroll.Position);
    xssReleaseTrack    : SetTopRow(FRows.TopRow);
  end;

  if CalcVScroll then begin
    Repaint;
    FRows.Paint;
  end;
end;

procedure TXLSBookSheet.Paint;
var
  i: integer;
  SA: TXLSSelectedArea;
begin
  BeginPaint;

  if FSheetChanged then
    DoSheetChanged;

  FSkin.GDI.BrushColor := $FFFFFF;
  FSkin.GDI.Rectangle(FCX1 - 1,FCY1 - 1,FCX2,FCY2);

  PaintCells(FColumns.LeftCol,FRows.TopRow,FColumns.RightCol,FRows.BottomRow,ctpAll);

  FColumns.ClearSelected;
  FRows.ClearSelected;

  for i := 0 to FXLSSheet.SelectedAreas.Count - 1 do begin
    SA := FXLSSheet.SelectedAreas[i];
    FColumns.SetSelectedState(SA.Col1 - FColumns.LeftCol,SA.Col2 - FColumns.LeftCol,(SA.Row1 = 0) and (SA.Row2 >= MAXSELROW));
    FRows.SetSelectedState(SA.Row1 - FRows.TopRow,SA.Row2 - FRows.TopRow,(SA.Col1 = 0) and (SA.Col2 >= XLS_MAXCOL));
    PaintCursor(SA.Col1,SA.Row1,SA.Col2,SA.Row2);
  end;

  if FCopyAreas.Count = 1 then
    PaintMarchingAnts(FCopyAreas[0].Col1,FCopyAreas[0].Row1,FCopyAreas[0].Col2,FCopyAreas[0].Row2);

{$ifdef SHAREWARE}
  FSkin.GDI.SetTextAlign(xhtaLeft,xvtaTop);
  FSkin.GDI.TextOut(FCX1,FCY1,'Copyright (c) 2017 Axolot Data');
{$endif}

  EndPaint;

//  FDrawing.Paint;
end;

procedure TXLSBookSheet.PaintCells(Col1,Row1,Col2,Row2: integer; CellsToPaint: TCellsToPaint);
var
  i,j: integer;
  xi,yi: integer;
  w,h,x,y: double;
  x1,y1,x2,y2: integer;
  cx1,cy1,cx2,cy2: integer;
  C,R,Col,Row: integer;
  Rect: TXYRect;
  CacheIndex: integer;
  Marg: integer;
  SelAreaHit: TSelectedAreaHit;
  SelEdgeHit: TSelectedEdgeHits;
  MergedCell: TXLSMergedCell;
begin
  // Moved here from EDIT060829
  if FCacheRequest then begin
    FCacheRequest := False;
    UpdateSheetArea;
    CacheCells;
  end;

  if FCopyAreas.Count = 1 then
    PaintMarchingAnts(FCopyAreas[0].Col1,FCopyAreas[0].Row1,FCopyAreas[0].Col2,FCopyAreas[0].Row2);

  ClipToSheet(Col1,Row1,Col2,Row2);

  y := FRows.AbsPos[Row1] + FRows.Offset;
  for Row := Row1 to Row2 do begin
    h := FRows.RowHeightFloat[Row - FRows.TopRow];
    if h <= 0 then
      Continue;
    x := FColumns.AbsPos[Col1] + FColumns.Offset;

    for Col := Col1 to Col2 do begin
      w := FColumns.ColWidthFloat[Col - FColumns.LeftCol];
      if w <= 0 then
        Continue;
      {HitArea :=} FXLSSheet.SelectedAreas.CellInAreas(Col,Row,SelEdgeHit,SelAreaHit);
      x1 := Round(x);
      y1 := Round(y);
      x2 := Round(x + w - 2);
      y2 := Round(y + h - 2);
{
      if x1 < FCX1 then x1 := FCX1;
      if y1 < FCY1 then y1 := FCY1;
      if x2 > FCX2 then x2 := FCX2;
      if y2 > FCY2 then y2 := FCY2;
}
      cx1 := x1;
      cy1 := y1;
      cx2 := x2;
      cy2 := y2;
      CacheIndex := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);

      if (FCellStyleCache[CacheIndex].InMerged >= 0) and
         FXLSSheet.MergedCells[FCellStyleCache[CacheIndex].InMerged].CellInArea(FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow) then
        SelAreaHit := sahActiveCell;
      if SelAreaHit in [sahInside,sahEdge] then begin
        if (CellsToPaint in [ctpAll,ctpSelected]) then begin
          PaintTheCell(Col,Row,x1,y1,x2,y2,True);

          Marg := SEL_CELL_MARGIN;
          if FXLSSheet.SelectedAreas.CursorVisible then
            Inc(Marg);

          Inc(cx2);
          Inc(cy2);

          if sehLeft   in SelEdgeHit then Inc(cx1,Marg);
          if sehTop    in SelEdgeHit then Inc(cy1,Marg);
          if sehRight  in SelEdgeHit then Dec(cx2,Marg + 1);
          if sehBottom in SelEdgeHit then Dec(cy2,Marg + 1);

          FSkin.GDI.PaintColor := FSkin.Colors.SheetSel;
          FSkin.GDI.PaintSelectBMP(cx1,cy1,cx2,cy2);
        end
        else if CellsToPaint = ctpInvSelected then begin
          FSkin.GDI.PaintColor := FSkin.Colors.SheetBkg;
          PaintTheCell(Col,Row,x1,y1,x2,y2);
        end
        else begin
          FSkin.GDI.PaintColor := FSkin.Colors.SheetSel;
          PaintTheCell(Col,Row,cx1,cy1,cx2,cy2);
        end;
      end
      else begin

        if CellsToPaint = ctpAll then begin
          FSkin.GDI.PaintColor := FSkin.Colors.SheetBkg;
          PaintTheCell(Col,Row,x1,y1,x2,y2);
        end;

        if not FXLSSheet.SelectedAreas.CursorVisible and (SelAreaHit = sahActiveCell) then begin
          if sehLeft   in SelEdgeHit then Inc(cx1);
          if sehTop    in SelEdgeHit then Inc(cy1);
          if sehRight  in SelEdgeHit then Dec(cx2);
          if sehBottom in SelEdgeHit then Dec(cy2);
          FSkin.PaintFocusCell(cx1,cy1,cx2,cy2);
        end;
      end;
      x :=  x + w;
    end;
    y := y + h;
  end;

//  if FSelectionEdit <> Nil then
//    FSelectionEdit.PaintRect
  if FEditRefs <> Nil then
    FEditRefs.Paint;

//  Exit;

  FXLSSheet.MergedCells.ClearPainted;

  y := FRows.AbsPos[Row1] + FRows.Offset;
  for Row := Row1 to Row2 do begin
    h := FRows.RowHeightFloat[Row - FRows.TopRow];
    if h <= 0 then
      Continue;
    x := FColumns.AbsPos[Col1] + FColumns.Offset;
    for Col := Col1 to Col2 do begin
      w := FColumns.ColWidthFloat[Col - FColumns.LeftCol];
      if w <= 0 then
        Continue;
//      HitArea := FXLSSheet.SelectedAreas.CellInAreas(Col,Row,SelEdgeHit,SelAreaHit);
      x1 := Round(x);
      y1 := Round(y);
      x2 := Round(x + w - 2);
      y2 := Round(y + h - 2);

      CacheIndex := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);

      if FCellStyleCache[CacheIndex].InMerged >= 0 then
        MergedCell := FXLSSheet.MergedCells[FCellStyleCache[CacheIndex].InMerged]
      else
        MergedCell := Nil;
      if (FCellStyleCache[CacheIndex].Ref.Data <> Nil) and (FCells.CellType(@FCellStyleCache[CacheIndex].Ref) <> xctBlank) and not (FCellStyleCache[CacheIndex].CondFmtRes in [xcfrtDataBarNoText,xcfrtIconNoText]) then begin
        FCellPaint.Clear(x1,y1,x2,y2,FCellStyleCache[CacheIndex]);
        with FCellStyleCache[CacheIndex] do
          FCellPaint.CalcPaintRect(StyleNW[ccS],StyleNW[ccE],StyleSE[ccN],StyleSE[ccW]);

        Rect := FCellPaint.FPaintRect;
        if FCellStyleCache[CacheIndex].CondFmtRes = xcfrtIcon then
          Rect.X1 := Rect.X1 + CONDFMT_ICON_SIZE + 1;
//        if (MergedCell <> Nil) and ((Col) = Integer(MergedCell.Col1)) and (Integer(MergedCell.Row1) = Row) then begin
//          C := Min(FColumns.RightCol,MergedCell.Col2);
//          R := Min(FRows.BottomRow,MergedCell.Row2);
//          Rect.X2 := FColumns.AbsPos[C] + FColumns.AbsWidth[C] + FColumns.Offset;
//          Rect.Y2 := FRows.AbsPos[R] + FRows.AbsHeight[R] + FRows.Offset;
//
//          PaintCellText(Col,Row,FCellStyleCache[CacheIndex].Ref,Rect,FCellStyleCache[CacheIndex].DXF);
//        end
        if (MergedCell <> Nil) and not MergedCell.Painted then begin

//          Rect.X2 := Rect.X1;
//          for j := MergedCell.Col1 to MergedCell.Col2 do
//            Rect.X2 := Rect.X2 + FXLSSheet.Columns[j].PixelWidth;

//          Rect.Y2 := Rect.Y1;
//          for j := MergedCell.Row1 to MergedCell.Row2 do
//            Rect.Y2 := Rect.Y2 + FXLSSheet.Rows[j].PixelHeight;

          if MergedCell.Col2 < FColumns.Count then
            C := MergedCell.Col2
          else
            C := FColumns.Count - 1;

          if MergedCell.Row2 < FRows.Count then
            R := MergedCell.Row2
          else
            R := FRows.Count - 1;

          Rect.X2 := FColumns.AbsPos[C] + FColumns.AbsWidth[C] + FColumns.Offset;
          Rect.Y2 := FRows.AbsPos[R] + FRows.AbsHeight[R] + FRows.Offset;

          PaintCellText(Col,Row,FCellStyleCache[CacheIndex].Ref,Rect,FCellStyleCache[CacheIndex].DXF);

          MergedCell.Painted := True;
        end
        else begin
          C := Col  - FColumns.LeftCol + 1;
          i := (Row - FRows.TopRow) * FColumns.Count + C;
          while (C < FColumns.Count) and (csfTextExpandRight in FCellStyleCache[i].Flags) do begin
            Inc(Rect.X2,FColumns.ColWidth[C]);
            Inc(C);
            Inc(i);
          end;
          C := Col - FColumns.LeftCol - 1;
          i := (Row - FRows.TopRow) * FColumns.Count + C;
          while (C >= 0) and (csfTextExpandLeft in FCellStyleCache[i].Flags) do begin
            Dec(Rect.X1,FColumns.ColWidth[C]);
            Dec(C);
            dec(i);
          end;
          PaintCellText(Col,Row,FCellStyleCache[CacheIndex].Ref,Rect,FCellStyleCache[CacheIndex].DXF);
        end;
      end;
      x := x + w;
    end;
    y := y + h;
  end;

  if (FRows.FrozenRow > 0) and (FRows.FrozenRow <= FRows.BottomRow) then begin
    FSkin.GDI.PenColor := FSkin.Colors.HeaderFrozenLine;
    yi := FRows.Pos[FRows.FrozenRow] + FRows.RowHeight[FRows.FrozenRow];
    FSkin.GDI.Line(FCX1,yi - 1,FCX2,yi - 1);
  end;
  if (FColumns.FrozenCol > 0) and (FColumns.FrozenCol <= FColumns.RightCol) then begin
    FSkin.GDI.PenColor := FSkin.Colors.HeaderFrozenLine;
    xi := FColumns.Pos[FColumns.FrozenCol] + FColumns.ColWidth[FColumns.FrozenCol];
    FSkin.GDI.Line(xi - 1,FCY1,xi - 1,FCY2);
  end;
end;

procedure TXLSBookSheet.PaintTheCell(Col,Row: integer; pX1, pY1, pX2, pY2: integer; Selected: boolean = False);
var
  i: integer;
  E: TCellEdges;
  MergedCell: TCellArea;
  GridColor: longword;
  BorderLeft,BorderTop,BorderRight,BorderBottom: TXc12CellBorderStyle;
  ColorRight,ColorBottom: longword;
  Ref: TXLSCellItem;
  XF: TXc12XF;
begin
  if FIsPrinting then begin
    if psoGridlines in FXLSSheet.PrintSettings.Options then
      GridColor := FSkin.Colors.SheetGridline
    else
      GridColor := FSkin.Colors.SheetBkg;
  end
  else begin
    if soGridlines in FXLSSheet.Options then
      GridColor := FSkin.Colors.SheetGridline
    else
      GridColor := FSkin.Colors.SheetBkg;
  end;

  i := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);

  FCellPaint.Clear(pX1,pY1,pX2,pY2,FCellStyleCache[i]);

  BorderLeft := FCellStyleCache[i].StyleNW[ccS];
  BorderTop := FCellStyleCache[i].StyleNW[ccE];
  BorderRight := FCellStyleCache[i].StyleSE[ccN];
  BorderBottom := FCellStyleCache[i].StyleSE[ccW];
//  BorderDiagUp
//  BorderDiagDown

  if (csfTextExpandLeft in FCellStyleCache[i].Flags) or ((Col <= FColumns.VisualSize) and (csfTextExpandRight in FCellStyleCache[i + 1].Flags)) then
    ColorRight := _Temp_XCColorToRGB(FCellStyleCache[i].ForeColor,FSkin.Colors.SheetBkg)
  else if BorderRight = cbsNone then begin
    if FCellStyleCache[i].ForeColor <> EX12_AUTO_COLOR then
      ColorRight := _Temp_XCColorToRGB(FCellStyleCache[i].ForeColor,0)
    else
      ColorRight := GridColor;
  end
  else
    ColorRight := _Temp_XCColorToRGB(FCellStyleCache[i].RightColor,FSkin.Colors.SheetGridline{COLOR_SHEET_BKG});

  if BorderBottom = cbsNone then begin
    if FCellStyleCache[i].ForeColor <> EX12_AUTO_COLOR then
      ColorBottom := _Temp_XCColorToRGB(FCellStyleCache[i].ForeColor,0)
    else
      ColorBottom := GridColor;
  end
  else
    ColorBottom := _Temp_XCColorToRGB(FCellStyleCache[i].BottomColor,FSkin.Colors.SheetGridline{COLOR_SHEET_BKG});

  FCellPaint.Calc(FCellStyleCache[i],BorderLeft,BorderTop,BorderRight,BorderBottom);

  if FCellStyleCache[i].InMerged >= 0 then
    MergedCell := FXLSSheet.MergedCells[FCellStyleCache[i].InMerged]
  else
    MergedCell := Nil;

  FCellPaint.PaintCell(FSkin);
  if FCellStyleCache[i].CondFmtRes in [xcfrtDataBar,xcfrtDataBarNoText] then
    FCellPaint.PaintDataBar(FSkin,FCellStyleCache[i].CondFmtVal / COND_FMT_VAL_SCALE,FCellStyleCache[i].CondFmtColor)
  else if FCellStyleCache[i].CondFmtRes in [xcfrtIcon,xcfrtIconNoText] then
    FCellPaint.PaintIcon(FSkin,TXc12IconSetType(FCellStyleCache[i].CondFmtVal),FCellStyleCache[i].CondFmtColor,FCellStyleCache[i].CondFmtRes = xcfrtIconNoText);

  if MergedCell <> Nil then begin
    E := MergedCell.Edges[Col,Row];
    if not (ceBottom in E) then
      ColorBottom := _Temp_XCColorToRGB(FCellStyleCache[i].ForeColor,FSkin.Colors.SheetBkg);
    if not (ceRight in E) then
      ColorRight := _Temp_XCColorToRGB(FCellStyleCache[i].ForeColor,FSkin.Colors.SheetBkg);

    FCellPaint.PaintBorders(FSkin,ColorRight,ColorBottom);
  end
  else begin
    FCellPaint.PaintBorders(FSkin,ColorRight,ColorBottom);

    FCellPaint.PaintDiagonal(FSkin);
  end;

  if FCellPaint.IsPrinting and ((Col = 0) or (Row = 0)) then begin
    Ref := FCells.FindCell(FColumns.VisualCol[Col],Row + FRows.TopRow);
    if Ref.Data <> Nil then begin
      XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Ref)];
      if (Col = 0) and (XF.Border.Left.Style <> cbsNone) then
        FCellPaint.PaintBorderLeft(FSkin,XF.Border.Left.Color.ARGB,XF.Border.Left.Style);
      if (Row = 0) and (XF.Border.Top.Style <> cbsNone) then
        FCellPaint.PaintBorderTop(FSkin,XF.Border.Top.Color.ARGB,XF.Border.Top.Style);
    end;
  end;

  if (FCellStyleCache[i].CondFmtRes = xcfrtDXF) and (FCellStyleCache[i].DXF.Border <> Nil) then
    FCellPaint.PaintBordersTopLeft(FSkin,FCellStyleCache[i].DXF.Border.Top,FCellStyleCache[i].DXF.Border.Left);

  if csfComment in FCellStyleCache[i].Flags then
    FSkin.GDI.PaintStockBMP(sxbCommentMark,FCellPaint.FPaintRect.X2,FCellPaint.FPaintRect.Y1);
end;

procedure TXLSBookSheet.PaintLine(Sender: TObject; LineType: TPaintLineType; var X,Y: integer; Execute: boolean);
begin
  BeginPaint;
  case LineType of
    pltHSize : FSkin.GDI.SizeLine(X,FY1,X + 2,FY2);
    pltVSize : FSkin.GDI.SizeLine(FX1,Y,FX2,Y + 2);
    pltHSplit: ;
    pltVSplit: ;
  end;
  EndPaint;
end;

procedure TXLSBookSheet.PaintMarchingAnts(C1, R1, C2, R2: integer);
var
  RectSel,RectDest: TXLSCellArea;
  RectXY: TXYRect;
begin
  if FIsPrinting then
    Exit;
  RectSel.Col1 := C1;
  RectSel.Col2 := C2;
  RectSel.Row1 := R1;
  RectSel.Row2 := R2;
  if IntersectCellArea(FSheetArea,RectSel,RectDest) then begin
    FColumns.ColPos(RectDest.Col1 - FColumns.LeftCol,RectDest.Col2 - FColumns.LeftCol,RectXY.X1,RectXY.X2);
    FRows.RowPos(RectDest.Row1 - FRows.TopRow,RectDest.Row2 - FRows.TopRow,RectXY.Y1,RectXY.Y2);
    if FCursorSizePos = cspCell then begin
      if RectDest.Col1 > RectSel.Col1 then Dec(RectXY.X1,10);
      if RectDest.Col2 < RectSel.Col2 then Inc(RectXY.X2,10);
      if RectDest.Row1 > RectSel.Row1 then Dec(RectXY.Y1,10);
      if RectDest.Row2 < RectSel.Row2 then Inc(RectXY.Y2,10);
    end;

    FSkin.GDI.LineStyled(RectXY.X1,RectXY.Y1,RectXY.X2,RectXY.Y1 + 1,$00000000,$00FFFFFF,pls_MediumDotted_,True);
    FSkin.GDI.LineStyled(RectXY.X2 - 3,RectXY.Y1,RectXY.X2 - 2,RectXY.Y2,$00000000,$00FFFFFF,pls_MediumDotted_,False);
    FSkin.GDI.LineStyled(RectXY.X1,RectXY.Y2 - 3,RectXY.X2,RectXY.Y2 - 2,$00000000,$00FFFFFF,pls_MediumDotted_,True);
    FSkin.GDI.LineStyled(RectXY.X1,RectXY.Y2 - 1,RectXY.X1 + 1,RectXY.Y1,$00000000,$00FFFFFF,pls_MediumDotted_,False);
  end;
end;

procedure TXLSBookSheet.PaintPrint(ASheet: TXLSWorkSheet; const ACol1,ARow1,ACol2,ARow2: integer);
var
  x,y: integer;
begin
  FCellPaint.IsPrinting := True;
  FCellPaint.SetCacheSize(ACol2 - ACol1 + 1,ARow2 - ARow1 + 1);

  FColumns.Visible := False;
  FRows.Visible := False;

  SetLeftCol(ACol1);
  SetTopRow(ARow1);

  SetSize(FX1,FY1,FX2,FY2);

  SheetChanged(ASheet);

  PaintCells(ACol1,ARow1,ACol2,ARow2,ctpAll);

  FCellPaint.PaintLineCache(FSkin);

  if psoGridlines in FXLSSheet.PrintSettings.Options then begin
    FSkin.GDI.PenColor := XSS_SYS_COLOR_BLACK;
    x := FColumns.AbsPos[ACol1] + FColumns.Offset;
    y := FRows.AbsPos[ARow1] + FRows.Offset;
    FSkin.GDI.Line(x,y,FColumns.PageWidth,y);
    FSkin.GDI.Line(x,y,x,FRows.PageHeight);
  end;

  FDrawing.PaintPrint;

{$ifdef SHAREWARE}
  FSkin.GDI.SetTextAlign(xhtaLeft,xvtaTop);
  FSkin.GDI.TextOut(FCX1,FCY1,'Copyright (c) 2017 Axolot Data');
{$endif}
end;

procedure TXLSBookSheet.PaintCursor(C1,R1,C2,R2: integer);
var
  RectSel,RectDest: TXLSCellArea;
begin
  if FIsPrinting then
    Exit;
  RectSel.Col1 := C1;
  RectSel.Col2 := C2;
  RectSel.Row1 := R1;
  RectSel.Row2 := R2;
  if FXLSSheet.SelectedAreas.CursorVisible and IntersectCellArea(FSheetArea,RectSel,RectDest) then begin
    // TODO Paint left and top edges of cursor when it's outside sheet
//            x1 := FCX1 - 10;
//            x2 := FCX1;
//            FRows.RowPos(RectSel.Y1 - FRows.TopRow,RectSel.Y2 - FRows.TopRow,y1,y2);
//            FSkin.PaintCursor(x1,y1,x2,y2);

    FColumns.ColPos(RectDest.Col1 - FColumns.LeftCol,RectDest.Col2 - FColumns.LeftCol,FCursorXY.X1,FCursorXY.X2);
    FRows.RowPos(RectDest.Row1 - FRows.TopRow,RectDest.Row2 - FRows.TopRow,FCursorXY.Y1,FCursorXY.Y2);
    if FCursorSizePos = cspCell then begin
      if RectDest.Col1 > RectSel.Col1 then Dec(FCursorXY.X1,10);
      if RectDest.Col2 < RectSel.Col2 then Inc(FCursorXY.X2,10);
      if RectDest.Row1 > RectSel.Row1 then Dec(FCursorXY.Y1,10);
      if RectDest.Row2 < RectSel.Row2 then Inc(FCursorXY.Y2,10);
    end;

    FSkin.PaintCursor(FCursorXY.X1,FCursorXY.Y1,FCursorXY.X2,FCursorXY.Y2,FCursorSizePos);
    FCursorXYMargO.X1 := FCursorXY.X1 - MARG_CELLCURSORHIT;
    FCursorXYMargO.Y1 := FCursorXY.Y1 - MARG_CELLCURSORHIT;
    FCursorXYMargO.X2 := FCursorXY.X2 + MARG_CELLCURSORHIT - 1;
    FCursorXYMargO.Y2 := FCursorXY.Y2 + MARG_CELLCURSORHIT - 1;

    FCursorXYMargI.X1 := FCursorXY.X1 + MARG_CELLCURSORHIT;
    FCursorXYMargI.Y1 := FCursorXY.Y1 + MARG_CELLCURSORHIT;
    FCursorXYMargI.X2 := FCursorXY.X2 - MARG_CELLCURSORHIT - 1;
    FCursorXYMargI.Y2 := FCursorXY.Y2 - MARG_CELLCURSORHIT - 1;
  end;
end;

procedure TXLSBookSheet.RepaintRequest(Sender: TObject);
begin
  FCacheRequest := True;
  Paint;
end;

procedure TXLSBookSheet.RowGroupBtnsClick(Sender: TObject; Hdr, Level: integer);
begin
  if Hdr >= 0 then
    FXLSSheet.Rows.ToggleGroupedAll(Level);
  FRows.CalcHeaders;
  Invalidate;
end;

procedure TXLSBookSheet.SetRightCol(const Value: integer);
begin
  FColumns.RightCol := Value;
end;

procedure TXLSBookSheet.SetSize(const pX1, pY1, pX2, pY2: integer);
var
  W: integer;
  pCX1,pCY1,pCX2,pCY2: integer;

procedure SetTabsSize;
begin
  if FTabs.Visible then begin
    W := Round((pCX2 - pX1 - SPLITTER_SIZE) * FTabs.PercentWidth + pX1);
    if (W - pX1) < SPLITTER_SIZE then
      W := pX1 + SPLITTER_SIZE;
    FTabs.SetSize(pX1,pCY2 + 1,W,pY2);
  end
  else
    W := pX1;
end;

procedure SetTopLeft;
var
  tlX,tlY: integer;
begin
  if FTopLeft.Visible then begin
    tlX := pX1;
    tlY := pY1;
    if FColumns.Visible and (FColumns.OutlineLevel > 0) then begin
      tlY := FColumns.CY1;
      FColGroupBtns.SetSize(FRows.CX1,pY1,FRows.CX2 + 1,FColumns.CY1 - 1);
    end;
    if FRows.Visible and (FRows.OutlineLevel > 0) then begin
      tlX := FRows.CX1;
      FRowGroupBtns.SetSize(pX1,FColumns.CY1,FRows.CX1,FRows.CY1);
    end;
    FTopLeft.SetSize(tlX,tlY,pCX1 - 1,pCY1 - 1);
    if FColGroupBtns.Visible and FRowGroupBtns.Visible then
      FTopTopLeft.SetSize(FX1,FY1,tlX - 1,tlY - 1);
  end;
end;

begin
  inherited SetSize(pX1, pY1, pX2, pY2);

  if FSheetChanged then
    DoSheetChanged;
//  H := Round((FXLSSheet.DefaultRowHeight / 20) * (FSkin.Font.PixelsPerInch / 72));
  pCX1 := pX1;
  PCY1 := pY1;
  pCX2 := pX2;
  PCY2 := pY2;

  FCSplit.SetSize(pX1 + 1,pY1 + 1,pX1,pY1);

  FHSplit.Small := FHideHSplitter;
  FVSplit.Small := FHideVSplitter;
  FCSplit.Visible := not FHideHSplitter and not FHideVSplitter;

  if FRows.Visible then
    Inc(pCX1,FRows.RowWidth);
  if FColumns.Visible then
    Inc(PCY1,FColumns.ColHeight);
  if FVScroll.Visible then
    Dec(pCX2,FVScroll.SysSize);
  if FTabs.Visible or FHScroll.Visible then
    Dec(pCY2,FHScroll.SysSize);
  SetClientSize(pCX1,pCY1,pCX2,pCY2);
  SetTabsSize;
  if FHScroll.Visible then begin
    FHScroll.SetSize(W,pCY2 + 1,pCX2 - SPLITTER_SIZE,pY2);
    FHSplit.SetSize(pCX2 - SPLITTER_SIZE,pCY2,pCX2,pY2 + 1);
  end;
  if FVScroll.Visible then begin
    FVScroll.SetSize(pCX2,pY1 + SPLITTER_SIZE,pX2,pCY2 + 1);
    FVSplit.SetSize(pCX2,pY1,pX2 + 1,pY1 + SPLITTER_SIZE);
  end;
  if FRows.Visible then
    FRows.SetSize(pX1,pCY1,pCX1 - 1,pCY2)
  else
    FRows.SetSize(0,pCY1,0,pCY2);
  if FColumns.Visible then
    FColumns.SetSize(pCX1,pY1,pCX2,pCY1 - 1)
  else
    FColumns.SetSize(pCX1,0,pCX2,0);
  if FHScroll.Visible or FTabs.Visible then begin
    FBottomRight.SetSize(pCX2 + 1,pCY2 + 1,pX2,pY2);
    CalcHScroll;
  end;
  if FVScroll.Visible then
    CalcVScroll;

  SetTopLeft;
  FDrawing.SetSize(pCX1,pCY1,pCX2,pCY2);
  FCacheRequest := True;
end;

procedure TXLSBookSheet.SetVisibleRects;
var
  i: integer;
  RowsColsVisible: boolean;

procedure SetGrpBtns;
begin
  FColGroupBtns.Visible := (FColumns.Visible) and (FColumns.OutlineLevel > 0);
  FColGroupBtns.Levels := FColumns.OutlineLevel;
  FRowGroupBtns.Visible := (FRows.Visible) and (FRows.OutlineLevel > 0);
  FRowGroupBtns.Levels := FRows.OutlineLevel;
end;

begin
  for i := 0 to FChilds.Count - 1 do
    FChilds[i].Visible := False;

  FHScroll.Visible := False;
  FVScroll.Visible := False;

  if FIsPrinting then
    RowsColsVisible := psoRowColHeading in FXLSSheet.PrintSettings.Options
  else
    RowsColsVisible := soRowColHeadings in FXLSSheet.Options;

  for i := 0 to FChilds.Count - 1 do
    FChilds[i].Visible := True;
  FColumns.Visible := RowsColsVisible;
  FRows.Visible := RowsColsVisible;
  FHScroll.Visible := True;
  FVScroll.Visible := True;
  SetGrpBtns;

  if FIsPrinting then begin
    FVScroll.Visible := False;
    FHScroll.Visible := False;
  end
  else begin
{
    FVScroll.Visible := soRowColHeadings in FXLSSheet.Options;
    FHScroll.Visible := soRowColHeadings in FXLSSheet.Options;
}
    FVScroll.Visible := FOptions.View.VerticalScroll;
    FHScroll.Visible := FOptions.View.HorizontalScroll;
  end;
  FTabs.Visible := FHScroll.Visible;
  FHSplit.Visible := FHScroll.Visible;
  FBottomRight.Visible := FHScroll.Visible;

  FTopTopLeft.Visible := FTopLeft.Visible;
end;

procedure TXLSBookSheet.SheetChanged(Sheet: TXLSWorkSheet);
begin
  HideForms;

  FXLSSheet := Sheet;
  FCells := FXLSSheet.MMUCells;
  FSheetChanged := True;
  FDrawing.LoadObjects(FXLSSheet.Xc12Sheet,FXLSSheet.Drawing);
end;

procedure TXLSBookSheet.ShowCursor;
begin
  with FXLSSheet.SelectedAreas.First do begin
    PaintCursor(Col1,Row1,Col2,Row2);
  end;
end;

procedure TXLSBookSheet.ShowInplaceEdit(AAssignCellValue: boolean; const AIsFormula: boolean);
var
  i: integer;
  S: AxUCString;
  C,R: integer;
  X,Y: integer;
  W,H: double;
  x1,y1,x2,y2: integer;
  Rect: TXYRect;
  CA: TCellArea;
  Cols: TDynIntegerArray;
  Rows: TDynIntegerArray;
//  Merged: TXc12MergedCell;
  MergedCell: TCellArea;
  CacheIndex: integer;
  Accept: boolean;
begin
  if (FInplaceEditor = Nil) and not FOPtions.ReadOnly then begin
    C := FXLSSheet.SelectedAreas[0].Col1;
    R := FXLSSheet.SelectedAreas[0].Row1;

    if XLSSheet.IsCellLocked(C,R) then begin
      if Assigned(FNotificationEvent) then
        FNotificationEvent(Self,xbnCellLocked);
      Exit;
    end;

    if Assigned(FEditorOpenEvent) then begin
      Accept := True;
      FEditorOpenEvent(Self,Accept);
      if not Accept then
        Exit;
    end;

    FInplaceEditor := TXBookInplaceEditor.Create(Self,FManager);
    FInplaceEditor.OnCharFmtChanged := IECharFormatChanged;
    Add(FInplaceEditor);

    SetLength(Cols,FColumns.Count);
    for i := 0 to FColumns.Count - 1 do
      Cols[i] := FColumns.ColWidth[i];

    SetLength(Rows,FRows.Count);
    for i := 0 to FRows.Count - 1 do
      Rows[i] := FRows.RowHeight[i];

    CA := FXLSSheet.SelectedAreas.First;

//    i := FXLSSheet.MergedCells.CellInAreas(CA.Col1,CA.Row1);
//    if i >= 0 then begin
//      Merged := FXLSSheet.MergedCells[i];
//      FInplaceEditor.SetColsRows(Cols,Rows,CA.Col1 - FColumns.LeftCol,Merged.Col2 - Merged.Col1 + 1,CA.Row1 - FRows.TopRow,Merged.Row2 - Merged.Row1 + 1);
//    end
//    else
      FInplaceEditor.SetColsRows(Cols,Rows,CA.Col1 - FColumns.LeftCol,1,CA.Row1 - FRows.TopRow,1);
    FInplaceEditor.SetCellPos(FXLSSheet.Index,CA.Col1,CA.Row1);

    X := FColumns.AbsPos[C] + FColumns.Offset;
    Y := FRows.AbsPos[R] + FRows.Offset;
    W := FColumns.ColWidthFloat[C - FColumns.LeftCol];
    H := FRows.RowHeightFloat[R - FRows.TopRow];

    x1 := Round(x);
    y1 := Round(y);
    x2 := Round(x + w - 2);
    y2 := Round(y + h - 2);

    CacheIndex := (R - FRows.TopRow) * FColumns.Count + (C - FColumns.LeftCol);
    FCellPaint.Clear(x1,y1,x2,y2,FCellStyleCache[CacheIndex]);
    with FCellStyleCache[CacheIndex] do
      FCellPaint.CalcPaintRect(StyleNW[ccS],StyleNW[ccE],StyleSE[ccN],StyleSE[ccW]);

    if FCellStyleCache[CacheIndex].InMerged >= 0 then begin
      MergedCell := FXLSSheet.MergedCells[FCellStyleCache[CacheIndex].InMerged];
      Rect := FCellPaint.FPaintRect;
      C := Min(FColumns.RightCol,MergedCell.Col2);
      R := Min(FRows.BottomRow,MergedCell.Row2);
      Rect.X2 := FColumns.AbsPos[C] + FColumns.AbsWidth[C] + FColumns.Offset;
      Rect.Y2 := FRows.AbsPos[R] + FRows.AbsHeight[R] + FRows.Offset;
    end
    else
      Rect := FCellPaint.FPaintRect;

    S := FInplaceEditor.Show(Rect.X1,Rect.Y1,Rect.X2 - Rect.X1,Rect.Y2 - Rect.Y1,AAssignCellValue);
    if AIsFormula or FCells.IsFormula(CA.Col1,CA.Row1) then begin
      if Assigned(FEditorDisableFmtEvent) then
        FEditorDisableFmtEvent(FInplaceEditor);

      FEditRefs := TXBookEditRefList.Create(FManager,FInplaceEditor,Self,FColumns,FRows);
      FEditRefs.OnPaintCells := IEPaintCells;
      FInplaceEditor.OnTextChanged := FEditRefs.SetFormula;
      FEditRefs.SetFormula(S,not FXLSSheet.MMUCells.IsFormulaCompiled(C,R));
      FEditRefs.Paint;
    end;
  end;
end;

procedure TXLSBookSheet.ShowPivotForm;
var
  PivTbl: TXLSPivotTable;
begin
  if (FOwner <> Nil) and (FFormPivot = Nil) then begin
    PivTbl := FXLSSheet.PivotTables.Hit(FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow);

    if (PivTbl = Nil) and (FXLSSheet.PivotTables.Count > 0) then
      PivTbl := FXLSSheet.PivotTables[0];

    if PivTbl <> Nil then
      DoShowPivotForm(PivTbl);
  end;
end;

procedure TXLSBookSheet.TabSizeChanged(Sender: TObject; X, Y: integer);
var
  W: integer;
begin
  FTabs.PercentWidth := (X - FX1) / (FCX2 - FX1 - SPLITTER_SIZE);

  W := Round((FCX2 - FX1 - SPLITTER_SIZE) * FTabs.PercentWidth + FX1);
  if (W - FX1) < SPLITTER_SIZE then
    W := FX1 + SPLITTER_SIZE;
  FTabs.SetSize(FX1,FCY2 + 1,W,FY2);
  
  FHScroll.X1 := W;
  FHScroll.Y1 := FCY2 + 1;
  FHScroll.X2 := FCX2 - SPLITTER_SIZE;
  FHScroll.Paint;

  FTabs.Paint;
end;

procedure TXLSBookSheet.ClearSelected(Col,Row: integer; DoPaint: boolean; CursorSizePos: TCursorSizePos);
var
  ActiveCellChanged: boolean;
  PivTbl: TXLSPivotTable;

procedure DebugShow;
var
  i: integer;
  S: string;
begin
  if FCellStyleCache = Nil then
    Exit;
  i := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);

  with FCellStyleCache[i] do begin
    S := Format('%.2d%.2d%.2d%.2d',[Integer(StyleNW[ccN]),Integer(StyleNW[ccS]),Integer(StyleNW[ccE]),Integer(StyleNW[ccW])]);
    S := S + Format('%.2d%.2d%.2d%.2d',[Integer(StyleNE[ccN]),Integer(StyleNE[ccS]),Integer(StyleNE[ccE]),Integer(StyleNE[ccW])]);
    S := S + Format('%.2d%.2d%.2d%.2d',[Integer(StyleSW[ccN]),Integer(StyleSW[ccS]),Integer(StyleSW[ccE]),Integer(StyleSW[ccW])]);
    S := S + Format('%.2d%.2d%.2d%.2d',[Integer(StyleSE[ccN]),Integer(StyleSE[ccS]),Integer(StyleSE[ccE]),Integer(StyleSE[ccW])]);
  end;
end;

begin
  ActiveCellChanged := (FXLSSheet.SelectedAreas.ActiveCol <> Col) or (FXLSSheet.SelectedAreas.ActiveRow <> Row);

  if DoPaint then begin
    BeginPaint;
    if FXLSSheet.SelectedAreas.CursorVisible then
      PaintCursor(FXLSSheet.SelectedAreas[0].Col1,FXLSSheet.SelectedAreas[0].Row1,FXLSSheet.SelectedAreas[0].Col2,FXLSSheet.SelectedAreas[0].Row2);
    PaintCells(FXLSSheet.SelectedAreas[0].Col1,FXLSSheet.SelectedAreas[0].Row1,FXLSSheet.SelectedAreas[0].Col2,FXLSSheet.SelectedAreas[0].Row2,ctpInvSelected);
//    PaintCells(FColumns.LeftCol,FRows.TopRow,FColumns.RightCol,FRows.BottomRow,ctpInvSelected);
    EndPaint;
  end;
  FSizeCol := Col;
  FSizeRow := Row;
  FColumns.ClearSelected;
  FRows.ClearSelected;

  FXLSSheet.SelectedAreas.Clear;

  FInMergedCells := FXLSSheet.MergedCells.CellInAreas(Col,Row);
  if FInMergedCells >= 0 then begin
    with FXLSSheet.MergedCells[FInMergedCells] do begin
      FXLSSheet.SelectedAreas.Init(Col1,Row1,Col2,Row2,Col,Row);
      FColumns.SetSelectedState(Integer(Col1) - FColumns.LeftCol,Integer(Col2) - FColumns.LeftCol,(Integer(Row1) = 0) and (Integer(Row2) >= MAXSELROW));
      FRows.SetSelectedState(Integer(Row1) - FRows.TopRow,Integer(Row2) - FRows.TopRow,(Col1 = 0) and (Integer(Col2) >= XLS_MAXCOL));
    end;
  end
  else begin
    FXLSSheet.SelectedAreas.Init(Col,Row);
    FColumns.SetSelectedState(Col - FColumns.LeftCol,Col - FColumns.LeftCol,False);
    FRows.SetSelectedState(Row - FRows.TopRow,Row - FRows.TopRow,False);
  end;

{
  FColumns.SetSelectedState(Col - FColumns.LeftCol,Col - FColumns.LeftCol,False);
  FRows.SetSelectedState(Row - FRows.TopRow,Row - FRows.TopRow,False);
}
  FColumns.Paint;
  FRows.Paint;

  if DoPaint then begin
    BeginPaint;
    FCursorSizePos := CursorSizePos;
    if (FCursorSizePos = cspCell) and FOptions.ReadOnly then
      FCursorSizePos := cspCellReadOnly;
    PaintCursor(FXLSSheet.SelectedAreas[0].Col1,FXLSSheet.SelectedAreas[0].Row1,FXLSSheet.SelectedAreas[0].Col2,FXLSSheet.SelectedAreas[0].Row2);

    if FCopyAreas.Count = 1 then
      PaintMarchingAnts(FCopyAreas[0].Col1,FCopyAreas[0].Row1,FCopyAreas[0].Col2,FCopyAreas[0].Row2);

    EndPaint;
  end;

  if ActiveCellChanged then begin
    if FOwner <> Nil then begin
      PivTbl := FXLSSheet.PivotTables.Hit(Col,Row);

      if PivTbl <> Nil then begin
        if (FFormPivot <> Nil) and (FFormPivot.PivotTable <> PivTbl) then
          HideForms;

        if FFormPivot = Nil then
          DoShowPivotForm(PivTbl);
      end
      else if FFormPivot <> Nil then
        HideForms;
    end;

    if Assigned(FCellChangedEvent) then
      FCellChangedEvent(Self,Col,Row);
  end;
end;

procedure TXLSBookSheet.AppendSelection(Col, Row: integer; DoPaint: boolean);
var
  i: integer;
  DColA,DColB,DRowA,DRowB: integer;
  OrigArea,NewArea,TotArea,MinArea,A: TXLSCellArea;
  SplitAreas: array[0..2] of TXLSCellArea;

function Split(SourceArea, Area: TXLSCellArea; var SplitAreas: array of TXLSCellArea): integer;
var
  TmpR1,TmpR2: word;
begin
  Result := 0;
  TmpR1 := SourceArea.Row1;
  TmpR2 := SourceArea.Row2;

  if Area.Row1 > SourceArea.Row1 then begin
    SplitAreas[Result].Col1 := SourceArea.Col1;
    SplitAreas[Result].Row1 := SourceArea.Row1;
    SplitAreas[Result].Col2 := SourceArea.Col2;
    SplitAreas[Result].Row2 := Area.Row1 - 1;
    Inc(Result);
    TmpR1 := Area.Row1;
  end;

  if Area.Row2 < SourceArea.Row2 then begin
    SplitAreas[Result].Col1 := SourceArea.Col1;
    SplitAreas[Result].Row1 := Area.Row2 + 1;
    SplitAreas[Result].Col2 := SourceArea.Col2;
    SplitAreas[Result].Row2 := SourceArea.Row2;
    Inc(Result);
    TmpR2 := Area.Row2;
  end;

  if Area.Col1 > SourceArea.Col1 then begin
    SplitAreas[Result].Col1 := SourceArea.Col1;
    SplitAreas[Result].Row1 := TmpR1;
    SplitAreas[Result].Col2 := Area.Col1 - 1;
    SplitAreas[Result].Row2 := TmpR2;
    Inc(Result);
  end;

  if Area.Col2 < SourceArea.Col2 then begin
    SplitAreas[Result].Col1 := Area.Col2 + 1;
    SplitAreas[Result].Row1 := TmpR1;
    SplitAreas[Result].Col2 := SourceArea.Col2;
    SplitAreas[Result].Row2 := TmpR2;
    Inc(Result);
  end;
end;

procedure ShrinkCol(var Area: TXLSCellArea);
begin
  if Area.Col1 = FXLSSheet.SelectedAreas.ActiveCol then
    Dec(Area.Col2)
  else
    Inc(Area.Col1);
end;

procedure ShrinkRow(var Area: TXLSCellArea);
begin
  if Area.Row1 = FXLSSheet.SelectedAreas.ActiveRow then
    Dec(Area.Row2)
  else {if Area.Row2 = FXLSSheet.SelectedAreas.ActiveRow then }
    Inc(Area.Row1);
end;

procedure ShrinkArea(var Area: TXLSCellArea);
begin
  ShrinkCol(Area);
  ShrinkRow(Area);
end;

procedure IncludeMerged;
var
  i,j,C,R,C1,R1,C2,R2: integer;
begin
   C1 := Max(FXLSSheet.SelectedAreas.Last.Col1 - FColumns.LeftCol,0);
   R1 := Max(FXLSSheet.SelectedAreas.Last.Row1 - FRows.TopRow,0);
   C2 := Min(FXLSSheet.SelectedAreas.Last.Col2 - FColumns.LeftCol,FColumns.Count - 1);
   R2 := Min(FXLSSheet.SelectedAreas.Last.Row2 - FRows.TopRow,FRows.Count - 1);
   for R := R1 to R2 do begin
     for C := C1 to C2 do begin
       i := R * FColumns.Count + C;
       if FCellStyleCache[i].InMerged >= 0 then begin
         j := FCellStyleCache[i].InMerged;
         FXLSSheet.SelectedAreas.Last.Intersect(FXLSSheet.MergedCells[j].Col1,FXLSSheet.MergedCells[j].Row1,FXLSSheet.MergedCells[j].Col2,FXLSSheet.MergedCells[j].Row2);
       end;
     end;
   end;
end;

begin
  with FXLSSheet.SelectedAreas.Last do begin
    OrigArea.Col1 := Col1;
    OrigArea.Row1 := Row1;
    OrigArea.Col2 := Col2;
    OrigArea.Row2 := Row2;

    Col1 := FXLSSheet.SelectedAreas.ActiveCol;
    Row1 := FXLSSheet.SelectedAreas.ActiveRow;
    Col2 := Col;
    Row2 := Row;

    Normalize;

    IncludeMerged;

    if (OrigArea.Col1 = Col1) and (OrigArea.Row1 = Row1) and (OrigArea.Col2 = Col2) and (OrigArea.Row2 = Row2) then
      Exit;

    FSizeCol := Col;
    FSizeRow := Row;
    if not DoPaint then
      Exit;

    NewArea.Col1 := Col1;
    NewArea.Row1 := Row1;
    NewArea.Col2 := Col2;
    NewArea.Row2 := Row2;

    TotArea.Col1 := Min(OrigArea.Col1,NewArea.Col1);
    TotArea.Row1 := Min(OrigArea.Row1,NewArea.Row1);
    TotArea.Col2 := Max(OrigArea.Col2,NewArea.Col2);
    TotArea.Row2 := Max(OrigArea.Row2,NewArea.Row2);
    MinArea.Col1 := Max(OrigArea.Col1,NewArea.Col1);
    MinArea.Row1 := Max(OrigArea.Row1,NewArea.Row1);
    MinArea.Col2 := Min(OrigArea.Col2,NewArea.Col2);
    MinArea.Row2 := Min(OrigArea.Row2,NewArea.Row2);

    FColumns.SetFocusedState(TotArea.Col1 - FColumns.LeftCol,TotArea.Col2 - FColumns.LeftCol);
    FColumns.Paint;
    FRows.SetFocusedState(TotArea.Row1 - FRows.TopRow,TotArea.Row2 - FRows.TopRow);
    FRows.Paint;

    DColA := Abs(Col2 - Col1);
    DColB := Abs(OrigArea.Col2 - OrigArea.Col1);
    DRowA := Abs(Row2 - Row1);
    DRowB := Abs(OrigArea.Row2 - OrigArea.Row1);

    BeginPaint;

    if FXLSSheet.SelectedAreas.CursorVisible then
      PaintCursor(OrigArea.Col1,OrigArea.Row1,OrigArea.Col2,OrigArea.Row2);

    if ((MinArea.Col2 - MinArea.Col1) = 0) or ((MinArea.Row2 - MinArea.Row1) = 0) then begin
      PaintCells(OrigArea.Col1,OrigArea.Row1,OrigArea.Col2,OrigArea.Row2,ctpAll);
      PaintCells(NewArea.Col1,NewArea.Row1,NewArea.Col2,NewArea.Row2,ctpAll);
    end
    else if (DColA > DColB) and (DRowA > DRowB) then begin
      ShrinkArea(OrigArea);
      Split(TotArea,OrigArea,SplitAreas);
      PaintCells(SplitAreas[0].Col1,SplitAreas[0].Row1,SplitAreas[0].Col2,SplitAreas[0].Row2,ctpAll);
      PaintCells(SplitAreas[1].Col1,SplitAreas[1].Row1,SplitAreas[1].Col2,SplitAreas[1].Row2,ctpAll);
    end
    else if (DColA < DColB) and (DRowA < DRowB) then begin
      ShrinkArea(NewArea);
      Split(TotArea,NewArea,SplitAreas);
      PaintCells(SplitAreas[0].Col1,SplitAreas[0].Row1,SplitAreas[0].Col2,SplitAreas[0].Row2,ctpAll);
      PaintCells(SplitAreas[1].Col1,SplitAreas[1].Row1,SplitAreas[1].Col2,SplitAreas[1].Row2,ctpAll);
    end
    else if (DColA > DColB) and (DRowA < DRowB) then begin
      A := MinArea;
      ShrinkRow(A);
      Split(OrigArea,A,SplitAreas);
      PaintCells(SplitAreas[0].Col1,SplitAreas[0].Row1,SplitAreas[0].Col2,SplitAreas[0].Row2,ctpAll);
      ShrinkCol(MinArea);
      Split(NewArea,MinArea,SplitAreas);
      PaintCells(SplitAreas[0].Col1,SplitAreas[0].Row1,SplitAreas[0].Col2,SplitAreas[0].Row2,ctpAll);
    end
    else if (DColA < DColB) and (DRowA > DRowB) then begin
      A := MinArea;
      ShrinkCol(A);
      Split(OrigArea,A,SplitAreas);
      PaintCells(SplitAreas[0].Col1,SplitAreas[0].Row1,SplitAreas[0].Col2,SplitAreas[0].Row2,ctpAll);
      ShrinkRow(MinArea);
      Split(NewArea,MinArea,SplitAreas);
      PaintCells(SplitAreas[0].Col1,SplitAreas[0].Row1,SplitAreas[0].Col2,SplitAreas[0].Row2,ctpAll);
    end
    else if (DColA > DColB) or (DRowA > DRowB) then begin
      if DColA > DColB then begin
        ShrinkCol(OrigArea);
        i := 0;
      end
      else begin
        ShrinkRow(OrigArea);
        i := 0;
      end;
      Split(TotArea,OrigArea,SplitAreas);
      PaintCells(SplitAreas[i].Col1,SplitAreas[i].Row1,SplitAreas[i].Col2,SplitAreas[i].Row2,ctpAll);
    end
    else if (DColA < DColB) or (DRowA < DRowB) then begin
      if DColA < DColB then begin
        ShrinkCol(NewArea);
        i := 0;
      end
      else begin
        ShrinkRow(NewArea);
        i := 0;
      end;
      Split(TotArea,NewArea,SplitAreas);
      PaintCells(SplitAreas[i].Col1,SplitAreas[i].Row1,SplitAreas[i].Col2,SplitAreas[i].Row2,ctpAll);
    end;

    if FXLSSheet.SelectedAreas.CursorVisible then
      PaintCursor(Col1,Row1,Col2,Row2);

    EndPaint;

    if Assigned(FSelectionChangedEvent) then
      FSelectionChangedEvent(Self);
  end;
end;

procedure TXLSBookSheet.ApplyCondFmt;
var
  i,j: integer;
  C,R: integer;
  Areas: TCellAreas;
  CondFmt: TXLSConditionalFormat;
begin
  Areas := TCellAreas.Create;
  try
    FXLSSheet.ConditionalFormats.FillHitList(FColumns.LeftCol,FRows.TopRow,FColumns.RightCol,FRows.BottomRow,Areas);

    for i := 0 to Areas.Count - 1 do begin
      CondFmt := TXLSConditionalFormat(Areas[i].Obj);
      CondFmt.BeginExecute(Areas[i].Tag);
      for R := Areas[i].Row1 to Areas[i].Row2 do begin
        for C := Areas[i].Col1 to Areas[i].Col2 do begin
          j := (R - FRows.TopRow) * FColumns.Count + (C - FColumns.LeftCol);

          FCellStyleCache[j].CondFmtRes := CondFmt.Execute(C,R);
          case FCellStyleCache[j].CondFmtRes of
            xcfrtDXF    : begin
              FCellStyleCache[j].DXF := CondFmt.ResultDXF;
              if FCellStyleCache[j].DXF.Fill <> Nil then begin
                FCellStyleCache[j].ForeColor := CondFmt.ResultDXF.Fill.BgColor.ARGB;
                FCellStyleCache[j].BackColor := CondFmt.ResultDXF.Fill.FgColor.ARGB;
              end;
              if FCellStyleCache[j].DXF.Border <> Nil then begin
                FCellStyleCache[j].RightColor := FCellStyleCache[j].DXF.Border.Right.Color.ARGB;
                FCellStyleCache[j].BottomColor := FCellStyleCache[j].DXF.Border.Bottom.Color.ARGB;
                FCellStyleCache[j].StyleSE[ccN] := FCellStyleCache[j].DXF.Border.Right.Style;
                FCellStyleCache[j].StyleSE[ccW] := FCellStyleCache[j].DXF.Border.Bottom.Style;
              end;
            end;
            xcfrtColor  : begin
              FCellStyleCache[j].ForeColor := CondFmt.ResultColor;
              FCellStyleCache[j].BackColor := CondFmt.ResultColor;
            end;
            xcfrtDataBar,
            xcfrtDataBarNoText: begin
              FCellStyleCache[j].CondFmtVal := Round(COND_FMT_VAL_SCALE * CondFmt.ResPercent);
              FCellStyleCache[j].CondFmtColor := CondFmt.ResultColor;
            end;
            xcfrtIcon,
            xcfrtIconNoText: begin
              FCellStyleCache[j].CondFmtVal := Integer(CondFmt.ResIconStyle);
              FCellStyleCache[j].CondFmtColor := CondFmt.ResIconIndex;
            end;
          end;
        end;
      end;
    end;
  finally
    Areas.Free;
  end;
end;

procedure TXLSBookSheet.BeginPrint(const AVertAdjustment: double; const AStdFontWidth: integer);
begin
  FPrintStdFontWidth := AStdFontWidth;

  FIsPrinting := True;
  FRows.PrintVertAdj := AVertAdjustment;

  FColumns.ColumnsChanged(FXLSSheet.Columns,0,FPrintStdFontWidth);
  FRows.RowsChanged(FXLSSheet.Rows,0);
end;

procedure TXLSBookSheet.BeginSelection(Col, Row: integer);
var
  C,R: integer;
begin
  if FXLSSheet.SelectedAreas.CursorVisible then
    PaintCursor(FXLSSheet.SelectedAreas[0].Col1,FXLSSheet.SelectedAreas[0].Row1,FXLSSheet.SelectedAreas[0].Col2,FXLSSheet.SelectedAreas[0].Row2);
  C := FXLSSheet.SelectedAreas.ActiveCol;
  R := FXLSSheet.SelectedAreas.ActiveRow;
  FXLSSheet.SelectedAreas.Add(Col,Row,Col,Row);
  FColumns.SetSelectedState(Col - FColumns.LeftCol,Col - FColumns.LeftCol,False);
  FColumns.Paint;
  FRows.SetSelectedState(Row - FRows.TopRow,Row - FRows.TopRow,False);
  FRows.Paint;

  BeginPaint;

  if CellInArea(C,R,FSheetArea) then
    PaintCells(C,R,C,R,ctpAll);
  PaintCells(Col,Row,Col,Row,ctpAll);

  EndPaint;

  FColumns.Paint;
  FRows.Paint;
end;

function TXLSBookSheet.GetLeftCol: integer;
begin
  Result := FColumns.LeftCol;
end;

function TXLSBookSheet.GetRightCol: integer;
begin
  Result := FColumns.RightCol;
end;

function TXLSBookSheet.GetTopRow: integer;
begin
  Result := FRows.TopRow;
end;

procedure TXLSBookSheet.SetBottomRow(Value: integer);
var
  W,V: integer;
begin
  Value := Fork(Value,0,XLS_MAXROW);
  V := Value - FRows.BottomRow;
  W := FRows.RowWidth;
  FRows.BottomRow := Value;
  // Width changed due to increase of row number (100 -> 1000)
  if W <> FRows.RowWidth then
    SetSize(FX1,FY1,FX2,FY2);
  FCacheRequest := True;
  if FEditRefs <> Nil then
    FEditRefs.ScrollUpdate(V,0);
end;

procedure TXLSBookSheet.SetLeftCol(Value: integer);
var
  V: integer;
begin
  Value := Fork(Value,0,XLS_MAXCOL);

  V := Value - FColumns.LeftCol;

  FColumns.LeftCol := Value;
  FCacheRequest := True;
  if FEditRefs <> Nil then
    FEditRefs.ScrollUpdate(V,0);
end;

procedure TXLSBookSheet.SetTopRow(Value: integer);
var
  W,V: integer;
begin
  Value := Fork(Value,0,XLS_MAXROW);
  V := Value - FRows.TopRow;
  W := FRows.RowWidth;
  FRows.SetTopRowPrint(Value,MAXINT);
  // Width changed due to increase of row number (100 -> 1000)
  if W <> FRows.RowWidth then
    SetSize(FX1,FY1,FX2,FY2);
  FCacheRequest := True;
  if FEditRefs <> Nil then
    FEditRefs.ScrollUpdate(V,0);
end;

procedure TXLSBookSheet.SetTopRowPrint(Value, MaxRowToCalc: integer);
var
  W: integer;
begin
  Value := Fork(Value,0,XLS_MAXROW);
  FRows.SetTopRowPrint(Value,MaxRowToCalc);
  W := FRows.RowWidth;
  // Width changed due to increase of row number (100 -> 1000)
  if W <> FRows.RowWidth then
    SetSize(FX1,FY1,FX2,FY2);
end;

function TXLSBookSheet.Scroll(Col, Row: integer): boolean;
begin
  Result := True;
  if Row >= 0 then begin
    if Row < FRows.TopRow then begin
      SetTopRow(Row);
      Result := False;
    end
    else if Row >= FRows.BottomRow then begin
      SetTopRow(FRows.TopRow + (Row - FRows.BottomRow) + 1);
      Result := False;
    end;
  end;
  if Col >= 0 then begin
    if Col < FColumns.LeftCol then begin
      SetLeftCol(Col);
      Result := False
    end
    else if Col >= FColumns.RightCol then begin
      SetLeftCol(FColumns.LeftCol + (Col - FColumns.RightCol) + 1);
      Result := False;
    end;
  end;
  if not Result then begin
    UpdateSheetArea;
    CalcVScroll;
    CalcHScroll;
    Repaint;
  end;
end;

procedure TXLSBookSheet.HandleKey(const Key: TXSSKey; const Shift: TXSSShiftKeys);
var
  dCol,dRow,Col,Row: integer;
  ShowArea: boolean;
begin
  dCol := 0;
  dRow := 0;

  ShowArea := False;
  if  Shift = [skCtrl] then begin
    case Key of
      kyNone: ;
      kyEscape: ;
      kyLeft: ;
      kyRight: ;
      kyUp: ;
      kyDown: ;
      kyPgUp: ;
      kyPgDown: ;
      kyHome: ;
      kyEnd: ;
      kyFirstCol: ;
      kyLastCol: ;
      kyFirstRow: ;
      kyLastRow: ;
      kyInplaceEdit: ;
    end;
  end;

  case Key of
    kyEscape     : begin
      FCopyAreas.Clear;
      HideInplaceEdit;
      PaintInvalidate;
      Exit;
    end;

    kyCopy       : begin
      if FXLSSheet.SelectedAreas.Count = 1 then begin
        FCopyAreas.Assign(FXLSSheet.SelectedAreas);
        FCopyAreas.Tag := Integer(kyCopy);

        PaintMarchingAnts(FCopyAreas[0].Col1,FCopyAreas[0].Row1,FCopyAreas[0].Col2,FCopyAreas[0].Row2);
      end;
    end;
    kyCut        : begin
      if FXLSSheet.SelectedAreas.Count = 1 then begin
        FCopyAreas.Assign(FXLSSheet.SelectedAreas);
        FCopyAreas.Tag := Integer(kyCut);

        PaintMarchingAnts(FCopyAreas[0].Col1,FCopyAreas[0].Row1,FCopyAreas[0].Col2,FCopyAreas[0].Row2);
      end;
    end;
    kyPaste,
    kyPasteSpecial: begin
      if FCopyAreas.Count = 1 then begin
        FWorkbook.CopyCells(FXLSSheet.Index,FCopyAreas,FXLSSheet.Index,FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow);
        dCol := FCopyAreas[0].Col2 - FCopyAreas[0].Col1;
        dRow := FCopyAreas[0].Row2 - FCopyAreas[0].Row1;
        if FCopyAreas.Tag = Integer(kyCut) then
          FWorkbook.DeleteCells(FXLSSheet.Index,FCopyAreas);
        // If Dest overlaps source the area can't be copied again.
        if (FCopyAreas.Tag = Integer(kyCut)) or FCopyAreas.AreaInAreas(FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow,FXLSSheet.SelectedAreas.ActiveCol + dCol,FXLSSheet.SelectedAreas.ActiveRow + dRow) then
          FCopyAreas.Clear;

        ShowArea := True;

        PaintInvalidate;
      end;
    end;
    kyDelete     : begin
      FWorkbook.DeleteCells(FXLSSheet.Index,FXLSSheet.SelectedAreas,FDeleteOptions);
      PaintInvalidate;
    end;

    kyTab        : begin
      if skShift in Shift then
        dCol :=  -1
      else
        dCol :=  1;
    end;

    kyLeft       : dCol := -1;
    kyRight      : dCol :=  1;
    kyUp         : dRow := -1;
    kyDown       : dRow :=  1;
    kyPgUp       : dRow := -(FRows.BottomRow - FRows.TopRow);
    kyPgDown     : dRow :=  FRows.BottomRow - FRows.TopRow;
    kyHome       : ;
    kyEnd        : ;
    kyFirstCol   : dCol := -XLS_MAXCOL;
    kyLastCol    : dCol :=  XLS_MAXCOL;
    kyFirstRow   : dRow := -XLS_MAXROW;
    kyLastRow    : dRow :=  XLS_MAXROW;

    kyInplaceEdit: begin
      FCopyAreas.Clear;
      ShowInplaceEdit(True);
      Exit;
    end;
  end;

  if FInMergedCells >= 0 then begin
    with FXLSSheet.MergedCells[FInMergedCells] do begin
      case Key of
        kyLeft:   dCol := Integer(Col1) - FXLSSheet.SelectedAreas.ActiveCol - 1;
        kyRight:  dCol := Integer(Col2) - FXLSSheet.SelectedAreas.ActiveCol + 1;
        kyUp:     dRow := Integer(Row1) - FXLSSheet.SelectedAreas.ActiveRow - 1;
        kyDown:   dRow := Integer(Row2) - FXLSSheet.SelectedAreas.ActiveRow + 1;
        kyPgUp:   ;
        kyPgDown: ;
        kyHome:   ;
        kyEnd:    ;
      end;
    end;
  end;

  if not FOptions.ReadOnly and (Key <> kyTab) and (skShift in Shift) then begin
    if dCol > 0 then
      FSizeCol := FColumns.NextVisible(FSizeCol + dCol)
    else
      FSizeCol := FColumns.PrevVisible(FSizeCol + dCol);
    if dRow > 0 then
      FSizeRow := FRows.NextVisible(FSizeRow + dRow)
    else
      FSizeRow := FRows.PrevVisible(FSizeRow + dRow);

    ClipToExtent(FSizeCol,FSizeRow);

    AppendSelection(FSizeCol,FSizeRow,True);
    MoveCursor(FSizeCol,FSizeRow);
  end
  else begin
    if dCol > 0 then
      Col := FColumns.NextVisible(FXLSSheet.SelectedAreas.ActiveCol + dCol)
    else
      Col := FColumns.PrevVisible(FXLSSheet.SelectedAreas.ActiveCol + dCol);
    if dRow > 0 then
      Row := FRows.NextVisible(FXLSSheet.SelectedAreas.ActiveRow + dRow)
    else
      Row := FRows.PrevVisible(FXLSSheet.SelectedAreas.ActiveRow + dRow);

    ClipToExtent(Col,Row);

//    if FCopyAreas.Count = 0 then
      ClearSelected(Col,Row,True);

    MoveCursor(Col,Row);

    if ShowArea then begin
      ClearSelected(FXLSSheet.SelectedAreas.ActiveCol,FXLSSheet.SelectedAreas.ActiveRow,False);
      AppendSelection(FXLSSheet.SelectedAreas.ActiveCol + dCol,FXLSSheet.SelectedAreas.ActiveRow + dRow,True);
    end;
  end;
end;

procedure TXLSBookSheet.HideCursor;
begin
  // XOR painting.
  ShowCursor;
end;

procedure TXLSBookSheet.HideForms;
begin
  if FFormPivot <> Nil then begin
    FFormPivot.Free;
    FFormPivot := Nil;
  end;
end;

procedure TXLSBookSheet.HideInplaceEdit(const ASaveText: boolean = False; AKey: TXSSKey = kyDown);
var
  S: AxUCString;
//  XF: TXc12XF;
  Col,Row: integer;
  Cell: TXLSCellItem;
  vFloat: double;
  vDateTime: TDateTime;
  FontRuns: TXc12DynFontRunArray;
  Accept: boolean;
begin
  if FInplaceEditor <> Nil then begin
    Col := FInplaceEditor.Col;
    Row := FInplaceEditor.Row;

    Accept := True;
    if Assigned(FEditorCloseEvent) then
      FEditorCloseEvent(Self,Accept);

    if ASaveText and Accept then begin

      Cell := FCells.FindCell(Col,Row);
//      if Cell.Data <> Nil then
//        XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Cell)]
//      else
//        XF := FManager.StyleSheet.XFs.DefaultXF;

      S := FInplaceEditor.Text;
      if (Length(S) > 1) and (Copy(S,1,1) = '''') then
        FXLSSheet.AsString[Col,Row] := Copy(S,2,MAXINT)
      else if (Length(S) > 1) and (Copy(S,1,1) = '=') then begin
        FXLSSheet.AsFormula[Col,Row] := Copy(S,2,MAXINT);
//{$ifdef _AXOLOT_DEBUG}
//{$message warn '_TODO_ Calculate minimal'}
//{$endif}
        if FManager.Workbook.CalcPr.CalcMode in [cmAutomatic,cmAutoExTables] then
          FWorkbook.Calculate;
      end
      else if (Uppercase(S) = G_StrTRUE) or (Uppercase(S) = G_StrFALSE) then
        FXLSSheet.AsBoolean[Col,Row] := Uppercase(S) = G_StrTRUE
      else if ErrorTextToCellError(S) <> errUnknown then
        FXLSSheet.AsError[Col,Row] := ErrorTextToCellError(S)
      else if TryStrToFloat(S,vFloat) then begin
        FXLSSheet.AsFloat[Col,Row] := vFloat;
        if FManager.Workbook.CalcPr.CalcMode in [cmAutomatic,cmAutoExTables] then
          FWorkbook.Calculate;
      end
      else if TryStrToDate(S,vDateTime) then begin
        FXLSSheet.AsFloat[Col,Row] := vDateTime;
        if FManager.Workbook.CalcPr.CalcMode in [cmAutomatic,cmAutoExTables] then
          FWorkbook.Calculate;
      end
      else if TryStrToTime(S,vDateTime) then begin
        FXLSSheet.AsFloat[Col,Row] := vDateTime;
        if FManager.Workbook.CalcPr.CalcMode in [cmAutomatic,cmAutoExTables] then
          FWorkbook.Calculate;
      end
      else if TryStrToDateTime(S,vDateTime) then begin
        FXLSSheet.AsFloat[Col,Row] := vDateTime;
        if FManager.Workbook.CalcPr.CalcMode in [cmAutomatic,cmAutoExTables] then
          FWorkbook.Calculate;
      end
      else begin
        if FInplaceEditor.IsFormatted then begin
          FInplaceEditor.GetFormatted(FontRuns);
          FXLSSheet.RichTextCell(Col,Row,S,FontRuns);
        end
        else
          FXLSSheet.AsString[Col,Row] := S;
      end;
    end;

    FInplaceEditor.Hide;
    DeleteChild(FInplaceEditor);
    FInplaceEditor := Nil;
    if FEditRefs <> Nil then begin
      FEditRefs.Free;
      FEditRefs := Nil;
    end;
    Invalidate;

    if ASaveText then
      HandleKey(AKey,[]);

    if Assigned(FCellChangedEvent) then
      FCellChangedEvent(Self,Col,Row);
  end;
end;

procedure TXLSBookSheet.HideComment(Col, Row: integer);
begin
  if (FVisibleComment <> Nil) and ((FVisibleComment.Col <> Col) or (FVisibleComment.Row <> Row)) then begin
    FDrawing.RemoveComment(FVisibleComment);
    FVisibleComment := Nil;
    Paint;
    FDrawing.Paint;
    FSystem.ProcessRequests;
  end;
end;

procedure TXLSBookSheet.ClipToExtent(var Col, Row: integer);
begin
  if Col < 0 then
    Col := 0
  else if Col > XLS_MAXCOL then
    Col := XLS_MAXCOL;
  if Row < 0 then
    Row := 0
  else if Row > XLS_MAXROW then
    Row := XLS_MAXROW;
end;

procedure TXLSBookSheet.HandleScrollTimer(ASender: TObject; AData: Pointer);
var
  Col,Row: integer;
begin
  Col := 0;
  Row := 0;
  case TXSSKey(AData) of
    kyUp:    Row := -1;
    kyDown:  Row :=  1;
    kyLeft:  Col := -1;
    kyRight: Col :=  1;
  end;

  if FEditRefs <> Nil then begin
    case Row of
      -1: FSizeRow := FRows.TopRow - 1;
       1: FSizeRow := FRows.BottomRow;
    end;
    case Col of
      -1: FSizeCol := FColumns.LeftCol - 1;
       1: FSizeCol := FColumns.RightCol;
    end;
  end
  else if FSelectionEdit = Nil then begin
    Inc(FSizeCol,Col);
    Inc(FSizeRow,Row);
    ClipToExtent(FSizeCol,FSizeRow);
    AppendSelection(FSizeCol,FSizeRow,True);
  end
  else begin
    FSelectionEdit.ScrollUpdate(Col,Row);
    case Row of
      -1: FSizeRow := FRows.TopRow - 1;
       1: FSizeRow := FRows.BottomRow;
    end;
    case Col of
      -1: FSizeCol := FColumns.LeftCol - 1;
       1: FSizeCol := FColumns.RightCol;
    end;
  end;

  if Col = 0 then
    Col := -1
  else
    Col :=  FSizeCol;
  if Row = 0 then
    Row := -1
  else
    Row := FSizeRow;
  Scroll(Col,Row);
  if ((Row >= 0) and (FRows.TopRow <= 0)) or ((Col >= 0) and (FColumns.LeftCol <= 0)) then begin
    FSystem.StopTimer(FScrollTimerId);
    FScrollTimerId := -1;
  end;
end;

procedure TXLSBookSheet.CacheCells;

var
  i,j,W: integer;
  S: AxUCString;
  XF: TXc12XF;
  Cols,Rows: integer;
  Col,Row: integer;
  FontRunsCount: integer;
  FR: PXc12FontRunArray;
  Ref: TXLSCellItem;
  SCF: TSetCellFormat;
//  CondFmt: TXLSConditionalFormat;
//  MergedCell: TCellArea;

procedure ExpandRight(Width,Col,Index: integer);
begin
  while (Width > 0) and (Col < Cols) and (FCellStyleCache[Index].Ref.Data = Nil) and (FCellStyleCache[Index].InMerged < 0) do begin
    FCellStyleCache[Index].Flags := FCellStyleCache[Index].Flags + [csfHorizOverride];
    FCellStyleCache[Index].Flags := FCellStyleCache[Index].Flags + [csfTextExpandRight];
    Inc(Col);
    Inc(Index);
    if Col < Cols then
//      Dec(Width,FColumns.Width[Col - FColumns.LeftCol]);
      Dec(Width,FColumns.ColWidth[Col]);
  end;
end;

procedure ExpandLeft(Width,Col,Index: integer);
begin
  while (Width > 0) and (Col >= 0) and (Index >= 0) and (FCellStyleCache[Index].Ref.Data = Nil) and (FCellStyleCache[Index].InMerged < 0) do begin
    FCellStyleCache[Index].Flags := FCellStyleCache[Index].Flags + [csfHorizOverride];
    FCellStyleCache[Index].Flags := FCellStyleCache[Index].Flags + [csfTextExpandLeft];
    Dec(Col);
    Dec(Index);
    if Col >= 0 then
//      Dec(Width,FColumns.Width[Col - FColumns.LeftCol]);
      Dec(Width,FColumns.ColWidth[Col]);
  end;
end;

function Ex12CheckAutoLineColor(LineStyle: TXc12CellBorderStyle; Color: longword): longword;
begin
  if (LineStyle > cbsNone) and (Color = EX12_AUTO_COLOR) then
    Result := $00000000
  else
    Result := Color;
end;

begin
  FR := Nil;
  Cols := FColumns.Count;
  Rows := FRows.Count;
  if (Cols * Rows) <> (FCellCacheCols * FCellCacheRows) then begin
    ReAllocMem(FCellStyleCache,(Cols + 1) * (Rows + 1) * SizeOf(TCellStyleRec));
    FCellCacheCols := Cols;
    FCellCacheRows := Rows;
  end;
  FillChar(FCellStyleCache^,Cols * Rows * SizeOf(TCellStyleRec),#0);
  for i := 0 to Cols * Rows - 1 do
    FCellStyleCache[i].InMerged := -1;
  for Row := 0 to Rows - 1 do begin
    for Col := 0 to Cols - 1 do begin
      i := Row * Cols + Col;
      if (FXLSSheet.VisibleAreas.Count > 0) and (FXLSSheet.VisibleAreas.CellInAreas(Col,Row) < 0) then
        Ref.Data := Nil
      else
//      Cell := FXLSSheet.Cell[Col + FColumns.LeftCol,Row + FRows.TopRow];
        Ref := FCells.FindCell(FColumns.VisualCol[Col],Row + FRows.TopRow); // FXLSSheet.Cell[FColumns.VisualCol[Col],Row + FRows.TopRow];
      XF := Nil;
      if (Ref.Data <> Nil) then begin
        XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Ref)];
//        XF := FManager.StyleSheet.XFs[0];
        // Optimize this...
        if FXLSSheet.Comments.Find(Col + FColumns.LeftCol,Row + FRows.TopRow) <> Nil then
          FCellStyleCache[i].Flags := FCellStyleCache[i].Flags + [csfComment];
        if FCells.CellType(@Ref) <> xctBlank then
          FCellStyleCache[i].Ref := Ref;
        if (FParent <> Nil) and (xsoHyperlink in FSelectOptions) then
          FCellStyleCache[i].Hyperlink := FXLSSheet.Hyperlinks.Find(Col + FColumns.LeftCol,Row + FRows.TopRow)
        else
          FCellStyleCache[i].Hyperlink := Nil
      end
      else
        FCellStyleCache[i].Hyperlink := Nil;

      if XF = Nil then begin
        XF := FRows.XLSRows[Row + FRows.TopRow].Style;
        if (XF = Nil) or (XF.Index = XLS_STYLE_DEFAULT_XF) then
          XF := FColumns.XLSColumns[Col + FColumns.LeftCol].Style;
      end;

      if (XF <> Nil) and (XF.Index <> XLS_STYLE_DEFAULT_XF) then begin
        SCF.ForeColor := XF.Fill.FgColor.ARGB;
        SCF.BackColor := XF.Fill.BgColor.ARGB;
        SCF.Pattern := XF.Fill.PatternType;
        SCF.LeftColor := XF.Border.Left.Color.ARGB;
        SCF.TopColor := XF.Border.Top.Color.ARGB;
        SCF.RightColor := XF.Border.Right.Color.ARGB;
        SCF.BottomColor := XF.Border.Bottom.Color.ARGB;
        SCF.LeftStyle := XF.Border.Left.Style;
        SCF.TopStyle := XF.Border.Top.Style;
        SCF.RightStyle := XF.Border.Right.Style;
        SCF.BottomStyle := XF.Border.Bottom.Style;
        if ecboDiagonalUp in XF.Border.Options then
          SCF.DiagLines := dlBoth
        else if (ecboDiagonalUp in XF.Border.Options) and (ecboDiagonalDown in XF.Border.Options) then
          SCF.DiagLines := dlUp
        else if ecboDiagonalDown in XF.Border.Options then
          SCF.DiagLines := dlDown
        else
          SCF.DiagLines := dlNone;
        SCF.DiagColor := XF.Border.Diagonal.Color.ARGB;
        SCF.DiagStyle := XF.Border.Diagonal.Style;

        if FCellStyleCache[i].InMerged < 0 then begin
          if SCF.Pattern = efpNone then
            FCellStyleCache[i].ForeColor := EX12_AUTO_COLOR
          else
            FCellStyleCache[i].ForeColor := SCF.ForeColor;
          FCellStyleCache[i].BackColor := SCF.BackColor;
          FCellStyleCache[i].Pattern := SCF.Pattern;
          FCellStyleCache[i].InMerged := FXLSSheet.MergedCells.CellInAreas(Col + FColumns.LeftCol,Row + FRows.TopRow);
          if FCellStyleCache[i].InMerged >= 0 then
            CacheMergedCells(FCellStyleCache[i].InMerged,FCellStyleCache[i].ForeColor,FCellStyleCache[i].BackColor);
          SCF.ForeColor := FCellStyleCache[i].ForeColor;
          SCF.BackColor := FCellStyleCache[i].BackColor;
        end;

        FCellStyleCache[i].StyleNW[ccS] := SCF.LeftStyle;
        FCellStyleCache[i].StyleNW[ccE] := SCF.TopStyle;

        FCellStyleCache[i].StyleNE[ccS] := SCF.RightStyle;
        FCellStyleCache[i].StyleNE[ccW] := SCF.TopStyle;

        FCellStyleCache[i].StyleSW[ccN] := SCF.LeftStyle;
        FCellStyleCache[i].StyleSW[ccE] := SCF.BottomStyle;

        FCellStyleCache[i].StyleSE[ccN] := SCF.RightStyle;
        FCellStyleCache[i].StyleSE[ccW] := SCF.BottomStyle;

        FCellStyleCache[i].RightColor := Ex12CheckAutoLineColor(SCF.RightStyle,SCF.RightColor);
        FCellStyleCache[i].BottomColor := Ex12CheckAutoLineColor(SCF.BottomStyle,SCF.BottomColor);

        FCellStyleCache[i].Diagonal := SCF.DiagLines;
        FCellStyleCache[i].DiagColor := SCF.DiagColor;
        FCellStyleCache[i].DiagStyle := SCF.DiagStyle;

        if (Col > 0) and (Row > 0) then begin
          j := (Row - 1) * Cols + Col - 1;
          if FCellStyleCache[j].BottomColor = SCF.TopColor then begin
            FCellStyleCache[j].Flags := FCellStyleCache[j].Flags + [csfHorizOverride];
            j := Row * Cols + Col - 1;
            FCellStyleCache[j].Flags := FCellStyleCache[j].Flags + [csfHorizOverride];
          end;
        end;
      end
      else begin
        FCellStyleCache[i].RightColor := EX12_AUTO_COLOR;
        FCellStyleCache[i].BottomColor := EX12_AUTO_COLOR;

        FCellStyleCache[i].ForeColor := EX12_AUTO_COLOR;
        FCellStyleCache[i].BackColor := EX12_AUTO_COLOR;

        if FCellStyleCache[i].InMerged < 0 then begin
          FCellStyleCache[i].InMerged := FXLSSheet.MergedCells.CellInAreas(Col + FColumns.LeftCol,Row + FRows.TopRow);
          if FCellStyleCache[i].InMerged >= 0 then
            CacheMergedCells(FCellStyleCache[i].InMerged,EX12_AUTO_COLOR,EX12_AUTO_COLOR);
        end;
      end;
      if Row > 0 then begin
        j := (Row - 1) * Cols + Col;
        if XF <> Nil then begin
          if (FCellStyleCache[j].BottomColor = EX12_AUTO_COLOR) and (SCF.TopStyle = cbsNone) and (FCellStyleCache[i].ForeColor <> EX12_AUTO_COLOR) then begin
            FCellStyleCache[j].BottomColor := SCF.ForeColor;
            FCellStyleCache[j].StyleSW[ccE] := cbsThin;
            FCellStyleCache[j].StyleSE[ccW] := cbsThin;
          end;
        end;
        if TBorderLinePri[Integer(FCellStyleCache[i].StyleNW[ccE])] > TBorderLinePri[Integer(FCellStyleCache[j].StyleSW[ccE])] then begin
          FCellStyleCache[j].StyleSW[ccE] := FCellStyleCache[i].StyleNW[ccE];
          FCellStyleCache[j].StyleSE[ccW] := FCellStyleCache[i].StyleNW[ccE];
          FCellStyleCache[j].BottomColor := Ex12CheckAutoLineColor(SCF.TopStyle,SCF.TopColor);
        end
        else begin
          FCellStyleCache[i].StyleNW[ccE] := FCellStyleCache[j].StyleSW[ccE];
          FCellStyleCache[i].StyleNE[ccW] := FCellStyleCache[j].StyleSW[ccE];
        end;
      end;

      if Col > 0 then begin
        j := Row * Cols + Col - 1;
        if XF <> Nil then begin
          if (FCellStyleCache[j].RightColor = EX12_AUTO_COLOR) and (SCF.LeftStyle = cbsNone) and (FCellStyleCache[i].ForeColor <> EX12_AUTO_COLOR) then begin
            FCellStyleCache[j].RightColor := SCF.ForeColor;
            FCellStyleCache[j].StyleNE[ccS] := cbsThin;
            FCellStyleCache[j].StyleSE[ccN] := cbsThin;
          end;
        end;
        if TBorderLinePri[Integer(FCellStyleCache[i].StyleNW[ccS])] > TBorderLinePri[Integer(FCellStyleCache[j].StyleNE[ccS])] then begin
          FCellStyleCache[j].StyleNE[ccS] := FCellStyleCache[i].StyleNW[ccS];
          FCellStyleCache[j].StyleSE[ccN] := FCellStyleCache[i].StyleNW[ccS];
          FCellStyleCache[j].RightColor := Ex12CheckAutoLineColor(SCF.LeftStyle,SCF.LeftColor);
        end
        else begin
          FCellStyleCache[i].StyleNW[ccS] := FCellStyleCache[j].StyleNE[ccS];
          FCellStyleCache[i].StyleSW[ccN] := FCellStyleCache[j].StyleNE[ccS];
        end;
      end;
    end;
  end;
  for Row := 0 to Rows - 2 do begin
    for Col := 0 to Cols - 2 do begin
      i := Row * Cols + Col;

      j := Row * Cols + Col + 1;
      FCellStyleCache[j].StyleNW[ccW] := FCellStyleCache[i].StyleNE[ccW];
      FCellStyleCache[i].StyleNE[ccE] := FCellStyleCache[j].StyleNW[ccE];
      FCellStyleCache[j].StyleSW[ccW] := FCellStyleCache[i].StyleSE[ccW];
      FCellStyleCache[i].StyleSE[ccE] := FCellStyleCache[j].StyleSW[ccE];

      j := (Row + 1) * Cols + Col;
      FCellStyleCache[j].StyleNW[ccN] := FCellStyleCache[i].StyleSW[ccN];
      FCellStyleCache[i].StyleSW[ccS] := FCellStyleCache[j].StyleNW[ccS];
      FCellStyleCache[j].StyleNE[ccN] := FCellStyleCache[i].StyleSE[ccN];
      FCellStyleCache[i].StyleSE[ccS] := FCellStyleCache[j].StyleNE[ccS];

      Ref := FCellStyleCache[i].Ref;
      if Ref.Data <> Nil then
        XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Ref)]
      else
        XF := Nil;
      if (FCellStyleCache[i].Ref.Data <> Nil) and
         (FCells.CellType(@FCellStyleCache[i].Ref) in [xctString,xctStringFormula]) and
         (FCellStyleCache[i].InMerged < 0) and
         (XF.Alignment.Options = []) and not
         (XF.Alignment.HorizAlignment in [chaFill,chaJustify]) and not
         (XF.Alignment.VertAlignment in [cvaJustify,cvaDistributed]) then begin

        FontRunsCount := 0;
        case FCells.CellType(@FCellStyleCache[i].Ref) of
          xctString: begin
            S := FCells.GetString(@Ref,FR,FontRunsCount);
          end;
          xctStringFormula:
            S := FCells.GetFormulaValString(@Ref);
        end;
        FSkin.SetCellFont(XF.Font);
        if FontRunsCount > 0 then
          W := RichTextWidth(Skin,XF,FR,FontRunsCount,S)
        else
          W := FSkin.GDI.TextWidth(S);
        if XF.Alignment.HorizAlignment = chaCenter then begin
          Inc(W,FSkin.CellLeftMarg + FSkin.CellRightMarg - 2);
          Dec(W,FColumns.ColWidth[Col]);
          if W > 0 then begin
            ExpandLeft(W div 2,Col,i - 1);
            ExpandRight(W div 2,Col,i + 1);
          end;
        end
        else if XF.Alignment.HorizAlignment = chaRight then begin
          Inc(W,FSkin.CellLeftMarg + FSkin.CellRightMarg - 2);
//          Inc(W,FSkin.CellRightMarg + 1);
          Dec(W,FColumns.ColWidth[Col]);
          if W > 0 then
            ExpandLeft(W,Col,i - 1);
        end
        else begin
          Inc(W,FSkin.CellLeftMarg + 1 + XF.Alignment.Indent * FSkin.GDI.TM.AveCharWidth);
          Dec(W,FColumns.ColWidth[Col]);
          if W > 0 then
            ExpandRight(W,Col,i + 1);
        end;
      end;
    end;
  end;
  ApplyCondFmt;
end;

procedure TXLSBookSheet.CacheMergedCells(const AMergedIndex: integer; const AFgColor, ABgColor: longword);
var
  i: integer;
  C,R: integer;
  MergedCell: TCellArea;
begin
  MergedCell := FXLSSheet.MergedCells[AMergedIndex];
  for R := Max(MergedCell.Row1,FRows.TopRow) to Min(MergedCell.Row2,FRows.BottomRow) do begin
    for C := Max(MergedCell.Col1,FColumns.LeftCol) to Min(MergedCell.Col2,FColumns.RightCol) do begin
      i := (R - FRows.TopRow) * FColumns.Count + (C - FColumns.LeftCol);
      FCellStyleCache[i].InMerged := AMergedIndex;
      FCellStyleCache[i].ForeColor := AFgColor;
      FCellStyleCache[i].BackColor := ABgColor;
//              FCellStyleCache[j].Pattern
    end;
  end;
end;

function TXLSBookSheet.GetBottomRow: integer;
begin
  Result := FRows.BottomRow;
end;

function TXLSBookSheet.GetCursorAreaHit(X, Y: integer): TCursorAreaHit;
begin
  if not FOptions.ReadOnly and PtInXYRect(X,Y,FCursorXYMargO) then begin
    if not PtInXYRect(X,Y,FCursorXYMargI) then begin
      case FCursorSizePos of
        cspColumn: begin
          if (X > (FCursorXYMargO.X2 - MARG_CELLCURSORSIZEHIT)) and (Y < (FCursorXYMargO.Y1 + MARG_CELLCURSORSIZEHIT + 5)) then
            Result := cahSize
          else
            Result := cahEdge;
        end;
        cspRow: begin
          if (X < (FCursorXYMargO.X1 + MARG_CELLCURSORSIZEHIT + 5)) and (Y > (FCursorXYMargO.Y2 - MARG_CELLCURSORSIZEHIT)) then
            Result := cahSize
          else
            Result := cahEdge;
        end;
        else begin
          if (X > (FCursorXYMargO.X2 - MARG_CELLCURSORSIZEHIT)) and (Y > (FCursorXYMargO.Y2 - MARG_CELLCURSORSIZEHIT)) then
            Result := cahSize
          else
            Result := cahEdge;
        end;
      end;
    end
    else
      Result := cahInside;
  end
  else
    Result := cahNo;
end;

procedure TXLSBookSheet.ClipToSheet(var C1, R1, C2, R2: integer);
begin
       if R1 < FRows.TopRow    then R1 := FRows.TopRow
  else if R1 > FRows.BottomRow then R1 := FRows.BottomRow;
       if R2 < FRows.TopRow    then R2 := FRows.TopRow
  else if R2 > FRows.BottomRow then R2 := FRows.BottomRow;

       if C1 < FColumns.LeftCol  then C1 := FColumns.LeftCol
  else if C1 > FColumns.RightCol then C1 := FColumns.RightCol;
       if C2 < FColumns.LeftCol  then C2 := FColumns.LeftCol
  else if C2 > FColumns.RightCol then C2 := FColumns.RightCol;
end;

procedure TXLSBookSheet.ClosePivotForm(ASender: TObject; var Action: TCloseAction);
begin
  Action := caFree;

  FFormPivot := Nil;
end;

procedure TXLSBookSheet.MarchingAntsTimer(Sender: TObject; TimerId, Data: integer);
begin
end;

procedure TXLSBookSheet.UpdateSheetArea;
begin
  FSheetArea.Col1 := FColumns.LeftCol;
  FSheetArea.Col2 := FColumns.RightCol;
  FSheetArea.Row1 := FRows.TopRow;
  FSheetArea.Row2 := FRows.BottomRow;
  if FSelectionEdit <> Nil then
    FSelectionEdit.SheetAreaChanged;
  FDrawing.CalcObjects;
end;

procedure TXLSBookSheet.PaintEditRefs(Sender: TObject; C1, R1, C2, R2: integer);
begin
  BeginPaint;
  PaintCells(C1,R1,C2,R2,ctpAll);
  EndPaint;
end;

procedure TXLSBookSheet.ScrollEditRefs(Sender: TObject; Direction: integer);
begin
  case Direction of
    0: FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyLeft));
    1: FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyRight));
    2: FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyUp));
    3: FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyDown));
    else begin
      FSystem.StopTimer(FScrollTimerId);
      FScrollTimerId := -1;
    end;
  end;
end;

procedure TXLSBookSheet.ColumnClick(Sender: TObject; Header: integer; SizeHit: TSizeHitType; Button: TXSSMouseButton; Shift: TXSSShiftState);
begin
  Inc(Header,TXLSBookHeaders(Sender).FirstHeader);
  if xssDouble in Shift  then begin
    FXLSSheet.AutoWidthCol(FColumns.ClickedCol);
    InvalidateAndReload;
  end
  else begin
    case SizeHit of
      shtNone: begin
        if xssCtrl in Shift then begin
          BeginSelection(Header,0);
          AppendSelection(Header,XLS_MAXROW,True);
          SetSelectedHeaders;
        end
        else if xssShift in Shift then begin
          AppendSelection(Header,XLS_MAXROW,True);
          SetSelectedHeaders;
        end
        else if not FOptions.ReadOnly then begin
          ClearSelected(Header,0,True,cspColumn);
  //        FCursorSizePos := cspColumn;
          AppendSelection(Header,XLS_MAXROW,True);
          SetSelectedHeaders;
        end;
      end;
      shtSize: begin
      end;
      shtHidden: begin
      end;
    end;
  end;
end;

procedure TXLSBookSheet.RowClick(Sender: TObject; Header: integer; SizeHit: TSizeHitType; Button: TXSSMouseButton; Shift: TXSSShiftState);
begin
  if xssDouble in Shift  then begin
    FXLSSheet.AutoHeightRow(FRows.ClickedRow);
    InvalidateAndReload;
  end
  else begin
    Inc(Header,TXLSBookHeaders(Sender).FirstHeader);
    case SizeHit of
      shtNone: begin
        if xssCtrl in Shift then begin
          BeginSelection(0,Header);
          AppendSelection(XLS_MAXCOL,Header,True);
          SetSelectedHeaders;
        end
        else if xssShift in Shift then begin
          AppendSelection(XLS_MAXCOL,Header,True);
          SetSelectedHeaders;
        end
        else if not FOptions.ReadOnly then begin
          ClearSelected(0,Header,True,cspRow);
  //        FCursorSizePos := cspRow;
          AppendSelection(XLS_MAXCOL,Header,True);
          SetSelectedHeaders;
        end;
      end;
      shtSize: begin
      end;
      shtHidden: begin
      end;
    end;
  end;
end;

procedure TXLSBookSheet.SetSelectedHeaders;
var
  i: integer;
begin
  FColumns.ClearSelected;
  FRows.ClearSelected;
  for i := 0 to FXLSSheet.SelectedAreas.Count - 1 do begin
    with FXLSSheet.SelectedAreas[i] do begin
      FColumns.SetSelectedState(Col1 - FColumns.LeftCol,Col2 - FColumns.LeftCol,(Row1 = 0) and (Row2 >= MAXSELROW));
      FRows.SetSelectedState(Row1 - FRows.TopRow,Row2 - FRows.TopRow,(Col1 = 0) and (Col2 >= XLS_MAXCOL));
    end;
  end;
  FColumns.Paint;
  FRows.Paint;
end;

procedure TXLSBookSheet.SelColChanged(Sender: TObject; Header,ScrollDirection: integer);
begin
  case ScrollDirection of
     -1: begin
        if FScrollTimerId < 0 then
          FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyLeft));
     end;
      0: begin
        if FScrollTimerId >= 0 then begin
          FSystem.StopTimer(FScrollTimerId);
          FScrollTimerId := -1;
        end;
        Inc(Header,TXLSBookHeaders(Sender).FirstHeader);
        AppendSelection(Header,XLS_MAXROW,True);
        SetSelectedHeaders;
     end;
      1: begin
        if FScrollTimerId < 0 then
          FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,50,Pointer(kyRight));
     end;
  end;
end;

procedure TXLSBookSheet.SelRowChanged(Sender: TObject; Header,ScrollDirection: integer);
begin
  case ScrollDirection of
     -1: begin
        if FScrollTimerId < 0 then
          FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,100,Pointer(kyUp));
     end;
      0: begin
        if FScrollTimerId >= 0 then begin
          FSystem.StopTimer(FScrollTimerId);
          FScrollTimerId := -1;
        end;
        Inc(Header,TXLSBookHeaders(Sender).FirstHeader);
        AppendSelection(XLS_MAXCOL,Header,True);
        SetSelectedHeaders;
     end;
      1: begin
        if FScrollTimerId < 0 then
          FScrollTimerId := FSystem.StartTimer(Self,HandleScrollTimer,100,Pointer(kyDown));
     end;
  end;
end;

procedure TXLSBookSheet.PaintCellText(Col,Row: integer; Ref: TXLSCellItem; Rect: TXYRect; ADXF: TXc12DXF);
var
  S: AxUCString;
  Cl: longword;
  FontRunsCount: integer;
  FR: PXc12FontRunArray;
  XF: TXc12XF;
begin
  FR := Nil;
  XF := FManager.StyleSheet.XFs[FCells.GetStyle(@Ref)];
  FontRunsCount := 0;
  FSkin.SetCellFont(XF.Font);
  FSkin.FontUnderline := XF.Font.Underline;

  if (ADXF <> Nil) and (ADXF.Font <> Nil) then begin
    FSkin.GDI.FontColor := RevRGB(Xc12ColorToRGB(ADXF.Font.Color));
    FSkin.GDI.SetFontStyle(xfsBold in ADXF.Font.Style,xfsItalic in ADXF.Font.Style,ADXF.Font.Underline <> xulNone);
  end
  else
    FSkin.GDI.FontColor := RevRGB(Xc12ColorToRGB(XF.Font.Color));
  case FCells.CellType(@Ref) of
    xctString: begin
      S := FCells.GetString(@Ref,FR,FontRunsCount);
    end
    else begin
      if (soShowFormulas in FXLSSheet.Options) and (FCells.CellType(@Ref) in XLSCellTypeFormulas) then
        S := '=' + FXLSSheet.AsFormula[Col,Row]
      else if (XF.NumFmt.Value = '') and (FCells.CellType(@Ref) in XLSCellTypeNum) then
        S := FitNumAsString((Rect.X2 - Rect.X1) - FSkin.GDI.CharWidth(WideChar('_')),FCells.GetFloat(Col,Row))
{$ifdef AXOLOT_DEBUG}
{$message warn '__TODO__ @Ref can not be used with compressed cells as this is a cached value'}
{$endif}
//        S := FitNumAsString((Rect.X2 - Rect.X1) - FSkin.GDI.CharWidth(WideChar('_')),FCells.GetFloat(@Ref))
      else begin
        S := FXLSSheet.GetFormatString(Ref,@Cl);
        if Cl <> EX12_AUTO_COLOR then
          FSkin.GDI.FontColor := Cl;
      end;
    end;
  end;

  FSkin.CellTextAlign(TPaintCellHorizAlign(XF.Alignment.HorizAlignment),TPaintCellVertAlign(XF.Alignment.VertAlignment),XF.Alignment.Rotation);
  if (soShowFormulas in FXLSSheet.Options) and (FCells.CellType(@Ref) in XLSCellTypeFormulas) then
    FSkin.CellTextRect(Rect,XF.Alignment.Indent,cttText,S)
  else begin
    case FCells.CellType(@Ref) of
      {xctCurrency,}xctFloat,xctFloatFormula:
        FSkin.CellTextRect(Rect,XF.Alignment.Indent,cttNumeric,S);
      xctBoolean,xctBooleanFormula,xctErrorFormula:
        FSkin.CellTextRect(Rect,XF.Alignment.Indent,cttBoolean,S);
      xctError:
        FSkin.CellTextRect(Rect,XF.Alignment.Indent,cttError,S);
      else begin
        if (FontRunsCount > 0){ or (XF.Alignment.HorizAlignment in [chaJustify,chaDistributed])} then
          RichTextRect(Skin,Rect,XF,FR,FontRunsCount,S)
        else begin
          if XF.Alignment.IsWrapText then
            FSkin.CellTextRectWrap(Rect,XF.Alignment.Indent,S)
          else
            FSkin.CellTextRect(Rect,XF.Alignment.Indent,cttText,S)
        end;
      end;
    end;
  end;
end;

procedure TXLSBookSheet.DoReloadSheet(ASender: TObject);
begin
  InvalidateAndReload;
end;

procedure TXLSBookSheet.DoSheetChanged;
begin
  FSheetChanged := False;
  FRows.DefaultHeight := FXLSSheet.DefaultRowHeight;

  FRows.RowsChanged(FXLSSheet.Rows,FXLSSheet.Xc12Sheet.SheetFormatPr.OutlineLevelRow);
  if FIsPrinting then begin
    FColumns.ColumnsChanged(FXLSSheet.Columns,FXLSSheet.Xc12Sheet.SheetFormatPr.OutlineLevelCol,FPrintStdFontWidth);
    SetLeftCol(0);
    SetTopRow(FXLSSheet.TopRow);
  end
  else begin
    FColumns.ColumnsChanged(FXLSSheet.Columns,FXLSSheet.Xc12Sheet.SheetFormatPr.OutlineLevelCol,FManager.StyleSheet.StdFontWidth);
    SetLeftCol(FXLSSheet.LeftCol);
    SetTopRow(FXLSSheet.TopRow);
  end;

  SetVisibleRects;
  SetSize(FX1,FY1,FX2,FY2);
end;

procedure TXLSBookSheet.DoShowPivotForm(APivTbl: TXLSPivotTable);
var
  Pt: TPoint;
begin
  FFormPivot := TfrmPivotTable.Create(FOwner);
  FFormPivot.OnClose := ClosePivotForm;
  Pt.X := FOwner.Left;
  Pt.Y := 0;
  Pt := FOwner.ClientToScreen(Pt);
  FFormPivot.Execute(APivTbl,-(Pt.X + CX2),Pt.Y + CY1,DoReloadSheet);
end;

procedure TXLSBookSheet.KeyPress(Key: AxUCChar);
begin
  if Word(Key) > $0020 then begin
    ShowInplaceEdit(False,Key = '=');
    if FInplaceEditor <> Nil then
      FInplaceEditor.KeyPress(Key);
  end;
end;

procedure TXLSBookSheet.PaintInvalidate;
begin
  FCacheRequest := True;
  FRows.CalcHeaders;
  FColumns.CalcHeaders;
  Paint;
end;

procedure TXLSBookSheet.IECharFormatChanged(ASender: TObject);
begin
  if Assigned(FIECharFmtEvent) then
    FIECharFmtEvent(FInplaceEditor);
end;

procedure TXLSBookSheet.IEPaintCells(Sender: TObject; const Col1, Row1, Col2, Row2: integer);
begin
  PaintCells(Col1,Row1,Col2,Row2,ctpAll);
end;

procedure TXLSBookSheet.Invalidate;
begin
  FCacheRequest := True;
  inherited Invalidate;
end;

procedure TXLSBookSheet.InvalidateAndReload;
begin
  FXLSSheet.CalcDimensions;

  SetSize(FX1,FY1,FX2,FY2);
  TopRow := TopRow;
  CalcVScroll;
  CalcHScroll;
  LeftCol := LeftCol;
  Paint;
  FColumns.Paint;
  FRows.Paint;

  PaintInvalidate;

//  FOwner.Invalidate;
  FOwner.Repaint;
end;

procedure TXLSBookSheet.InvalidateArea(C1, R1, C2, R2: integer);
begin
  C1 := Max(C1,FColumns.LeftCol);
  R1 := Max(R1,FRows.TopRow);
  C2 := Min(C2,FColumns.RightCol);
  R2 := Min(R2,FRows.BottomRow);
  if (C1 >= FColumns.LeftCol) and (C2 <= FColumns.RightCol) and (R1 >= FRows.TopRow) and (R2 <= FRows.BottomRow) then begin
    BeginPaint;

    CacheCells;
    PaintCells(C1,R1,C2,R2,ctpAll);

    EndPaint;
  end;
end;

procedure TXLSBookSheet.InvalidateCell(Col,Row: integer);
var
  CacheIndex: integer;
  px1,py1,px2,py2: integer;
begin
  if (Col >= FColumns.LeftCol) and (Col <= FColumns.RightCol) and (Row >= FRows.TopRow) and (Row <= FRows.BottomRow) then begin
    CacheIndex := (Row - FRows.TopRow) * FColumns.Count + (Col - FColumns.LeftCol);
    FCellStyleCache[CacheIndex].Ref := FCells.FindCell(Col,Row);

    px1 := FColumns.AbsPos[Col];
    py1 := FRows.AbsPos[Row];
    px2 := px1 + FColumns.AbsWidth[Col];
    py2 := py1 + FRows.AbsHeight[Row];
    FSkin.GDI.InvalidateRect(px1,py1,px2 - 1,py2 - 1);

    PaintCells(Col,Row,Col,Row,ctpAll);
    FLayers.ReleaseHandle;
  end;
end;

{ TCellPaintData }

procedure TCellPaintData.Calc(CSR: TCellStyleRec; Left, Top, Right, Bottom: TXc12CellBorderStyle);
const
  BorderWidthtsSub: array[cbsNone..cbsSlantedDashDot] of integer = (0,0,1,0,0,1,1,0,1,0,1,0,1,1);
  BorderWidthtsAdd: array[cbsNone..cbsSlantedDashDot] of integer = (0,0,0,0,0,1,1,0,0,0,0,0,0,0);
var
  X1,Y1,X2,Y2: integer;

procedure CalcCorner(Corner: TCellCorner; X,Y: integer);
var
  Directions: integer;
begin
  Directions := (Integer(Corner[ccN] <> cbsNone) shl 3) + (Integer(Corner[ccS] <> cbsNone) shl 2) + (Integer(Corner[ccW] <> cbsNone) shl 1) + (Integer(Corner[ccE] <> cbsNone) shl 0);
              // NSWE
  case Directions of
    $0: begin // ----
      X1 := X - 1;
      X2 := X + 1;
      Y1 := Y;
      Y2 := Y + 1;
    end;
    $1: begin // ---E
      X1 := X - 1;
      X2 := X;
      Y1 := Y - BorderWidthtsSub[Corner[ccE]] - 1;
      Y2 := Y + BorderWidthtsAdd[Corner[ccE]] + 1;
    end;
    $2: begin // --W-
      X1 := X;
      X2 := X + 1;
      Y1 := Y - BorderWidthtsSub[Corner[ccW]] - 1;
      Y2 := Y + BorderWidthtsAdd[Corner[ccW]] + 1;
    end;
    $3: begin // --WE
      X1 := X;
      X2 := X + 1;
      Y1 := Y - BorderWidthtsSub[Corner[ccW]] - 1;
      Y2 := Y + BorderWidthtsAdd[Corner[ccW]] + 1;
    end;
    $4: begin // --S--
      X1 := X - BorderWidthtsSub[Corner[ccS]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccS]] + 1;
      Y1 := Y - 1;
      Y2 := Y;
    end;
    $5: begin // -S-E
      X1 := X - BorderWidthtsSub[Corner[ccS]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccS]] + 1;
      Y1 := Y - BorderWidthtsSub[Corner[ccE]] - 1;
      Y2 := Y - BorderWidthtsSub[Corner[ccE]];
    end;
    $6: begin // -SW-
      X1 := X - BorderWidthtsSub[Corner[ccS]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccS]] + 1;
      Y1 := Y - BorderWidthtsSub[Corner[ccW]] - 1;
      Y2 := Y - BorderWidthtsSub[Corner[ccW]];
    end;
    $7: begin // -SWE
      if csfHorizOverride in FCellStyle.Flags then begin
        X1 := X;
        X2 := X + 1;
        Y1 := Y - BorderWidthtsSub[Corner[ccE]];
        Y2 := Y + BorderWidthtsAdd[Corner[ccE]] + 1;
      end
      else begin
        X1 := X - BorderWidthtsSub[Corner[ccS]] - 1;
        X2 := X + BorderWidthtsAdd[Corner[ccS]] + 1;
        Y1 := Y - BorderWidthtsSub[StyleMax(Corner[ccW],Corner[ccE])] - 1;
        Y2 := Y1 + 1;
      end;
    end;
    $8: begin // N---
      X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
      Y1 := Y;
      Y2 := Y + 1;
    end;
    $9: begin // N--E
      X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
      Y1 := Y + BorderWidthtsAdd[Corner[ccE]];
      Y2 := Y + BorderWidthtsAdd[Corner[ccE]] + 1;
    end;
    $A: begin // N-W-
      X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
      Y1 := Y + BorderWidthtsAdd[Corner[ccW]];
      Y2 := Y + BorderWidthtsAdd[Corner[ccW]] + 1;
    end;
    $B: begin // N-WE
      if csfHorizOverride in FCellStyle.Flags then begin
        X1 := X;
        X2 := X + 1;
        Y1 := Y - BorderWidthtsSub[Corner[ccE]];
        Y2 := Y + BorderWidthtsAdd[Corner[ccE]] + 1;
      end
      else begin
        X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
        X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
        Y1 := Y + BorderWidthtsAdd[StyleMax(Corner[ccW],Corner[ccE])];
        Y2 := Y1 + 1;
      end;
    end;
    $C: begin // NS--
      X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
      Y1 := Y;
      Y2 := Y + 1;
    end;
    $D: begin // NS-E
      X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
      Y1 := Y - BorderWidthtsSub[Corner[ccE]] - 1;
      Y2 := Y - BorderWidthtsSub[Corner[ccE]];
    end;
    $E: begin // NSW-
      X1 := X - BorderWidthtsSub[Corner[ccN]] - 1;
      X2 := X + BorderWidthtsAdd[Corner[ccN]] + 1;
      Y1 := Y - BorderWidthtsSub[Corner[ccW]] - 1;
      Y2 := Y - BorderWidthtsSub[Corner[ccW]];
    end;
    $F: begin // NSWE
      if csfHorizOverride in FCellStyle.Flags then begin
        X1 := X;
        X2 := X + 1;
        Y1 := Y - BorderWidthtsSub[Corner[ccE]];
        Y2 := Y + BorderWidthtsAdd[Corner[ccE]] + 1;
      end
      else begin
        X1 := X - BorderWidthtsSub[StyleMax(Corner[ccN],Corner[ccS])] - 1;
        X2 := X + BorderWidthtsAdd[StyleMax(Corner[ccN],Corner[ccS])] + 1;
        Y1 := Y - BorderWidthtsSub[StyleMax(Corner[ccW],Corner[ccE])] - 1;
        Y2 := Y1 + 1;
      end;
    end;
  end;
end;

begin
  CalcCorner(CSR.StyleNE,FCellRect.X2,FCellRect.Y1);
  FRightPos.X1 := FCellRect.X2 - BorderWidthtsSub[CSR.StyleSE[ccN]];
  FRightPos.Y1 := Y2;
  // Set border line priority to vertical lines if there is no horizontal lines.
  // This is not perfect...
  if (CSR.RightColor <> EX12_AUTO_COLOR) and (CSR.StyleNE[ccS] = CSR.StyleSE[ccS]) and (CSR.StyleNE[ccS] <> cbsNone) and ((CSR.StyleSW[ccE] = cbsNone) and (CSR.StyleSE[ccE] = cbsNone)) then
    Dec(FRightPos.Y1);

  CalcCorner(CSR.StyleSE,FCellRect.X2,FCellRect.Y2);
  FRightPos.X2 := FCellRect.X2 + BorderWidthtsAdd[CSR.StyleSE[ccN]];
  FRightPos.Y2 := Y1;

  FBottomPos.X2 := X1;
  FBottomPos.Y2 := FCellRect.Y2 + BorderWidthtsAdd[CSR.StyleSE[ccW]];
  CalcCorner(CSR.StyleSW,FCellRect.X1,FCellRect.Y2);
  FBottomPos.X1 := X2;
  FBottomPos.Y1 := FCellRect.Y2 - BorderWidthtsSub[CSR.StyleSE[ccW]];

  case CellBorderWidthts[Left] of
    3: Inc(FPaintRect.X1);
  end;
  case CellBorderWidthts[Top] of
    3: Inc(FPaintRect.Y1);
  end;
  case CellBorderWidthts[Right] of
    2: Dec(FPaintRect.X2);
    3: Dec(FPaintRect.X2);
  end;
  case CellBorderWidthts[Bottom] of
    2: Dec(FPaintRect.Y2);
    3: Dec(FPaintRect.Y2);
  end;
end;

procedure TCellPaintData.CalcPaintRect(Left, Top, Right, Bottom: TXc12CellBorderStyle);
begin
  case CellBorderWidthts[Left] of
    3: Inc(FPaintRect.X1);
  end;
  case CellBorderWidthts[Top] of
    3: Inc(FPaintRect.Y1);
  end;
  case CellBorderWidthts[Right] of
    2: Dec(FPaintRect.X2);
    3: Dec(FPaintRect.X2);
  end;
  case CellBorderWidthts[Bottom] of
    2: Dec(FPaintRect.Y2);
    3: Dec(FPaintRect.Y2);
  end;
end;

procedure TCellPaintData.Clear(X1, Y1, X2, Y2: integer; CellStyle: TCellStyleRec);
begin
  FCellStyle := CellStyle;

  FCellRect.X1 := X1 - 1;
  FCellRect.Y1 := Y1 - 1;
  FCellRect.X2 := X2 + 1;
  FCellRect.Y2 := Y2 + 1;

  FPaintRect.X1 := X1;
  FPaintRect.Y1 := Y1;
  FPaintRect.X2 := X2;
  FPaintRect.Y2 := Y2;
end;

procedure TCellPaintData.PaintBorderTop(Skin: TXLSBookSkin; AColor: longword; AStyle: TXc12CellBorderStyle);
begin
  if FIsPrinting then begin
    if FLineCacheCount < Length(FVertLCache) then begin
      FHorizLCache[FLineCacheCount].Pos.X1 := FBottomPos.X1;
      FHorizLCache[FLineCacheCount].Pos.Y1 := 0;
      FHorizLCache[FLineCacheCount].Pos.X2 := FBottomPos.X2;
      FHorizLCache[FLineCacheCount].Pos.Y2 := 0;

      FHorizLCache[FLineCacheCount].PenColor := _Temp_XCColorToRGB(AColor,Skin.Colors.SheetBkg);
      FHorizLCache[FLineCacheCount].BackColor := _Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg);
      FHorizLCache[FLineCacheCount].Style := TPaintLineStyle(AStyle);
      Inc(FLineCacheCount);
    end;
  end;
end;

procedure TCellPaintData.PaintBorderLeft(Skin: TXLSBookSkin; AColor: longword; AStyle: TXc12CellBorderStyle);
begin
  if FIsPrinting then begin
    if FLineCacheCount < Length(FVertLCache) then begin
      FVertLCache[FLineCacheCount].Pos.X1 := 0;
      FVertLCache[FLineCacheCount].Pos.Y1 := FRightPos.Y1;
      FVertLCache[FLineCacheCount].Pos.X2 := 0;
      FVertLCache[FLineCacheCount].Pos.Y2 := FRightPos.Y2;

      FVertLCache[FLineCacheCount].PenColor := _Temp_XCColorToRGB(AColor,Skin.Colors.SheetBkg);
      FVertLCache[FLineCacheCount].BackColor := _Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg);
      FVertLCache[FLineCacheCount].Style := TPaintLineStyle(AStyle);
      Inc(FLineCacheCount);
    end;
  end;
end;

procedure TCellPaintData.PaintBorders(Skin: TXLSBookSkin; Right, Bottom: longword);
begin
  if FIsPrinting then begin
    if FLineCacheCount < Length(FVertLCache) then begin
      FVertLCache[FLineCacheCount].Pos := FRightPos;
      FVertLCache[FLineCacheCount].PenColor := Right;
      FVertLCache[FLineCacheCount].BackColor := _Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg);
      FVertLCache[FLineCacheCount].Style := TPaintLineStyle(FCellStyle.StyleSE[ccN]);

      FHorizLCache[FLineCacheCount].Pos := FBottomPos;
      FHorizLCache[FLineCacheCount].PenColor := Bottom;
      FHorizLCache[FLineCacheCount].BackColor := _Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg);
      FHorizLCache[FLineCacheCount].Style := TPaintLineStyle(FCellStyle.StyleSE[ccW]);

      Inc(FLineCacheCount);
    end;

//    if FLineCacheCount > Length(FRightLCache) then
//      raise XLSRWException.Create('Line cache overlow');
  end
  else begin
    Skin.GDI.LineStyled(FRightPos.X1,FRightPos.Y1,FRightPos.X2,FRightPos.Y2 + 1,Right,_Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg),TPaintLineStyle(FCellStyle.StyleSE[ccN]),False);
    Skin.GDI.LineStyled(FBottomPos.X1,FBottomPos.Y1,FBottomPos.X2 + 1,FBottomPos.Y2,Bottom,_Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg),TPaintLineStyle(FCellStyle.StyleSE[ccW]),True);
  end;
{
  Skin.GDI.LineStyled(FRightPos.X1,FRightPos.Y1,FRightPos.X2,FRightPos.Y2,Right,FCellStyle.ForeColor,TPaintLineStyle(FCellStyle.StyleSE[ccN]),False);
  Skin.GDI.LineStyled(FBottomPos.X1,FBottomPos.Y1,FBottomPos.X2 + 1,FBottomPos.Y2,Bottom,FCellStyle.ForeColor,TPaintLineStyle(FCellStyle.StyleSE[ccW]),True);
}
end;

procedure TCellPaintData.PaintBordersTopLeft(Skin: TXLSBookSkin; ATop, ALeft: TXc12BorderPr);
begin
  Skin.GDI.LineStyled(FBottomPos.X1,FRightPos.Y1,FRightPos.X2,FRightPos.Y1,_Temp_XCColorToRGB(ATop.Color.ARGB,Skin.Colors.SheetBkg),_Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg),TPaintLineStyle(ATop.Style),True);
  Skin.GDI.LineStyled(FBottomPos.X1,FRightPos.Y1,FBottomPos.X1,FRightPos.Y2,_Temp_XCColorToRGB(ALeft.Color.ARGB,Skin.Colors.SheetBkg),_Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg),TPaintLineStyle(ALeft.Style),False);
end;

procedure TCellPaintData.PaintCell(Skin: TXLSBookSkin);
var
  R: TXYRect;
begin
  Skin.GDI.BrushColor := _Temp_XCColorToRGB(FCellStyle.ForeColor,Skin.Colors.SheetBkg);
  R := FPaintRect;
  Inc(R.X2);
  Inc(R.Y2);
//  Inc(R.X1,10);
//  Inc(R.Y1,10);
//  Dec(R.X2,10);
//  Dec(R.Y2,10);
  Skin.GDI.FillRect(R);
end;

procedure TCellPaintData.PaintDataBar(Skin: TXLSBookSkin; const APercent: double; const AColor: longword);
var
  Marg: integer;
begin
  Marg := 2;
  Skin.GDI.GradientFillRect(FPaintRect.X1,FPaintRect.Y1 + Marg,FPaintRect.X1 + Round((FPaintRect.X2 - FPaintRect.X1) * APercent * 0.90),FPaintRect.Y2 - Marg,RevRGB(AColor),$FFFFFF,True);
end;

procedure TCellPaintData.PaintDiagonal(ASkin: TXLSBookSkin);
begin
  case FCellStyle.Diagonal of
    dlUp  : ASkin.GDI.LineStyled(FCellRect.X1,FCellRect.Y2,FCellRect.X2,FCellRect.Y1,FCellStyle.DiagColor,_Temp_XCColorToRGB(FCellStyle.ForeColor,ASkin.Colors.SheetBkg),TPaintLineStyle(FCellStyle.DiagStyle),False,True);
    dlDown: ASkin.GDI.LineStyled(FCellRect.X1,FCellRect.Y1,FCellRect.X2,FCellRect.Y2,FCellStyle.DiagColor,_Temp_XCColorToRGB(FCellStyle.ForeColor,ASkin.Colors.SheetBkg),TPaintLineStyle(FCellStyle.DiagStyle),False,True);
    dlBoth: begin
      ASkin.GDI.LineStyled(FCellRect.X1,FCellRect.Y1,FCellRect.X2,FCellRect.Y2,FCellStyle.DiagColor,_Temp_XCColorToRGB(FCellStyle.ForeColor,ASkin.Colors.SheetBkg),TPaintLineStyle(FCellStyle.DiagStyle),False,True);
      ASkin.GDI.LineStyled(FCellRect.X1,FCellRect.Y2,FCellRect.X2,FCellRect.Y1,FCellStyle.DiagColor,_Temp_XCColorToRGB(FCellStyle.ForeColor,ASkin.Colors.SheetBkg),TPaintLineStyle(FCellStyle.DiagStyle),False,True);
    end;
  end;
end;

procedure TCellPaintData.PaintIcon(Skin: TXLSBookSkin; const AIconSet: TXc12IconSetType; const AIndex: integer; const ACentre: boolean);
begin
  if ACentre then
    Skin.GDI.PaintCondFmtIcon(AIconSet,AIndex,FPaintRect.X1 + ((FPaintRect.X2 - FPaintRect.X1) div 2) - (CONDFMT_ICON_SIZE div 2),FPaintRect.Y1)
  else
    Skin.GDI.PaintCondFmtIcon(AIconSet,AIndex,FPaintRect.X1 + 2,FPaintRect.Y2 - CONDFMT_ICON_SIZE - 1);
end;

procedure TCellPaintData.PaintLineCache(ASkin: TXLSBookSkin);
var
  i: integer;
begin
  for i := 0 to FLineCacheCount - 1 do begin
    if FVertLCache[i].PenColor <> FVertLCache[i].BackColor then
      ASkin.GDI.LineStyled(FVertLCache[i].Pos.X1,FVertLCache[i].Pos.Y1,FVertLCache[i].Pos.X2,FVertLCache[i].Pos.Y2 + 1,FVertLCache[i].PenColor,FVertLCache[i].BackColor,FVertLCache[i].Style,False);
    if FHorizLCache[i].PenColor <> FHorizLCache[i].BackColor then
      ASkin.GDI.LineStyled(FHorizLCache[i].Pos.X1,FHorizLCache[i].Pos.Y1,FHorizLCache[i].Pos.X2 + 1,FHorizLCache[i].Pos.Y2,FHorizLCache[i].PenColor,FHorizLCache[i].BackColor,FHorizLCache[i].Style,True);
  end;
end;

procedure TCellPaintData.SetCacheSize(ACols,ARows: integer);
begin
  SetLength(FVertLCache,(ACols + 1) * (ARows + 1) + ARows + 2);
  SetLength(FHorizLCache,(ACols + 1) * (ARows + 1) + ACols + 2);
end;

procedure TCellPaintData.SetIsPrinting(const Value: boolean);
begin
  FIsPrinting := Value;
  FLineCacheCount := 0;
end;

end.

