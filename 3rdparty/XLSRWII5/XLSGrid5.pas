unit XLSGrid5;

{
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

uses Classes, SysUtils, Math, vcl.Controls, vcl.Grids, vcl.Tabs, Windows,
{$ifdef DELPHI_XE5_OR_LATER}
     UITypes, vcl.Graphics,
{$else}
     Graphics,
{$endif}
     Xc12Utils5, Xc12DataStylesheet5, xpgParseDrawingCommon,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSColumn5, XLSRichPainter5, XLSDrawing5,
     XLSSheetData5, XLSReadWriteII5, XLSTools5;

type TXLSTabSet = class(TTabSet)
protected
     // Do nothing when WriteState is called.
     // The tab set must be outside of the grid window and the only way I found
     // to do so is to set the parent of the tab set to the same parent as the
     // grid. This has the side effect that the tab set published properties
     // are streamed to the form. This will create a non existing tab set on
     // the form. By overriding WriteState and doing nothing when it is called,
     // this is prevented.
     procedure WriteState(Writer: TWriter); override;
public
     end;

type TXLSReadWriteIIGrid5 = class(TXLSReadWriteII5)
protected
     procedure AfterRead; override;
public
     constructor Create(AOwner: TComponent); override;
     end;

type TXLSGrid = class(TCustomDrawGrid)
protected
     FTabs          : TXLSTabSet;
     FXLS           : TXLSReadWriteIIGrid5;
     FSheet         : TXLSWorksheet;
     FHeaderColor   : TColor;
     FGridlineColor : TColor;
     FNoCellPaintCnt: integer;
     FEditText      : AxUCString;
     FEditCol       : integer;
     FEditRow       : integer;

     procedure TopLeftChanged; override;
     procedure SetParent(AParent: TWinControl); override;
     procedure DrawCell(ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState); override;
     procedure Paint; override;
     function  CanEditAcceptKey(Key: Char): Boolean; override;
     function  GetEditText(ACol, ARow: Longint): string; override;
     procedure SetEditText(ACol, ARow: Longint; const Value: string); override;
     function  SelectCell(ACol, ARow: Longint): Boolean; override;

     procedure TabChanged(Sender: TObject; NewTab: Integer; var AllowChange: Boolean);

     function  RowPixelHeight(AHeight: integer): integer;

     procedure GetShapeRect(AShape: TXLSDrwTwoPosShape; out ARect: TRect);

     function  IsCellVisible(ACol,ARow: integer): boolean;

     procedure ColsChanged;
     procedure RowsChanged;

     function  FindXF(ACol,ARow: integer): TXc12XF;

     function  DoMerged(var ARect: TRect; ACol,ARow,AIndex: integer): boolean;
     procedure DoImage(AImage: TXLSDrawingImage);
     procedure DoTextBox(ATextBox: TXLSDrawingTextBox);
     procedure DoGraphics;

     procedure PaintHeader(ARect: TRect; AText: AxUCString);
     procedure PaintBorder(X1,Y1,X2,Y2: integer; ABorder: TXc12BorderPr);
     procedure PaintBorders(ARect: TRect; ABorder: TXc12Border);
     procedure PaintDefaultCell(ARect: TRect; AXF: TXc12XF; AFillCell: boolean);
     procedure PaintCell(ARect: TRect; AXF,AXFLeft,AXFTop: TXc12XF; AFillCell: boolean);
     procedure ExpandWidthLeft(var ARect: TRect; ACol,ARow,AWidth: integer);
     procedure ExpandWidthRight(var ARect: TRect; ACol,ARow,AWidth: integer);
     procedure PaintText(ARect: TRect; AXF: TXc12XF; ACol,ARow: integer; AText: AxUCString; ACellType: TXLSCellType; ACanExpand: boolean);
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure XLSChanged(AUpdateTabs: boolean = True);

     procedure SetCellText(ACol, ARow: integer; AText: AxUCString);

     procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override;

     property Sheet        : TXLSWorksheet read FSheet;
published
     property XLS          : TXLSReadWriteIIGrid5 read FXLS;
     property HeaderColor  : TColor read FHeaderColor write FHeaderColor;
     property GridlineColor: TColor read FGridlineColor write FGridlineColor;

     property Align;
     property Anchors;
     property BevelEdges;
     property BevelInner;
     property BevelKind;
     property BevelOuter;
     property BevelWidth;
     property BiDiMode;
     property BorderStyle;
     property Color;
     property ColCount;
     property Constraints;
     property Ctl3D;
     property DefaultColWidth;
     property DefaultRowHeight;
     property DoubleBuffered;
     property DragCursor;
     property DragKind;
     property DragMode;
//     property DrawingStyle;
     property Enabled;
     property FixedColor;
     property RowCount;
//     property GradientEndColor;
//     property GradientStartColor;
     property GridLineWidth;
     property Options;
     property ParentBiDiMode;
     property ParentColor;
     property ParentCtl3D;
//     property ParentDoubleBuffered;
     property ParentFont;
     property ParentShowHint;
     property PopupMenu;
     property ScrollBars;
     property ShowHint;
     property TabOrder;
//     property Touch;
     property Visible;
//     property StyleElements;
     property VisibleColCount;
     property VisibleRowCount;
     property OnClick;
     property OnColumnMoved;
     property OnContextPopup;
     property OnDblClick;
     property OnDragDrop;
     property OnDragOver;
     property OnEndDock;
     property OnEndDrag;
     property OnEnter;
     property OnExit;
//     property OnFixedCellClick;
//     property OnGesture;
     property OnGetEditMask;
     property OnGetEditText;
     property OnKeyDown;
     property OnKeyPress;
     property OnKeyUp;
//     property OnMouseActivate;
     property OnMouseDown;
//     property OnMouseEnter;
//     property OnMouseLeave;
     property OnMouseMove;
     property OnMouseUp;
     property OnMouseWheelDown;
     property OnMouseWheelUp;
     property OnSelectCell;
     property OnSetEditText;
     property OnStartDock;
     property OnStartDrag;
     property OnTopLeftChanged;
     end;

implementation

{ TXLSGrid }

function TXLSGrid.CanEditAcceptKey(Key: Char): Boolean;
begin
  Result := True;

  if Key = #13 then begin
    SetCellText(FEditCol,FEditRow,FEditText);

    FEditText := '';
  end;
end;

procedure TXLSGrid.ColsChanged;
var
  c: integer;
  W: integer;
  n: integer;
begin
  DefaultColWidth := FSheet.Columns.DefaultPixelWidth - 1;

  ColCount := Max(32,FSheet.LastCol + 2);

  if Parent <> Nil then begin
    FXLS.Manager.StyleSheet.Fonts[0].CopyToTFont(Canvas.Font);
    W := Canvas.TextWidth('8');
    n := Max(Trunc(Log10(TopRow + VisibleRowCount) + 1),2);

    ColWidths[0] := W * (n + 1);
  end;

  for c := 0 to FSheet.LastCol do
    ColWidths[c + 1] := FSheet.Columns[c].PixelWidth - 1;
end;

constructor TXLSGrid.Create(AOwner: TComponent);
begin
  inherited;

  FXLS := TXLSReadWriteIIGrid5.Create(Self);
  FSheet := FXLS[FXLS.SelectedTab];

  ColCount := 64;
  RowCount := 255;
  FixedCols := 1;
  FixedRows := 1;
  Options := Options - [goHorzLine];
  Options := Options - [goVertLine];
  Options := Options + [goColSizing];
  Options := Options + [goRowSizing];
  Options := Options + [goEditing];
  DefaultDrawing := False;
  VirtualView := True;
//  DoubleBuffered := True;

  FHeaderColor := $F7ECE4;
  FGridlineColor := $E5D7D0;
end;

destructor TXLSGrid.Destroy;
begin
  FXLS.Free;

  inherited;
end;

procedure TXLSGrid.DoGraphics;
var
  i: integer;
  h: longword;
begin
  h := Windows.CreateRectRgn(ColWidths[0] + 1,RowHeights[0] + 1,Width,Height);
  Windows.SelectClipRgn(Canvas.Handle,h);

  for i := 0 to FSheet.Drawing.Images.Count - 1 do
    DoImage(FSheet.Drawing.Images[i]);
  for i := 0 to FSheet.Drawing.TextBoxes.Count - 1 do
    DoTextBox(FSheet.Drawing.TextBoxes[i]);

  Windows.SelectClipRgn(Canvas.Handle,0);
  Windows.DeleteObject(h);
end;

procedure TXLSGrid.DoImage(AImage: TXLSDrawingImage);
var
  Rect: TRect;
begin
  AImage.CacheBitmap;

  if IsCellVisible(AImage.Col1,AImage.Row1) or IsCellVisible(AImage.Col2 + 1,AImage.Row2 + 1) then begin
    GetShapeRect(AImage,Rect);
    Canvas.StretchDraw(Rect,AImage.Bitmap);
  end;
end;

function TXLSGrid.DoMerged(var ARect: TRect; ACol,ARow,AIndex: integer): boolean;
var
  i  : integer;
  Opt: TXLSSheetOptions;
  XF : TXc12XF;
begin
  XF := FindXF(ACol,ARow);

  if (FSheet.MergedCells[AIndex].Col1 = ACol) and (FSheet.MergedCells[AIndex].Row1 = ARow) then begin
    for i := ACol + 1 to FSheet.MergedCells[AIndex].Col2 do
      Inc(ARect.Right,ColWidths[i + 1] + 1);
    for i := ARow + 1 to FSheet.MergedCells[AIndex].Row2 do
      Inc(ARect.Bottom,RowHeights[i + 1] + 1);

    if XF <> Nil then
      Canvas.Brush.Color := RevRGB(Xc12ColorToRGB(XF.Fill.FgColor))
    else
      Canvas.Brush.Color := RevRGB(Xc12ColorToRGB(FXLS.Manager.StyleSheet.XFs[XLS_STYLE_DEFAULT_XF].Fill.FgColor));
    Canvas.Pen.Color := Canvas.Brush.Color;
    Canvas.Rectangle(ARect);

    Result := True;
  end
  else
    Result := False;

  Opt := FSheet.Options;
  FSheet.Options := FSheet.Options - [soGridlines];
  PaintCell(ARect,XF,FindXF(ACol - 1,ARow),FindXF(ACol,ARow - 1),False);
  FSheet.Options := Opt;
end;

procedure TXLSGrid.DoTextBox(ATextBox: TXLSDrawingTextBox);
var
  i,j    : integer;
  Rect   : Trect;
  TxtBody: TCT_TextBody;
  Para   : TCT_TextParagraph;
  Run    : TCT_RegularTextRun;
  Font   : TXc12Font;
  Text   : TXLSFormattedText;
begin
  if IsCellVisible(ATextBox.Col1,ATextBox.Row1) or IsCellVisible(ATextBox.Col2 + 1,ATextBox.Row2 + 1) then begin
    GetShapeRect(ATextBox,Rect);

    Canvas.Pen.Color := clBlack;
    Canvas.Brush.Color := clWhite;
    Canvas.Rectangle(Rect);

    Text := TXLSFormattedText.Create;

    TxtBody := ATextBox.Anchor.Objects.Sp.TxBody;

    for i := 0 to TxtBody.Paras.Count - 1 do begin
      Para := TxtBody.Paras[i];
      for j := 0 to Para.TextRuns.Count - 1 do begin
        Font := TXc12Font.Create(Nil);
        Font.Assign(FXLS.Manager.StyleSheet.Fonts.DefaultFont);

        Run := Para.TextRuns[j].Run;
        if Run.RPr.Available then begin
          Font.Assign(Run.RPr);
          if Run.RPr.FillProperties.SolidFill <> Nil then
            Font.PColor.ARGB := RevRGB(SolidFillToRGB(Run.RPr.FillProperties.SolidFill));
        end;

        Text.Add(Run.T,Font);
      end;
    end;

    Windows.SetBkMode(Canvas.Handle,TRANSPARENT);
    RichTextRect(Canvas,Rect.Left + 6,Rect.Top + 6,Rect.Right - 6,Rect.Bottom - 6,Text);
  end;
end;

procedure TXLSGrid.DrawCell(ACol, ARow: Integer; ARect: TRect; AState: TGridDrawState);
var
  i  : integer;
  S  : AxUCString;
  XF : TXc12XF;
  XF2: TXc12XF;
begin
  Dec(ARow);
  Dec(ACol);

  if (ACol < 0) and (ARow < 0) then
    PaintHeader(ARect,'')
  else if ACol < 0 then begin
    PaintHeader(ARect,IntToStr(ARow + 1));
  end
  else if ARow < 0 then begin
    PaintHeader(ARect,ColToRefStr(ACol,False));
  end
  else begin
    S := FSheet.AsFmtString[ACol,ARow];
    XF := FindXF(ACol,ARow);
    XF2 := XF;
    if XF2 = Nil then
      XF2 := FXLS.Manager.StyleSheet.XFs[XLS_STYLE_DEFAULT_XF];

    i := FSheet.MergedCells.FindArea(ACol,ARow);
    if i >= 0 then begin
      if DoMerged(ARect,ACol,ARow,i) then
        PaintText(ARect,XF2,ACol,ARow,S,FSheet.CellType[ACol,ARow],False);
    end
    else begin
      PaintCell(ARect,XF,FindXF(ACol - 1,ARow),FindXF(ACol,ARow - 1),True);
      if S <> '' then
        PaintText(ARect,XF2,ACol,ARow,S,FSheet.CellType[ACol,ARow],True);
    end;
  end;

  if gdFocused in AState then begin
    Inc(ARect.Left);
    Inc(ARect.Top);
    Canvas.Pen.Color := clBlack;
    Canvas.Pen.Width := 2;
    Canvas.Brush.Style := bsClear;
    Canvas.Rectangle(ARect);
    Canvas.Brush.Style := bsSolid;
    Canvas.Pen.Width := 1;
  end;


//  gdSelected, gdFocused, gdFixed, gdRowSelected, gdHotTrack, gdPressed
end;

procedure TXLSGrid.ExpandWidthLeft(var ARect: TRect; ACol, ARow, AWidth: integer);
begin
  while (AWidth > 0) and (FSheet.AsString[ACol,ARow] = '') and (ACol >= (LeftCol - 1)) do begin
    Dec(ARect.Left,ColWidths[ACol + 1] + 1);
    Dec(AWidth,ColWidths[ACol + 1] + 1);

    Dec(ACol);
  end;
end;

procedure TXLSGrid.ExpandWidthRight(var ARect: TRect; ACol, ARow, AWidth: integer);
var
  n: integer;
  R: TRect;
begin
  FNoCellPaintCnt := 0;

  n := 0;
  R := ARect;
  while (AWidth > 0) and (FSheet.AsString[ACol,ARow] = '') and (ACol <= (LeftCol + VisibleColCount)) do begin
    Inc(ARect.Right,ColWidths[ACol + 1] + 1);
    Dec(AWidth,ColWidths[ACol + 1] + 1);
    Inc(n);

    R := CellRect(ACol + 1,ARow + 1);
    PaintCell(R,FindXF(ACol,ARow),FindXF(ACol - 1,ARow),FindXF(ACol,ARow - 1),True);

    Inc(ACol);
  end;
  // Can't assign FNoCellPaintCnt in loop as PaintCell will decrement it.
  FNoCellPaintCnt := n;
end;

function TXLSGrid.FindXF(ACol, ARow: integer): TXc12XF;
var
  i : integer;
  CI: TXLSCellItem;
  C : TXLSColumn;
begin
  i := XLS_STYLE_DEFAULT_XF;
  if (ACol >= 0) and (ARow >= 0) and FSheet.MMUCells.FindCell(ACol,ARow,CI) then
    i := FSheet.MMUCells.GetStyle(@CI);

  if (i = XLS_STYLE_DEFAULT_XF) and (ARow >= 0) then
    i := FSheet.MMUCells.GetRowStyle(ARow);

  if (i = XLS_STYLE_DEFAULT_XF) and (ACol >= 0) then begin
    C := FSheet.Columns[ACol];
    i := C.XF.Index;
  end;

  if i <> XLS_STYLE_DEFAULT_XF  then
    Result := FXLS.Manager.StyleSheet.XFs[i]
  else
    Result := Nil;
end;

function TXLSGrid.GetEditText(ACol, ARow: Integer): string;
begin
  if FSheet.CellType[ACol,ARow] in XLSCellTypeFormulas then
    Result := FSheet.AsFormula[ACol - 1,ARow - 1]
  else
    Result := FSheet.AsString[ACol - 1,ARow - 1];
end;

procedure TXLSGrid.GetShapeRect(AShape: TXLSDrwTwoPosShape; out ARect: TRect);
var
  c,r: integer;
begin
  ARect.Left := ColWidths[0];
  ARect.Right := ColWidths[0];
  ARect.Top := RowHeights[0];
  ARect.Bottom := RowHeights[0];

  c := AShape.Col1 - LeftCol + 1;

  if c < 0 then begin
    while c < 0 do begin
      Dec(ARect.Left,ColWidths[LeftCol + c] + 1);
      Inc(c);
    end;
  end
  else begin
    while c > 0 do begin
      Inc(ARect.Left,ColWidths[LeftCol + c] + 1);
      Dec(c);
    end;
  end;
  Inc(ARect.Left,Round((ColWidths[AShape.Col1] + 1) * AShape.Col1Offs));

  c := AShape.Col2 - LeftCol + 1;

  if c < 0 then begin
    while c < 0 do begin
      Dec(ARect.Right,ColWidths[LeftCol + c] + 1);
      Inc(c);
    end;
  end
  else begin
    while c > 0 do begin
      Inc(ARect.Right,ColWidths[LeftCol + c] + 1);
      Dec(c);
    end;
  end;
  Inc(ARect.Right,Round((ColWidths[AShape.Col2] + 1) * AShape.Col2Offs));

  r := AShape.Row1 - TopRow + 1;
  if r < 0 then begin
    while r < 0 do begin
      Dec(ARect.Top,RowHeights[TopRow + r] + 1);
      Inc(r);
    end;
  end
  else begin
    while r > 0 do begin
      Inc(ARect.Top,RowHeights[TopRow + r] + 1);
      Dec(r);
    end;
  end;
  Inc(ARect.Top,Round((RowHeights[AShape.Row1] + 1) * AShape.Row1Offs));

  r := AShape.Row2 - TopRow + 1;
  if r < 0 then begin
    while r < 0 do begin
      Dec(ARect.Bottom,RowHeights[TopRow + r] + 1);
      Inc(r);
    end;
  end
  else begin
    while r > 0 do begin
      Inc(ARect.Bottom,RowHeights[TopRow + r] + 1);
      Dec(r);
    end;
  end;
  Inc(ARect.Bottom,Round((RowHeights[AShape.Row2] + 1) * AShape.Row2Offs));
end;

function TXLSGrid.IsCellVisible(ACol, ARow: integer): boolean;
begin
  Result := (ACol >= LeftCol) and (ACol <= (LeftCol + VisibleColCount)) and (ARow >= TopRow) and (ARow <= (TopRow + VisibleRowCount));
end;

procedure TXLSGrid.Paint;
begin
  inherited Paint;

  DoGraphics;
end;

procedure TXLSGrid.PaintBorder(X1,Y1,X2,Y2: integer; ABorder: TXc12BorderPr);
begin
  Canvas.Pen.Color := RevRGB(Xc12ColorToRGB(ABorder.Color,$000000));
  Canvas.MoveTo(X1,Y1);
  Canvas.LineTo(X2,Y2);
end;

procedure TXLSGrid.PaintBorders(ARect: TRect; ABorder: TXc12Border);
begin
  if ABorder.Left.Style <> cbsNone   then begin
    PaintBorder(ARect.Left - 1,ARect.Top - 1,ARect.Left - 1,ARect.Bottom,ABorder.Left);
    if ABorder.Left.Style in [cbsMedium,cbsThick,cbsDouble,cbsMediumDashed,cbsMediumDashDot] then
      PaintBorder(ARect.Left,ARect.Top - 1,ARect.Left,ARect.Bottom,ABorder.Left);
  end;
  if ABorder.Top.Style <> cbsNone    then begin
    PaintBorder(ARect.Left,ARect.Top - 1,ARect.Right,ARect.Top - 1,ABorder.Top);
    if ABorder.Top.Style in [cbsMedium,cbsThick,cbsDouble,cbsMediumDashed,cbsMediumDashDot] then
      PaintBorder(ARect.Left,ARect.Top,ARect.Right,ARect.Top,ABorder.Top);
  end;
  if ABorder.Right.Style <> cbsNone  then begin
    PaintBorder(ARect.Right,ARect.Top,ARect.Right,ARect.Bottom,ABorder.Right);
    if ABorder.Right.Style in [cbsMedium,cbsThick,cbsDouble,cbsMediumDashed,cbsMediumDashDot] then
      PaintBorder(ARect.Right - 1,ARect.Top,ARect.Right - 1,ARect.Bottom,ABorder.Right);
  end;
  if ABorder.Bottom.Style <> cbsNone then begin
    PaintBorder(ARect.Right,ARect.Bottom,ARect.Left - 1,ARect.Bottom,ABorder.Bottom);
    if ABorder.Bottom.Style in [cbsMedium,cbsThick,cbsDouble,cbsMediumDashed,cbsMediumDashDot] then
      PaintBorder(ARect.Right,ARect.Bottom - 1,ARect.Left - 1,ARect.Bottom - 1,ABorder.Bottom);
  end;
end;

procedure TXLSGrid.PaintCell(ARect: TRect; AXF,AXFLeft,AXFTop: TXc12XF; AFillCell: boolean);
var
  R: TRect;
begin
  if FNoCellPaintCnt > 0 then begin
    Dec(FNoCellPaintCnt);
    Exit;
  end;

  if AXF = Nil then
    PaintDefaultCell(ARect,FXLS.Manager.StyleSheet.XFs.DefaultXF,AFillCell)
  else begin
    R := ARect;

    if AXF.Fill.PatternType = efpNone then
      PaintDefaultCell(ARect,FXLS.Manager.StyleSheet.XFs.DefaultXF,AFillCell)
    else if AFillCell then begin
      if (AXFLeft = Nil) or (AXFLeft.Border.Right.Style = cbsNone) then
        Dec(ARect.Left);

      if (AXFTop = Nil) or (AXFTop.Border.Bottom.Style = cbsNone) then
        Dec(ARect.Top);

      Inc(ARect.Right);
      Inc(ARect.Bottom);
      Canvas.Brush.Color := RevRGB(Xc12ColorToRGB(AXF.Fill.FgColor,$00FFFFFF));
      Canvas.Pen.Color := Canvas.Brush.Color;
      Canvas.Rectangle(ARect);
    end;

    PaintBorders(R,AXF.Border);
  end;
end;

procedure TXLSGrid.PaintDefaultCell(ARect: TRect; AXF: TXc12XF; AFillCell: boolean);
begin
  if soGridlines in FSheet.Options then begin
    Canvas.Pen.Color := FGridlineColor;
    Canvas.MoveTo(ARect.Left,ARect.Bottom);
    Canvas.LineTo(ARect.Right,ARect.Bottom);
    Canvas.LineTo(ARect.Right,ARect.Top - 1);
  end;

  if AFillCell then begin
    Canvas.Brush.Color := RevRGB(Xc12ColorToRGB(AXF.Fill.FgColor,$00FFFFFF));
    Canvas.Pen.Color := Canvas.Brush.Color;
    Canvas.Rectangle(ARect);
  end;
end;

procedure TXLSGrid.PaintHeader(ARect: TRect; AText: AxUCString);
var
  Sz: TSize;
begin
  Windows.SetTextAlign(Canvas.Handle,TA_LEFT + TA_BOTTOM);

  FXLS.Manager.StyleSheet.Fonts[0].CopyToTFont(Canvas.Font);

  Canvas.Brush.Color := FHeaderColor;
  Canvas.Pen.Color := Canvas.Brush.Color;
  Canvas.Rectangle(ARect);

  Sz := Canvas.TextExtent(AText);
  Canvas.TextOut(ARect.Left + ((ARect.Right - ARect.Left) div 2) - (Sz.cx div 2),ARect.Bottom - (Sz.cy div 7),AText);

  Canvas.Pen.Color := FGridlineColor;
  Canvas.MoveTo(ARect.Left,ARect.Bottom);
  Canvas.LineTo(ARect.Right,ARect.Bottom);
  Canvas.LineTo(ARect.Right,ARect.Top - 1);

  Windows.SetTextAlign(Canvas.Handle,TA_LEFT + TA_TOP);
end;

procedure TXLSGrid.PaintText(ARect: TRect; AXF: TXc12XF; ACol,ARow: integer; AText: AxUCString; ACellType: TXLSCellType; ACanExpand: boolean);
var
//  TF  : TTextFormat;
  Opts: longword;
  w   : integer;
  R   : TRect;
begin
  Inc(ARect.Left,2);
  Inc(ARect.Top);
  Dec(ARect.Right,2);

  Opts := 0;

  AXF.Font.CopyToTFont(Canvas.Font);
  if AXF.Alignment.IsWrapText then
    Opts := Opts + DT_WORDBREAK
  else
    Opts := Opts + DT_SINGLELINE;

  case AXF.Alignment.HorizAlignment of
    chaGeneral         : begin
      case ACellType of
        xctNone,
        xctBlank,
        xctString,
        xctError,
        xctStringFormula,
        xctErrorFormula  : Opts := Opts + DT_LEFT;
        xctFloat,
        xctFloatFormula  : Opts := Opts + DT_RIGHT;
        xctBoolean,
        xctBooleanFormula: Opts := Opts + DT_CENTER;
      end;
//      TF := TF + [tfLeft];
    end;
    chaLeft            : Opts := Opts + DT_LEFT;
    chaCenter          : Opts := Opts + DT_CENTER;
    chaRight           : Opts := Opts + DT_RIGHT;
    chaFill,
    chaJustify,
    chaCenterContinuous,
    chaDistributed     : Opts := Opts + DT_LEFT;
  end;
  case AXF.Alignment.VertAlignment of
    cvaTop        : Opts := Opts + DT_TOP;
    cvaCenter     : Opts := Opts + DT_VCENTER;
    cvaBottom     : Opts := Opts + DT_BOTTOM;
    cvaJustify,
    cvaDistributed: Opts := Opts + DT_BOTTOM;
  end;
  w := Canvas.TextWidth(AText);
  if AXF.Alignment.Indent > 0 then
    Inc(w,Canvas.TextWidth('w') * AXF.Alignment.Indent);
  if ACanExpand and (w > (ARect.Right - ARect.Left)) and not ((DT_WORDBREAK and Opts) <> 0) then begin
    Dec(w,ARect.Right - ARect.Left);
    if (DT_LEFT and Opts) <> 0 then
      ExpandWidthRight(ARect,ACol + 1,ARow,w)
    else if (DT_CENTER and Opts) <> 0 then begin
      R := ARect;
      ExpandWidthLeft(R,ACol - 1,ARow,w div 2);
      ExpandWidthRight(R,ACol + 1,ARow,w div 2);
      Dec(ARect.Left,w div 2);
      Inc(ARect.Right,w div 2);
    end
    else if (DT_RIGHT and Opts) <> 0 then
      ExpandWidthLeft(ARect,ACol - 1,ARow,w);
  end;
  if AXF.Alignment.Indent > 0 then
    Inc(ARect.Left,Canvas.TextWidth(' ') * AXF.Alignment.Indent);
  Canvas.Brush.Style := bsClear;

  Windows.DrawTextExW(Canvas.Handle, PWideChar(AText), Length(AText), ARect, Opts, nil);

  Canvas.Brush.Style := bsSolid;
end;

function TXLSGrid.RowPixelHeight(AHeight: integer): integer;
var
  PPIX,PPIY: integer;
begin
  FXLS.Manager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

  Result := Round((AHeight / 20) / (72 / PPIY));
end;

procedure TXLSGrid.RowsChanged;
var
  i   : integer;
  r   : integer;
  c   : integer;
  h,h2: integer;
  Row : PXLSMMURowHeader;
begin
  DefaultRowHeight := RowPixelHeight(XLS_DEFAULT_ROWHEIGHT);

  RowCount := Max(255,FSheet.LastRow + 2);

  for r := 0 to FSheet.LastRow do begin
    Row := FSheet.MMUCells.FindRow(r);
    if Row <> Nil then begin
      if Row.Height > 60000 then
        h := RowPixelHeight(XLS_DEFAULT_ROWHEIGHT)
      else
        h := RowPixelHeight(Row.Height);

      h2 := 0;
      if xroCustomHeight in Row.Options then begin
        for c := 0 to FSheet.LastCol do begin
          i := FSheet.MMUCells.GetStyle(c,r);
          if (i <> XLS_STYLE_DEFAULT_XF) and (FXLS.Manager.StyleSheet.XFs[i].Border.Bottom.Style in [cbsMedium,cbsThick,cbsDouble,cbsMediumDashed,cbsMediumDashDot]) then begin
            h2 := 2;
            Break;
          end;
        end;
      end;

      if h >= 0 then
        RowHeights[r + 1] := h + h2;
    end;
  end;
end;

function TXLSGrid.SelectCell(ACol, ARow: Integer): Boolean;
begin
  Result := inherited SelectCell(ACol,ARow);

   if FEditText <> '' then begin
     SetCellText(FEditCol,FEditRow,FEditText);

     FEditText := '';
   end;
end;

procedure TXLSGrid.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
begin
  if FTabs <> Nil then begin
    Dec(AHeight,FTabs.Height);
    FTabs.Left := ALeft;
    FTabs.Top := ATop + AHeight;
    FTabs.Width := AWidth;
  end;

  inherited;
end;

procedure TXLSGrid.SetCellText(ACol, ARow: integer; AText: AxUCString);
var
  vD : double;
  vDT: TDateTime;
  S  : AxUCString;
begin
  inherited;

  Dec(ACol);
  Dec(ARow);

  S := Trim(AText);
  if S = '' then
    Exit;

  if S[1] = '''' then
    FSheet.AsString[Acol,ARow] := Copy(S,2,MAXINT)
  else if S[1] = '=' then
    FSheet.AsFormula[Acol,ARow] := S
  else if Lowercase(S) = 'TRUE' then
    FSheet.AsBoolean[Acol,ARow] := True
  else if Lowercase(S) = 'FALSE' then
    FSheet.AsBoolean[Acol,ARow] := False
  else if TryStrToFloat(S,vD) then
    FSHeet.AsFloat[ACol,ARow] := vD
  else if TryStrToDateTime(S,vDT) then
    FSHeet.AsDateTime[ACol,ARow] := vDT
  else
    FSHeet.AsString[ACol,ARow] := AText;
end;

procedure TXLSGrid.SetEditText(ACol, ARow: Integer; const Value: string);
begin
  inherited;

  FEditText := Value;
  FEditCol := ACol;
  FEditRow := ARow;
end;

procedure TXLSGrid.SetParent(AParent: TWinControl);
begin
  inherited;

  if FTabs = Nil then begin
    FTabs := TXLSTabSet.Create(Owner);
    FTabs.Parent := Parent;
    FTabs.Tabs.Add('Sheet1');
    FTabs.TabIndex := 0;

    FTabs.OnChange := TabChanged;

    Height := Height - FTabs.Height;
    FTabs.Left := Left;
    FTabs.Top := Top + Height;
    FTabs.Width := Width;

//    FTabs.Left := 0;
//    FTabs.Top := 160;

    XLSChanged;
  end;
end;

procedure TXLSGrid.TabChanged(Sender: TObject; NewTab: Integer; var AllowChange: Boolean);
begin
  FXLS.SelectedTab := NewTab;

  XLSChanged(False);
end;

procedure TXLSGrid.TopLeftChanged;
begin
  inherited;

//  ColsChanged;
//  RowsChanged;
end;

procedure TXLSGrid.XLSChanged(AUpdateTabs: boolean = True);
var
  i: integer;
begin
  FSheet := FXLS[FXLS.SelectedTab];

  FSheet.CalcDimensions;

  if AUpdateTabs then begin
    FTabs.Tabs.Clear;
    for i := 0 to FXLS.Count - 1 do
       FTabs.Tabs.Add(FXLS[i].Name);

    FTabs.TabIndex := FXLS.SelectedTab;
  end;

  ColsChanged;
  RowsChanged;

  Invalidate;
end;

{ TXLSReadWriteIIGrid5 }

procedure TXLSReadWriteIIGrid5.AfterRead;
begin
  inherited AfterRead;

  TXLSGrid(Owner).XLSChanged;
end;

constructor TXLSReadWriteIIGrid5.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

end;

{ TXLSTabSet }

procedure TXLSTabSet.WriteState(Writer: TWriter);
begin

end;

end.
