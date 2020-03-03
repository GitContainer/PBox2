unit XBookRows2;

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

uses { Delphi  } Classes, SysUtils, vcl.Controls,
     { XLSRWII } Xc12Utils5,
                 XLSUtils5, XLSRow5, XLSAutofilter5,
     { XBook   } XBookUtils2, XBookSysVar2, XBookPaintGDI2, XBookPaint2,
                 XBookSkin2, XBookOptions2, XBookHeaders2, XBookWindows2;

type TRowData = record
     SelState: THeaderSelState;
     Height: word;
     end;

type TXRowsGutter = class(TXHeadersGutter)
protected
     procedure PaintOutline(P1,P2,Level: integer; LineType: TOutlineLineType); override;
     procedure CentrePos(P1,P2,Delta: integer; var cx,cy: integer); override;
public
     end;

type TXLSBookRows = class(TXLSBookHeaders)
private
     FXLSRows       : TXLSRows;
     FDefaultHeight : integer;
     FWidth         : integer;
     FCurrY         : integer;
     FAutofilter    : TXLSAutofilter;
     FFrozenFirstHdr: integer;
     FFrozenHdr     : integer;
     FPrintVertAdj  : double;

     function  GetWidth: integer;
     function  GetHeight(Index: integer): integer;
     function  GetBottomRow: integer;
     function  GetTopRow: integer;
     procedure SetTopRow(const Value: integer);
     function  GetAbsHeight(Index: integer): integer;
     function  GetHeightFloat(Index: integer): double;
     function  GetFrozenHdr: integer;
     procedure SetBottomRow(const Value: integer);
protected
     procedure GetHeaderData(Hdr: integer; var Size,Outline: integer; var Options: THeaderOptions); override;
     procedure GetHeaderDataFloat(Hdr: integer; var Size: double; var Outline: integer; var Options: THeaderOptions); override;
     procedure PaintSizing; override;
     function  FirstPos: integer; override;
     function  LastPos: integer; override;
     function  XorY(X,Y: integer): integer; override;
     function  GetCursor(HitType: TSizeHitType): TXLSCursorType; override;
     function  ClipY(Y: integer): integer;
     procedure PaintHeader(Hdr: integer); override;
public
     constructor Create(Parent: TXSSClientWindow; Rows: TXLSRows; Autofilter: TXLSAutofilter);
     destructor Destroy; override;
     procedure Paint; override;
     procedure RowsChanged(Rows: TXLSRows; OutlineLevel: integer);
     procedure CalcHeaders(MaxHdrToCalc: integer = MAXINT); override;
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseLeave; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure ButtonClick(Sender: TObject; Hdr,Level: integer);
     procedure RowPos(Row: integer; var Y1,Y2: integer); overload;
     procedure RowPos(Row1,Row2: integer; var Y1,Y2: integer); overload;
     function  HeightToPixels(H: integer): integer;
     function  PageHeight: integer;
     procedure SetTopRowPrint(Row, MaxRowToCalc: integer);

     property RowWidth: integer read GetWidth;
     property RowHeight[Index: integer]: integer read GetHeight;
     property RowHeightFloat[Index: integer]: double read GetHeightFloat;
     property AbsHeight[Index: integer]: integer read GetAbsHeight;
     property DefaultHeight: integer read FDefaultHeight write FDefaultHeight;
     property TopRow: integer read GetTopRow write SetTopRow;
     property BottomRow: integer read GetBottomRow write SetBottomRow;
     property FrozenRow: integer read GetFrozenHdr write FFrozenHdr;
     property FrozenTopRow: integer read FFrozenFirstHdr write FFrozenFirstHdr;
     property XLSRows: TXLSRows read FXLSRows;
     property PrintVertAdj: double read FPrintVertAdj write FPrintVertAdj;
     property ClickedRow: integer read FCurrHeader;
     end;

implementation

{ TXLSBookRows }

procedure TXLSBookRows.CalcHeaders(MaxHdrToCalc: integer = MAXINT);
var
  W: integer;
begin
  inherited CalcHeaders(MaxHdrToCalc);
  if FLastHdr < 1000 then
    W := FSkin.GDI.TextWidth('000') + FSkin.HeaderMargin * 2
  else if FLastHdr < 10000 then
    W := FSkin.GDI.TextWidth('0000') + FSkin.HeaderMargin * 2
  else if FLastHdr < 100000 then
    W := FSkin.GDI.TextWidth('00000') + FSkin.HeaderMargin * 2
  else if FLastHdr < 1000000 then
    W := FSkin.GDI.TextWidth('000000') + FSkin.HeaderMargin * 2
  else
    W := FSkin.GDI.TextWidth('0000000') + FSkin.HeaderMargin * 2;
  if W <> FWidth then
    FWidth := W;
end;

constructor TXLSBookRows.Create(Parent: TXSSClientWindow; Rows: TXLSRows; Autofilter: TXLSAutofilter);
begin
  inherited Create(Parent,XLS_MAXROW);
  FPrintVertAdj := 1.0;
  FAutofilter := Autofilter;
  FGutter := TXRowsGutter.Create(Self);
  FGutter.Buttons.OnButtonClick := ButtonClick;
  Add(FGutter);
  FDefaultHeight := 255;
  RowsChanged(Rows,0);
end;

function TXLSBookRows.FirstPos: integer;
begin
  Result := FCY1;
end;

function TXLSBookRows.GetHeight(Index: integer): integer;
begin
  Result := VisibleSize[Index];
end;

function TXLSBookRows.GetHeightFloat(Index: integer): double;
begin
  Result := VisibleSizeFloat[Index];
end;

function TXLSBookRows.GetWidth: integer;
begin
  if FVisible then
    Result := FWidth + FGutter.Levels * GUTTER_SIZE
  else
    Result := 0;
end;

function TXLSBookRows.LastPos: integer;
begin
  Result := FCY2 - FOffset;
end;

function TXLSBookRows.PageHeight: integer;
var
  i: integer;
  V: double;
begin
  V := 0;
  for i := FFirstHdr to FLastHdr do
    V := V + RowHeightFloat[i];
  Result := Round(V);
end;

procedure TXLSBookRows.Paint;

procedure DoPaint(Row1,Row2: integer);
var
  i,Row: integer;
  y: integer;
  Hidden: boolean;
  F: boolean;
begin
  y := FCY1;
  Hidden := False;
  Row := FFirstHdr;
  for i := 0 to High(FHeaders) do begin
    if (FHeaders[i].Size <= 0) or (hoHidden in FHeaders[i].Options) then
      Hidden := True
    else begin
      F := FAutofilter.IsFiltered and ((i + FFirstHdr) >= (FAutofilter.Row1 + 1)) and ((i + FFirstHdr) <= FAutofilter.Row2);
      FSkin.PaintHeader(FCX1,y,FCX2,y + FHeaders[i].Size - 1,FHeaders[i].SelState,xshsRow,FHeaders[i].Hoover,IntToStr(Row + 1),F);
      if Hidden then begin
        FSkin.GDI.Line(FCX1,y,FCX2,y);
        Hidden := False;
      end;
      Inc(y,FHeaders[i].Size);
    end;
    Inc(Row);
  end;
end;

begin
  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);

  FGutter.Paint(FHeaders);

  BeginPaint;

  FSkin.GDI.BrushColor := $FFFFFF;
  FSkin.GDI.Rectangle(FCX1 - 1,FCY1 - 1,FCX2 + 1,FCY2 + 1);

{
  if FFrozenHdr > 0 then begin
    DoPaint(0,FFrozenHdr);
    DoPaint(FFrozenHdr + 1 + FFirstHdr,High(FHeaders));
  end
  else
}
  DoPaint(FFirstHdr,High(FHeaders));
  EndPaint;
end;

procedure TXLSBookRows.PaintSizing;
var
  Row: integer;
  y,LastR: integer;

procedure PaintRows(R1,R2: integer);
var
  i: integer;
  Hidden: boolean;
  F: boolean;
begin
  Hidden := False;
  for i := R1 to R2 do begin
    if (FHeaders[i].SizingSize <= 0) or (hoHidden in FHeaders[i].Options) then
      Hidden := True
    else begin
      F := FAutofilter.IsFiltered and ((i + FFirstHdr) >= (FAutofilter.Row1 + 1)) and ((i + FFirstHdr) <= FAutofilter.Row2);
      FSkin.PaintHeader(FCX1,y,FCX2,y + FHeaders[i].SizingSize - 1,FHeaders[i].SelState,xshsRow,FHeaders[i].Hoover,IntToStr(Row + 1),F);
      if Hidden then begin
        FSkin.GDI.Line(FCX1,y,FCX2,y);
        Hidden := False;
      end;
    end;
    if not (hoHidden in FHeaders[i].Options) then
      Inc(y,FHeaders[i].SizingSize);
    Inc(Row);
  end;
end;

begin
  BeginPaint;

  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);
  y := FCY1;
  Row := FFirstHdr;
  LastR := High(FHeaders);
  PaintRows(0,LastR);
  if y < FCY2 then begin
    AddSizingHeaders;
    PaintRows(LastR + 1,High(FHeaders));
  end;

  EndPaint;
end;

procedure TXLSBookRows.GetHeaderData(Hdr: integer; var Size,Outline: integer; var Options: THeaderOptions);
var
  XRow: TXLSRow;
  H: integer;
begin
  XRow := FXLSRows[Hdr];
  if XRow <> Nil then begin
    H := XRow.Height;
    Outline := XRow.OutlineLevel;
    Options := [];
    if XRow.Hidden then
      Options := Options + [hoHidden];
    if XRow.CollapsedOutline then
      Options := Options + [hoCollapsed];
  end
  else begin
    H := FDefaultHeight;
    Outline := 0;
    Options := [];
  end;
  Size := Round((H / 20) * (FSkin.GDI.PixelsPerInchY / 72));
end;

procedure TXLSBookRows.GetHeaderDataFloat(Hdr: integer; var Size: double; var Outline: integer; var Options: THeaderOptions);
var
  XRow: TXLSRow;
begin
  XRow := FXLSRows[Hdr];
  if XRow <> Nil then begin
    Size := XRow.PixelHeight * FPrintVertAdj;

    Outline := XRow.OutlineLevel;
    Options := [];
    if XRow.Hidden then
      Options := Options + [hoHidden];
    if XRow.CollapsedOutline then
      Options := Options + [hoCollapsed];
  end
  else begin
    Size := FDefaultHeight;
    Outline := 0;
    Options := [];
  end;
end;

procedure TXLSBookRows.RowsChanged(Rows: TXLSRows; OutlineLevel: integer);
begin
  FXLSRows := Rows;

  FXLSRows.CalcLastVisibleRow;

  if OutlineLevel > 0 then
    FGutter.Levels := OutlineLevel + 1
  else
    FGutter.Levels := 0;
  CalcHeaders;
end;

function TXLSBookRows.XorY(X, Y: integer): integer;
begin
  Result := Y;
end;

function TXLSBookRows.GetCursor(HitType: TSizeHitType): TXLSCursorType;
begin
  case HitTYpe of
    shtSize   : Result := xctVertSplit;
    shtHidden : Result := xctVert2Split;
    else        Result := xctRowSelect;
  end;
end;

function TXLSBookRows.GetFrozenHdr: integer;
begin
//  Result := FFrozenHdr;
  Result := 0;
end;

function TXLSBookRows.GetAbsHeight(Index: integer): integer;
begin
  Result := AbsVisibleSize[Index];
end;

function TXLSBookRows.GetBottomRow: integer;
begin
  Result := FLastHdr;
end;

function TXLSBookRows.GetTopRow: integer;
begin
  Result := FFirstHdr;
end;

procedure TXLSBookRows.SetTopRow(const Value: integer);
begin
  FFirstHdr := Value;
  CalcHeaders;
//  CalcGutter;
end;

procedure TXLSBookRows.SetTopRowPrint(Row, MaxRowToCalc: integer);
begin
  FFirstHdr := Row;
  CalcHeaders(MaxRowToCalc);
end;

procedure TXLSBookRows.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
var
  h: integer;
begin
  inherited MouseUp(Button,Shift,X,Y);;

  if FSizingHdr >= 0 then begin
    FPaintLineEvent(Self,pltVSize,X,FCurrY,False);
//    for i := 0 to High(FHeaders) do begin
      h := Round((FHeaders[FSizingHdr].SizingSize / (FSkin.GDI.PixelsPerInchY / 72)) * 20);
      if h <> FDefaultHeight then
        FXLSRows[FFirstHdr + FSizingHdr].Height := h
      else if FXLSRows[FFirstHdr + FSizingHdr] <> Nil then
        FXLSRows[FFirstHdr + FSizingHdr].Height := h;
//    end;
    CalcHeaders;
    FChangedEvent(Self);
  end;
  FSizingHdr := -1;
end;

procedure TXLSBookRows.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  if FSizingHdr >= 0 then begin
    FCurrY := ClipY(Y - DeltaPos - 1);
    FPaintLineEvent(Self,pltVSize,X,FCurrY,False);
  end;
end;

procedure TXLSBookRows.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);;

  if FSizingHdr >= 0 then begin
    FPaintLineEvent(Self,pltVSize,X,FCurrY,False);
    FCurrY := ClipY(Y - DeltaPos - 1);
    FPaintLineEvent(Self,pltVSize,X,FCurrY,False);
  end;
end;

function TXLSBookRows.ClipY(Y: integer): integer;
begin
  if Y < FCY1 then
    Result := FCY1
  else if Y > FCY2 then
    Result := FCY2
  else
    Result := Y;
end;

destructor TXLSBookRows.Destroy;
begin
  inherited;
end;

procedure TXLSBookRows.SetBottomRow(const Value: integer);
begin
  FLastHdr := Value;
  CalcHeadersFromLast;
end;

procedure TXLSBookRows.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited SetSize(pX1, pY1 ,pX2, pY2);
  SetClientSize(pX1 + FGutter.Levels * GUTTER_SIZE, pY1, pX2, pY2);
  FGutter.SetSize(pX1, pY1 - 1, pX1 + FGutter.Levels * GUTTER_SIZE, pY2 + 2);
  CalcHeaders;
end;

procedure TXLSBookRows.MouseLeave;
begin
  inherited;
  FGutter.MouseLeave;
end;

procedure TXLSBookRows.ButtonClick(Sender: TObject; Hdr, Level: integer);
begin
  FXLSRows.ToggleGrouped(Hdr);
  CalcHeaders;
  if (Hdr >= 0) and Assigned(FSyncButtonClickEvent) then
    FSyncButtonClickEvent(Self,-1,-1);
  FParent.Invalidate;
end;

procedure TXLSBookRows.RowPos(Row: integer; var Y1, Y2: integer);
begin
  Y1 := FHeaders[Row].Pos;
  Y2 := Y1;
  if not (hoHidden in FHeaders[Row].Options) then
    Inc(Y2,FHeaders[Row].Size);
end;

procedure TXLSBookRows.RowPos(Row1, Row2: integer; var Y1, Y2: integer);
begin
  Y1 := FHeaders[Row1].Pos;
  Y2 := FHeaders[Row2].Pos;
  if not (hoHidden in FHeaders[Row2].Options) then
    Inc(Y2,FHeaders[Row2].Size);
end;

procedure TXLSBookRows.PaintHeader(Hdr: integer);
var
  y: integer;
  F: boolean;
begin
  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);
  y := FHeaders[Hdr].Pos;
  F := FAutofilter.IsFiltered and ((Hdr + FFirstHdr) >= (FAutofilter.Row1 + 1)) and ((Hdr + FFirstHdr) <= FAutofilter.Row2);
  FSkin.PaintHeader(FCX1,y,FCX2,y + FHeaders[Hdr].Size - 1,FHeaders[Hdr].SelState,xshsRow,FHeaders[Hdr].Hoover,IntToStr(FFirstHdr + Hdr + 1),F);
end;

function TXLSBookRows.HeightToPixels(H: integer): integer;
begin
  Result := Round((H / 20) * (FSkin.GDI.PixelsPerInchY / 72));
end;

{ TXRowsGutter }

procedure TXRowsGutter.CentrePos(P1, P2,Delta: integer; var cx, cy: integer);
begin
  cx := FCX1 + (GUTTER_SIZE div 2) + Delta;
  cy := FCY1 + P1 + (P2 - P1) div 2;
end;

procedure TXRowsGutter.PaintOutline(P1,P2,Level: integer; LineType: TOutlineLineType);
var
  cx,cy: integer;
begin
  Inc(P1,FCY1);
  Inc(P2,FCY1);
  cx := FCX1 + Level * GUTTER_SIZE + GUTTER_SIZE div 2;
  cy := P1 + (P2 - P1) div 2;
  case LineType of
    olpFirst: begin
      FSkin.GDI.Rectangle(cx - 1,P1 + 1,cx,P2 + 1);
      FSkin.GDI.Rectangle(cx + 1,P1 + 1,cx + 4,P1 + 2);
    end;
    olpMiddle: FSkin.GDI.Rectangle(cx - 1,P1 + 2,cx,P2 + 1);
    olpLast  : FSkin.GDI.Rectangle(cx - 1,P1 + 2,cx,cy);
    olpDot   : FSkin.GDI.Rectangle(cx - 1,cy - 1,cx + 1,cy + 1);
  end;
end;

end.
