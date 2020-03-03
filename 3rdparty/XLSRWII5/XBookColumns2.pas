unit XBookColumns2;

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
                 XLSUtils5, XLSColumn5,
     { XBook   } XBookUtils2, XBookSysVar2, XBookPaintGDI2, XBookPaint2, XBookWindows2,
                 XBookSkin2, XBookOptions2, XBookHeaders2;

type TXColsGutter = class(TXHeadersGutter)
protected
     procedure PaintOutline(P1,P2,Level: integer; LineType: TOutlineLineType); override;
     procedure CentrePos(P1,P2,Delta: integer; var cx,cy: integer); override;
public
     end;

type TXLSBookColumns = class(TXLSBookHeaders)
private
     FXLSColumns     : TXLSColumns;
     FStdFontWidth   : integer;
     FDefaultWidth   : integer;
     FDefaultWidthPix: integer;
     FDefaultHeight  : integer;
     FCurrX          : integer;
     FFrozenFirstHdr : integer;
     FFrozenHdr      : integer;

     function  GetWidth(Index: integer): integer;
     function  GetLeftCol: integer;
     function  GetRightCol: integer;
     procedure SetLeftCol(const Value: integer);
     function  GetHeight: integer;
     function  GetAbsWidth(Index: integer): integer;
     function  GetWidthFloat(Index: integer): double;
     function  GetVisualCol(Index: integer): integer; {$ifdef ver170} inline; {$endif}
     function  GetFrozenHdr: integer;
     procedure SetRightCol(const Value: integer);
protected
     function  PixelsToColWidth(Pixels: integer): integer;
     function  ColToRefStr(ACol: integer): AxUCString;
     procedure PaintSizing; override;
     function  FirstPos: integer; override;
     function  LastPos: integer; override;
     function  XorY(X,Y: integer): integer; override;
     function  GetCursor(HitType: TSizeHitType): TXLSCursorType; override;
     function  GetDefaultWidth(FontSize: integer): integer;
     procedure PaintHeader(Hdr: integer); override;
public
     constructor Create(Parent: TXSSClientWindow; Columns: TXLSColumns);
     destructor Destroy; override;
     procedure GetHeaderData(Hdr: integer; var Size,Outline: integer; var Options: THeaderOptions); override;
     procedure GetHeaderDataFloat(Hdr: integer; var Size: double; var Outline: integer; var Options: THeaderOptions); override;
     procedure Paint; override;
     procedure ColumnsChanged(AColumns: TXLSColumns; const AOutlineLevel, AStdFontWidth: integer);
     procedure MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseMove(Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer); override;
     procedure MouseLeave; override;
     procedure SetSize(const pX1,pY1,pX2,pY2: integer); override;
     procedure ButtonClick(Sender: TObject; Hdr,Level: integer);
     procedure ColPos(Col: integer; var X1,X2: integer); overload;
     procedure ColPos(Col1,Col2: integer; var X1,X2: integer); overload;
     function  WidthToPixels(W: integer): integer;
     procedure ScrollLeft;
     procedure ScrollPageLeft;
     procedure ScrollRight;
     procedure ScrollPageRight;
     function  PageWidth: integer;

     property DefaultWidth: integer read FDefaultWidth write FDefaultWidth;
     property DefaultWidthPix: integer read FDefaultWidthPix write FDefaultWidthPix;
     property ColWidth[Index: integer]: integer read GetWidth;
     property ColWidthFloat[Index: integer]: double read GetWidthFloat;
     property AbsWidth[Index: integer]: integer read GetAbsWidth;
     property LeftCol: integer read GetLeftCol write SetLeftCol;
     property RightCol: integer read GetRightCol write SetRightCol;
     property FrozenCol: integer read GetFrozenHdr write FFrozenHdr;
     property FrozenLeftCol: integer read FFrozenFirstHdr write FFrozenFirstHdr;
     property ColHeight: integer read GetHeight;
     property DefaultHeight: integer read FDefaultHeight write FDefaultHeight;
     property XLSColumns: TXLSColumns read FXLSColumns;
     property VisualCol[Index: integer]: integer read GetVisualCol;
     property ClickedCol: integer read FCurrHeader;
     end;

implementation

{ TXLSBookColumns }

function TXLSBookColumns.ColToRefStr(ACol: integer): AxUCString;
var
  C: integer;
begin
  if ACol < 26 then
    Result := Char(Ord('A') + ACol)
  else if ACol < 676 then
    Result := Char(Ord('@') + ACol div 26) + Char(Ord('A') + ACol mod 26)
  else begin
    C := ACol div 676;
    Result := Char(Ord('A') + C);
    Dec(ACol,C * 676);
    C := ACol div 26;
    Result := Result + Char(Ord('A') + C);
    Result := Result + Char(Ord('A') + ACol mod 26);
  end;
end;

procedure TXLSBookColumns.ColumnsChanged(AColumns: TXLSColumns; const AOutlineLevel, AStdFontWidth: integer);
begin
  FXLSColumns := AColumns;
  if AOutlineLevel > 0 then
    FGutter.Levels := AOutlineLevel + 1
  else
    FGutter.Levels := 0;
  FStdFontWidth := AStdFontWidth;

  FDefaultWidth := GetDefaultWidth(FSkin.SysFontSize);
  FDefaultWidthPix := Round((FDefaultWidth / 256) * FStdFontWidth);

  FSkin.SetSystemFont(False);
  FSkin.GDI.CheckMinFontSize;
  FSkin.GetCellFontData;

  CalcHeaders;
end;

procedure TXLSBookColumns.GetHeaderData(Hdr: integer; var Size,Outline: integer; var Options: THeaderOptions);
var
  XCol: TXLSColumn;
begin
  XCol := FXLSColumns[Hdr];
  Options := [];
  if XCol <> Nil then begin
    Size := Round((XCol.Width / 256) * FStdFontWidth);
    Outline := XCol.OutlineLevel;
    if XCol.Hidden then
      Options := Options + [hoHidden];
    if XCol.CollapsedOutline then
      Options := Options + [hoCollapsed];
  end
  else begin
    Size := FDefaultWidthPix;
    Outline := 0;
  end;
end;

procedure TXLSBookColumns.GetHeaderDataFloat(Hdr: integer; var Size: double; var Outline: integer; var Options: THeaderOptions);
var
  XCol: TXLSColumn;
begin
  XCol := FXLSColumns[Hdr];
  Options := [];
  if XCol <> Nil then begin
    Size := Round((XCol.Width / 256) * FStdFontWidth);
    Outline := XCol.OutlineLevel;
    if XCol.Hidden then
      Options := Options + [hoHidden];

    if XCol.CollapsedOutline then
      Options := Options + [hoCollapsed];
  end
  else begin
//    Size := Round(FDefaultWidth * FStdCharWidth);
    Size := FDefaultWidthPix;
    Outline := 0;
  end;
{
  if FSkin.GDI.Printing then
    Size := Size * 7;
}
end;

constructor TXLSBookColumns.Create(Parent: TXSSClientWindow; Columns: TXLSColumns);
begin
  inherited Create(Parent,XLS_MAXCOL);
  FGutter := TXColsGutter.Create(Self);
  FGutter.Buttons.OnButtonClick := ButtonClick;
  FDefaultWidthPix := 64;
  Add(FGutter);
//  ColumnsChanged(Columns,0);
end;

function TXLSBookColumns.GetLeftCol: integer;
begin
  Result := FFirstHdr;
end;

function TXLSBookColumns.GetRightCol: integer;
begin
  Result := FLastHdr;
end;

function TXLSBookColumns.GetWidth(Index: integer): integer;
begin
  Result := VisibleSize[Index];
end;

function TXLSBookColumns.GetWidthFloat(Index: integer): double;
begin
  Result := VisibleSizeFloat[Index];
end;

function TXLSBookColumns.GetVisualCol(Index: integer): integer;
begin
  Result := FHeaders[Index].Header;
end;

function TXLSBookColumns.PageWidth: integer;
var
  i: integer;
begin
  Result := 0;
  for i := FFirstHdr to FLastHdr do
    Inc(Result,ColWidth[i]);
end;

procedure TXLSBookColumns.Paint;
var
  x,i: integer;
  Hidden: boolean;
begin
  FSkin.SetSystemFont(False);
  FSkin.TextAlign(ptaCenter);

  FGutter.Paint(FHeaders);

  BeginPaint;

  FSkin.GDI.BrushColor := $FFFFFF;
  FSkin.GDI.Rectangle(FCX1 - 1,FCY1 - 1,FCX2 + 1,FCY2 + 1);

  Hidden := False;
  x := FirstPos;
  for i := 0 to High(FHeaders) do begin
    if (FHeaders[i].Size <= 0) or (hoHidden in FHeaders[i].Options) then
      Hidden := True
    else begin
      FSkin.PaintHeader(x,FCY1,x + FHeaders[i].Size - 1,FCY2,FHeaders[i].SelState,xshsNormal,FHeaders[i].Hoover,ColToRefStr(FHeaders[i].Header));
      if Hidden then begin
        FSkin.GDI.Line(x,FCY1,x,FCY2);
        Hidden := False;
      end;
      Inc(x,FHeaders[i].Size);
    end;
  end;
  EndPaint;
end;

procedure TXLSBookColumns.PaintSizing;
var
  Col: integer;
  x,LastC: integer;

procedure PaintCols(C1,C2: integer);
var
  i: integer;
  Hidden: boolean;
begin
  Hidden := False;
  for i := C1 to C2 do begin
    if (FHeaders[i].SizingSize <= 0) or (hoHidden in FHeaders[i].Options) then
      Hidden := True
    else begin
      FSkin.PaintHeader(x,FY1,x + FHeaders[i].SizingSize - 1,FY2,FHeaders[i].SelState,xshsNormal,FHeaders[i].Hoover,ColToRefStr(Col));
      if Hidden then begin
        FSkin.GDI.Line(x,FCY1,x,FCY2);
        Hidden := False;
      end;
    end;
    if not (hoHidden in FHeaders[i].Options) then
      Inc(x,FHeaders[i].SizingSize);
    Inc(Col);
  end;
end;

begin
  BeginPaint;

  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);
  x := FCX1;
  Col := FFirstHdr;
  LastC := High(FHeaders);
  PaintCols(0,LastC);
  if x < FCX2 then begin
    AddSizingHeaders;
    PaintCols(LastC + 1,High(FHeaders));
  end;

  EndPaint;
end;

function TXLSBookColumns.PixelsToColWidth(Pixels: integer): integer;
begin
  Result := Round((Pixels / FStdFontWidth) * 256) - 1;
end;

procedure TXLSBookColumns.ScrollLeft;
begin
  Scroll(hsInc);
end;

procedure TXLSBookColumns.ScrollPageLeft;
begin
  Scroll(hsIncPage);
end;

procedure TXLSBookColumns.ScrollPageRight;
begin
  Scroll(hsDecPage);
end;

procedure TXLSBookColumns.ScrollRight;
begin
  Scroll(hsDec);
end;

procedure TXLSBookColumns.SetLeftCol(const Value: integer);
begin
  FFirstHdr := Value;
  CalcHeaders;
//  CalcGutter;
end;

procedure TXLSBookColumns.SetRightCol(const Value: integer);
begin
  FLastHdr := Value;
  CalcHeadersFromLast;
end;

function TXLSBookColumns.FirstPos: integer;
begin
  Result := FCX1;
end;

function TXLSBookColumns.LastPos: integer;
begin
  Result := FCX2 - FOffset;
end;

function TXLSBookColumns.XorY(X, Y: integer): integer;
begin
  Result := X;
end;

function TXLSBookColumns.GetAbsWidth(Index: integer): integer;
begin
  Result := AbsVisibleSize[Index];
end;

function TXLSBookColumns.GetCursor(HitType: TSizeHitType): TXLSCursorType;
begin
  case HitTYpe of
    shtSize   : Result := xctHorizSplit;
    shtHidden : Result := xctHoriz2Split;
    else        Result := xctColSelect;
  end;
end;

function TXLSBookColumns.GetDefaultWidth(FontSize: integer): integer;
begin
       if FontSize <=  4 then Result := 2730
  else if FontSize <=  6 then Result := 2457
  else if FontSize <=  8 then Result := 2389
  else if FontSize <= 10 then Result := 2340
  else if FontSize <= 11 then Result := 2304
  else if FontSize <= 12 then Result := 2275
  else if FontSize <= 14 then Result := 2234
  else if FontSize <= 16 then Result := 2218
  else if FontSize <= 20 then Result := 2321
  else if FontSize <= 24 then Result := 2275
  else Result := 2199;
end;

function TXLSBookColumns.GetFrozenHdr: integer;
begin
//  Result := FFrozenHdr;
  Result := 0;
end;

procedure TXLSBookColumns.MouseUp(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseUp(Button,Shift,X,Y);;
  if (FSizingHdr >= 0) and not (xssDouble in Shift) then begin
    FPaintLineEvent(Self,pltHSize,FCurrX,Y,False);
//    for i := 0 to High(FHeaders) do begin
      if FHeaders[FSizingHdr].SizingSize <> FDefaultWidthPix then begin
//        FXLSColumns.AddIfNone(FFirstHdr + FSizingHdr,1);
        FXLSColumns[FFirstHdr + FSizingHdr].Width := PixelsToColWidth(FHeaders[FSizingHdr].SizingSize);
      end
      else if FXLSColumns[FFirstHdr + FSizingHdr] <> Nil then
        FXLSColumns[FFirstHdr + FSizingHdr].Width := PixelsToColWidth(FHeaders[FSizingHdr].SizingSize);
//    end;
    CalcHeaders;
    FChangedEvent(Self);
  end;
  FSizingHdr := -1;
end;

procedure TXLSBookColumns.MouseDown(Button: TXSSMouseButton; Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseDown(Button,Shift,X,Y);
  if FSizingHdr >= 0 then begin
{
    if FSizingHdr > 0 then
      FPaintLineEvent(Self,pltHSize,FCurrX);
}
    FCurrX := X - DeltaPos - 1;
    FPaintLineEvent(Self,pltHSize,FCurrX,Y,False);
  end;
end;

procedure TXLSBookColumns.MouseMove(Shift: TXSSShiftState; X, Y: Integer);
begin
  inherited MouseMove(Shift,X,Y);;

  if FSizingHdr >= 0 then begin
    FPaintLineEvent(Self,pltHSize,FCurrX,Y,False);
    FCurrX := X - DeltaPos - 1;
    FPaintLineEvent(Self,pltHSize,FCurrX,Y,False);
  end;
end;

procedure TXLSBookColumns.MouseLeave;
begin
  inherited;
  FGutter.MouseLeave;
end;

function TXLSBookColumns.GetHeight: integer;
begin
  Result := Round((FDefaultHeight / 20) * (FSkin.GDI.PixelsPerInchY / 72)) + FGutter.Levels * GUTTER_SIZE;
end;

destructor TXLSBookColumns.Destroy;
begin
  inherited;
end;

procedure TXLSBookColumns.SetSize(const pX1, pY1, pX2, pY2: integer);
begin
  inherited SetSize(pX1, pY1 ,pX2, pY2);
  SetClientSize(pX1, pY1 + FGutter.Levels * GUTTER_SIZE, pX2, pY2);
  FGutter.SetSize(pX1 - 1, pY1, pX2, pY1 + FGutter.Levels * GUTTER_SIZE - 1);
  CalcHeaders;
end;

procedure TXLSBookColumns.ButtonClick(Sender: TObject; Hdr, Level: integer);
begin
  if Hdr >= 0 then
    FXLSColumns.ToggleGrouped(Hdr);
  CalcHeaders;
  if (Hdr >= 0) and Assigned(FSyncButtonClickEvent) then
    FSyncButtonClickEvent(Self,-1,-1);
  FParent.Invalidate;
end;

procedure TXLSBookColumns.ColPos(Col: integer; var X1, X2: integer);
begin
  X1 := FHeaders[Col].Pos;
  X2 := X1;
  if not (hoHidden in FHeaders[Col].Options) then
    Inc(X2,FHeaders[Col].Size);
end;

procedure TXLSBookColumns.ColPos(Col1, Col2: integer; var X1, X2: integer);
begin
  X1 := FHeaders[Col1].Pos;
  X2 := FHeaders[Col2].Pos;
  if not (hoHidden in FHeaders[Col2].Options) then
    Inc(X2,FHeaders[Col2].Size);
end;

procedure TXLSBookColumns.PaintHeader(Hdr: integer);
var
  x: integer;
begin
  FSkin.SetSystemFont(False);
  FSkin.GetCellFontData;
  FSkin.TextAlign(ptaCenter);
  x := FHeaders[Hdr].Pos;
  FSkin.PaintHeader(x,FCY1,x + FHeaders[Hdr].Size - 1,FCY2,FHeaders[Hdr].SelState,xshsNormal,FHeaders[Hdr].Hoover,ColToRefStr(FFirstHdr + Hdr));
end;

function TXLSBookColumns.WidthToPixels(W: integer): integer;
begin
  Result := Round((W / 256) * FStdFontWidth);
end;

{ TXColsGutter }

procedure TXColsGutter.CentrePos(P1, P2, Delta: integer; var cx, cy: integer);
begin
  cx := FCX1 + P1 + (P2 - P1) div 2;
  cy := FCY1 + GUTTER_SIZE div 2 + Delta;
end;

procedure TXColsGutter.PaintOutline(P1,P2,Level: integer; LineType: TOutlineLineType);
var
  cx,cy: integer;
begin
  Inc(P1,FCX1);
  Inc(P2,FCX1);
  cx := P1 + (P2 - P1) div 2;
  cy := FCY1 + Level * GUTTER_SIZE + GUTTER_SIZE div 2;
  case LineType of
    olpFirst: begin
      FSkin.GDI.Rectangle(P1 + 2,cy - 1,P2 + 1,cy);
      FSkin.GDI.Rectangle(P1 + 2,cy + 1,P1 + 3,cy + 4);
    end;
    olpMiddle: FSkin.GDI.Rectangle(P1 + 2,cy - 1,P2 + 1,cy);
    olpLast  : FSkin.GDI.Rectangle(P1 + 2,cy - 1,cx,cy);
    olpDot   : FSkin.GDI.Rectangle(cx - 1,cy - 1,cx + 1,cy + 1);
  end;
end;

end.
