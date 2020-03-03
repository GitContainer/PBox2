unit XBookPaint2;

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

uses { Delphi  } Classes, SysUtils, Types, vcl.Graphics, Contnrs, Math, Windows,
     { XLSRWII } XLSUtils5,
                 Xc12DataStyleSheet5,
     { XBook   } XBookTypes2, XBookPaintGDI2;

const SIZE_LINEWIDTH_2 = 2;

const FONT_SMALL = 'small fonts';

type TPaintTextAlign = (ptaLeft,ptaCenter,ptaRight);

type TPaintCellHorizAlign = (pchaGeneral,pchaLeft,pchaCenter,pchaRight,pchaFill,pchaJustify,pchaCenterAcross);
type TPaintCellVertAlign = (pcvaTop,pcvaCenter,pcvaBottom,pcvaJustify,pcvaDistributed);

type TCellTextType = (cttText,cttTextWordBreak,cttNumeric,cttBoolean,cttError);

type TXLSBookPicture = class(TObject)
protected
     FId: integer;
public
     procedure Draw(GDI: TAXWGDI; X1,Y1,X2,Y2: integer); virtual; abstract;
     property Id: integer read FId write FId;
     end;

type TXLSBookPictureBMP = class(TXLSBookPicture)
protected
     FBitmap: vcl.Graphics.TBitmap;
public
     constructor Create;
     destructor Destroy; override;
     procedure Draw(GDI: TAXWGDI; X1,Y1,X2,Y2: integer); override;

     property Bitmap: vcl.Graphics.TBitmap read FBitmap;
     end;

type TXLSBookPictureMetafile = class(TXLSBookPicture)
protected
     FIsEMF: boolean;
     FBuffer: Pointer;
     FBufferSize: integer;
public
     procedure Draw(GDI: TAXWGDI; X1,Y1,X2,Y2: integer); override;
     procedure SetBuffer(Buf: Pointer; BufSize: integer; EMF: boolean);
     end;

type TXLSBookPictures = class(TObjectList)
private
     function GetItems(Index: integer): TXLSBookPicture;
public
     destructor Destroy; override;
     procedure Add(Picture: TXLSBookPicture);

     property Items[Index: integer]: TXLSBookPicture read GetItems; default;
     end;

type TXStockFontHandle = (xsfhCurrCell,xsfhSmallCell,xsfhSys,xsfhSysBold,xsfhTab,xsfhTabBold);

type THeaderFooterData = record
     PageNumber: integer;
     PageCount: integer;
     SheetName: AxUCString;
     Filename: AxUCString;
     end;

type TXLSBookPaint = class(TObject)
private
     function  GetFontRotation: integer;
     procedure SetFontRotation(Value: integer);
protected
     FGDI: TAXWGDI;
     FCanvasWidth,FCanvasHeight: integer;
     FBMPWidth,FBMPHeight: integer;
     FBmpBits: Pointer;
     FFontHandles: array[Low(TXStockFontHandle)..High(TXStockFontHandle)] of longword;
     FFontRotateHandle: longword;
     FCurrFont: TXStockFontHandle;
     FPrintFontHandle: longword;
     FSysFontSize: integer;
     FFontColor: longword;
     FFontRotationRad: double;
     // Text is stacked when rotation is set to 255.
     FStackedText: boolean;
     FFontUnderline: TXc12Underline;
     FPictures: TXLSBookPictures;
     FHorizTextAlign: TPaintCellHorizAlign;
     FVertTextAlign: TPaintCellVertAlign;
     FCellLeftMarg,FCellRightMarg,
     FCellTopMarg,FCellBottomMarg: integer;

     procedure IntPrintHeaderFooter(HFData: THeaderFooterData; S: AxUCString; Margins: TRect; IsHeader: boolean; PassCount: integer; RowSizes: array of TPoint);
public
     constructor Create(AGDI: TAXWGDI);
     destructor Destroy; override;
     procedure Border(x1,y1,x2,y2,Width: integer);
     procedure SetDefaultFont;
     procedure SetSystemFont(Bold: boolean);
     procedure SetTabFont(Bold: boolean);
     procedure AssignSystemFont(Name: AxUCString; Size, CharSet: integer);
     procedure SetCellFont(AFont: TXc12Font);
     procedure TextAlign(Align: TPaintTextAlign);
     procedure CellTextAlign(HorizAlign: TPaintCellHorizAlign; VertAlign: TPaintCellVertAlign; Rotation: integer);
     procedure PaintHMoveLine(x1, y, x2: integer);
     procedure PaintVMoveLine(x, y1, y2: integer);
     procedure CellTextRect(XYRect: TXYRect; Indent: integer; TextType: TCellTextType; Text: AxUCString);
     procedure CellTextRectWrap(XYRect: TXYRect; Indent: integer; Text: AxUCString);
{
     procedure RichTextRect(X1,Y1,X2,Y2: integer; Text: AxUCString; FontRuns: TDynTXORUNArray); overload;
     procedure RichTextRect(XYRect: TXYRect; Cell: TCell; FR: PFontRun; FontRunCount: integer; Text: AxUCString); overload;
     function  RichTextWidth(Cell: TCell; FR: PFontRun; FontRunCount: integer; Text: AxUCString): integer;
}
     procedure PrintHeaderFooter(HFData: THeaderFooterData; Text: AxUCString; Margins: TRect; IsHeader: boolean);
     procedure PaintImage(Id,X1,Y1,X2,Y2: integer);
     function  ImageValid(Id: integer): boolean;
//     procedure CopyArea(BMP: TXLSBitmap; X1,Y1,X2,Y2: integer);
//     procedure RestoreArea(BMP: TXLSBitmap; X,Y: integer);
     procedure GetCellFontData;
     procedure SetSize(Width, Height: integer);
     procedure Clear;
     procedure PaintArrow(X,Y,Scale: integer; Angle: double); overload;
     procedure PaintArrow(P1,P2: TXYPoint; Scale: integer); overload;

     procedure Test;

     property GDI: TAXWGDI read FGDI write FGDI;
     property Pictures: TXLSBookPictures read FPictures;
     property CellLeftMarg: integer read FCellLeftMarg;
     property CellRightMarg: integer read FCellRightMarg;
     property CellTopMarg: integer read FCellTopMarg;
     property CellBottomMarg: integer read FCellBottomMarg;
     property FontRotation: integer read GetFontRotation write SetFontRotation;
     property FontUnderline: TXc12Underline read FFontUnderline write FFontUnderline;
     property SysFontSize: integer read FSysFontSize;
     end;

implementation

type PPoints = ^TPoints;
     TPoints = array[0..0] of TPoint;


{ TXLSBookPaint }

constructor TXLSBookPaint.Create(AGDI: TAXWGDI);
begin
  FGDI := AGDI;
  FFontColor := $000000;

  FFontHandles[xsfhTab] := FGDI.CreateFont('Arial',8,False,False,False,0);
  FFontHandles[xsfhTabBold] := FGDI.CreateFont('Arial',8,True,False,False,0);

  FPictures := TXLSBookPictures.Create;
end;

destructor TXLSBookPaint.Destroy;
var
  i: integer;
begin
  FPictures.Free;
  for i := 1 to Length(FFontHandles) - 1 do
    FGDI.DeleteFont(FFontHandles[TXStockFontHandle(i)]);
  FGDI.DeleteFont(FFontRotateHandle);
  if FPrintFontHandle <> 0 then begin
    Windows.DeleteObject(FPrintFontHandle);
    FPrintFontHandle := 0;
  end;
  inherited;
end;

procedure TXLSBookPaint.SetCellFont(AFont: TXc12Font);
var
  H: longword;
  LF: LOGFONT;
begin
  if FFontRotateHandle <> 0 then begin
    FGDI.DeleteFont(FFontRotateHandle);
    FFontRotateHandle := 0;
    FFontHandles[xsfhCurrCell] := 0;
    FFontRotationRad := 0;
  end;

  if FPrintFontHandle <> 0 then begin
    DeleteObject(FPrintFontHandle);
    FPrintFontHandle := 0;
  end;

  H := AFont.GetHandle(FGDI.PixelsPerInchY);
  Windows.GetObject(H,SizeOf(LOGFONT),@LF);

  // This check causes problems if the zoom not is 100%
//  if Round(LF.lfHeight * FGDI.Zoom) <> Cell.FontHeight(FGDI.PixelsPerInch) then
//    raise XLSRWException.CreateFmt('Font error in SetCellFont (%d,%d)',[Cell.FontHeight(FGDI.PixelsPerInch),FGDI.PixelsPerInch]);
  FCurrFont := xsfhCurrCell;
  if H <> FGDI.GetFont {FFontHandles[xsfhCurrCell]} then begin
    FFontHandles[xsfhCurrCell] := H;
//    Move(FFontTmp,FFont,SizeOf(FFont));
    SetDefaultFont;
    FGDI.CheckMinFontSize;
    GetCellFontData;
  end;

{
  Cell.CopyToLOGFONT(FFontTmp,FFontColor);
  FFontSize := Round(FFontTmp.lfHeight / 20);
  FFontTmp.lfHeight := -Round((FFontTmp.lfHeight / 20) * PixelsPerInch / 72);
  if not CompareMem(@FFontTmp,@FFont,SizeOf(FFont)) then begin
    Move(FFontTmp,FFont,SizeOf(FFont));
    SetDefaultFont;
    CheckMinFontSize;
    GetCellFontData;
  end;
}
end;

procedure TXLSBookPaint.SetDefaultFont;
begin
  FCurrFont := xsfhCurrCell;
  if FFontHandles[xsfhCurrCell] <> 0 then
    FGDI.SelectFont(FFontHandles[xsfhCurrCell]);
end;

procedure TXLSBookPaint.SetFontRotation(Value: integer);
begin
  // Check this carefully
  if Value <> FGDI.FontRotation then begin
    if FFontRotateHandle <> 0 then begin
      FGDI.DeleteFont(FFontRotateHandle);
      FFontRotateHandle := 0;
    end;

    FStackedText := Value = 255;
    if not FStackedText then begin
      if Value > 90 then
        Value := -(Value - 90);
      FFontRotationRad := DegToRad(Value);
      FFontRotateHandle := FGDI.SetFontRotation(FFontHandles[xsfhCurrCell],Value);
    end;
  end;
end;

procedure TXLSBookPaint.CellTextAlign(HorizAlign: TPaintCellHorizAlign; VertAlign: TPaintCellVertAlign; Rotation: integer);
begin
  FHorizTextAlign := HorizAlign;
  FVertTextAlign := VertAlign;
  FontRotation := Rotation;
end;

procedure TXLSBookPaint.SetSize(Width, Height: integer);
begin
  FCanvasWidth := Width;
  FCanvasHeight := Height;
end;

procedure TXLSBookPaint.SetSystemFont(Bold: boolean);
begin
  if Bold then begin
    FGDI.SelectFont(FFontHandles[xsfhSysBold]);
    FCurrFont := xsfhSysBold;
  end
  else begin
    FGDI.SelectFont(FFontHandles[xsfhSys]);
    FCurrFont := xsfhSys;
  end;
  FGDI._GetTextMetricSys;
end;

procedure TXLSBookPaint.SetTabFont(Bold: boolean);
begin
  if Bold then begin
    FGDI.SelectFont(FFontHandles[xsfhTabBold]);
    FCurrFont := xsfhTabBold;
  end
  else begin
    FGDI.SelectFont(FFontHandles[xsfhTab]);
    FCurrFont := xsfhTab;
  end;
end;

procedure TXLSBookPaint.PaintArrow(X, Y, Scale: integer; Angle: double);
const
  Arrow1: array[0..2] of TPoint = ((X:0; Y:0),(X:7; Y:-3),(X:7; Y:3));
var
  i: integer;
  Arr: array[0..2] of TPoint;
begin
  for i := 0 to 2 do begin
    Arr[i].X := (Arrow1[i].X * Scale) + X;
    Arr[i].Y := (Arrow1[i].Y * Scale) + Y;
  end;
  FGDI.Trans.Clear;
  FGDI.Trans.Rotate(X,Y,Angle);
  FGDI.Trans.Transform(Arr);
  FGDI.Polygon(Arr);
end;

procedure TXLSBookPaint.PaintArrow(P1, P2: TXYPoint; Scale: integer);
var
  A: double;
begin
  if (P2.X - P1.X) = 0 then
    A := 0
  else
    A := (ArcTan2(P2.X - P1.X,P2.Y - P1.Y)  * (180 / PI));
  if P1.X < P2.X then
    A := A - 90
  else
    A := A + 270;
  PaintArrow(P1.X,P1.Y,Scale,A);
end;

procedure TXLSBookPaint.PaintHMoveLine(x1, y, x2: integer);
begin
  FGDI.Rectangle(x1,y - SIZE_LINEWIDTH_2,x2 - 1,y + SIZE_LINEWIDTH_2 - 1);
end;

procedure TXLSBookPaint.PaintImage(Id,X1, Y1, X2, Y2: integer);
begin
  if FPictures[Id] <> Nil then
    FPictures[Id].Draw(FGDI,X1, Y1, X2, Y2);
end;

procedure TXLSBookPaint.PaintVMoveLine(x, y1, y2: integer);
begin
  FGDI.Rectangle(x - SIZE_LINEWIDTH_2,y1,x + SIZE_LINEWIDTH_2 - 1,y2 - 1);
end;

procedure TXLSBookPaint.PrintHeaderFooter(HFData: THeaderFooterData; Text: AxUCString; Margins: TRect; IsHeader: boolean);
var
  S: AxUCString;
  p,PassCount: integer;
  RowSizes: array[0..9] of TPoint;

function PosEndSection(S: AxUCString): integer;
var
  p1,p2,p3: integer;
begin
  S := Copy(S,2,MAXINT);
  p1 := Pos(AxUCString('&L'),S);
  if p1 < 1 then p1 := MAXINT;
  p2 := Pos(AxUCString('&C'),S);
  if p2 < 1 then p2 := MAXINT;
  p3 := Pos(AxUCString('&R'),S);
  if p3 < 1 then p3 := MAXINT;
  Result := Min(p1,Min(p2,p3));
end;


begin
  if Text = '' then
    Exit;
  PassCount := 0;
  if Copy(Text,1,1) <> '&' then
    Text := '&L' + Text;
  while Text <> '' do begin
    p := PosEndSection(Text);
    S := Copy(Text,1,p);
    IntPrintHeaderFooter(HFData,S,Margins,IsHeader,PassCount,RowSizes);
    if p = MAXINT then
      Text := ''
    else
      Text := Copy(Text,p + 1,MAXINT);
  end;
end;

procedure TXLSBookPaint.IntPrintHeaderFooter(HFData: THeaderFooterData; S: AxUCString; Margins: TRect; IsHeader: boolean; PassCount: integer; RowSizes: array of TPoint);
var
  p,FontSz: integer;
  YPos,XPos: integer;
  SSPos,SSSize: TPoint;
  C,AlignC: Char;
  T,S2: AxUCString;
  OrigS: AxUCString;
  Text: AxUCString;
  FntBold,FntItalic,FntUnderline,FntStrikeOut: boolean;
  FntSubSuperscript: integer;
  Row: integer;
  VertAlign: TXGDIVertTextAlign;

procedure ReserAlign;
begin
  case AlignC of
    'L': begin
      FGDI.SetTextAlign(xhtaLeft,VertAlign);
      XPos := Margins.Left;
     end;
    'C': begin
//      FGDI.SetTextAlign(xhtaLeft,VertAlign);
//      XPos := FGDI._Width div 2 - RowSizes[Row].X div 2;
    end;
    'R': begin
      FGDI.SetTextAlign(xhtaRight,VertAlign);
      XPos := Margins.Right;
    end;
  end;
end;

procedure ResetFontStyles;
begin
  FntBold := False;
  FntItalic := False;
  FntUnderline := False;
  FntStrikeOut := False;
  FntSubSuperscript := 0;
  SSPos.X := 0;
  SSPos.Y := 0;
end;

procedure WriteText(Txt: AxUCString);
var
  p: integer;
begin
  p := CPos(AXUCChar($000A),Txt);
  if p > 0 then begin
    WriteText(Copy(Txt,1,p - 1));
    Txt := Copy(Txt,p + 1,MAXINT);
    if IsHeader then
      Inc(YPos,RowSizes[Row + 1].Y)
    else
      Inc(YPos,RowSizes[Row].Y);
    Inc(Row);
    ReserAlign;
  end;
  if PassCount > 0 then
    FGDI.TextOut(XPos + SSPos.X,YPos + SSPos.Y,Txt)
  else begin
    Inc(RowSizes[Row].X,FGDI.TextWidth(Txt));
    RowSizes[Row].Y := Max(FGDI.TextHeight(Txt),RowSizes[Row].Y);
  end;
  case AlignC of
    'L': Inc(XPos,FGDI.TextWidth(Txt));
    'C': Inc(XPos,FGDI.TextWidth(Txt));
    'R': Dec(XPos,FGDI.TextWidth(Txt));
  end;
end;

begin
  Text := '';
  Row := 0;
  SetDefaultFont;
  FontSz := FGDI.GetFontSize;
  if IsHeader then
    VertAlign := xvtaBaseline
  else
    VertAlign := xvtaTop;
  if PassCount <= 0 then begin
    OrigS := S;
    for p := 0 to High(RowSizes) do begin
      RowSizes[p].X := 0;
      RowSizes[p].Y := 0;
    end;
  end;
  if IsHeader then
    YPos := Margins.Top + RowSizes[0].Y
  else begin
    YPos := Margins.Bottom;
    p := 0;
    while p <= High(RowSizes) do begin
      Dec(YPos,RowSizes[p].Y);
      Inc(p);
    end;
  end;
  AlignC := 'L';
  ResetFontStyles;

  p := CPos('&',S);
  while p > 0 do begin
    if Copy(S,2,1) <> '' then
      C := Char(Copy(S,2,1)[1])
    else
      C := #0;

    S := Copy(S,p + 2,MAXINT);
    p := CPos('&',S);
    if p > 0 then begin
      T := Copy(S,1,p - 1);
      S := Copy(S,p,MAXINT);
    end
    else
      T := S;
    case C of
      '&': WriteText('&' + T);
      'L','C','R': begin
        AlignC := C;
        ResetFontStyles;
        ReserAlign;
        WriteText(T);
      end;
      'P': WriteText(IntToStr(HFData.PageNumber) + T);
      'N': WriteText(IntToStr(HFData.PageCount) + T);
      'D': WriteText(DateToStr(Now) + T);
      'T': WriteText(TimeToStr(Now) + T);
      'A': WriteText(HFData.SheetName + T);
      'F': WriteText(ExtractFilename(HFData.Filename));
      'Z': WriteText(ExtractFilePath(HFData.Filename));
      'S': begin
        FntStrikeOut := not FntStrikeOut;
        FGDI.SetFontStrikeOut(FntStrikeOut);
        WriteText(T);
      end;
      'X': begin
        if FntSubSuperscript = 0 then begin
          FGDI.GetFontSuperscriptData(SSPos,SSSize);
          FntSubSuperscript := 1;
          FGDI.SetFontSize(FGDI.FntPixHeightToPt(SSSize.X));
        end
        else begin
          SSPos.X := 0;
          SSPos.Y := 0;
          FGDI.SetFontSize(FontSz);
        end;
        WriteText(T);
      end;
      'Y': begin
        if FntSubSuperscript = 0 then begin
          FGDI.GetFontSubscriptData(SSPos,SSSize);
          FntSubSuperscript := -1;
          FGDI.SetFontSize(FGDI.FntPixHeightToPt(SSSize.Y));
        end
        else begin
          SSPos.X := 0;
          SSPos.Y := 0;
          FGDI.SetFontSize(FontSz);
        end;
        WriteText(T);
      end;
      '"': begin
        T := Copy(T,1,Length(T) - 1);
        S2 := SplitAtChar(',',T);
        FGDI.SelectFont(FGDI.CreateFontByStr(S2,T,FGDI.GetFontSize));
      end;
      'B': begin
        FntBold := not FntBold;
        FGDI.SetFontStyle(FntBold,FntItalic,FntUnderline);
        WriteText(T);
      end;
      'I': begin
        FntItalic := not FntItalic;
        FGDI.SetFontStyle(FntBold,FntItalic,FntUnderline);
        WriteText(T);
      end;
      // TODO E shall be double underlining
      'E','U': begin
        FntUnderline := not FntUnderline;
        FGDI.SetFontStyle(FntBold,FntItalic,FntUnderline);
        WriteText(T);
      end;
      '1'..'9': begin
        p := 1;
        while (p <= Length(T)) and CharInSet(T[p],['0'..'9']) do
          Inc(p);
        FontSz := StrToIntDef(C + Copy(T,1,p - 1),8);
        T := Copy(T,p,MAXINT);
        if FontSz > 0 then
          FGDI.SetFontSize(FontSz);
        WriteText(T);
      end;
    end;
    T := '';
    p := CPos('&',S);
  end;
  if PassCount = 0 then begin
    Inc(PassCount);
    IntPrintHeaderFooter(HFData,OrigS,Margins,IsHeader,PassCount,RowSizes);
  end;
end;

{
procedure TXLSBookPaint.RichTextRect(XYRect: TXYRect; Cell: TCell; FR: PFontRun; FontRunCount: integer; Text: AxUCString);
var
  IE: TXBookInplaceEditor;
  R: TRect;
begin
  R := Rect(XYRect.X1,XYRect.Y1,XYRect.X2,XYRect.Y2);
  IE := TXBookInplaceEditor.Create(FGDI);
  try
    IE.AssignText(Text,Cell.FontIndex,FR,FontRunCount);
    IE.DrawText(XYRect.X1,XYRect.Y1,XYRect.X2,XYRect.Y2);
    FGDI._SetFontByFontIndex(0);
  finally
    IE.Free;
  end;
end;

procedure TXLSBookPaint.RichTextRect(X1, Y1, X2, Y2: integer; Text: AxUCString; FontRuns: TDynTXORUNArray);
var
  IE: TXBookInplaceEditor;
begin
  IE := TXBookInplaceEditor.Create(FGDI);
  try
    IE.AssignText(Text,FontRuns);
    IE.DrawText(X1,Y1,X2,Y2);
    FGDI._SetFontByFontIndex(0);
  finally
    IE.Free;
  end;
end;

function TXLSBookPaint.RichTextWidth(Cell: TCell; FR: PFontRun; FontRunCount: integer; Text: AxUCString): integer;
var
  IE: TXBookInplaceEditor;
begin
  IE := TXBookInplaceEditor.Create(FGDI);
  try
    IE.AssignText(Text,Cell.FontIndex,FR,FontRunCount);
    Result := IE.TextWidth;
    FGDI._SetFontByFontIndex(0);
  finally
    IE.Free;
  end;
end;
}

procedure TXLSBookPaint.Test;
begin
end;

procedure TXLSBookPaint.TextAlign(Align: TPaintTextAlign);
begin
  case Align of
    ptaLeft:   FGDI.SetTextAlign(xhtaLeft,xvtaNone);
    ptaCenter: FGDI.SetTextAlign(xhtaCenter,xvtaBaseline);
    ptaRight:  FGDI.SetTextAlign(xhtaRight,xvtaNone);
  end;
end;

procedure TXLSBookPaint.AssignSystemFont(Name: AxUCString; Size, CharSet: integer);
begin
  FGDI.DeleteFont(FFontHandles[xsfhSys]);
  FFontHandles[xsfhSys] := FGDI.CreateFont(Name,Size,False,False,False,CharSet);
  FGDI.DeleteFont(FFontHandles[xsfhSysBold]);
  FFontHandles[xsfhSysBold] := FGDI.CreateFont(Name,Size,True,False,False,CharSet);
  FSysFontSize := Size;
end;

procedure TXLSBookPaint.Border(x1, y1, x2, y2, Width: integer);
begin
  if Width <= 1 then begin
    FGDI.BrushSolid := False;
    FGDI.Rectangle(x1,y1,x2,y2);
    FGDI.BrushSolid := True;
  end
  else begin
    FGDI.Rectangle(x1,y1,x2,y1 + Width - 1);
    FGDI.Rectangle(x2 - Width + 1,y1,x2,y2);
    FGDI.Rectangle(x1,y2 - Width,x2,y2);
    FGDI.Rectangle(x1,y1,x1 + Width - 1,y2 - 1);
  end;
end;

{
procedure TXLSBookPaint.CopyArea(BMP: TXLSBitmap; X1, Y1, X2, Y2: integer);
begin
  BMP.Width := (X2 - X1) + 1;
  BMP.Height := (Y2 - Y1) + 1;
  FCanvas.CopyMode := cmSrcCopy;
  BMP.Canvas.CopyRect(Rect(0,0,BMP.Width,BMP.Height),FCanvas,Rect(X1,Y1,X2,Y2));
end;

procedure TXLSBookPaint.RestoreArea(BMP: TXLSBitmap; X, Y: integer);
begin
  FCanvas.CopyMode := cmSrcCopy;
  FCanvas.CopyRect(Rect(X,Y,X + BMP.Width,Y + BMP.Height),BMP.Canvas,Rect(0,0,BMP.Width,BMP.Height));
end;
}

procedure TXLSBookPaint.CellTextRect(XYRect: TXYRect; Indent: integer; TextType: TCellTextType; Text: AxUCString);
var
  i: integer;
  X,Y: integer;
  TW: integer;
  R: TRect;
  HorizAlign: TPaintCellHorizAlign;
begin
  X := 0;
  Y := 0;
  if (FHorizTextAlign = pchaGeneral) and (FFontRotationRad = 0) then begin
    if FStackedText then
      HorizAlign := pchaCenter
    else begin
      case TextType of
        cttNumeric : HorizAlign := pchaRight;
        cttBoolean,
        cttError   : HorizAlign := pchaCenter;
        else begin
          if (Text <> '') and (FGDI.CharIsRTL(Text[1])) then
            HorizAlign := pchaRight
          else
            HorizAlign := pchaLeft;
        end;
      end;
    end;
  end
  else
    HorizAlign := FHorizTextAlign;
  FGDI.TransparentMode := True;
  if FFontRotationRad <> 0 then begin
    if HorizAlign > pchaRight then
      HorizAlign := pchaLeft;
    if FGDI.FontRotation <= -90 then begin
      case HorizAlign of
        pchaLeft,pchaGeneral: begin
          FGDI.SetTextAlign(xhtaRight,xvtaBaseline);
          X := XYRect.X1 + FCellLeftMarg;
          Y := XYRect.Y2 - FCellBottomMarg;
        end;
        pchaCenter: begin
          FGDI.SetTextAlign(xhtaRight,xvtaBottom);
          X := XYRect.X1 + ((XYRect.X2 - XYRect.X1) div 2) - FGDI.TM.Height div 2;
          Y := XYRect.Y2 - FCellBottomMarg;
        end;
        pchaRight: begin
          FGDI.SetTextAlign(xhtaRight,xvtaTop);
          X := XYRect.X2 - FCellRightMarg;
          Y := XYRect.Y2 - FCellBottomMarg;
        end;
      end;
      case FVertTextAlign of
        pcvaTop         : Y := XYRect.Y1 + FGDI.TextWidth(Text) + FCellTopMarg;
        pcvaCenter      : Y := XYRect.Y1 + ((XYRect.Y2 - XYRect.Y1) div 2) + FGDI.TextWidth(Text) div 2;
        pcvaBottom      : Y := XYRect.Y2 - FCellBottomMarg;
      end;
    end
    else if FFontRotationRad < 0 then begin
      case HorizAlign of
        pchaLeft: begin
          FGDI.SetTextAlign(xhtaRight,xvtaBaseline);
          X := XYRect.X1 + Round(FGDI.TextWidth(Text) * Cos(FFontRotationRad)) + FCellLeftMarg;
        end;
        pchaCenter: begin
          FGDI.SetTextAlign(xhtaRight,xvtaBaseline);
          X := XYRect.X1 + Round(FGDI.TM.Height * Cos(FFontRotationRad)) + ((XYRect.X2 - XYRect.X1) div 2) - FCellLeftMarg
        end;
        pchaRight,pchaGeneral: begin
          FGDI.SetTextAlign(xhtaRight,xvtaBaseline);
          X := XYRect.X2 + Round(FGDI.TM.Height * Sin(FFontRotationRad)) - FCellRightMarg;
        end;
      end;
      case FVertTextAlign of
        pcvaTop         : Y := XYRect.Y1 + Round(FGDI.TextWidth(Text) * Sin(-FFontRotationRad) + FGDI.TM.Ascent * Cos(-FFontRotationRad)) +  FCellTopMarg;
        pcvaCenter      : Y := XYRect.Y1 + ((XYRect.Y2 - XYRect.Y1) div 2) - Round((FGDI.TextWidth(Text) div 2) * Sin(FFontRotationRad)) + FCellLeftMarg;
        pcvaBottom      : Y := XYRect.Y2 - FCellBottomMarg;
      end;
    end
    else if FGDI.FontRotation < 90 then begin
      case HorizAlign of
        pchaLeft,pchaGeneral: begin
          FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
          X := XYRect.X1 + Round(FGDI.TM.Height * Sin(FFontRotationRad)) + FCellLeftMarg;
        end;
        pchaCenter: begin
          FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
          X := XYRect.X1 - Round(FGDI.TM.Height * Cos(FFontRotationRad)) + ((XYRect.X2 - XYRect.X1) div 2) + FCellLeftMarg;
        end;
        pchaRight: begin
          FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
          X := XYRect.X1 + Round(FGDI.TM.Ascent * Sin(FFontRotationRad));
        end;
      end;
      case FVertTextAlign of                      //
        pcvaTop         : Y := XYRect.Y1 + Round(FGDI.TextWidth(Text) * Sin(FFontRotationRad) + FGDI.TM.Ascent * Cos(FFontRotationRad)) + FCellTopMarg;
        pcvaCenter      : Y := XYRect.Y1 + ((XYRect.Y2 - XYRect.Y1) div 2) + Round((FGDI.TextWidth(Text) div 2) * Sin(FFontRotationRad)) + FCellTopMarg;
        pcvaBottom      : Y := XYRect.Y2 - FCellBottomMarg;
      end;
    end
    else begin
      case HorizAlign of
        pchaLeft: begin
          FGDI.SetTextAlign(xhtaLeft,xvtaTop);
          X := XYRect.X1 + FCellLeftMarg;
        end;
        pchaCenter: begin
          FGDI.SetTextAlign(xhtaLeft,xvtaBottom);
          X := XYRect.X1 + ((XYRect.X2 - XYRect.X1) div 2) + FGDI.TM.Height div 2;
        end;
        pchaRight,pchaGeneral: begin
          FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
          X := XYRect.X2 - FGDI.TM.Descent;
        end;
      end;
      case FVertTextAlign of
        pcvaTop         : Y := XYRect.Y1 + FGDI.TextWidth(Text) + FCellTopMarg;
        pcvaCenter      : Y := XYRect.Y1 + ((XYRect.Y2 - XYRect.Y1) div 2) + FGDI.TextWidth(Text) div 2;
        pcvaBottom      : Y := XYRect.Y2 - FCellBottomMarg;
      end;
    end;
  end
  else begin
    case HorizAlign of
      pchaLeft: begin
        FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
        X := XYRect.X1 + FCellLeftMarg + Indent * FGDI.TM.AveCharWidth;
      end;
      pchaCenter: begin
        FGDI.SetTextAlign(xhtaCenter,xvtaBaseline);
        X := XYRect.X1 + ((XYRect.X2 - XYRect.X1) div 2) + 1;
//        FGDI.SetTextAlign(xvtaLeft,xhtaBaseline);
//        X := XYRect.X1 + ((XYRect.X2 - XYRect.X1) div 2) - FGDI.TextWidth(Text) div 2;
      end;
      pchaRight: begin
        FGDI.SetTextAlign(xhtaRight,xvtaBaseline);
        X := XYRect.X2 - FCellRightMarg - Indent * FGDI.TM.AveCharWidth;
      end;
      pchaFill: begin
        FGDI.SetTextAlign(xhtaLeft,xvtaBaseline);
        X := XYRect.X1 + FCellLeftMarg;
      end;
      pchaJustify: begin
        FGDI.SetTextAlign(xhtaLeftNoUpdateCP,xvtaTop);
      end;
      pchaCenterAcross: begin
        FGDI.SetTextAlign(xhtaCenter,xvtaBaseline);
        X := XYRect.X1 + ((XYRect.X2 - XYRect.X1) div 2);
      end;
    end;
    case FVertTextAlign of
      pcvaTop         : Y := XYRect.Y1 + FGDI.TM.Ascent{ - FCellTopMarg};
      pcvaCenter      : Y := XYRect.Y1 + ((XYRect.Y2 - XYRect.Y1) div 2) + FGDI.TM.Ascent div 2 - FGDI.TM.Descent div 2;
      pcvaBottom      : Y := XYRect.Y2 - FCellBottomMarg;
      pcvaJustify     : Y := XYRect.Y2 - FCellBottomMarg;
      pcvaDistributed : Y := XYRect.Y2 - FCellBottomMarg;
    end;
  end;
  R := Rect(XYRect.X1,XYRect.Y1,XYRect.X2,XYRect.Y2);
  if FStackedText then begin
    case FVertTextAlign of
      pcvaJustify,
      pcvaTop        : begin
        for i := 1 to Length(Text) do begin
          FGDI.TextOut(X,Y,R,Text[i]);
          Inc(Y,FGDI.TM.Height);
        end;
      end;
      pcvaDistributed,
      pcvaCenter     : begin
        Dec(Y,(Length(Text) * FGDI.TM.Height) div 2 - FGDI.TM.Ascent);
        for i := 1 to Length(Text) do begin
          FGDI.TextOut(X,Y,R,Text[i]);
          Inc(Y,FGDI.TM.Height);
        end;
      end;
      pcvaBottom     : begin
        for i := Length(Text) downto 1 do begin
          FGDI.TextOut(X,Y,R,Text[i]);
          Dec(Y,FGDI.TM.Height);
        end;
      end;
    end;
  end
  else begin
    case HorizAlign of
      pchaFill: begin
        TW := FGDI.TextWidth(Text);
        while (X + TW) < XYRect.X2 do begin
          FGDI.TextOut(X,Y,R,Text);
          Inc(X,TW);
        end;
      end;
      pchaJustify: begin
        FGDI.DrawText(R,Text,[dtoLeft,dtoWordbreak]);
      end
      else begin
  {
        FGDI.BrushColor := $FF0000;
        FGDI.PenColor := $FF0000;
        FGDI.Rectangle(XYRect);
  }
        FGDI.TextOut(X,Y,R,Text);
  {
        FGDI.PenColor := $000000;
        FGDI.Line(X,XYRect.Y1,X,XYRect.Y2);
  }
      end;
    end;
  end;
  if (FFontUnderline > xulSingle) and (FGDI.FontRotation = 0) and not FStackedText then begin
    TW := FGDI.TextWidth(Text);
    FGDI.GetTextMetric;
    FGDI.PenColor := FGDI.FontColor;
    FGDI.PenWidth := FGDI.TM.UnderscoreSize;
    case FFontUnderline of
      xulDouble       : begin
        Dec(Y,FGDI.TM.UnderscorePosition - 1);
        FGDI.Line(X,Y,X + TW,Y);
        Dec(Y,FGDI.TM.UnderscorePosition);
        FGDI.Line(X,Y,X + TW,Y);
      end;
      xulSingleAccount: begin
        Inc(Y,FGDI.TM.Descent);
        FGDI.Line(X,Y,X + TW,Y);
      end;
      xulDoubleAccount: begin
        Inc(Y,FGDI.TM.Descent);
        FGDI.Line(X,Y,X + TW,Y);
        Dec(Y,FGDI.TM.UnderscorePosition);
        FGDI.Line(X,Y,X + TW,Y);
      end;
    end;
  end;
end;

procedure TXLSBookPaint.CellTextRectWrap(XYRect: TXYRect; Indent: integer; Text: AxUCString);
var
  R: TRect;
  Opts: TDrawTextOptions;
begin
  Opts := [dtoWordbreak];
  case FVertTextAlign of
    pcvaTop         : Opts := Opts + [dtoTop];
    pcvaCenter      : Opts := Opts + [dtoVCenter];
    pcvaBottom      : Opts := Opts + [dtoBottom];
    pcvaDistributed : Opts := Opts + [dtoVDistributed];
  end;
  case FHorizTextAlign of
    pchaLeft        : Opts := Opts + [dtoLeft];
    pchaCenter      : Opts := Opts + [dtoHCenter];
    pchaRight       : Opts := Opts + [dtoRight];
    else              Opts := Opts + [dtoLeft];
  end;
  FGDI.SetTextAlign(xhtaLeft,xvtaTop);
  R := Rect(XYRect.X1 + FCellLeftMarg,XYRect.Y1,XYRect.X2 - FCellRightMarg,XYRect.Y2);
  FGDI.DrawText(R,Text,Opts);
end;

{
function TXLSBookPaint.CharWidth10(C: WideChar): double;
const
  Mat : tmat2 = ( eM11 : (fract :  0; value : 10);
                  eM12 : (fract :  0; value : 0);
                  eM21 : (fract :  0; value : 0);
                  eM22 : (fract :  0; value : -10); );
var
  GM : TGLYPHMETRICS;
begin
  if GetGlyphOutlineW(FDC,Longword(C),GGO_METRICS,GM,0, Nil,Mat) <> GDI_ERROR then
    Result := GM.gmCellIncX
  else
    Result := -1;
end;
}

procedure TXLSBookPaint.Clear;
begin
  FPictures.Clear;
end;

procedure TXLSBookPaint.GetCellFontData;
begin
  FGDI.GetTextMetric;
  // The value 0.2 is found by testing...
  FCellLeftMarg := Round(FGDI.TM.AveCharWidth * 0.2);
  FCellRightMarg := FCellLeftMarg;
  FCellTopMarg := FGDI.TM.InternalLeading;
  FCellBottomMarg := FGDI.TM.Descent;
end;

function TXLSBookPaint.GetFontRotation: integer;
begin
  Result := FGDI.FontRotation;
end;

function TXLSBookPaint.ImageValid(Id: integer): boolean;
begin
  Result := (Id >= 0) and (Id < FPictures.Count) and (FPictures[Id] <> Nil);
end;

{ TXLSBookPictureBMP }

constructor TXLSBookPictureBMP.Create;
begin
  FBitmap := vcl.Graphics.TBitmap.Create;
end;

destructor TXLSBookPictureBMP.Destroy;
begin
  FBitmap.Free;
  inherited;
end;

procedure TXLSBookPictureBMP.Draw(GDI: TAXWGDI; X1,Y1,X2,Y2: integer);
begin
//  GDI.PaintBMPImage(FBitmap.Canvas.Handle,FBitmap.Width,FBitmap.Height,X1,Y1,X2,Y2);
  GDI.PaintBMPImage(FBitmap,X1,Y1,X2,Y2);
end;

{ TXLSBookPictures }

procedure TXLSBookPictures.Add(Picture: TXLSBookPicture);
begin
  inherited Add(Picture);
end;

destructor TXLSBookPictures.Destroy;
begin

  inherited;
end;

function TXLSBookPictures.GetItems(Index: integer): TXLSBookPicture;
begin
  Result := TXLSBookPicture(inherited Items[Index]);
end;

{ TXLSBookPictureMetafile }

procedure TXLSBookPictureMetafile.Draw(GDI: TAXWGDI; X1, Y1, X2, Y2: integer);
begin
  if FIsEMF then
    GDI.PaintEMFImage(FBuffer,FBufferSize,X1,Y1,X2,Y2)
  else
    GDI.PaintWMFImage(FBuffer,FBufferSize,X1,Y1,X2,Y2);
end;

procedure TXLSBookPictureMetafile.SetBuffer(Buf: Pointer; BufSize: integer; EMF: boolean);
begin
  FBuffer := Buf;
  FBufferSize := BufSize;
  FIsEMF := EMF;
end;

end.
