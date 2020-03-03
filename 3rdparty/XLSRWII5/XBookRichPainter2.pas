unit XBookRichPainter2;

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

uses {* Delphi  *} Classes, SysUtils, Contnrs, Types, Math,
     {* XLSRWII *} Xc12DataStyleSheet5,
                   XLSUtils5, XLSTools5,
     {* XLSBook *} XBookTypes2, XBookPaintGDI2, XBookSkin2;

type PXBookRichRow = ^TXBookRichRow;
     TXBookRichRow = record
     CR1,CR2: integer;
     CP1,CP2: integer;
     Width  : integer;
     Height : integer;
     // Not implemented yet. Need to update the inplace editor as well.
     LineGap: integer;
     Ascent : integer;
     Descent: integer;
     end;

type TXBookRichPainter = class(TObject)
protected
     FSkin      : TXLSBookSkin;
     FXF        : TXc12XF;
     FHorizAlign: TXc12HorizAlignment;
     FVertAlign : TXc12VertAlignment;
     FTotHeight : integer;
     FPaintWidth: integer;

     FCharRuns  : TXLSFormattedText;
     FRows      : array of TXBookRichRow;
     FRowCount  : integer;

     procedure BeginRow(const ACR, ACP: integer);
     procedure UpdateRow(const ACR, ACP,AWidth,AHeight,ALineGap,AAscent,ADescent: integer);
     procedure EndRow;
     procedure Paginate;
     function  CalcX(const AX1,AX2: integer; ARow: PXBookRichRow): integer;
     function  CalcY(const AY1,AY2: integer): integer;
     function  DrawCharRun(ACR: TXLSFormattedTextItem; const ACP,ACharCount,AX,AY: integer): AxUCString;
public
     constructor Create(ASkin: TXLSBookSkin);
     destructor Destroy; override;

     procedure DrawText(const AX1,AY1,AX2,AY2: integer);
     function  TextWidth: integer;
     procedure AssignText(S: AxUCString; XF: TXc12XF; FR: PXc12FontRunArray; FontRunCount: integer);

     property HorizAlign: TXc12HorizAlignment read FHorizAlign write FHorizAlign;
     property VertAlign : TXc12VertAlignment read FVertAlign write FVertAlign;
     end;

procedure RichTextRect(ASkin: TXLSBookSkin; const AX1,AY1,AX2,AY2: double; AText: TXLSFormattedText; AHorizAlign: TXc12HorizAlignment = chaLeft; AVertAlign : TXc12VertAlignment = cvaTop); overload;
procedure RichTextRect(ASkin: TXLSBookSkin; const AX1,AY1,AX2,AY2: integer; AFontRuns: PXc12FontRunArray; const AFRCount: integer; const AText: AxUCString); overload;
procedure RichTextRect(ASkin: TXLSBookSkin; const AXYRect: TXYRect; AXF: TXc12XF; AFontRuns: PXc12FontRunArray; const AFRCount: integer; const AText: AxUCString); overload;
function  RichTextWidth(ASkin: TXLSBookSkin; AXF: TXc12XF; AFontRuns: PXc12FontRunArray; const AFRCount: integer; const AText: AxUCString): integer;

implementation

{ TXBookRichPainter }

procedure TXBookRichPainter.AssignText(S: AxUCString; XF: TXc12XF; FR: PXc12FontRunArray; FontRunCount: integer);
begin
  FXF := XF;

  if FXF <> Nil then begin
    FHorizAlign := FXF.Alignment.HorizAlignment;
    FVertAlign := FXF.Alignment.VertAlignment;
    FCharRuns.Add(S,XF.Font,FR,FontRunCount);
  end
  else begin
    FHorizAlign := chaGeneral;
    FVertAlign := cvaTop;
    FCharRuns.Add(S,Nil,FR,FontRunCount);
  end;
end;

procedure TXBookRichPainter.BeginRow(const ACR, ACP: integer);
begin
  if FRowCount >= LengtH(FRows) then
    SetLength(FRows,LengtH(FRows) + 32);
  FRows[FRowCount].CR1 := ACR;
  FRows[FRowCount].CP1 := ACP;
  FRows[FRowCount].Width := 0;
end;

function TXBookRichPainter.CalcX(const AX1, AX2: integer; ARow: PXBookRichRow): integer;
begin
  Result := 0;
  case FHorizAlign of
    chaGeneral,
    chaLeft            : Result := AX1;
    chaCenter          : Result := AX1 + ((AX2 - AX1) div 2) - (ARow.Width div 2);
    chaRight           : Result := AX2 - ARow.Width;
    chaFill            : ;
    chaJustify         : ;
    chaCenterContinuous: ;
    chaDistributed     : ;
  end;
end;

function TXBookRichPainter.CalcY(const AY1,AY2: integer): integer;
begin
  Result := 0;
  case FVertAlign of
    cvaTop        : Result := AY1;
    cvaCenter     : begin
      if FRowCount > 1 then
        Result := AY1 + ((AY2 - AY1) div 2) - (FTotHeight div 2) + Frows[0].Ascent
      else
        Result := AY1 + ((AY2 - AY1) div 2) + (Frows[0].Ascent div 2) - (Frows[0].Descent div 2);
    end;
    cvaBottom     : Result := AY2 - FTotHeight;
    cvaJustify    : ;
    cvaDistributed: ;
  end;
end;

constructor TXBookRichPainter.Create(ASkin: TXLSBookSkin);
begin
  FSkin := ASkin;
  FCharRuns := TXLSFormattedText.Create;
  FCharRuns.SplitAtCR := True;
  SetLength(FRows,64);
end;

destructor TXBookRichPainter.Destroy;
begin
  FCharRuns.Free;
  inherited;
end;

procedure TXBookRichPainter.Paginate;
var
  i: integer;
  CR: TXLSFormattedTextItem;
  CP: integer;
  W: integer;
  H: integer;
  LG: integer;
  A: integer;
  D: integer;
  S,S2: AxUCString;
begin
  if FCharRuns.Count < 1 then
    Exit;

  FTotHeight := 0;
  BeginRow(0,1);
  for i := 0 to FCharRuns.Count - 1 do begin
    CR := FCharRuns[i];

    CP := 0;
    FSkin.SetFont(CR.Font);
    FSkin.GDI.GetTextMetric;
    H := FSkin.GDI.TM.Height;
    LG := FSkin.GDI.TM.LineGap;
    A := FSkin.GDI.TM.Ascent;
    D := FSkin.GDI.TM.Descent;

    S := CR.Text;
    repeat
      S2 := SplitAtChar(' ',S);
      Inc(CP);
      W := FSkin.GDI.TextWidth(S2);
      if (FRows[FRowCount].Width + W) > FPaintWidth then begin
        EndRow;
        BeginRow(i,CP);
      end;
      Inc(CP,Length(S2));
      UpdateRow(i,CP,W,H,LG,A,D);
    until S = '';

    if CR.NewLine and (i < (FCharRuns.Count - 1)) then begin
      EndRow;
      BeginRow(i + 1,1);
    end;
  end;
  EndRow;
end;

function TXBookRichPainter.DrawCharRun(ACR: TXLSFormattedTextItem; const ACP, ACharCount, AX, AY: integer): AxUCString;
begin
  FSkin.SetFont(ACR.Font);
  Result := Copy(ACR.Text,ACP,ACharCount);
  FSkin.GDI.TextOut(AX,AY,Result);
end;

procedure TXBookRichPainter.DrawText(const AX1,AY1,AX2,AY2: integer);
var
  i,j: integer;
  S: AxUCString;
  R: PXBookRichRow;
  X: integer;
  Y: integer;
begin
  FPaintWidth := AX2 - AX1 + 1;

  Paginate;

  FSkin.GDI.SetTextAlign(xhtaLeft,xvtaBaseline);

  Y := CalcY(AY1,AY2);

  for i := 0 to FRowCount - 1 do begin
    R := @FRows[i];
    X := CalcX(AX1,AX2,R);
    Inc(Y,R.Ascent);
    if R.CR1 = R.CR2 then
      DrawCharRun(FCharRuns[R.CR1],R.CP1,R.CP2 - R.CP1 + 1,X,Y)
    else begin
      S := DrawCharRun(FCharRuns[R.CR1],R.CP1,MAXINT,X,Y);
      Inc(X,FSkin.GDI.TextWidth(S));
      for j := R.CR1 + 1 to R.CR2 - 1 do begin
        S := DrawCharRun(FCharRuns[j],1,MAXINT,X,Y);
        Inc(X,FSkin.GDI.TextWidth(S));
      end;
      DrawCharRun(FCharRuns[R.CR2],1,R.CP2,X,Y);
    end;
    Inc(Y,R.Descent);
  end;
end;

procedure TXBookRichPainter.EndRow;
begin
  Inc(FTotHeight,FRows[FRowCount].Height);
  Inc(FRowCount);
end;

function TXBookRichPainter.TextWidth: integer;
var
  i: integer;
  W: integer;
  CR: TXLSFormattedTextItem;
begin
  Result := 0;
  W := 0;
  for i := 0 to FCharRuns.Count - 1 do begin
    CR := FCharRuns[i];
    FSkin.SetFont(CR.Font);
    Inc(W,FSkin.GDI.TextWidth(CR.Text));
    if CR.NewLine then begin
      Result := Max(Result,W);
      W := 0;
    end;
  end;
  Result := Max(Result,W);
end;

procedure TXBookRichPainter.UpdateRow(const ACR, ACP, AWidth, AHeight, ALineGap, AAscent, ADescent: integer);
begin
  FRows[FRowCount].CR2 := ACR;
  FRows[FRowCount].CP2 := ACP;
  FRows[FRowCount].Width := FRows[FRowCount].Width + AWidth;
  FRows[FRowCount].Height := Max(FRows[FRowCount].Height,AHeight);
  FRows[FRowCount].LineGap := Max(FRows[FRowCount].LineGap,ALineGap);
  FRows[FRowCount].Ascent := Max(FRows[FRowCount].Ascent,AAscent);
  FRows[FRowCount].Descent := Max(FRows[FRowCount].Descent,ADescent);
end;

procedure RichTextRect(ASkin: TXLSBookSkin; const AX1,AY1,AX2,AY2: integer; AFontRuns: PXc12FontRunArray; const AFRCount: integer; const AText: AxUCString);
var
  XRP: TXBookRichPainter;
begin
  XRP := TXBookRichPainter.Create(ASkin);
  try
    XRP.AssignText(AText,Nil,AFontRuns,AFRCount);
    XRP.DrawText(AX1,AY1,AX2,AY2);
    ASkin.SetDefaultFont;
  finally
    XRP.Free;
  end;
end;

procedure RichTextRect(ASkin: TXLSBookSkin; const AX1,AY1,AX2,AY2: double; AText: TXLSFormattedText; AHorizAlign: TXc12HorizAlignment = chaLeft; AVertAlign : TXc12VertAlignment = cvaTop);
var
  XRP: TXBookRichPainter;
begin
  XRP := TXBookRichPainter.Create(ASkin);
  try
    XRP.HorizAlign :=  AHorizAlign;
    XRP.VertAlign := AVertAlign;

    XRP.FCharRuns.Assign(AText);
    XRP.DrawText(Round(AX1),Round(AY1),Round(AX2),Round(AY2));
    ASkin.SetDefaultFont;
  finally
    XRP.Free;
  end;
end;

procedure RichTextRect(ASkin: TXLSBookSkin; const AXYRect: TXYRect; AXF: TXc12XF; AFontRuns: PXc12FontRunArray; const AFRCount: integer; const AText: AxUCString); overload;
var
  XRP: TXBookRichPainter;
begin
  XRP := TXBookRichPainter.Create(ASkin);
  try
    XRP.AssignText(AText,AXF,AFontRuns,AFRCount);
//    XRP.DrawText(AXYRect.X1 + ASkin.CellLeftMarg,AXYRect.Y1 + ASkin.CellTopMarg,AXYRect.X2 - ASkin.CellRightMarg,AXYRect.Y2 - ASkin.CellBottomMarg);
    XRP.DrawText(AXYRect.X1 + ASkin.CellLeftMarg,AXYRect.Y1,AXYRect.X2 - ASkin.CellRightMarg,AXYRect.Y2);
    ASkin.SetDefaultFont;
  finally
    XRP.Free;
  end;
end;

function RichTextWidth(ASkin: TXLSBookSkin; AXF: TXc12XF; AFontRuns: PXc12FontRunArray; const AFRCount: integer; const AText: AxUCString): integer;
var
  XRP: TXBookRichPainter;
begin
  XRP := TXBookRichPainter.Create(ASkin);
  try
    XRP.AssignText(AText,AXF,AFontRuns,AFRCount);
    Result := XRP.TextWidth;
    ASkin.SetDefaultFont;
  finally
    XRP.Free;
  end;
end;


end.
