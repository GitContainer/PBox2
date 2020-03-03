unit XSSIEGDIText;

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

uses { Delphi } Classes, SysUtils, Types, Contnrs, Math, Windows,
                XBookPaintGDI2,
     { AxWord } XSSIEDefs, XSSIEUtils, XSSIEDocProps, XSSIECharRun, XSSIEPhyRow, XSSIELogParas,
                XSSIELogEditor;

type TAXWULData = class(TObject)
private
     FColor: longword;
     FIsULColor: boolean;
     FSize: integer;
     FPosition: integer;
     FDescent: integer;
     FX,FWidth: double;
public
     property Size: integer read FSize write FSize;
     property Position: integer read FPosition write FPosition;
     property Descent: integer read FDescent write FDescent;
     property Color: longword read FColor write FColor;
     property IsULColor: boolean read FIsULColor write FIsULColor;
     property X: double read FX write FX;
     property Width: double read FWidth write FWidth;
     end;

type TAXWUnderlineList = class(TObjectList)
private
     function  GetItems(Index: integer): TAXWULData;
protected
     FGDI         : TAXWGDI;
     FNewULColor  : longword;
     FNewUnderline,
     FUnderline   : TAXWChpUnderline;
     FY           : double;
     FAveSize,
     FAvePos,
     FAveDescent  : double;

     procedure PaintUnderline;
public
     constructor Create(AGDI: TAXWGDI);
     procedure Clear; override;

     function  Add: TAXWULData;
     function  Active: boolean; {$ifdef D2006PLUS} inline; {$endif}
     procedure UpdatePos(const X,Y,W: double);
     procedure BeginProps; {$ifdef D2006PLUS} inline; {$endif}
     procedure EndProps;
     procedure Calc;
     procedure BeginRow;
     procedure EndRow;

     property Items[Index: integer]: TAXWULData read GetItems; default;
     property NewUnderline: TAXWChpUnderline read FNewUnderline write FNewUnderline;
     property NewULColor: longword read FNewULColor write FNewULColor;
     property Y: double read FY write FY;
     end;

type TAXWProcPrint = procedure(X,Y: integer; Text: AxUCString) of object;

type TAXWTextPrint = class(TObject)
protected
     FGDI              : TAXWGDI;
     FArea             : TAXWClientArea;
     FDOP              : TAXWDOP;
     FSelections       : TAXWSelections;
     FIsPrintRow       : boolean;
     FShowNonprintChars: boolean;
     FLastWidth        : integer;
     FULList           : TAXWUnderlineList;
     FTabCount         : integer;

     procedure PrintTextSelected(const S: AxUCString; var X,Y: double; Row: TAXWPhyRow); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure PrintTextMarker(CHP: TAXWCHP; CR: TAXWCharRun; X, Y: double; S: AxUCString; var P: integer; Row: TAXWPhyRow);
     procedure BeginRow;
     procedure EndRow;
     procedure Print(X,Y: double; Text: AxUCString);
public
     constructor Create(AGDI: TAXWGDI; AArea: TAXWClientArea; ASelections: TAXWSelections; ADOP: TAXWDOP);
     destructor Destroy; override;

     procedure PrintRow(Para: TAXWLogPara; CHP: TAXWCHP; PAP: TAXWPAP; Row: TAXWPhyRow; Y: double);

     procedure PrintNoDoc(X,Y: integer; Text: AxUCString);
     procedure SetupFormat(CHP: TAXWCHP; CHPX: TAXWCHPX);
     function  CurrBreakChar: AxUCChar; {$ifdef D2006PLUS} inline; {$endif}
     function  SetTextJustification(const BreakExtra, BreakCount: Integer): Integer;

     property ShowNonprintChars: boolean read FShowNonprintChars write FShowNonprintChars;
     property GDI: TAXWGDI read FGDI;
     end;

implementation

{ TAXWTextPrint }

procedure TAXWTextPrint.BeginRow;
begin
//  if FIsPrintRow then
//    raise XLSRWException.Create('FIsPrintRow is True');  // TODO Remove when ok
  FIsPrintRow := True;
  FULList.BeginRow;
  FTabCount := 0;
end;

constructor TAXWTextPrint.Create(AGDI: TAXWGDI; AArea: TAXWClientArea; ASelections: TAXWSelections; ADOP: TAXWDOP);
begin
  FGDI := AGDI;
  FArea := AArea;
  FSelections := ASelections;
  FDOP := ADOP;
  FULList := TAXWUnderlineList.Create(FGDI);
end;

function TAXWTextPrint.CurrBreakChar: AxUCChar;
begin
  // Correct break char is in TEXTMETRIC
  Result := AXW_CHAR_BREAKCHAR;
end;

destructor TAXWTextPrint.Destroy;
begin
  FULList.Free;
  inherited;
end;

procedure TAXWTextPrint.Print(X, Y: double; Text: AxUCString);
var
  i: integer;
  S: AxUCString;
begin
  if FShowNonprintChars then begin
    SetLength(S,Length(Text));
    for i := 1 to Length(Text) do begin
      if (Word(Text[i]) and $FF80) = 0 then
        // TODO Will this work for all fonts? Have all fonts the first 127 chars
        // the same as US-ASCII? Has all fonts the replacemant chars?
        S[i] := AxUCChar(NonPrintableCharMap[Integer(Text[i])])
      else
        S[i] := Text[i];
    end;
  end
  else
    S := Text;
  FGDI.TextOutF(X,Y,S);

  FLastWidth := FGDI.TextWidth(S);

  if FULList.Active then
    FULList.UpdatePos(X,Y,FLastWidth);
end;

procedure TAXWTextPrint.EndRow;
begin
  FIsPrintRow := False;
  FULList.EndRow;
end;

procedure TAXWTextPrint.PrintNoDoc(X, Y: integer; Text: AxUCString);
var
  i: integer;
  S: AxUCString;
begin
  if FShowNonprintChars then begin
    SetLength(S,Length(Text));
    for i := 1 to Length(Text) do begin
      if (Word(Text[i]) and $FF80) = 0 then
        // TODO Will this work for all fonts? Have all fonts the first 127 chars
        // the same as US-ASCII? Has all fonts the replacemant chars?
        S[i] := AxUCChar(NonPrintableCharMap[Integer(Text[i])])
      else
        S[i] := Text[i];
    end;
  end
  else
    S := Text;
  FGDI.TextOut(X,Y,S);
end;

procedure TAXWTextPrint.PrintRow(Para: TAXWLogPara; CHP: TAXWCHP; PAP: TAXWPAP; Row: TAXWPhyRow; Y: double);
var
  X: double;
  Offs: integer;
  RowCx: TAXWPhyRowComplex;
  S: AxUCString;
  CR: TAXWCharRun;
  JustifyWidth: double;

procedure DoComplexJustify(S: AxUCString);
var
  Cnt: integer;
  Percent: double;
begin
  Cnt := CountChars(CurrBreakChar,S);
  if Cnt > 0 then begin
    Percent := RowCx.IterateNext(Cnt);
    FGDI.SetTextJustification(Round(JustifyWidth * Percent),Cnt);
  end;
end;

begin
  BeginRow;

  Offs := 1;

  X := 0;
  RowCx := Nil;
  case PAP.Alignment of
    axptaLeft:    X := FArea.CX1 + Row.XOffset;
    axptaCenter:  X := FArea.CX1 + Row.XOffset + (FArea.ClientWidth - Row.XOffset) / 2 - Row.Width / 2;
    axptaRight:   X := FArea.CX2 - Row.Width;
    axptaJustify: begin
      X := FArea.CX1;
      if not (aprfLast in Row.Flags) then begin
        JustifyWidth := FArea.Width - Row.Width;
        if aprfComplex in Row.Flags then begin
          RowCx := TAXWPhyRowComplex(Row);
          RowCx.BeginIterate;
        end
        else begin
          FGDI.SetTextJustificationF(JustifyWidth,Row.BreakCharCount);
        end;
      end;
    end;
  end;

  if not Row.Empty then begin
    if Row.CR1 = Row.CR2 then begin
      S := Copy(Row.CR1.Text,Row.CRPos1,Row.CRPos2 - Row.CRPos1 + 1);
      PrintTextMarker(CHP,Row.CR1,X,Y,S,Offs,Row);
    end
    else begin
      S := Copy(Row.CR1.Text,Row.CRPos1,MAXINT);
      if RowCx <> Nil then
        DoComplexJustify(S);
      PrintTextMarker(CHP,Row.CR1,X,Y,S,Offs,Row);
      X := X + FGDI.TextWidthF(S);

      CR := Row.CR1.Next;
      while (CR <> Nil) and (CR <> Row.CR2) do begin
        if RowCx <> Nil then
          DoComplexJustify(CR.Text);
        PrintTextMarker(CHP,CR,X,Y,CR.Text,Offs,Row);
        X := X + FGDI.TextWidthF(CR.Text);

        CR := CR.Next;
      end;

      S := Copy(Row.CR2.Text,1,Row.CRPos2);
      if RowCx <> Nil then
        DoComplexJustify(TrimRight(S));
      PrintTextMarker(CHP,Row.CR2,X,Y,S,Offs,Row);
    end;
  end;

  FGDI.SetTextJustification(0,0);
end;

procedure TAXWTextPrint.PrintTextMarker(CHP: TAXWCHP; CR: TAXWCharRun; X, Y: double; S: AxUCString; var P: integer; Row: TAXWPhyRow);
type
  TSelectionInCharRun = (sicrOutside,sicrLeft,sicrRight,sicrMid);
var
  L,n: integer;
  P1,P2: integer;
  HP1,HP2: integer;
  S2: AxUCString;
  SICR: TSelectionInCharRun;
begin
  SetupFormat(CHP,CR.CHPX);

  if FSelections.HitResult = axshNone then
    Print(X,Y,S)
  else if FSelections.HitResult = axshFullRow then
    PrintTextSelected(S,X,Y,Row)
  else begin
    L := Length(S);
    P1 := 1;
    P2 := L;
    HP1 := FSelections.HitP1Offs - P;
    HP2 := FSelections.HitP2Offs - P;

    if (HP1 <= 0) and (HP2 <= 0) then begin
      SICR := sicrOutside;
    end
    else if HP1 <= 0 then
      SICR := sicrLeft
    else if HP2 <= 0 then
      SICR := sicrRight
    else
      SICR := sicrMid;

    HP1 := Fork(1,L,HP1);
    HP2 := Fork(1,L,HP2);
    case FSelections.HitResult of
      axshFirstOnRow: begin
        if SICR in [sicrOutside,sicrLeft] then
          Dec(HP1)
        else begin
          S2 := Copy(S,1,HP1);
          Print(X,Y,S2);
          X := X + FGDI.TextWidthF(S2);
        end;

        S2 := Copy(S,HP1 + 1,MAXINT);
        if S2 <> '' then
          PrintTextSelected(S2,X,Y,Row);
      end;
      axshLastOnRow: begin
        if SICR in [sicrOutside,sicrRight] then
          Dec(HP2)
        else begin
          S2 := Copy(S,1,HP2);
          if S2 <> '' then
            PrintTextSelected(S2,X,Y,Row);
        end;
        S2 := Copy(S,HP2 + 1,MAXINT);
        Print(X,Y,S2);
      end;
      axshBothOnRow: begin
        if SICR = sicrLeft then
          Dec(HP1)
        else begin
          n := HP1 - P1 + 1;
          S2 := Copy(S,P1,n);
          Print(X,Y,S2);
          X := X + FGDI.TextWidthF(S2);
        end;

        n := HP2 - HP1;
        S2 := Copy(S,HP1 + 1,n);
        if S2 <> '' then
          PrintTextSelected(S2,X,Y,Row);
        n := P2 - HP2;
        S2 := Copy(S,HP2 + 1,n);
        if S2 <> '' then
          Print(X,Y,S2);
      end;
    end;
    Inc(P,L);
  end;
end;

procedure TAXWTextPrint.PrintTextSelected(const S: AxUCString; var X, Y: double; Row: TAXWPhyRow);
var
  W: double;
begin
  W := FGDI.TextWidth(S);
  FGDI.PaintColor := AXW_COLOR_SELECTION;
  FGDI.RectangleF(X,Y - Row.Ascent,X + W,Y + Row.Descent);
  Print(X,Y,S);
  X := X + W;
end;

function TAXWTextPrint.SetTextJustification(const BreakExtra, BreakCount: Integer): Integer;
begin
  Result := FGDI.SetTextJustification(BreakExtra, BreakCount);
end;

procedure TAXWTextPrint.SetupFormat(CHP: TAXWCHP; CHPX: TAXWCHPX);
var
  i: integer;
  Bold,
  Italic: boolean;
begin
  { TODO Only include changes }

  if FIsPrintRow then
    FULList.BeginProps;

  FGDI.SetFontName(CHP.FontName);
  FGDI.SetFontSize(Round(CHP.Size));
  FGDI.FontColor := RevRGB(CHP.Color);
  Bold := CHP.Bold;
  Italic := CHP.Italic;

  if CHPX <> Nil then begin
    for i := 0 to CHPX.Count - 1 do begin
      case TAXWChpId(CHPX[i].Id) of
        axciBold      : if CHPX[i].ValBool then
                         Bold := True;
        axciItalic    : if CHPX[i].ValBool then
                         Italic := True;

        axciSize      : FGDI.SetFontSize(Round(CHPX[i].ValFloat));
        axciColor     : FGDI.FontColor := RevRGB(CHPX[i].ValInt);
        axciUnderline : begin
          if not FIsPrintRow then
            Break;
          FULList.NewUnderline := TAXWChpUnderline(CHPX[i].ValInt);
          if FULList.NewUnderline < axcuSingle then
            FULList.NewUnderline := axcuSingle;
        end;

        axciFontName  : begin
          if Pos('ding',CHPX.FontName) > 1 then
            FGDI.SetFontName(CHPX.FontName,SYMBOL_CHARSET)
          else
            FGDI.SetFontName(CHPX.FontName);
        end;
      end;
    end;
  end;

  FGDI.SetFontStyle(Bold,Italic,False);

  if FIsPrintRow then
    FULList.EndProps;
end;

{ TAXWUnderlineList }

function TAXWUnderlineList.Active: boolean;
begin
  Result := Count > 0;
end;

function TAXWUnderlineList.Add: TAXWULData;
begin
  Result := TAXWULData.Create;
  inherited Add(Result);
end;

procedure TAXWUnderlineList.BeginProps;
begin
  FNewUnderline := axcuNone;
  NewULColor := AXW_COLOR_DEFAULT;
end;

procedure TAXWUnderlineList.BeginRow;
begin
  FUnderline := axcuNone;
  Clear;
end;

procedure TAXWUnderlineList.Calc;
var
  i: integer;
  TotW: double;
  Pos,Sz,Desc: double;
  M: double;
  ULD: TAXWULData;
begin
  if Count = 1 then begin
    FAveSize := Items[0].Size;
    FAvePos := -Items[0].Position;
    FAveDescent := Items[0].Descent;
  end
  else begin
    TotW := 0;
    for i := 0 to Count - 1 do
      TotW := TotW + Items[i].Width;

    Sz := 0;
    Pos := 0;
    Desc := 0;
    for i := 0 to Count - 1 do begin
      ULD := Items[i];
      M := ULD.Width / TotW;
      Sz := Sz + M * ULD.Size;
      Pos := Pos + M * ULD.Position;
      Desc := Desc + M * ULD.Descent;
    end;
    FAveSize := Round(Sz);
    FAvePos := -Round(Pos);
    FAveDescent := Round(Desc);
  end;
end;

procedure TAXWUnderlineList.Clear;
begin
  inherited Clear;
  FUnderline := axcuNone;
  FY := -1;
end;

constructor TAXWUnderlineList.Create(AGDI: TAXWGDI);
begin
  inherited Create;
  FGDI := AGDI;
  FY := -1;
end;

procedure TAXWUnderlineList.EndProps;
var
  ULD: TAXWULData;
//  OTM: OUTLINETEXTMETRICW;
begin
  if (Count <= 0) and (FNewUnderline = axcuNone) then
    Exit;
  if Count > 0 then begin
    if FNewUnderline <> FUnderline then begin
      PaintUnderline;
      Clear;
    end;
  end;
  if FNewUnderline <> axcuNone then begin
    FUnderline := FNewUnderline;
//    GetOutlineTextMetricsW(FCanvas.Handle,SizeOf(OUTLINETEXTMETRICW),@OTM);
    ULD := Add;
    ULD.Size := FGDI.TM.UnderscoreSize;
    ULD.Position := FGDI.TM.UnderscorePosition;
    ULD.Descent := FGDI.TM.Descent;
    if FNewULColor = AXW_COLOR_DEFAULT then
      ULD.Color := FGDI.FontColor
    else
      ULD.Color := FNewULColor;
  end;
end;

procedure TAXWUnderlineList.EndRow;
begin
  if Count > 0 then begin
    PaintUnderline;
    Clear;
  end;
end;

function TAXWUnderlineList.GetItems(Index: integer): TAXWULData;
begin
  Result := TAXWULData(inherited Items[Index]);
end;

procedure TAXWUnderlineList.PaintUnderline;
var
  i: integer;
  Y2,W: double;
  ULD: TAXWULData;
begin
  if FY < 0 then
    Exit;

  Calc;

  ULD := Items[0];
  case FUnderline of
    axcuSingle: begin
      W := FAveSize;
      // The <10pt limit is by comparing with Excel.
      // TODO Average size...
      if FGDI.GetFontSize < 10 then
        Y2 := FAvePos
      else
        Y2 := FAveDescent - W;
      FGDI.PenWidthF := W;
      FGDI.MoveToF(ULD.X,FY + Y2);
      for i := 0 to Count - 1 do begin
        ULD := Items[i];
        FGDI.PenColor := ULD.Color;
        FGDI.LineToF(ULD.X + ULD.Width,FY + Y2);
      end;
    end;
    axcuDouble: begin
      W := Round(FAveSize * 0.7);
      Y2 := FAveDescent - W;
      FGDI.PenWidthF := W;

      FGDI.MoveToF(ULD.X,FY + Y2);
      for i := 0 to Count - 1 do begin
        ULD := Items[i];
        FGDI.PenColor := ULD.Color;
        FGDI.LineToF(ULD.X + ULD.Width,FY + Y2);
      end;

      Y2 := Y2 - W * 2;
      FGDI.MoveToF(ULD.X,FY + Y2);
      for i := 0 to Count - 1 do begin
        ULD := Items[i];
        FGDI.PenColor := ULD.Color;
        FGDI.LineToF(ULD.X + ULD.Width,FY + Y2);
      end;

    end;
  end;
  FGDI.PenWidth := 1;
end;

procedure TAXWUnderlineList.UpdatePos(const X,Y,W: double);
var
  ULD: TAXWULData;
begin
  FY := Y;
  ULD := Items[Count - 1];
  ULD.X := X;
  ULD.Width := W;
end;

end.
