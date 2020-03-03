unit XSSIECaretRow;

{-
********************************************************************************
******* XLSSpreadSheet V2.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2013 Lars Arvidsson, Axolot Data               *******
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

uses Classes, SysUtils,
     XSSIEUtils, XSSIEDefs, XSSIECharRun, XSSIELogParas, XSSIEPhyRow, XSSIELogPhyPosition,
     XSSIEGDIText, XSSIEDocProps;

type TAXWPhyCharData = record
     C: AxUCChar;
     X: double;
     Width: double;
     end;

type TAXWCaretRowEditor = class(TObject)
private
     function  GetChar(Index: integer): TAXWPhyCharData;
     procedure SetCaretPos(const Value: TAXWEventLogPhyPosition);
protected
     FArea      : TAXWClientArea;
     FPrint     : TAXWTextPrint;
     FParas     : TAXWLogParas;
     FHeight    : double;
     FPrevCaretX: double;
     FTabCount  : integer;

     FScrCaretX : double;
     FScrCaretY : double;
     FScrCaretH : double;

     FCaretPos: TAXWEventLogPhyPosition;
     FChars: array of TAXWPhyCharData;
     FBreakCharsWidth: array of integer;

     procedure AddEOL(X: double);
     procedure DoCharRun(var AX: double; ACR: TAXWCharRun; const APos1,APos2: integer; ARowCx: TAXWPhyRowComplex; const AJustifyWidth: double);
public
     constructor Create(Print: TAXWTextPrint; Paras: TAXWLogParas; Area: TAXWClientArea);
     procedure Clear;

     procedure RowChanged;
     function  Equal(Para: TAXWLogPara; RowIndex: integer): boolean; overload; {$ifdef D2006PLUS} inline; {$endif}
     function  Equal(Pos: TAXWLogPhyPosition): boolean; overload;
     function  XToRelCharPos(X: double): integer;
     function  MatchPrevCP: integer;
     function  IsEOL: boolean;

     procedure SaveX;
     procedure RestoreX;

     function  CaretChar: TAXWPhyCharData;

     property CaretPos: TAXWEventLogPhyPosition read FCaretPos write SetCaretPos;
     property Char[Index: integer]: TAXWPhyCharData read GetChar; default;

     property ScrCaretX : double read FScrCaretX write FScrCaretX;
     property ScrCaretY : double read FScrCaretY write FScrCaretY;
     property ScrCaretH : double read FScrCaretH write FScrCaretH;
     end;

implementation

{ TAXWCaretRowEditor }

procedure TAXWCaretRowEditor.AddEOL(X: double);
var
  i: integer;
begin
  SetLength(FChars,Length(FChars) + 1);
  i := High(FChars);
  FChars[i].C := AXW_CHAR_EOL;
  FChars[i].Width := 0;
  FChars[i].X := X;
end;

function TAXWCaretRowEditor.CaretChar: TAXWPhyCharData;
begin
  if Length(FChars) > 0 then
    Result := FChars[FCaretPos.RelativePos]
  else
    Result := FChars[0];
end;

procedure TAXWCaretRowEditor.Clear;
begin
  SetLength(FChars,1);
  FChars[0].C := #0;
  FChars[0].X := 0;
  FChars[0].Width := 0;

//  FPrevCaretX := 0;
  FTabCount := 0;
end;

constructor TAXWCaretRowEditor.Create(Print: TAXWTextPrint; Paras: TAXWLogParas; Area: TAXWClientArea);
begin
  FPrint := Print;
  FParas := Paras;
  FArea := Area;

  Clear;
end;

procedure TAXWCaretRowEditor.DoCharRun(var AX: double; ACR: TAXWCharRun; const APos1,APos2: integer; ARowCx: TAXWPhyRowComplex; const AJustifyWidth: double);
var
  i,j: integer;
  w: double;
  S: AxUCString;

procedure DoComplexJustify(S: AxUCString);
var
  Cnt: integer;
  Percent: double;
begin
  Cnt := CountChars(FPrint.CurrBreakChar,S);
  if Cnt > 0 then begin
    Percent := ARowCx.IterateNext(Cnt);
    FPrint.SetTextJustification(Round(AJustifyWidth * Percent),Cnt);
  end;
end;

begin
  FPrint.SetupFormat(FParas[FCaretPos.Para.Index].CHP,ACR.CHPX);
  S := Copy(ACR.Text,APos1,APos2);
  if ARowCX <> Nil then
    DoComplexJustify(S);

  j := Length(FChars) - 1;
  SetLength(FChars,Length(FChars) + Length(S));
  for i := 1 to Length(S) do begin
    if S[i] = FPrint.CurrBreakChar then begin
      if i = 1 then
        w := FPrint.GDI.TextWidthF(S[1])
      else
        w := FPrint.GDI.TextWidthF(Copy(S,1,i)) - FPrint.GDI.TextWidthF(Copy(S,1,i - 1));
    end
    else
      w := FPrint.GDI.TextWidthF(S[i]);
    FChars[j + i].C := S[i];
    FChars[j + i].Width := w;
  end;
  for i := 1 to Length(S) do begin
    FChars[j + i].X := AX;
    AX := AX + FChars[j + i].Width;
  end;
end;

function TAXWCaretRowEditor.Equal(Pos: TAXWLogPhyPosition): boolean;
begin
  Result := FCaretPos.Equal(Pos);
end;

function TAXWCaretRowEditor.Equal(Para: TAXWLogPara; RowIndex: integer): boolean;
begin
  Result := (FCaretPos.Para = Para) and (FCaretPos.RowIndex = RowIndex) and (Length(FChars) > 0);
end;

function TAXWCaretRowEditor.GetChar(Index: integer): TAXWPhyCharData;
begin
  if Index <= High(FChars) then
    Result := FChars[Index]
  else
    Result := FChars[0];
end;

function TAXWCaretRowEditor.IsEOL: boolean;
begin
  Result := FCaretPos.CharPos = Length(FChars);
end;

procedure TAXWCaretRowEditor.SaveX;
begin
  if Length(FChars) > 0 then
    FPrevCaretX := Round(FChars[FCaretPos.RelativePos].X)
  else
    FPrevCaretX := 0;
end;

procedure TAXWCaretRowEditor.SetCaretPos(const Value: TAXWEventLogPhyPosition);
begin
  FCaretPos := Value;
end;

function TAXWCaretRowEditor.XToRelCharPos(X: double): integer;
var
  i: integer;
  X2: double;
begin
  for i := 1 to High(FChars) do begin
    if FChars[i].X > X then begin
      // TODO Test this
      X2 := FChars[i - 1].X + (FChars[i - 1].Width / 2);
      if X < X2 then
        Result := i - 1
      else
        Result := i;
      Exit;
    end;
  end;
  if Length(FChars) > 0 then
    Result := Length(FChars) - 1
  else
    Result := 0;
end;

function TAXWCaretRowEditor.MatchPrevCP: integer;
begin
  Result := XToRelCharPos(FPrevCaretX);
end;

procedure TAXWCaretRowEditor.RestoreX;
var
  CP: integer;
begin
  CP := XToRelCharPos(FPrevCaretX);
  FCaretPos.UpdatePosRelative(CP);
end;

procedure TAXWCaretRowEditor.RowChanged;
var
  i: integer;
  x: double;
  Row: TAXWPhyRow;
  RowCx: TAXWPhyRowComplex;
  PAP: TAXWPAP;
  IsJustify: boolean;
  JustifyWidth: double;

procedure DoComplexJustify(S: AxUCString);
var
  Cnt: integer;
  Percent: double;
begin
  Cnt := CountChars(FPrint.CurrBreakChar,S);
  if Cnt > 0 then begin
    Percent := RowCx.IterateNext(Cnt);
    FPrint.SetTextJustification(Round(JustifyWidth * Percent),Cnt);
  end;
end;

begin
  Clear;

  RowCx := Nil;
  Row := FCaretPos.Row;

  PAP := TAXWPAP.Create(FParas.MasterPAP);
  try
    PAP.Assign(FCaretPos.Para.PAPX);
    case PAP.Alignment of
      axptaLeft    : x := FCaretPos.Row.XOffset;
      axptaCenter  : x := FCaretPos.Row.XOffset + ((FArea.ClientWidth - FCaretPos.Row.XOffset) - Row.Width) / 2;
      axptaRight   : x := FArea.ClientWidth - Row.Width;
      axptaJustify : begin
        // Last space is wrong width when justify. Shall be font's default width,
        // not the justified width.
        JustifyWidth := FArea.ClientWidth - Row.Width;
        IsJustify := not (aprfLast in Row.Flags);
        if IsJustify and (aprfComplex in Row.Flags) then begin
          RowCx := TAXWPhyRowComplex(Row);
          RowCx.BeginIterate;
        end;
        x := 0;
      end;
    end;
  finally
    PAP.Free;
  end;

  if not Row.Empty then begin
    if Row.CR1 = Row.CR2 then
      DoCharRun(X,Row.CR1,Row.CRPos1,Row.CRPos2 - Row.CRPos1 + 1,RowCx,JustifyWidth)
    else begin
      DoCharRun(X,Row.CR1,Row.CRPos1,MAXINT,RowCx,JustifyWidth);
      for i := Row.CR1.Index + 1 to Row.CR2.Index - 1 do
        DoCharRun(X,FCaretPos.Para[i],1,MAXINT,RowCx,JustifyWidth);
      DoCharRun(X,Row.CR2,1,Row.CRPos2,RowCx,JustifyWidth);
    end;
  end;

  AddEOL(x);
  FPrint.SetTextJustification(0,0);
end;

end.
