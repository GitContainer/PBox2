unit XLSTokenizer5;

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

{$I AxCompilers.inc}
{$ifdef AXW}
  {$I AXW.Inc}
{$else}
  {$I XLSRWII.inc}
{$endif}

uses Classes, SysUtils,
     xpgPUtils,
{$ifdef XLSRWII}
     Xc12Utils5,
{$endif}
     XLSUtils5;

type TXLSTokenValType = (xtvtNone,xtvtInteger,xtvtFloat,xtvtChar,xtvtString,xtvtIdetifier,xtvtTokenList);
type TXLSTokenType = (xttNone,
                      xttWhitespace,

                      xttComma,
                      xttColon,
                      xttSemicolon,
                      xttSpaceOp,
                      xttExclamation,
                      xttLPar,
                      xttRPar,
                      xttLSqBracket,
                      xttRSqBracket,
                      xttAphostrophe,
                      xttSpace,

                      // Order must match order in Xc12Utils5
                      xttConstErrNull,
                      xttConstErrDiv0,
                      xttConstErrValue,
                      xttConstErrRef,
                      xttConstErrName,
                      xttConstErrNum,
                      xttConstErrNA,
                      xttConstErrGettingData,

                      xttConstLogicalTrue,
                      xttConstLogicalFalse,

                      xttConstString,

                      xttConstArrayBegin,
                      xttConstArrayEnd,

                      xttConstNumber,

                      xttOpHat,
                      xttOpMult,
                      xttOpDiv,
                      xttOpPlus,
                      xttOpMinus,
                      xttOpUMinus,
                      xttOpUPlus,
                      xttOpAmp,
                      xttOpEqual,
                      xttOpNotEqual,
                      xttOpLess,
                      xttOpLessEqual,
                      xttOpGreater,
                      xttOpGreaterEqual,
                      xttOpPercent,
                      xttOpRange,

                      xttOpISect,

                      xttRelCol,
                      xttAbsCol,
                      xttRelRow,
                      xttAbsRow,

                      xttR1C1RelCol,
                      xttR1C1AbsCol,
                      xttR1C1RelRow,
                      xttR1C1AbsRow,

                      xttName,
                      xttFuncName,

                      xttSheetName,
                      xttWorkbookName,

                      xttTable,
                      xttTableCol,
                      xttTableSpecial
                      );

{$ifdef AXW}
const XLS_MAXCOL  = 255;
const XLS_MAXCOLS = 256;
const XLS_MAXROW  = 65535;
const XLS_MAXROWS = 65536;
{$endif}

// Without xttOpUnion
const TXLSTokenTypeOperator = [xttOpHat..xttOpRange];

type TXLSTokenPos = packed record
     P1,P2: word;
     end;

type PXLSToken = ^TXLSToken;
     TXLSToken = record
     Pos: TXLSTokenPos;
     TokenType: TXLSTokenType;
     case ValueType: TXLSTokenValType of
       xtvtInteger  : (vInteger: integer);
       xtvtFloat    : (vFloat: double);
       xtvtChar     : (vChar: AxUCChar);
       xtvtString   : (vPStr: PWordArray);
     end;

type TXLSTokenList = class(TList)
private
     function  GetItems(const Index: integer): PXLSToken;
     function  GetAsChar(const Index: integer): AxUCChar;
     function  GetAsFloat(const Index: integer): double;
     function  GetAsInteger(const Index: integer): integer;
     function  GetAsString(const Index: integer): AxUCString;
protected
     procedure Notify(Ptr: Pointer; Action: TListNotification); override;
     procedure AddStringVal(const AValue: AxUCString; const AToken: TXLSTokenType; const APos: TXLSTokenPos);
public
//     procedure Add(const AValue: TXLSTokenType); overload;
     procedure Add(const AValue: TXLSTokenType; const APos: integer); overload;
     procedure Add(const AValue: TXLSTokenType; const APos: TXLSTokenPos); overload;
     procedure AddUnknown(const AValue: AxUCChar);
{$ifdef XLSRWII}
     procedure AddConstError(const AValue: TXc12CellError; const APos: TXLSTokenPos);
{$endif}
     procedure AddConstLogical(const AValue: boolean; const APos: TXLSTokenPos);
     procedure AddConstNumber(const AValue: double; const APos: TXLSTokenPos);
     procedure AddConstString(const AValue: AxUCString; const APos: TXLSTokenPos);
     procedure AddSpace(const ACount: integer; const AIsWhitespace: boolean);
     procedure AddCol(const ACol: integer; const AAbsolute: boolean; const APos: TXLSTokenPos);
     procedure AddRow(const ARow: integer; const AAbsolute: boolean; const APos: TXLSTokenPos);
     procedure AddR1C1Col(const ACol: integer; const AAbsolute: boolean);
     procedure AddR1C1Row(const ARow: integer; const AAbsolute: boolean);
     procedure AddName(const AValue: AxUCString; const APos: TXLSTokenPos);
     procedure AddTableCol(const AValue: AxUCString; const APos: TXLSTokenPos);
     procedure AddTableSpecial(const AValue: AxUCString; const APos: TXLSTokenPos);
     procedure AddTableEnd;

     procedure AddSheetName(const AValue: AxUCString; const APos: TXLSTokenPos);
     procedure AddWorkbookName(const AValue: AxUCString; const APos: TXLSTokenPos);

     function  Last: PXLSToken; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  LastTokenType: TXLSTokenType; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  NextTokenType(AIndex: integer): TXLSTokenType; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     procedure Dump(AList: TStrings);

     property AsInteger[const Index: integer]: integer read GetAsInteger;
     property AsFloat[const Index: integer]: double read GetAsFloat;
     property AsChar[const Index: integer]: AxUCChar read GetAsChar;
     property AsString[const Index: integer]: AxUCString read GetAsString;
     property Items[const Index: integer]: PXLSToken read GetItems; default;
     end;

type TXLSTokenizer = class(TObject)
private
     procedure  SetFormula(const Value: AxUCString);
     procedure  SetStrictMode(const Value: boolean);
protected
     FErrors    : TXLSErrorManager;

     FR1C1Mode  : boolean;
     FFormula   : AxUCString;
     FIndex     : integer;
     FTokenList : TXLSTokenList;

     FStrictMode: boolean;
     FDecimalSep: AxUCChar;
     FListSep   : AxUCChar;

     function  CurrChar: AxUCChar; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  CurrChars(const ACount: integer): AxUCString; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  NextChar: AxUCChar; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  NextChars(const AOffset, ACount: integer): AxUCString; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  NextCharsIs(const AText: AxUCString): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  PrevChars(const ACount: integer): AxUCString; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  SkipUntilNot(AOkChars: AxUCString): AxUCString;
     function  DropAnchor: integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure GotoAnchor(AAnchor: integer); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  AnchorLink(AAnchor: integer): AxUCString;  {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  AnchorPos(AAnchor: integer): TXLSTokenPos;  {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure Eat(const ACount: integer = 1); {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure EatWhitespaces;
     function  EatAndSaveChar(AChar: AxUCChar; AToken: TXLSTokenType): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  IsEOF: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     procedure DeleteChar;

     function  Xc12StrToFloat(AValue: AxUCString): double;
     // All quotes are assumed to be double.
     function  RemoveDoubleQuotes(AValue: AxUCString): AxUCString;
     // True if CurrChar is a single operator character. Combinations such
     // as '<> are not tested.
     function  IsOperatorChar: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  IsA1Column(out ACol: integer): boolean;
     function  IsA1Row(out ARow: integer): boolean;
     function  IsNameChar(AChar: AxUCChar): boolean;
     function  IsFuncName(AName: AxUCString): boolean;

     function  TryChar(AChar: AxUCChar): boolean;

     function  TrySpace: boolean;
     function  TryStringChars: boolean;
     function  TryComma: boolean;
     function  TryListSep: boolean;
     function  TryColon: boolean;
     function  TrySemicolon: boolean;
     function  TrySpaceOp: boolean;

     function  TryConstantListRow: boolean;
     function  TryConstantListRows: boolean;
     function  TryFractionalPart: boolean;
     function  TrySign: boolean;
     function  TryExponentPart: boolean;
     function  TryFullStop: boolean;
     function  TryDecimalSep: boolean;
     function  TryDecimalDigit: boolean;
     function  TryDigitSequence: boolean;
     function  TryWholePartNumber: boolean;

{$ifdef XLSRWII}
     function  TryErrorConstant: boolean;
{$endif}
     function  TryLogicalConstant: boolean;
     function  TryNumericConstant: boolean;
     function  TryStringConstant: boolean;
     function  TryArrayConstant: boolean;
     function  TryConstant: boolean;

     function  TryInfixOperator: boolean;
     function  TryPostfixOperator: boolean;
     function  TryPrefixOperator: boolean;

     function  TryLetter: boolean;
     function  TryNameStartCharacter: boolean;
     function  TryNameCharacters: boolean;
     function  TryTableColNameCharacters: boolean;
     function  TryTableColStartNameCharacter: boolean;
     function  TryName: boolean;

     function  TrySheetNameCharacter: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  TrySheetNameCharacters: boolean;
     function  TryBookNameCharacter: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  TryBookNameCharacters: boolean;
     function  TryWorkbookName: boolean;
     function  TrySheetName(const APending: AxUCString; const APendingPos: TXLSTokenPos): boolean;
     // Testing is not complete. All characters within aphostrophes are accepted.
     // As rules in docs are a bit complicated, do the correct testing when a
     // sheet is renamed. If an invalid sheet name is accepted here, it will
     // fail anyway as the sheet don't exists.
     function  TryWorkSheetPrefixSpecial: boolean;
     function  TryA1Reference: boolean;
     function  TryRRReference: boolean;
     function  TryR1C1Reference: boolean;
     function  TryWorkSheetPrefix: boolean;
     function  TryCellReference: boolean;
     function  TryTable: boolean;
     function  TryTableParts: boolean;
     function  TryTableCol: boolean;

     function  TryFunctionName: boolean;

     function  TryFormula: boolean;
     function  TryExpression: boolean;

     procedure DoToken;

     procedure Postcheck;
public
     constructor Create(AErrors: TXLSErrorManager);
     destructor Destroy; override;

     // TODO Return Nil when error.
     function  Parse: TXLSTokenList;

     property R1C1Mode : boolean read FR1C1Mode write FR1C1Mode;
     property Formula: AxUCString read FFormula write SetFormula;
     // StrictMode = decimal separator, list separator must be as defined in the XML files.
     property StrictMode: boolean read FStrictMode write SetStrictMode;
     end;

function TokenAsString(AToken: PXLSToken): AxUCString;

implementation

// TODO Localized values when formulas are entered in code.
const Xc12Chr_WhitespaceChars = [' ',#13,#10];
const Xc12Chr_DecimalDigits = ['0'..'9'];
const Xc12Chr_FullStop = '.';
const Xc12Chr_Signs = ['+','-'];
const Xc12Chr_ExponentChars = ['e','E'];
const Xc12Chr_ErrorChar = '#';
const Xc12Chr_TRUE = 'TRUE';
const Xc12Chr_FALSE = 'FALSE';
const Xc12Chr_QuoteChars = ['"'];
const Xc12Chr_Hash = '#';
const Xc12Chr_Quote = '"';
const Xc12Chr_Aphostrophe = '''';
const Xc12Chr_Comma = ',';
const Xc12Chr_Colon = ':';
const Xc12Chr_Semicolon = ';';
const Xc12Chr_Space = ' ';
const Xc12Chr_BeginArray = '{';
const Xc12Chr_EndArray = '}';
const Xc12Chr_LPar = '(';
const Xc12Chr_RPar = ')';
const Xc12Chr_Underscore = '_';
const Xc12Chr_Backslash = '\';

function TokenAsString(AToken: PXLSToken): AxUCString;
begin
  SetLength(Result,AToken.vPStr[0]);
  System.Move(AToken.vPStr[1],Pointer(Result)^,AToken.vPStr[0] * 2);
end;

{ TXLSTokenizer }

procedure TXLSTokenizer.Eat(const ACount: integer);
begin
  Inc(FIndex,ACount);
end;

function TXLSTokenizer.EatAndSaveChar(AChar: AxUCChar; AToken: TXLSTokenType): boolean;
begin
  Result := CurrChar = AChar;
  if Result then begin
    FTokenList.Add(AToken,FIndex);
    Eat;
  end;
end;

procedure TXLSTokenizer.EatWhitespaces;
begin
  while CurrChar = Xc12Chr_Space do
    Eat;
end;

function TXLSTokenizer.AnchorLink(AAnchor: integer): AxUCString;
begin
  Result := Copy(FFormula,AAnchor,FIndex - AAnchor);
end;

function TXLSTokenizer.AnchorPos(AAnchor: integer): TXLSTokenPos;
begin
  Result.P1 := AAnchor;
  Result.P2 := FIndex - 1;
end;

constructor TXLSTokenizer.Create(AErrors: TXLSErrorManager);
begin
  FErrors := AErrors;
  SetStrictMode(True);
end;

function TXLSTokenizer.CurrChar: AxUCChar;
begin
  if IsEOF then
    Result := #0
  else
    Result := FFormula[FIndex];
end;

function TXLSTokenizer.CurrChars(const ACount: integer): AxUCString;
begin
  Result := Copy(FFOrmula,FIndex,ACount);
end;

procedure TXLSTokenizer.DeleteChar;
begin
  if not IsEOF then
    System.Delete(FFormula,FIndex,1);
end;

destructor TXLSTokenizer.Destroy;
begin

  inherited;
end;

procedure TXLSTokenizer.DoToken;
begin
  if TrySpace then
    Exit;
  if TryFormula then
    Exit;

 FTokenList.AddUnknown(CurrChar);
 Eat;
end;

function TXLSTokenizer.DropAnchor: integer;
begin
  Result := FIndex;
end;

function TXLSTokenizer.IsEOF: boolean;
begin
  Result := FIndex > Length(FFormula);
end;

procedure TXLSTokenizer.GotoAnchor(AAnchor: integer);
begin
  FIndex := AAnchor;
end;

function TXLSTokenizer.IsOperatorChar: boolean;
begin
  Result := CharInSet(CurrChar,[':',Xc12Chr_Comma,Xc12Chr_Space,'^','*','/','+','-','&','=','<','>']);
end;

function TXLSTokenizer.NextChar: AxUCChar;
begin
  if FIndex < Length(FFormula) then
    Result := FFormula[FIndex + 1]
  else
    Result := #0;
end;

function TXLSTokenizer.NextChars(const AOffset, ACount: integer): AxUCString;
begin
  Result := Copy(FFormula,FIndex + AOffset,ACount);
end;

function TXLSTokenizer.NextCharsIs(const AText: AxUCString): boolean;
begin
  Result := Copy(FFormula,FIndex,Length(AText)) = AText;
end;

function TXLSTokenizer.Parse: TXLSTokenList;
var
  i: integer;
  n: integer;
begin
  Result := Nil;
  if (FFormula = '') or (Trim(FFormula) = '') then
    Exit;

//  FFormula := 'IF(AND(A1="LV",$AR14="x"),"OBS!! ¤","")';
//  FFormula := 'AND(1,1)';

  FTokenList := TXLSTokenList.Create;
  Result := FTokenList;

  n := 0;
  for i := 1 to Length(FFormula) do begin
    if CharInSet(FFormula[i],Xc12Chr_WhitespaceChars) then
      Inc(n)
    else
      Break;
  end;
  if n > 0 then begin
    FTokenList.AddSpace(n,True);
    FFormula := Copy(FFormula,n + 1,MAXINT);
  end;

  n := 0;
  for i := Length(FFormula) downto 1 do begin
    if CharInSet(FFormula[i],Xc12Chr_WhitespaceChars) then
      Inc(n)
    else
      Break;
  end;

  SetLength(FFormula,Length(FFormula) - n);

  while not IsEOF do
    DoToken;

  Postcheck;

  if (n > 0) and (n < Length(FFormula)) then
    FTokenList.AddSpace(n,False);

//  FTokenList.Dump(FErrors.ErrorList);
end;

procedure TXLSTokenizer.Postcheck;
var
  i: integer;
begin
  for i := 1 to FTokenList.Count - 2 do begin
    if (FTokenList[i].TokenType = xttSpace) then begin
      if not (FTokenList[i - 1].TokenType in [xttComma,xttLPar,xttFuncName] + TXLSTokenTypeOperator) and not (FTokenList[i + 1].TokenType in [xttComma,xttRPar] + TXLSTokenTypeOperator) then
        FTokenList[i].TokenType := xttOpISect
      else
        FTokenList[i].TokenType := xttWhitespace;
    end;
  end;
end;

function TXLSTokenizer.PrevChars(const ACount: integer): AxUCString;
begin
  Result := Copy(FFormula,FIndex - ACount,ACount);
end;

function TXLSTokenizer.RemoveDoubleQuotes(AValue: AxUCString): AxUCString;
var
  i,j: integer;
begin
  i := 1;
  j := 1;
  SetLength(Result,Length(AValue));
  while i <= Length(AValue) do begin
    if CharInSet(AValue[i],Xc12Chr_QuoteChars) then
      Inc(i);
    Result[j] := AValue[i];
    Inc(i);
    Inc(j);
  end;
  SetLength(Result,j - 1);
end;

procedure TXLSTokenizer.SetFormula(const Value: AxUCString);
begin
  FFormula := Value;
  FIndex := 1;
end;

function TXLSTokenizer.SkipUntilNot(AOkChars: AxUCString): AxUCString;
var
  p: integer;
begin
  P := FIndex;
  while not IsEOF do begin
    if CPos(CurrChar,AOkChars) < 0 then begin
      Result := Copy(FFormula,p,FIndex - p);
      Exit;
    end;
    Eat;
  end;
  Result := '';
end;

procedure TXLSTokenizer.SetStrictMode(const Value: boolean);
begin
  FStrictMode := Value;
  if FStrictMode then begin
    FDecimalSep := Xc12Chr_FullStop;
    FListSep := Xc12Chr_Comma;
  end
  else begin
    FDecimalSep := FormatSettings.DecimalSeparator;
    FListSep := FormatSettings.ListSeparator;
  end;
end;

function TXLSTokenizer.IsA1Column(out ACol: integer): boolean;
var
  Anchor: integer;
  Cnt: integer;
  m: integer;
  C: AxUCChar;
begin
  Result := False;
  Cnt := 0;
  Anchor := DropAnchor;
  while not IsEOF do begin
    if not CharInSet(CurrChar,['a'..'z','A'..'Z']) then
      Break;
    Eat;
    Inc(Cnt);
    if Cnt > 3 then
      Exit;
  end;
  ACol := 0;
  m := 1;
  Dec(Cnt);
  while Cnt >= 0 do begin
    C := FFormula[Anchor + Cnt];
    if CharInSet(C,['a'..'z']) then
      C := AxUCChar(Word(C) - 32);
    ACol := ACol + (Ord(C) - Ord('A') + 1) * m;
    m := m * 26;
    Dec(Cnt);
  end;
  Dec(ACol);
  Result := (ACol >= 0) and (ACol <= XLS_MAXCOL);
end;

function TXLSTokenizer.TryA1Reference: boolean;
var
  C,R: integer;
  P1,P2: TXLSTokenPos;
  AbsC,AbsR: boolean;
  Anchor: integer;

function TryAbsolute: boolean;
begin
  Result := CurrChar = '$';
  if Result then
    Eat;
end;

begin
  Result := True;
  Anchor := DropAnchor;
  P1.P1 := FIndex;
  AbsC := TryAbsolute;
  if IsA1Column(C) then begin
    P1.P2 := FIndex;
    if CurrChar = ':' then begin
      FTokenList.AddCol(C,AbsC,P1);
      FTokenList.Add(xttColon,FIndex);
      Eat;
      P1.P1 := FIndex;
      AbsC := TryAbsolute;
      if IsA1Column(C) then begin
        P1.P2 := FIndex;
        FTokenList.AddCol(C,AbsC,P1);
        Exit;
      end;
      FErrors.Error('',XLSERR_FMLA_BADCELLREF);
    end;
    P2.P1 := FIndex;
    AbsR := TryAbsolute;
    if IsA1Row(R) and not IsNameChar(CurrChar) then begin
      P2.P2 := FIndex - 1;
      FTokenList.AddCol(C,AbsC,P1);
      FTokenList.AddRow(R,AbsR,P2);
      Exit;
    end;
  end;

  GotoAnchor(Anchor);
  P1.P1 := FIndex;
  AbsR := TryAbsolute;
  if IsA1Row(R) then begin
    if CurrChar = ':' then begin
      P1.P2 := FIndex;
      FTokenList.AddRow(R,AbsR,P1);
      FTokenList.Add(xttColon,FIndex);
      Eat;
      P1.P1 := FIndex;
      AbsR := TryAbsolute;
      if IsA1Row(R) then begin
        P1.P2 := FIndex - 1;
        FTokenList.AddRow(R,AbsR,P1);
        Exit;
      end;
    end;
    FErrors.Error('',XLSERR_FMLA_BADCELLREF);
  end;

  GotoAnchor(Anchor);
  Result := False;
end;

function TXLSTokenizer.IsA1Row(out ARow: integer): boolean;
var
  Anchor: integer;
  Cnt: integer;
begin
  Result := False;
  Cnt := 0;
  Anchor := DropAnchor;
  while not IsEOF and CharInSet(CurrChar,['0'..'9']) do begin
    Eat;
    Inc(Cnt);
  end;
  if (Cnt <= 0) or (Cnt > 9) then
    Exit;
  ARow := StrToIntDef(Copy(FFormula,Anchor,Cnt),-1) - 1;
  Result := (ARow >= 0) and (ARow <= XLS_MAXROWS);
end;

function TXLSTokenizer.IsFuncName(AName: AxUCString): boolean;
var
  i: integer;
begin
  Result := AName <> '';
  if Result then begin
    if (AName[1] = '_') and (Copy(AName,1,6) = '_xlfn.') then
      AName := Copy(AName,7,MAXINT);

    Result := CharInSet(AName[1],['a'..'z','A'..'Z']);
    if Result then begin
      for i := 2 to Length(AName) do begin
        Result := CharInSet(AName[i],['0'..'9','a'..'z','A'..'Z','_','.']);
        if not Result then
          Exit;
      end;
    end;
  end;
end;

function TXLSTokenizer.IsNameChar(AChar: AxUCChar): boolean;
begin
  Result := CharInSet(AChar,['a'..'z','A'..'Z','0'..'9'] + Xc12Chr_DecimalDigits + [Xc12Chr_Underscore] + [Xc12Chr_FullStop]);
  if not Result then
    Result := Ord(AChar) > $7F;
//{$ifdef _AXOLOT_DEBUG}
//{$MESSAGE WARN 'Localize'}
//{$endif}
//    Result := CharInSet(AChar,['å','ä','ö','Å','Ä','Ö']);
end;

function TXLSTokenizer.TryArrayConstant: boolean;
begin
  Result := CurrChar = Xc12Chr_BeginArray;
  if Result then begin
    Eat;
    FTokenList.Add(xttConstArrayBegin,FIndex);

    if not TryConstantListRows then
      FErrors.Error('',XLSERR_FMLA_BADARRAYCONST);

    if CurrChar <> Xc12Chr_EndArray then
      FErrors.Error('',XLSERR_FMLA_BADARRAYCONST);

    Eat;
    FTokenList.Add(xttConstArrayEnd,FIndex);
  end;
end;

function TXLSTokenizer.TryBookNameCharacter: boolean;
var
  C: AxUCChar;
begin
  Result := not IsOperatorChar;
  if Result then begin
    C := CurrChar;
    Result := not IsEOF and not CharInSet(C,[Xc12Chr_Aphostrophe,'[',']','?']);
    if Result then
      Eat;
  end;
end;

function TXLSTokenizer.TryBookNameCharacters: boolean;
begin
  Result := TryBookNameCharacter;
  if Result then begin
    while TryBookNameCharacter do ;
  end;
end;

function TXLSTokenizer.TryCellReference: boolean;
var
  Anchor: integer;
begin
  Result := True;

  TryWorksheetPrefix;

  Anchor := DropAnchor;
  if not FR1C1Mode then begin
{$ifdef XLSRWII}
    if TryErrorConstant then
      Exit;
{$endif}
    if TryA1Reference then begin
      if CurrChar = ':' then begin
        EatAndSaveChar(':',xttColon);
        if TryA1Reference then
          Exit;
//        FErrors.Error('',XLSERR_FMLA_BADCELLREF);
      end;
      Exit;
    end;
  end
  else begin
    if TryR1C1Reference then begin
{$ifdef XLSRWII}
      if TryErrorConstant then
        Exit;
{$endif}
      if CurrChar = ':' then begin
        EatAndSaveChar(':',xttColon);
        if TryR1C1Reference then begin
          Exit;
        end;
//        FErrors.Error('',XLSERR_FMLA_BADCELLREF);
      end;
      Exit;
    end;
  end;

  GotoAnchor(Anchor);
  if TryName then
    Exit;

  Result := False;
end;

function TXLSTokenizer.TryChar(AChar: AxUCChar): boolean;
begin
  Result := CurrChar = AChar;
  if Result then
    Eat;
end;

function TXLSTokenizer.TryColon: boolean;
begin
  Result := CurrChar = Xc12Chr_Colon;
  if Result then begin
    FTokenList.Add(xttColon,FIndex);
    Eat;
  end;
end;

function TXLSTokenizer.TryComma: boolean;
begin
  Result := CurrChar = Xc12Chr_Comma;
  if Result then begin
    FTokenList.Add(xttComma,FIndex);
    Eat;
  end;
end;

function TXLSTokenizer.TryConstant: boolean;
begin
  Result := True;

{$ifdef XLSRWII}
  if TryErrorConstant then
    Exit;
{$endif}
  if TryLogicalConstant then
    Exit;
  if TryRRReference then
    Exit;
  if TryNumericConstant then
    Exit;
  if TryStringConstant then
    Exit;
  if TryArrayConstant then
    Exit;

  Result := False;
end;

function TXLSTokenizer.TryConstantListRow: boolean;
begin
  Result := False;
  TryPrefixOperator;
  if not TryConstant then begin
    FErrors.Error('',XLSERR_FMLA_BADARRAYCONST);
    Exit;
  end;
  while TryListSep do begin
    TryPrefixOperator;
    if not TryConstant then begin
      FErrors.Error('',XLSERR_FMLA_BADARRAYCONST);
      Exit;
    end;
  end;
  Result := True;
end;

function TXLSTokenizer.TryConstantListRows: boolean;
begin
  Result := False;
  if not TryConstantListRow then begin
    FErrors.Error('',XLSERR_FMLA_BADARRAYCONST);
    Exit;
  end;
  while TrySemicolon do begin
    if not TryConstantListRow then begin
      FErrors.Error('',XLSERR_FMLA_BADARRAYCONST);
      Exit;
    end;
  end;
  Result := True;
end;

function TXLSTokenizer.TryDecimalDigit: boolean;
begin
  Result := CharInSet(CurrChar,Xc12Chr_DecimalDigits);
  if Result then
    Eat;
end;

function TXLSTokenizer.TryDigitSequence: boolean;
begin
  Result := CharInSet(CurrChar,Xc12Chr_DecimalDigits);
  if Result then begin
    while not IsEOF and CharInSet(CurrChar,Xc12Chr_DecimalDigits) do
      Eat;
  end;
end;


{$ifdef XLSRWII}
function TXLSTokenizer.TryErrorConstant: boolean;
var
  e: TXc12CellError;
  P: TXLSTokenPos;
begin
  Result := False;
  if CurrChar = Xc12Chr_ErrorChar then begin
    P.P1 := FIndex;
    for e := Succ(Low(TXc12CellError)) to High(TXc12CellError) do begin
      if NextCharsIs(Xc12CellErrorNames[e]) then begin
        P.P2 := FIndex;
        FTokenList.AddConstError(e,P);
        Eat(Length(Xc12CellErrorNames[e]));
        Result := True;
        Exit;
      end;
    end;
    FErrors.Error(SkipUntilNot('#/!?01234567890_ABCDEFGHIJKLMNOPQRSTUVWXYZ'),XLSERR_FMLA_BADERRORCONST);
  end;
end;
{$endif}

function TXLSTokenizer.TryExponentPart: boolean;
begin
  Result := CharInSet(CurrChar,Xc12Chr_ExponentChars);
  if Result then begin
    Eat;
    TrySign;
    if not TryDigitSequence then
      FErrors.Error('',XLSERR_FMLA_BADEXPONENT);
  end;
end;

function TXLSTokenizer.TryExpression: boolean;
var
  P: TXLSTokenPos;
begin
  Result := True;
  if CurrChar = Xc12Chr_LPar then begin
    Eat;

    // TODO Can user function names contain any characters (localize)?
    if (FTokenList.Last <> Nil) and (FTokenList.Last.TokenType = xttName) and IsFuncName(FTokenList.AsString[FTokenList.Count - 1]) then
      FTokenList.Last.TokenType := xttFuncName
    else
      FTokenList.Add(xttLPar,FIndex);

    while TryExpression do ;

    Result := CurrChar = Xc12Chr_RPar;
    if Result then begin
      P.P1 := FIndex;
      P.P2 := FIndex;
      FTokenList.Add(xttRPar,P);
    end
    else
      FErrors.Error(FFormula,XLSERR_FMLA_MISSINGRPAR);
    Eat;
  end
  else begin
    if TryConstant then
      Exit;

    if (FTokenList.LastTokenType in [xttNone,xttLPar,xttComma,xttFuncName,xttWhitespace,xttOpHat..xttOpGreaterEqual]) and TryPrefixOperator then begin
      if TryExpression then
        Exit;
      FErrors.Error('-',XLSERR_FMLA_INVALIDUSEOF);
    end;

    if TryListSep then begin
      TryExpression;
      Exit;
    end;

//    Anchor := DropAnchor;
    if TryInfixOperator then begin
      // TODO check that there not is two infix operator after each other.
//      if FTokenList.Last.TokenType in [xttOpHat,xttOpMult,xttOpDiv,xttOpPlus,xttOpMinus,xttOpAmp,xttOpEqual,xttOpNotEqual,xttOpLess,xttOpLessEqual,xttOpGreater,xttOpGreaterEqual] then begin
//        FErrors.Error('',XLSERR_FMLA_MISSINGOPERAND);
//        Exit;
//      end;
      if TryExpression then begin
        Exit;
      end;
//      FErrors.Error(AnchorLink(Anchor),XLSERR_FMLA_MISSINGEXPR);
    end;
    if TryPostfixOperator then
      Exit;

    if TryCellReference then
      Exit;

    if TryTableParts then
      Exit;

    Result := False;
  end;
end;

function TXLSTokenizer.TryFormula: boolean;
begin
  Result := TryExpression;
  if not Result then
    FErrors.Error('',XLSERR_FMLA_FORMULA);
end;

function TXLSTokenizer.TryFractionalPart: boolean;
begin
  Result := TryDigitSequence;
end;

function TXLSTokenizer.TryDecimalSep: boolean;
begin
  Result := CurrChar = FDecimalSep;
  if Result then
    Eat;
end;

function TXLSTokenizer.TryFullStop: boolean;
begin
  Result := CurrChar = Xc12Chr_FullStop;
  if Result then
    Eat;
end;

function TXLSTokenizer.TryFunctionName: boolean;
begin
  Result := CharInSet(CurrChar,['a'..'z','A'..'Z']);
  if Result then begin
    Eat;
    while not IsEOF and CharInSet(CurrChar,['0'..'9','a'..'z','A'..'Z']) do
      Eat;
  end;
end;

function TXLSTokenizer.TryInfixOperator: boolean;

function DoOperator(AOp: AxUCString; AToken: TXLSTokenType): boolean;
begin
  Result := CurrChars(Length(AOp)) = AOp;
  if Result then begin
    FTokenList.Add(AToken,FIndex);
    Eat;
  end;
end;

begin
  Result := True;
//  Result := FTokenList.LastTokenType <> xttLPar;
//  if not Result then
//    Exit;
  if DoOperator('^',xttOpHat) then
    Exit;
  if DoOperator('*',xttOpMult) then
    Exit;
  if DoOperator('/',xttOpDiv) then
    Exit;
  if DoOperator('+',xttOpPlus) then
    Exit;
  if DoOperator('-',xttOpMinus) then
    Exit;
  if DoOperator('&',xttOpAmp) then
    Exit;
  if DoOperator('=',xttOpEqual) then
    Exit;
  if DoOperator('<>',xttOpNotEqual) then begin
    Eat;
    Exit;
  end;
  if DoOperator('<=',xttOpLessEqual) then begin
    Eat;
    Exit;
  end;
  if DoOperator('>=',xttOpGreaterEqual) then begin
    Eat;
    Exit;
  end;
  if DoOperator('<',xttOpLess) then
    Exit;
  if DoOperator('>',xttOpGreater) then
    Exit;
  if DoOperator(':',xttOpRange) then
    Exit;

//  if DoOperator(Xc12Chr_Space,xttOpSpace) then
//    Exit;
  if CurrChar = Xc12Chr_Space then begin
    Eat;
    if (FTokenList.Last <> Nil) and (FTokenList.Last.TokenType = xttSpace) then
      FTokenList.Last.vInteger := FTokenList.Last.vInteger + 1
    else begin
      FTokenList.Add(xttSpace,FIndex);
      FTokenList.Last.vInteger := 1;
    end;
    Exit;
  end;

  Result := False;
end;

function TXLSTokenizer.TryLetter: boolean;
begin
  Result := CharInSet(CurrChar,['a'..'z','A'..'Z','0'..'9']);
  if Result then
    Eat
  else begin
// Localize better
    Result := (Ord(CurrChar) >= $00C0);
//    Result := CharInSet(CurrChar,['å','ä','ö','Å','Ä','Ö','é','É','ô','Ô','ü','Ü','ï','Ï','ß']);
    if Result then
      Eat
  end;
end;

function TXLSTokenizer.TryListSep: boolean;
begin
  Result := CurrChar = FListSep;
  if Result then begin
    FTokenList.Add(xttComma,FIndex);
    Eat;
  end;
end;

function TXLSTokenizer.TryLogicalConstant: boolean;
var
  P: TXLSTokenPos;
begin
  P.P1 := FINdex;
  Result := NextCharsIs(Xc12Chr_TRUE);
  if Result then begin
    P.P2 := P.P1 + Length(Xc12Chr_TRUE);
    FTokenList.AddConstLogical(True,P);
    Eat(Length(Xc12Chr_TRUE));
    Exit;
  end;

  Result := NextCharsIs(Xc12Chr_FALSE);
  if Result then begin
    P.P2 := P.P1 + Length(Xc12Chr_FALSE);
    FTokenList.AddConstLogical(False,P);
    Eat(Length(Xc12Chr_FALSE));
  end;
end;

// B2*500-SUM(C2:G2)+SIN(1,2)
// 123456789012345678901234567
function TXLSTokenizer.TryName: boolean;
var
  Anchor: integer;
begin
  Anchor := DropAnchor;
  Result := TryNameStartCharacter;
  if Result then begin
    TryNameCharacters;
    FTokenList.AddName(AnchorLink(Anchor),AnchorPos(Anchor));

    TryTable;

    Exit;
  end;

  Result := False;
end;

function TXLSTokenizer.TryNameCharacters: boolean;
var
  Anchor: integer;
begin
  Anchor := DropAnchor;
  while TryLetter or TryDecimalDigit or TryChar(Xc12Chr_Underscore) or TryFullStop do ;
  Result := FIndex > Anchor;
end;

function TXLSTokenizer.TryNameStartCharacter: boolean;
begin
  Result := TryLetter or TryChar(Xc12Chr_Underscore) or TryChar(Xc12Chr_Backslash);
end;

function TXLSTokenizer.TryNumericConstant: boolean;
var
  Anchor: integer;
begin
  Anchor := DropAnchor;
  Result := TryWholePartNumber;
  if Result then begin
    if TryDecimalSep then begin
      if TryFractionalPart then begin
        TryExponentPart;
        FTokenList.AddConstNumber(Xc12StrToFloat(AnchorLink(Anchor)),AnchorPos(Anchor));
        Exit;
      end;
    end;
    TryExponentPart;
    FTokenList.AddConstNumber(Xc12StrToFloat(AnchorLink(Anchor)),AnchorPos(Anchor));
  end
  else begin
    Result := TryDecimalSep;
    if Result then begin
      if not TryFractionalPart then
        FErrors.Error('',XLSERR_FMLA_BADFRACTIONAL);
      TryExponentPart;
      FTokenList.AddConstNumber(Xc12StrToFloat(AnchorLink(Anchor)),AnchorPos(Anchor));
    end;
  end;
end;

function TXLSTokenizer.TryPostfixOperator: boolean;
begin
  Result := CurrChar = '%';
  if Result then begin
    FTokenList.Add(xttOpPercent,FIndex);
    Eat;
  end;
end;

function TXLSTokenizer.TryPrefixOperator: boolean;
begin
  Result := CurrChar = '-';
  if Result then begin
    FTokenList.Add(xttOpUMinus,FIndex);
    Eat;
  end
  else begin
    Result := CurrChar = '+';
    if Result then begin
      FTokenList.Add(xttOpUPlus,FIndex);
      Eat;
    end
  end;
end;

function TXLSTokenizer.TryR1C1Reference: boolean;

function TryR1C1(AR1C1: AxUCChar): boolean;
var
  Val: integer;
  Anchor,A: integer;
begin
  Result := CurrChar = AR1C1;
  if Result then begin
    Anchor := DropAnchor;
    Eat;
    if CurrChar = '[' then begin
      Eat;
      A := DropAnchor;
      Result := CharInSet(CurrChar,['+','-','0'..'9']);
      if Result then begin
        Eat;
        while not IsEOF and CharInSet(CurrChar,['0'..'9']) do
          Eat;
        Result := CurrChar = ']';
        if Result then begin
          Val := StrToInt(AnchorLink(A));
          if AR1C1 = 'C' then
            FTokenList.AddR1C1Col(Val,False)
          else
            FTokenList.AddR1C1Row(Val,False);
          Eat;
        end
        else
          FErrors.Error('',XLSERR_FMLA_BADR1C1REF);
      end
      else
        FErrors.Error('',XLSERR_FMLA_BADR1C1REF);
    end
    else begin
      A := DropAnchor;
      while not IsEOF and CharInSet(CurrChar,['0'..'9']) do
        Eat;
      Result := (CurrChar = 'C') or IsEOF or IsOperatorChar;
      if Result then begin
        if AnchorLink(A) <> '' then
          Val := StrToInt(AnchorLink(A))
        else
          Val := 0;
        if AR1C1 = 'C' then
          FTokenList.AddR1C1Col(Val,True)
        else
          FTokenList.AddR1C1Row(Val,True);
      end;
    end;
    if not Result then
      GotoAnchor(Anchor);
  end;
end;

begin
  Result := TryR1C1('R');
  if Result then begin
    TryR1C1('C');
  end
  else
    Result := TryR1C1('C');
end;

function TXLSTokenizer.TryRRReference: boolean;
var
  Anchor: integer;
  R1Abs,R2Abs: boolean;
begin
  Anchor := DropAnchor;
  R1Abs := CurrChar = '$';
  if R1Abs then
    Eat;
  if TryDigitSequence then begin
    if CurrChar = Xc12Chr_Colon then begin
      FTokenList.AddRow(StrToInt(AnchorLink(Anchor)) - 1,R1Abs,AnchorPos(Anchor));
      FTokenList.Add(xttColon,FIndex);
      Eat;
      R2Abs := CurrChar = '$';
      if R2Abs then
        Eat;
      Anchor := DropAnchor;
      if TryDigitSequence then begin
        FTokenList.AddRow(StrToInt(AnchorLink(Anchor)) - 1,R2Abs,AnchorPos(Anchor));
        Result := True;
        Exit;
      end
      else
        FErrors.Error('',XLSERR_FMLA_BADCELLREF);
    end
    else
      GotoAnchor(Anchor);
  end
  else
    GotoAnchor(Anchor);
  Result := False;
end;

function TXLSTokenizer.TrySemicolon: boolean;
begin
  Result := CurrChar = Xc12Chr_Semicolon;
  if Result then begin
    FTokenList.Add(xttSemicolon,FIndex);
    Eat;
  end;
end;

function TXLSTokenizer.TrySheetName(const APending: AxUCString; const APendingPos: TXLSTokenPos): boolean;
var
  SheetName1: AxUCString;
  P: TXLSTokenPos;
  Anchor,A: integer;
begin
  SheetName1 := '';
  Anchor := DropAnchor;
  A := Anchor;
  while not IsEOF do begin
    Result := IsOperatorChar;
    if Result then begin
      if CurrChar = ':' then begin
        SheetName1 := AnchorLink(Anchor);
        P := AnchorPos(Anchor);
        Eat;
        A := DropAnchor;
        Continue;
      end
      else begin
        GotoAnchor(Anchor);
        // If needed to be changed to True, or removed, check that tables cols [Col1]*[Col2] works
        Result := False;
        Exit;
      end;
    end;
    Result := not CharInSet(CurrChar,[Xc12Chr_Aphostrophe,'(',')','[',']','\','?']);
    if not Result then begin
      GotoAnchor(Anchor);
      Exit;
    end;
    if CurrChar = '!' then begin
      if APending <> '' then
        FTokenList.AddWorkbookName(APending,APendingPos);
      if SheetName1 <> '' then begin
        FTokenList.AddSheetName(SheetName1,P);
        FTokenList.Add(xttColon,FIndex);
      end;
      FTokenList.AddSheetName(AnchorLink(A),AnchorPos(Anchor));
      Eat;
      FTokenList.Add(xttExclamation,FIndex);
      Result := True;
      Exit;
    end;
    Eat;
  end;
  GotoAnchor(Anchor);
  Result := False;
end;

function TXLSTokenizer.TrySheetNameCharacter: boolean;
var
  C: AxUCChar;
begin
  Result := not IsOperatorChar;
  if Result then begin
    C := CurrChar;
    Result := not IsEOF and not CharInSet(C,[Xc12Chr_Aphostrophe,'[',']','\','?']);
    if Result then
      Eat;
  end;
end;

function TXLSTokenizer.TrySheetNameCharacters: boolean;
begin
  Result := TrySheetNameCharacter;
  if Result then begin
    while TrySheetNameCharacter do ;
  end;
end;

function TXLSTokenizer.TrySign: boolean;
begin
  Result := CharInSet(CurrChar,Xc12Chr_Signs);
  if Result then
    Eat;
end;

function TXLSTokenizer.TryStringChars: boolean;
begin
  Result := True;

  while not IsEOF do begin
    if CharInSet(CurrChar,Xc12Chr_QuoteChars) then begin
      if not CharInSet(NextChar,Xc12Chr_QuoteChars) then
        Exit;
      Eat;
    end;
    Eat;
  end;
end;

function TXLSTokenizer.TryStringConstant: boolean;
var
  S: AxUCString;
  P: TXLSTokenPos;
  Anchor: integer;
begin
  Result := CharInSet(CurrChar,Xc12Chr_QuoteChars);
  if Result then begin
    P.P1 := FIndex;
    Eat;
    S := '';
    Anchor := DropAnchor;
    if TryStringChars then
      S := RemoveDoubleQuotes(AnchorLink(Anchor));
    if not CharInSet(CurrChar,Xc12Chr_QuoteChars) then
      FErrors.Error('',XLSERR_FMLA_MISSINGQUOTE);
    P.P2 := FIndex;
    Eat;

    FTokenList.AddConstString(S,P);
  end;
end;

function TXLSTokenizer.TryTable: boolean;
begin
  Result := (CurrChar = '[') and (FTokenList.Last.TokenType = xttName);
  if Result then begin
    FTokenList.Last.TokenType := xttTable;
    Eat;
    if not TryTableParts then
      TryTableCol;
    Result := CurrChar = ']';
    if Result then begin
      Eat;
      FTokenList.AddTableEnd;
    end;
  end;
end;

function TXLSTokenizer.TryTableCol: boolean;
var
  Anchor: integer;
begin
  if CurrChar = '#' then begin
    Eat;
    Anchor := DropAnchor;
    while CurrChar <> ']' do
      Eat;
    Result := CurrChar = ']';
    if Result then
      FTokenList.AddTableSpecial(AnchorLink(Anchor),AnchorPos(Anchor));
  end
  else begin
    Anchor := DropAnchor;
    Result := TryTableColStartNameCharacter;
    if Result then begin
      Result := TryTableColNameCharacters;
      if Result then
        FTokenList.AddTableCol(AnchorLink(Anchor),AnchorPos(Anchor));
    end;
  end;
end;

function TXLSTokenizer.TryTableColNameCharacters: boolean;
var
  Anchor: integer;
begin
  Anchor := DropAnchor;
  repeat
    Result := CurrChar = Xc12Chr_Aphostrophe;
    if Result then begin
      Eat;
      Result := CharInSet(CurrChar,[Xc12Chr_Aphostrophe,'#','[',']']);
      if Result then
        Eat;
    end
    else begin
      // Accept any printable char
      Result := (CurrChar >= #32) and not CharInSet(CurrChar,['[',']','#']);
      if Result then
        Eat;
    end;
  until not Result;
  Result := FIndex > Anchor;
end;

function TXLSTokenizer.TryTableColStartNameCharacter: boolean;
begin
  Result := CurrChar = Xc12Chr_Aphostrophe;
  if Result then begin
    Eat;
    Result := CharInSet(CurrChar,[Xc12Chr_Aphostrophe,'#','[',']']);
    if Result then
      Eat;
  end
  else begin
    // Accept any printable char except '-','+'
    Result := (CurrChar > #32) and not CharInSet(CurrChar,['-','+','[',']','#']);
    if Result then
      Eat;
  end;
end;

function TXLSTokenizer.TryTableParts: boolean;
begin
  EatWhitespaces;
  Result := CurrChar = '[';
  if Result then begin
    Eat;
    while Result do begin
      Result := TryTableCol;
      if Result then begin
        Result := CurrChar = ']';
        if Result then begin
          Eat;
          EatWhitespaces;
          case CurrChar of
            Xc12Chr_Comma    : begin
              Result := TryListSep;
              EatWhitespaces;
            end;
            Xc12Chr_Colon    : Result := TryColon;
            Xc12Chr_Semicolon: Result := TrySemicolon;
            Xc12Chr_Space    : Result := TrySpaceOp;
            ']': Break;
            else begin
              if not IsEOF then
                TryExpression;
              Exit;
            end;
          end;
          if CurrChar <> '[' then
            Break;
          Eat;
        end;
      end;
    end;
  end;
end;

function TXLSTokenizer.TrySpace: boolean;
var
  n: integer;
begin
  n := 0;
  while CharInSet(CurrChar,Xc12Chr_WhitespaceChars) do begin
    Inc(n);
    Eat;
  end;
  Result := n > 0;
  if Result then
    FTokenList.AddSpace(Length(PrevChars(n)),False);
end;

function TXLSTokenizer.TrySpaceOp: boolean;
begin
  Result := CurrChar = Xc12Chr_Space;
  if Result then begin
    FTokenList.Add(xttSpaceOp,FIndex);
    Eat;
  end;
end;

function TXLSTokenizer.TryWholePartNumber: boolean;
begin
  Result := TryDigitSequence;
end;

function TXLSTokenizer.TryWorkbookName: boolean;
var
  Anchor: integer;
begin
  Anchor := DropAnchor;
  Result := TryBookNameCharacters;
  if not Result then
    GotoAnchor(Anchor);
end;

function TXLSTokenizer.TryWorkSheetPrefix: boolean;
var
  S: AxUCString;
  P: TXLSTokenPos;
  Anchor: integer;
begin
  Result := True;

  if TryWorkSheetPrefixSpecial then
    Exit;
  if TrySheetName('',P) then
    Exit;

  if CurrChar = '[' then begin
    Eat;
    Anchor := DropAnchor;
    if TryWorkbookName then begin
      S := AnchorLink(Anchor);
      P := AnchorPos(Anchor);
      if CurrChar = ']' then begin
        Eat;
        if TrySheetName(S,P) then
          Exit;
        GotoAnchor(Anchor - 1);
      end
      else
        GotoAnchor(Anchor - 1);
    end
    else
      GotoAnchor(Anchor - 1);
  end;

  Result := False;
end;

function TXLSTokenizer.TryWorkSheetPrefixSpecial: boolean;
var
  p: integer;
  P1: TXLSTokenPos;
  S: AxUCString;
  Anchor: integer;
begin
  Result := CurrChar = Xc12Chr_Aphostrophe;
  if Result then begin
    Eat;
//    FTokenList.Add(xttAphostrophe);
    Anchor := DropAnchor;

//    while not IsEOF and (CurrChar <> Xc12Chr_Aphostrophe) do
//      Eat;

    while not IsEOF do begin
      if (CurrChar = Xc12Chr_Aphostrophe) and (NextChar = Xc12Chr_Aphostrophe) then begin
        DeleteChar;
        Eat;
      end
      else if CurrChar = Xc12Chr_Aphostrophe then
        Break;

      Eat;
    end;

    Result := CurrChar = Xc12Chr_Aphostrophe;
    if Result then begin
      S := AnchorLink(Anchor);
      P1 := AnchorPos(Anchor);
      Eat;
      if S = '' then
        FErrors.Error('',XLSERR_FMLA_BADSHEETNAME)
      else begin
        Result := CurrChar = '!';
        if Result then begin
          Eat;
          if S[1] = '[' then begin
            p := CPos(']',S);
            if (p = Length(S)) or (p < 3) then begin
              FErrors.Error('',XLSERR_FMLA_BADWORKBOOKNAME);
              Exit;
            end;
//            FTokenList.Add(xttLSqBracket);
            FTokenList.AddWorkbookName(Copy(S,2,p - 2),P1);
//            FTokenList.Add(xttRSqBracket);
            S := Copy(S,p + 1,MAXINT);
          end;

          p := CPos(':',S);
          if p > 0 then begin
            if (p = 1) or (p = Length(S)) then begin
              FErrors.Error('',XLSERR_FMLA_BADSHEETNAME);
              Exit;
            end;
            FTokenList.AddSheetName(Copy(S,1,p - 1),P1);
            FTokenList.Add(xttColon,FIndex);
            S := Copy(S,p + 1,MAXINT);
          end;

          if Length(S) > 0 then begin
            FTokenList.AddSheetName(S,P1);
//            FTokenList.Add(xttAphostrophe);
            FTokenList.Add(xttExclamation,FIndex);
          end
          else
            FErrors.Error('',XLSERR_FMLA_BADSHEETNAME);
        end
        else
          FErrors.Error('!',XLSERR_FMLA_MISSING);
      end;
    end
    else
      FErrors.Error('''',XLSERR_FMLA_MISSING);
  end;
end;

function TXLSTokenizer.Xc12StrToFloat(AValue: AxUCString): double;
begin
  if FStrictMode then
    Result := XmlStrToFloat(AValue)
  else
    Result := StrToFloatDef(AValue,0);
end;

{ TXLSTokenList }

procedure TXLSTokenList.Add(const AValue: TXLSTokenType; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  PT.TokenType := AValue;
  PT.ValueType := xtvtNone;
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;
  inherited Add(PT);
end;

procedure TXLSTokenList.Add(const AValue: TXLSTokenType; const APos: integer);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  PT.TokenType := AValue;
  PT.ValueType := xtvtNone;
  PT.Pos.P1 := APos;
  PT.Pos.P2 := APos;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddCol(const ACol: integer; const AAbsolute: boolean; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  if AAbsolute then
    PT.TokenType := xttAbsCol
  else
    PT.TokenType := xttRelCol;
  PT.ValueType := xtvtInteger;
  PT.vInteger := ACol;
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;
  inherited Add(PT);
end;

{$ifdef XLSRWII}
procedure TXLSTokenList.AddConstError(const AValue: TXc12CellError; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  case AValue of
    errNull       : PT.TokenType := xttConstErrNull;
    errDiv0       : PT.TokenType := xttConstErrDiv0;
    errValue      : PT.TokenType := xttConstErrValue;
    errRef        : PT.TokenType := xttConstErrRef;
    errName       : PT.TokenType := xttConstErrName;
    errNum        : PT.TokenType := xttConstErrNum;
    errNA         : PT.TokenType := xttConstErrNA;
    errGettingData: PT.TokenType := xttConstErrGettingData;

    else            PT.TokenType := xttConstErrNull;
  end;
  PT.ValueType := xtvtNone;
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;
  inherited Add(PT);
end;
{$endif}

procedure TXLSTokenList.AddConstLogical(const AValue: boolean; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  if AValue then
    PT.TokenType := xttConstLogicalTrue
  else
    PT.TokenType := xttConstLogicalFalse;
  PT.ValueType := xtvtNone;
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddConstNumber(const AValue: double; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  PT.TokenType := xttConstNumber;
  PT.ValueType := xtvtFloat;
  PT.vFloat := AValue;
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddConstString(const AValue: AxUCString; const APos: TXLSTokenPos);
begin
  AddStringVal(AValue,xttConstString,APos);
end;

procedure TXLSTokenList.AddName(const AValue: AxUCString; const APos: TXLSTokenPos);
begin
  AddStringVal(AValue,xttName,APos);
end;

procedure TXLSTokenList.AddR1C1Col(const ACol: integer; const AAbsolute: boolean);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  if AAbsolute then
    PT.TokenType := xttR1C1AbsCol
  else
    PT.TokenType := xttR1C1RelCol;
  PT.ValueType := xtvtInteger;
  PT.vInteger := ACol;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddR1C1Row(const ARow: integer; const AAbsolute: boolean);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  if AAbsolute then
    PT.TokenType := xttR1C1AbsRow
  else
    PT.TokenType := xttR1C1RelRow;
  PT.ValueType := xtvtInteger;
  PT.vInteger := ARow;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddRow(const ARow: integer; const AAbsolute: boolean; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  if AAbsolute then
    PT.TokenType := xttAbsRow
  else
    PT.TokenType := xttRelRow;
  PT.ValueType := xtvtInteger;
  PT.vInteger := ARow;
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddSheetName(const AValue: AxUCString; const APos: TXLSTokenPos);
begin
  AddStringVal(AValue,xttSheetName,APos);
end;

procedure TXLSTokenList.AddStringVal(const AValue: AxUCString; const AToken: TXLSTokenType; const APos: TXLSTokenPos);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  PT.TokenType := AToken;
  PT.ValueType := xtvtString;
  GetMem(PT.vPStr,Length(AValue) * 2 + 2);
  System.Move(Pointer(AValue)^,PT.vPStr[1],Length(AValue) * 2);
  PT.vPStr[0] := Length(AValue);
  PT.Pos.P1 := APos.P1;
  PT.Pos.P2 := APos.P2;

  inherited Add(PT);
end;

procedure TXLSTokenList.AddTableCol(const AValue: AxUCString; const APos: TXLSTokenPos);
begin
  AddStringVal(AValue,xttTableCol,APos);
end;

procedure TXLSTokenList.AddTableEnd;
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  PT.TokenType := xttRSqBracket;
  PT.ValueType := xtvtNone;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddTableSpecial(const AValue: AxUCString; const APos: TXLSTokenPos);
begin
  AddStringVal(AValue,xttTableSpecial,APos);
end;

procedure TXLSTokenList.AddUnknown(const AValue: AxUCChar);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  PT.TokenType := xttNone;
  PT.ValueType := xtvtChar;
  PT.vChar := AValue;
  inherited Add(PT);
end;

procedure TXLSTokenList.AddWorkbookName(const AValue: AxUCString; const APos: TXLSTokenPos);
begin
  AddStringVal(AValue,xttWorkbookName,APos);
end;

procedure TXLSTokenList.AddSpace(const ACount: integer; const AIsWhitespace: boolean);
var
  PT: PXLSToken;
begin
  PT := New(PXLSToken);
  if AIsWhitespace then
    PT.TokenType := xttWhitespace
  else
    PT.TokenType := xttSpace;
  PT.ValueType := xtvtInteger;
  PT.vInteger := ACount;
  inherited Add(PT);
end;

procedure TXLSTokenList.Dump(AList: TStrings);
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    case Items[i].TokenType of
      xttWhitespace         : AList.Add('[WS] "' + IntToStr(AsInteger[i]) + '"');

      xttComma              : AList.Add('[arg sep]');
      xttColon              : AList.Add(':');
      xttSemicolon          : AList.Add(Xc12Chr_Semicolon);
      xttLPar               : AList.Add(Xc12Chr_LPar);
      xttRPar               : AList.Add(Xc12Chr_RPar);
      xttLSqBracket         : AList.Add('[');
      xttRSqBracket         : AList.Add(']');
      xttExclamation        : AList.Add('!');
      xttAphostrophe        : AList.Add('''');

      xttSpace              : AList.Add('[space * ' + IntToStr(AsInteger[i]) + ']');

      xttConstErrNull       : AList.Add('#NULL!');
      xttConstErrDiv0       : AList.Add('#DIV/0!');
      xttConstErrValue      : AList.Add('#VALUE!');
      xttConstErrRef        : AList.Add('#REF!');
      xttConstErrName       : AList.Add('#NAME?');
      xttConstErrNum        : AList.Add('#NUM!');
      xttConstErrNA         : AList.Add('#N/A');
      xttConstErrGettingData: AList.Add('#GETTING_DATA');

      xttConstLogicalTrue   : AList.Add('[bool] TRUE');
      xttConstLogicalFalse  : AList.Add('[bool] FALSE');

      xttConstString        : AList.Add('[const] ' + AsString[i]);

      xttConstArrayBegin    : AList.Add('[array] {');
      xttConstArrayEnd      : AList.Add('[array] }');

      xttConstNumber        : AList.Add('[const num] ' + FloatToStr(AsFloat[i]));

      xttOpHat              : AList.Add('[op] ^');
      xttOpMult             : AList.Add('[op] *');
      xttOpDiv              : AList.Add('[op] /');
      xttOpPlus             : AList.Add('[op] +');
      xttOpMinus            : AList.Add('[op] -');
      xttOpUMinus           : AList.Add('[op] u-');
      xttOpUPlus            : AList.Add('[op] u+');
      xttOpAmp              : AList.Add('[op] &');
      xttOpEqual            : AList.Add('[op] =');
      xttOpNotEqual         : AList.Add('[op] <>');
      xttOpLess             : AList.Add('[op] <');
      xttOpLessEqual        : AList.Add('[op] <=');
      xttOpGreater          : AList.Add('[op] >');
      xttOpGreaterEqual     : AList.Add('[op] >=');
      xttOpPercent          : AList.Add('[op] %');
      xttOpRange            : AList.Add('[op] :');
      xttOpISect            : AList.Add('[op] {Intersect}');

{$ifdef XLSRWII}
      xttRelCol             : AList.Add('[col] ' + ColToRefStr(AsInteger[i],False));
      xttAbsCol             : AList.Add('[col] $' + ColToRefStr(AsInteger[i],False));
{$endif}
      xttRelRow             : AList.Add('[row] ' + IntToStr(AsInteger[i] + 1));
      xttAbsRow             : AList.Add('[row] $' + IntToStr(AsInteger[i] + 1));

      xttR1C1RelCol         : AList.Add('[col] C[' + IntToStr(AsInteger[i]) + ']');
      xttR1C1AbsCol         : AList.Add('[col] C' + IntToStr(AsInteger[i]));
      xttR1C1RelRow         : AList.Add('[row] R[' + IntToStr(AsInteger[i]) + ']');
      xttR1C1AbsRow         : AList.Add('[row] R' + IntToStr(AsInteger[i]));

      xttName               : AList.Add('[name] ' + AsString[i]);
      xttFuncName           : AList.Add('[function] ' + AsString[i] + '(');
      xttSheetName          : AList.Add('[sheet name] ' + AsString[i]);
      xttWorkbookName       : AList.Add('[workbook name] ' + AsString[i]);

      xttTable              : AList.Add('[table] ' + AsString[i]);
      xttTableCol           : AList.Add('[table col] ' + AsString[i]);
      xttTableSpecial       : AList.Add('[table special] ' + AsString[i]);

      else                    AList.Add('_DUMP_ERROR_');
    end;
  end;
end;

function TXLSTokenList.GetAsChar(const Index: integer): AxUCChar;
begin
  Result := Items[Index].vChar;
end;

function TXLSTokenList.GetAsFloat(const Index: integer): double;
begin
  Result := Items[Index].vFloat;
end;

function TXLSTokenList.GetAsInteger(const Index: integer): integer;
begin
  Result := Items[Index].vInteger;
end;

function TXLSTokenList.GetAsString(const Index: integer): AxUCString;
begin
  SetLength(Result,Items[Index].vPStr[0]);
  System.Move(Items[Index].vPStr[1],Pointer(Result)^,Items[Index].vPStr[0] * 2);
end;

function TXLSTokenList.GetItems(const Index: integer): PXLSToken;
begin
  Result := PXLSToken(inherited Items[Index]);
end;

function TXLSTokenList.Last: PXLSToken;
begin
  if Count > 0 then
    Result := Items[Count - 1]
  else
    Result := Nil;
end;

function TXLSTokenList.LastTokenType: TXLSTokenType;
var
  i: integer;
begin
  i := Count - 1;
  while (i >= 0) and (Items[i].TokenType = xttSpace) do
    Dec(i);

  if i >= 0 then
    Result := Items[i].TokenType
  else
    Result := xttNone;
end;

function TXLSTokenList.NextTokenType(AIndex: integer): TXLSTokenType;
begin
  Inc(AIndex);
  while (AIndex < Count) and (Items[AIndex].TokenType = xttSpace) do
    Inc(AIndex);

  if AIndex >= 0 then
    Result := Items[AIndex].TokenType
  else
    Result := xttNone;
end;

procedure TXLSTokenList.Notify(Ptr: Pointer; Action: TListNotification);
begin
  inherited;
  case Action of
    lnDeleted: begin
      case PXLSToken(Ptr).ValueType of
        xtvtString   : FreeMem(PXLSToken(Ptr).vPStr);
      end;
      FreeMem(Ptr);
    end;
  end;
end;

end.
