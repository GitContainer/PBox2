unit XLSMask5;

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

uses SysUtils, Classes, Contnrs, Math,
{$ifdef MSWINDOWS}
  {$ifndef BABOON}
     vcl.Graphics,
  {$endif}
{$endif}
     XLSUtils5;

type TMaskEntity = (meNone,meLitteral,meString,mePlace,meSpace,meZero,
                    // Date
                    meYear2,meYear4,
                    meMonthDig1,meMonthDig2,meMonthShort,meMonthName,meMonthChar,
                    meDayDig1,meDayDig2,meDayShort,meDayName,
                    // Time
                    meHourDig1,meHourDig2,
                    meHourDig1AmPm,meHourDig2AmPm,
                    meHourElapsed,
                    meMinuteDig1,meMinuteDig2,
                    meMinuteElapsed,
                    meSecondDig1,meSecondDig2,
                    meSecondElapsed,
                    meStrAmPmUpp,meStrAmPmLow,meStrap,
                    // Control
                    meDecimalPos,
                    mePercentPos,
                    meFirstDigitPlace,
                    meRightIndent
                    );

type TMaskFlag = (mfGeneral,mfThousand,mfDecimals,mfPercent,mfScientific,mfIsDateTime,mfZero);
type TMaskFlags = set of TMaskFlag;

type TMaskData = class(TObject)
protected
     FEntity: TMaskEntity;
     FValue: AxUCString;
     FC: AxUCChar;
     FS: AxUCString;
public
     property Entity: TMaskEntity read FEntity write FEntity;
     property Value: AxUCString read FValue write FValue;
     property C: AxUCChar read FC write FC;
     property S: AxUCString read FS write FS;
     end;

type TMaskDataList = class(TObjectList)
private
     function  GetItems(Index: integer): TMaskData;
protected
public
     function Add: TMaskData;
     function Insert(AIndex: integer): TMaskData;

     property Items[Index: integer]: TMaskData read GetItems; default;
     end;

type TFormatData = record
     Data: TMaskDataList;
     Color: TColor;
     Flags: TMaskFlags;
     DataDecimalPos: integer;
     DecimalCount: integer;
     Div1000Cnt: integer;
     end;

type TExcelMask = class (TObject)
private
     FMask: AxUCString;
     FIndex: integer;
     Formats: array[0..3] of TFormatData;
     FColor: TColor;

     procedure SetMask(Value: AxUCString);
     procedure AddMask(Index: integer; const Value: AxUCString);
     procedure ClearData;
     function  AddEntity(Index: integer; E: TMaskEntity): TMaskData;
protected
     function FormatNumberDateTime(FD: TFormatData; Value: double): AxUCString;
     function FormatNumberNumber(FD: TFormatData; Value: double): AxUCString;
public
     destructor Destroy; override;
     function FormatNumber(Value: double): AxUCString;
     function ValueColor(Value: double): TColor;
     function IsDateTime: boolean;

     property Mask: AxUCString read FMask write SetMask;
     property Index: integer read FIndex write FIndex;
     property Color: TColor read FColor;
     end;

function XLSGetCurrencyFormat(DecStr,ThStr: AxUCString): AxUCString;

implementation

function _MyStrAlloc(Size: Cardinal): PWideChar;
begin
  GetMem(Result,Size * 2 + 2);
end;

procedure _MyStrPCopy(Dest: PWideChar; const Source: AxUCString);
begin
  Move(Source[1],Dest[0],Length(Source) * 2);
  Dest[Length(Dest) + 1] := #0;
end;

{ TExcelMask }

{$ifndef DELPHI_XE2_OR_LATER}
function IntPower(Base: double; Exponent: Integer): double;
asm
        mov     ecx, eax
        cdq
        fld1
        xor     eax, edx
        sub     eax, edx
        jz      @@3
        fld     Base
        jmp     @@2
@@1:    fmul    ST, ST
@@2:    shr     eax,1
        jnc     @@1
        fmul    ST(1),ST
        jnz     @@1
        fstp    st
        cmp     ecx, 0
        jge     @@3
        fld1
        fdivrp
@@3:
        fwait
end;
{$endif}

function Power(Base, Exponent: double): double;
begin
  if Exponent = 0.0 then
    Result := 1.0
  else if (Base = 0.0) and (Exponent > 0.0) then
    Result := 0.0
  else if (Frac(Exponent) = 0.0) and (Abs(Exponent) <= MaxInt) then
    Result := IntPower(Base, Integer(Trunc(Exponent)))
  else
    Result := Exp(Exponent * Ln(Base))
end;

function XLSGetCurrencyFormat(DecStr,ThStr: AxUCString): AxUCString;
begin
  Result := FormatSettings.CurrencyString;
  case FormatSettings.CurrencyFormat of
    0: Result := FormatSettings.CurrencyString + ' ' + ThSTr + DecStr;
    1: Result := ThStr + DecStr + ' ' + FormatSettings.CurrencyString;
    2: Result := FormatSettings.CurrencyString + '  ' + ThStr + DecStr;
    3: Result := ThSTr + DecStr + '  ' + FormatSettings.CurrencyString;
  end;
  case FormatSettings.NegCurrFormat of
     0: Result := Result + ';(' + FormatSettings.CurrencyString + ThStr + DecStr + ')';
     1: Result := Result + ';-' + FormatSettings.CurrencyString + ThStr + DecStr;
     2: Result := Result + ';' + FormatSettings.CurrencyString + '-' + ThStr + DecStr;
     3: Result := Result + ';' + FormatSettings.CurrencyString + ThStr + DecStr + '-';
     4: Result := Result + ';(' + ThStr + DecStr + FormatSettings.CurrencyString + ')';
     5: Result := Result + ';-' + ThStr + DecStr + FormatSettings.CurrencyString;
     6: Result := Result + ';' + ThStr + DecStr + '-' + FormatSettings.CurrencyString;
     7: Result := Result + ';' + ThStr + DecStr + FormatSettings.CurrencyString + '-';
     8: Result := Result + ';-' + ThStr + DecStr + ' ' + FormatSettings.CurrencyString;
     9: Result := Result + ';-' + FormatSettings.CurrencyString + ' ' + ThSTr + DecStr;
    10: Result := Result + ';' + ThStr + DecStr + ' ' + FormatSettings.CurrencyString + '-';
    11: Result := Result + ';' + FormatSettings.CurrencyString + ' ' + ThStr + DecStr + '-';
    12: Result := Result + ';' + FormatSettings.CurrencyString + ' -' + ThSTr + DecStr;
    13: Result := Result + ';' + ThSTr + DecStr + '- ' + FormatSettings.CurrencyString;
    14: Result := Result + ';(' + FormatSettings.CurrencyString + ' ' + ThStr + DecStr + ')';
    15: Result := Result + ';(' + ThStr + DecStr + ' ' + FormatSettings.CurrencyString + ')';
  end;
end;

destructor TExcelMask.Destroy;
begin
  ClearData;
  inherited Destroy;
end;

procedure TExcelMask.ClearData;
var
  i: integer;
begin
  for i := 0 to High(Formats) do begin
    if Formats[i].Data <> Nil then begin
      Formats[i].Data.Clear;
      Formats[i].Data.Free;
      Formats[i].Data := Nil;
    end;
  end;
end;

function TExcelMask.IsDateTime: boolean;
begin
  if Formats[0].Data <> Nil then
    Result := mfIsDateTime in Formats[0].Flags
  else
    Result := False;
end;

function TExcelMask.AddEntity(Index: integer; E: TMaskEntity): TMaskData;
begin
  Result := Formats[Index].Data.Add;
  Result.Entity := E;
end;

procedure TExcelMask.SetMask(Value: AxUCString);
var
  i,j,p: integer;
  SepPos: array[0..4] of integer;
  InsideQuotes: boolean;
  S: AxUCString;
begin
  if Length(Value) > 255 then
    raise XLSRWException.Create('Format string more than 255 characthers');

  Value := StringReplace(Value,'general','@',[rfReplaceAll,rfIgnoreCase]);

  // Masks like '[$-407]...' causes an av
  i := CPos('[',Value);
  j := CPos(']',Value);
  if (i >= 1) and (J > i) then
    System.Delete(Value,i,j - i + 1);

  FMask := Value;
  ClearData;
  InsideQuotes := False;
  j := 0;
  for i := 1 to Length(FMask) do begin
    if FMask[i] = '"' then
      InsideQuotes := not InsideQuotes;
    if (FMask[i] = ';') and not InsideQuotes then begin
      SepPos[j] := i - 1;
      Inc(j);
      if j > (High(SepPos) - 1) then
        Break;
    end;
  end;
  SepPos[j] := Length(FMask) + 1;
  if j > 3 then
    j := 3;
  p := 1;
  for i := 0 to j do begin
    S := Copy(FMask,p,SepPos[i] - p + 1);
    if S <> '' then begin
      AddMask(i,S);
    end;
    p := SepPos[i] + 2;
  end;
end;

procedure TExcelMask.AddMask(Index: integer; const Value: AxUCString);
var
  i,j: integer;
  C: char;
  pMD, LastTimeMD: TMaskData;
  S: string;
begin
  with Formats[Index] do begin
    Data := TMaskDataList.Create;
    Flags := [];
    if Index = 2 then
      Flags := Flags + [mfZero];
    Color := clBlack;
    DecimalCount := 0;
    DataDecimalPos := -1;
    Div1000Cnt := 0;
    LastTimeMD := Nil;
    i := 1;
    while i <= Length(Value) do begin
      case Value[i] of
        '@':
        begin
          AddEntity(Index,meNone);
          Flags := [mfGeneral];
        end;
        '#':
        begin
          AddEntity(Index,mePlace);
          if mfDecimals in Flags then
            Inc(DecimalCount);
          Div1000Cnt := 0;
        end;
        '?':
        begin
          AddEntity(Index,meSpace);
          Div1000Cnt := 0;
        end;
        '0':
        begin
          AddEntity(Index,meZero);
          if mfDecimals in Flags then
            Inc(DecimalCount);
          Div1000Cnt := 0;
        end;
        'e','E': begin
          if (i < Length(Mask)) and CharInSet(Mask[i + 1],['-','+']) then begin
            Inc(i);
            AddEntity(Index,meNone);
            Flags := Flags + [mfScientific];
          end
          else begin
            pMD := AddEntity(Index,meLitteral);
            pMD.C := Value[i];
          end;
        end;
        '%':
        begin
          Flags := Flags + [mfPercent];
          AddEntity(Index,mePercentPos);
        end;
        '\':
        begin
          Inc(i);
          if i <= Length(Mask) then begin
            pMD := AddEntity(Index,meLitteral);
            pMD.C := Value[i];
          end;
        end;
        '*': ;
        '_':
        begin
          Inc(i);
          if i <= Length(Mask) then begin

            pMD := AddEntity(Index,meLitteral);
            pMD.C := ' ';

//            pMD := AddEntity(Index,meRightIndent);
          end;
        end;
        '"':
        begin
          j := i + 1;
          Inc(i);
          while (i <= Length(Value)) and (Value[i] <> '"') do
            Inc(i);
          if i <= Length(Value) then begin
            pMD := AddEntity(Index,meString);
           pMD.S := Copy(Value,j,i - j);
          end;
        end;
        '.':
        begin
          Flags := Flags + [mfDecimals];
          AddEntity(Index,meDecimalPos);
          DataDecimalPos := Data.Count;
        end;
        ',': begin
          Flags := Flags + [mfThousand];
          Inc(Div1000Cnt);
        end;
        '[':
        begin
          j := i + 1;
          Inc(i);
          while (i <= Length(Value)) and (Value[i] <> ']') do
            Inc(i);
          if i <= Length(Value) then begin
            S := Uppercase(Copy(Value,j,i - j));
            if Copy(S,1,1) = '$' then begin
              j := Pos('-',S);
              if j > 0 then
                S := Copy(S,2,j - 2);
              pMD := AddEntity(Index,meString);
              pMD.S := S;
            end
            else if (S = 'H') or (S = 'HH') then
              AddEntity(Index,meHourElapsed)
            else if (S = 'M') or (S = 'MM') then
              AddEntity(Index,meMinuteElapsed)
            else if (S = 'S') or (S = 'SS') then
              AddEntity(Index,meSecondElapsed)
            else if S = 'BLACK'   then Color := clBlack
            else if S = 'CYAN'    then Color := clFuchsia
            else if S = 'MAGENTA' then Color := clPurple
            else if S = 'WHITE'   then Color := clWhite
            else if S = 'BLUE'    then Color := clBlue
            else if S = 'GREEN'   then Color := clGreen
            else if S = 'RED'     then Color := clRed
            else if S = 'YELLOW'  then Color := clYellow;
          end;
          if (Data.Count > 0) and (Data[Data.Count - 1].Entity in [meHourElapsed,meMinuteElapsed,meSecondElapsed]) then
            Flags := Flags + [mfIsDateTime];
        end;
        'Y','M','D','y','m','d','H','S','h','s':
        begin
          Flags := Flags + [mfIsDateTime];
          S := Lowercase(Value);
          C := S[i];
          j := i;
          while (i <= Length(Value)) and (S[i] = C) do
            Inc(i);
          j := i - j;
          case C of
            'y': if j < 3 then AddEntity(Index,meYear2)
                 else AddEntity(Index,meYear4);
            'm': if (LastTimeMD <> Nil) and (LastTimeMD.Entity in [meHourDig1,meHourDig2]) then begin
                   if j = 1 then AddEntity(Index,meMinuteDig1)
                   else AddEntity(Index,meMinuteDig2);
                 end
                 else begin
                   if j = 1 then AddEntity(Index,meMonthDig1)
                   else if j = 2 then AddEntity(Index,meMonthDig2)
                   else if j = 3 then AddEntity(Index,meMonthShort)
                   else if j = 4 then AddEntity(Index,meMonthName)
                   else AddEntity(Index,meMonthChar);
                 end;
            'd': if j = 1 then AddEntity(Index,meDayDig1)
                 else if j = 2 then AddEntity(Index,meDayDig2)
                 else if j = 3 then AddEntity(Index,meDayShort)
                 else AddEntity(Index,meDayName);
            'h': if j = 1 then AddEntity(Index,meHourDig1)
                 else AddEntity(Index,meHourDig2);
            's': begin
                   if (LastTimeMD <> Nil) and (LastTimeMD.Entity = meMonthDig1) then
                     LastTimeMD.Entity := meMinuteDig1
                   else if (LastTimeMD <> Nil) and (LastTimeMD.Entity = meMonthDig2) then
                     LastTimeMD.Entity := meMinuteDig2;
                   if j = 1 then AddEntity(Index,meSecondDig1)
                   else AddEntity(Index,meSecondDig2);
                 end;
          end;
          LastTimeMD := Data[Data.Count - 1];
          Dec(i);
        end;
        else begin
          pMD := AddEntity(Index,meLitteral);
          pMD.C := Value[i];
          if CharInSet(Value[i],['a','A']) then begin
            if Copy(Value,i,3) = 'a/p' then begin
              pMD.Entity := meStrap;
              Inc(i,2);
            end
            else if Copy(Value,i,5) = 'am/pm' then begin
              pMD.Entity := meStrAmPmLow;
              Inc(i,4);
            end
            else if Copy(Value,i,5) = 'AM/PM' then begin
              pMD.Entity := meStrAmPmUpp;
              Inc(i,4);
            end;
            if pMD.Entity in [meStrap,meStrAmPmLow,meStrAmPmUpp] then begin
              for j := Data.Count - 2 downto 0 do begin
                if Data[j].Entity = meHourDig1 then begin
                  Data[j].Entity := meHourDig1AmPm;
                  Break;
                end
                else if Data[j].Entity = meHourDig2 then begin
                  Data[j].Entity := meHourDig2AmPm;
                  Break;
                end
              end;
            end;
          end;
        end;
      end;
      Inc(i);
    end;
    for i := 0 to Data.Count - 1 do begin
      if Data[i].Entity in [mePlace,meSpace,meZero,meDecimalPos] then begin
        if not (Data[i].Entity = meDecimalPos) then
          Flags := Flags - [mfIsDateTime];
        pMD := Data.Insert(i);
        pMD.Entity := meFirstDigitPlace;
        Break;
      end;
    end;
    if Div1000Cnt > 4 then
      Div1000Cnt := 4;
  end;
end;

function TExcelMask.FormatNumberDateTime(FD: TFormatData; Value: double): AxUCString;
var
  i,j: integer;
  S: string;
  YY,MM,DD,HH,NN,SS,MS: word;
begin
  S := '';
  i := 0;
  DecodeDate(Value,YY,MM,DD);
  DecodeTime(Value,HH,NN,SS,MS);
  with FD do begin
    while i < Data.Count do begin
      case Data[i].Entity of
        meLitteral: begin
          if Data[i].C = '/' then
            S := S + FormatSettings.DateSeparator
          else if Data[i].C = ':' then
            S := S + FormatSettings.TimeSeparator
          else
            S := S + Data[i].C;
        end;
        meString: S := S + Data[i].S;
        meYear2: S := S + FormatDateTime('yy',Value);
        meYear4: S := S + FormatDateTime('yyyy',Value);
        meMonthDig1: S := S + IntToStr(MM);
        meMonthDig2: if MM < 10 then S := S + '0' + IntToStr(MM)
                     else S := S + IntToStr(MM);
        meMonthShort: S := S + FormatSettings.ShortMonthNames[MM];
        meMonthName: S := S + FormatSettings.LongMonthNames[MM];
        meMonthChar: begin
          DecodeDate(Value,YY,MM,DD);
          S := S + FormatSettings.ShortMonthNames[MM][1];
        end;
        meDayDig1: S := S + IntToStr(DD);
        meDayDig2: if DD < 10 then S := S + '0' + IntToStr(DD)
                   else S := S + IntToStr(DD);
        meDayShort: S := S + FormatSettings.ShortDayNames[DayOfWeek(Value)];
        meDayName: S := S + FormatSettings.LongDayNames[DayOfWeek(Value)];
        meHourDig1: S := S + IntToStr(HH);
        meHourDig2: if HH < 10 then S := S + '0' + IntToStr(HH)
                    else S := S + IntToStr(HH);
        meHourDig1AmPm: if HH > 11 then S := S + IntToStr(HH - 12)
                        else S := S + IntToStr(HH);
        meHourDig2AmPm: if HH > 11 then S := S + Format('%.2d',[HH - 12])
                        else S := S + Format('%.2d',[HH]);
        meHourElapsed: S := S + IntToStr(Trunc(Value) * 24 + Round(Frac(Value) / (1 / 24)));
        meMinuteDig1: S := S + IntToStr(NN);
        meMinuteDig2: if NN < 10 then S := S + '0' + IntToStr(NN)
                      else S := S + IntToStr(NN);
        meMinuteElapsed: S := S + IntToStr(Trunc(Value) * (24 * 60) + Round(Frac(Value) / (1 / (24 * 60))));
        meSecondDig1: S := S + IntToStr(SS);
        meSecondDig2: if SS < 10 then S := S + '0' + IntToStr(SS)
                      else S := S + IntToStr(SS);
        meSecondElapsed: S := S + IntToStr(Trunc(Value) * (24 * 60 * 60) + Round(Frac(Value) / (1 / (24 * 60 * 60))));
        meStrAmPmUpp: if HH > 11 then S := S + 'PM'
                      else S := S + 'AM';
        meStrAmPmLow: if HH > 11 then S := S + 'pm'
                      else S := S + 'am';
        meStrap:      if HH > 11 then S := S + 'p'
                      else S := S + 'a';
        meDecimalPos: begin
          Inc(i);
          j := 0;
          while (i < Data.Count) and (Data[i].Entity = meZero) do begin
            Inc(j);
            Inc(i);
          end;
          if j > 3 then
            j := 3;
          if j > 0 then
            S := S + Copy(Format('%.*f',[j,MS / 1000]),2,MAXINT);
        end;
      end;
      Inc(i);
    end;
    Result := S;
  end;
end;

function TExcelMask.FormatNumberNumber(FD: TFormatData; Value: double): AxUCString;
const
  SZ_RESBUF = 300;
var
  i,j,ValPos,ResPos: integer;
  C: WideChar;
  sVal,Res: AxUCString;

function GetNextDigit(DefaultChar: WideChar): WideChar;
begin
  if ValPos > 0 then begin
    Result := sVal[ValPos];
    Dec(ValPos);
  end
  else
    Result := DefaultChar;
end;

function GetNextDecimal(DefaultChar: WideChar): WideChar;
begin
  if ValPos <= Length(sVal) then begin
    Result := sVal[ValPos];
    Inc(ValPos);
  end
  else
    Result := DefaultChar;
end;

begin
  if FD.Data = Nil then begin
    Result := FloatToStr(Value);
    Exit;
  end;
  if FD.Data.Count < 1 then begin
    Result := FloatToStr(Value);
    Exit;
  end;
  if mfGeneral in FD.Flags then begin
    Result := FloatToStr(Value);
    for i := FD.Data.Count - 1 downto 0 do begin
      if FD.Data[i].Entity = meLitteral then
        Result := FD.Data[i].C + Result;
    end;
    Exit;
  end
  else if mfIsDateTime in FD.Flags then begin
    Result := FormatNumberDateTime(FD,Value);
    Exit;
  end

  else if mfScientific in FD.Flags then begin
    Result := FloatToStrF(Value,ffExponent,FD.DecimalCount - 1,FD.DataDecimalPos);
    Exit;
  end;
  SetLength(Res,SZ_RESBUF);
  ResPos := SZ_RESBUF - 1;
  if mfZero in FD.Flags then
    sVal := ''
  else begin
    if FD.Div1000Cnt > 0 then
      Value := Value / Power(10,FD.Div1000Cnt * 3);
    if mfPercent in FD.Flags then
      Value := Value * 100;
    if mfScientific in FD.Flags then
      sVal := Format('%e',[Value])
    else if mfThousand in FD.Flags then
      sVal := Format('%.*n',[FD.DecimalCount,Value])
    else
      sVal := Format('%.*f',[FD.DecimalCount,Value]);
  end;
  if FD.DataDecimalPos < 0 then
    i := FD.Data.Count - 1
  else
    i := FD.DataDecimalPos;
  ValPos := CPos(FormatSettings.DecimalSeparator,sVal) - 1;
  if ValPos < 1 then
    ValPos := Length(sVal);
  while i >= 0 do begin
    case FD.Data[i].Entity of
      meLitteral: begin
        Res[ResPos] := FD.Data[i].C;
        Dec(ResPos);
      end;
      meString: begin
        j := Length(FD.Data[i].S);
        while (j >= 1) and (ResPos > 0) do begin
          Res[ResPos] := FD.Data[i].S[j];
          Dec(j);
          Dec(ResPos);
        end;
      end;
      mePlace: begin
        C := GetNextDigit('!');
        if C <> '!' then begin
          Res[ResPos] := C;
          Dec(ResPos);
        end;
      end;
      meSpace: begin
        Res[ResPos] := GetNextDigit(' ');
        Dec(ResPos);
      end;
      meZero: begin
        Res[ResPos] := GetNextDigit('0');
        Dec(ResPos);
      end;
      mePercentPos: begin
        Res[ResPos] := '%';
        Dec(ResPos);
      end;
      meFirstDigitPlace:
      begin
        while (ValPos > 0) and (ResPos > 0) do begin
//            if sVal[ValPos] <> '0' then begin
            Res[ResPos] := sVal[ValPos];
            Dec(ResPos);
//            end;
          Dec(ValPos);
        end;
      end;
      meRightIndent: begin
        Res[ResPos - 0] := ' ';
        Res[ResPos - 1] := ' ';
        Res[ResPos - 2] := ' ';
        Dec(ResPos,3);
      end;
    end;
    Dec(i);
  end;
  Result := Copy(Res,ResPos + 1,SZ_RESBUF - ResPos - 1);
  if FD.DataDecimalPos >= 0 then begin
    ValPos := CPos(FormatSettings.DecimalSeparator,sVal) + 1;
    if ValPos < 1 then
      Exit;
    ResPos := 1;
    i := FD.DataDecimalPos;
    while i < FD.Data.Count do begin
      case FD.Data[i].Entity of
        meLitteral: begin
          Res[ResPos] := FD.Data[i].C;
          Inc(ResPos);
        end;
        meString: begin
          j := 1;
          while (j <= Length(FD.Data[i].S)) and (ResPos < SZ_RESBUF) do begin
            Res[ResPos] := FD.Data[i].S[j];
            Inc(j);
            Inc(ResPos);
          end;
        end;
        mePlace: begin
          C := GetNextDecimal('!');
          if C <> '!' then begin
            Res[ResPos] := C;
            Inc(ResPos);
          end;
        end;
        meSpace: begin
          Res[ResPos] := GetNextDecimal(' ');
          Inc(ResPos);
        end;
        meZero: begin
          Res[ResPos] := GetNextDecimal('0');
          Inc(ResPos);
        end;
        meDecimalPos: begin
          Res[ResPos] := FormatSettings.DecimalSeparator;
          Inc(ResPos);
        end;
        mePercentPos: begin
          Res[ResPos] := '%';
          Inc(ResPos);
        end;
        meFirstDigitPlace:
        begin
          while (ValPos > 0) and (ResPos > 0) do begin
            Res[ResPos] := sVal[ValPos];
            Dec(ValPos);
            Dec(ResPos);
          end;
        end;
      end;
      Inc(i);
    end;
    SetLength(Res,ResPos - 1);
    Result := Result + Res;
  end;
end;

function TExcelMask.FormatNumber(Value: double): AxUCString;
begin
  if (Formats[1].Data <> Nil) and (Value < 0) then begin
    Result := FormatNumberNumber(Formats[1],Abs(Value));
    FColor := Formats[1].Color;
  end
  else if (Formats[2].Data <> Nil) and (Value = 0) then begin
    Result := FormatNumberNumber(Formats[2],Value);
    FColor := Formats[2].Color;
  end
  else begin
    Result := FormatNumberNumber(Formats[0],Value);
    FColor := Formats[0].Color;
  end;
end;

function TExcelMask.ValueColor(Value: double): TColor;
begin
  if (Formats[1].Data <> Nil) and (Value < 0) then
    Result := Formats[1].Color
  else if (Formats[2].Data <> Nil) and (Value = 0) then
    Result := Formats[2].Color
  else
    Result := Formats[0].Color;
end;

{ TMaskDataList }

function TMaskDataList.Add: TMaskData;
begin
  Result := TMaskData.Create;
  inherited Add(Result);
end;

function TMaskDataList.GetItems(Index: integer): TMaskData;
begin
  Result := TMaskData(inherited Items[Index]);
end;

function TMaskDataList.Insert(AIndex: integer): TMaskData;
begin
  Result := TMaskData.Create;
  inherited Insert(AIndex,Result);
end;

end.
