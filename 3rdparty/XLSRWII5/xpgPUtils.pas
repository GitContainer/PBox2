unit xpgPUtils;

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils,
	   XLSUtils5; 

{$ifdef DELPHI_2009_OR_LATER}
type AxUCChar = char;
type AxUCString = string;
type AxPUCChar = PChar;
{$else}
type AxUCChar = WideChar;
type AxUCString = WideString;
type AxPUCChar = PWideChar;
{$endif}

const XPG_UNKNOWN_ENUM = $000000FF;

type TXpgErrorMsg = (xemUnknownAttribute,xemUnknownElement);

type TWriteElementEvent = procedure (Sender: TObject; var WriteElement: boolean) of object;

function  CPos(const C: AxUCChar; const S: AxUCString): integer;
function  RCPos(const C: AxUCChar; const S: AxUCString): integer;
function  SplitAtChar(const C: AxUCChar; var S: AxUCString): AxUCString;

function XmlStrToFloat(const AValue: AxUCString): extended;
function XmlStrToBool(const AValue: AxUCString): boolean;

function XmlStrToIntDef(const AValue: AxUCString; ADefault: integer): integer; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function XmlStrToBoolDef(const AValue: AxUCString; ADefault: boolean): boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function XmlStrToFloatDef(const AValue: AxUCString; ADefault: extended): extended; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
function XmlStrHexToIntDef(const AValue: AxUCString; ADefault: integer): integer;// {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

function XmlStrToDateTime(const AValue: AxUCString): TDateTime;
function XmlStrToTime(const AValue: AxUCString): TDateTime;
function XmlStrToDate(const AValue: AxUCString): TDateTime;

function XmlTryStrToInt(const AValue: AxUCString; out AResult: integer): boolean; overload;
function XmlTryStrToInt(const AValue: AxUCString; out AResult: longword): boolean; overload;
function XmlTryStrToBool(const AValue: AxUCString; out AResult: boolean): boolean;
function XmlTryStrToFloat(const AValue: AxUCString; out AResult: extended): boolean;
function XmlTryStrToDateTime(const AValue: AxUCString; out AResult: TDateTime): boolean;
function XmlTryStrToTime(const AValue: AxUCString; out AResult: TDateTime): boolean;
function XmlTryStrToDate(const AValue: AxUCString; out AResult: TDateTime): boolean;

function XmlIntToStr(const AValue: integer): AxUCString;
function XmlIntToHexStr(const AValue: integer): AxUCString; overload;
function XmlIntToHexStr(const AValue: integer; ADigits: integer): AxUCString; overload;
function XmlFloatToStr(const AValue: extended): AxUCString;
function XmlDateTimeToStr(const AValue: TDateTime): AxUCString;
function XmlBoolToStr(const AValue: boolean): AxUCString;

function XmlStrToBoolToStr(const AValue: AxUCString): AxUCString;

procedure HtmlCheckText(var AText: AxUCString);

function  NameOfQName(AQName: AxUCString): AxUCString;

function XPGDecodeBase64(const Value: AnsiString): AnsiString;

type TXpgPErrors = class(TObject)
private
     procedure SetErrorEvent(const Value: TNotifyEvent);
protected
     FNoDuplicates: boolean;
     FErrorText: AxUCString;

     FErrorEvent: TNotifyEvent;
public
     constructor Create;

     procedure Error(Msg: TXpgErrorMsg; const Args: AxUCString; ARow: integer = 0);
     function  Count: integer;
     procedure GetDuplicates(AStrings: TStrings);

     property NoDuplicates: boolean read FNoDuplicates write FNoDuplicates;
     property LastErrorText: AxUCString read FErrorText;

     property OnError: TNotifyEvent read FErrorEvent write SetErrorEvent;
     end;

const HASH_FUNC_COUNT = 5;
function CalcHash_A(const Str : AxUCString): longword;
function CalcHash_A_P(const Str : PAnsiChar): longword;
function CalcHash_B(const Str : AxUCString): longword;
function CalcHash_C(const Str : AxUCString): longword;
function CalcHash_D(const Str : AxUCString): longword;
function CalcHash_CRC32(const Str : AxUCString): longword;

implementation

const
  CRC32_POLYNOMIAL = $EDB88320;

  ReTablebase64 =
    #$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$3E +#$40
    +#$40 +#$40 +#$3F +#$34 +#$35 +#$36 +#$37 +#$38 +#$39 +#$3A +#$3B +#$3C
    +#$3D +#$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$40 +#$00 +#$01 +#$02 +#$03
    +#$04 +#$05 +#$06 +#$07 +#$08 +#$09 +#$0A +#$0B +#$0C +#$0D +#$0E +#$0F
    +#$10 +#$11 +#$12 +#$13 +#$14 +#$15 +#$16 +#$17 +#$18 +#$19 +#$40 +#$40
    +#$40 +#$40 +#$40 +#$40 +#$1A +#$1B +#$1C +#$1D +#$1E +#$1F +#$20 +#$21
    +#$22 +#$23 +#$24 +#$25 +#$26 +#$27 +#$28 +#$29 +#$2A +#$2B +#$2C +#$2D
    +#$2E +#$2F +#$30 +#$31 +#$32 +#$33 +#$40 +#$40 +#$40 +#$40 +#$40 +#$40;

var
  Ccitt32Table: array[0..255] of longword;

procedure BuildCRCTable;
var i, j: longint;
    value: longword;
begin
  for i := 0 to 255 do begin
    value := i;
    for j := 8 downto 1 do
      if ((value and 1) <> 0) then
        value := (value shr 1) xor CRC32_POLYNOMIAL
      else
        value := value shr 1;
    Ccitt32Table[i] := value;
  end
end;

// From Synapse synacode
function Decode4to3Ex(const Value, Table: AnsiString): AnsiString;
type
  TDconvert = record
    case byte of
      0: (a0, a1, a2, a3: AnsiChar);
      1: (i: integer);
  end;
var
  x, y, l, lv: Integer;
  d: TDconvert;
  dl: integer;
  c: byte;
  p: ^AnsiChar;
begin
  lv := Length(Value);
  SetLength(Result, lv);
  x := 1;
  dl := 4;
  d.i := 0;
  p := pointer(result);
  while x <= lv do
  begin
    y := Ord(Value[x]);
    if y in [33..127] then
      c := Ord(Table[y - 32])
    else
      c := 64;
    Inc(x);
    if c > 63 then
      continue;
    d.i := (d.i shl 6) or c;
    dec(dl);
    if dl <> 0 then
      continue;
    p^ := d.a2;
    inc(p);
    p^ := d.a1;
    inc(p);
    p^ := d.a0;
    inc(p);
    d.i := 0;
    dl := 4;
  end;
  case dl of
    1:
      begin
        d.i := d.i shr 2;
        p^ := d.a1;
        inc(p);
        p^ := d.a0;
        inc(p);
      end;
    2:
      begin
        d.i := d.i shr 4;
        p^ := d.a0;
        inc(p);
      end;
  end;
  l := integer(p) - integer(pointer(result));
  SetLength(Result, l);
end;

function CPos(const C: AxUCChar; const S: AxUCString): integer;
begin
  for Result := 1 to Length(S) do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function RCPos(const C: AxUCChar; const S: AxUCString): integer;
begin
  for Result := Length(S) downto 1 do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function SplitAtChar(const C: AxUCChar; var S: AxUCString): AxUCString;
var
  p: integer;
begin
  p := CPos(C,S);
  if p > 0 then begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

function XmlStrToFloat(const AValue: AxUCString): extended;
var
  TempDS: AxUCChar;
begin
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    if not TryStrToFloat(AValue,Result) then
      Result := 0;
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

function XmlStrToBool(const AValue: AxUCString): boolean;
var
  S: AxUCString;
begin
  S := Lowercase(AValue);
  if (S = 'true') or (S = '1') or (S = 't') then
    Result := True
  else if (S = 'false') or (S = '0')  or (S = 'f') then
    Result := False
  else
    raise Exception.CreateFmt('Invalid boolean value "%s"',[AValue]);
end;

function XmlStrToIntDef(const AValue: AxUCString; ADefault: integer): integer;
begin
//  Inc(G_Counter);
  Result := StrToIntDef(AValue,ADefault);
end;

function XmlStrToBoolDef(const AValue: AxUCString; ADefault: boolean): boolean;
var
  S: AxUCString;
begin
  S := Lowercase(AValue);
  if (S = 'true') or (S = '1')  or (S = 't') then
    Result := True
  else if (S = 'false') or (S = '0')  or (S = 'f') then
    Result := False
  else
    Result := ADefault;
end;

function XmlStrToFloatDef(const AValue: AxUCString; ADefault: extended): extended;
var
  TempDS: AxUCChar;
begin
//  Inc(G_Counter);
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    Result := StrToFloatDef(AValue,ADefault);
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

function XmlStrHexToIntDef(const AValue: AxUCString; ADefault: integer): integer;
begin
  if (AValue <> '') and (AValue[1] = '#') then
    Result := StrToIntDef('$' + Copy(AValue,2,MAXINT),ADefault)
  else
    Result := StrToIntDef(AValue,ADefault);
end;

function XmlStrToDateTime(const AValue: AxUCString): TDateTime;
begin
  if not TryStrToDateTime(AValue,Result) then
    Result := 0;
end;

function XmlStrToTime(const AValue: AxUCString): TDateTime;
begin
  raise Exception.Create('Not implemented: XmlStrToTime');
end;

function XmlStrToDate(const AValue: AxUCString): TDateTime;
begin
  raise Exception.Create('Not implemented: XmlStrToDate');
end;

function XmlTryStrToInt(const AValue: AxUCString; out AResult: integer): boolean;
begin
//  Inc(G_Counter);
  Result := TryStrToInt(AValue,AResult);
end;

// TODO fix this without try..except
function XmlTryStrToInt(const AValue: AxUCString; out AResult: longword): boolean;
begin
//  Inc(G_Counter);
  try
    AResult := StrToInt(AValue);
    Result := True;
  except
    Result := False;
  end;
end;

function XmlTryStrToBool(const AValue: AxUCString; out AResult: boolean): boolean;
var
  S: AxUCString;
begin
  S := Lowercase(AValue);
  if S = 'true' then begin
    AResult := True;
    Result := True;
  end
  else if S = 'false' then begin
    AResult := False;
    Result := True;
  end
  else
    Result := False;
end;

function XmlTryStrToFloat(const AValue: AxUCString; out AResult: extended): boolean;
var
  TempDS: AxUCChar;
begin
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    Result := TryStrToFloat(AValue,AResult);
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

function XmlTryStrToDateTime(const AValue: AxUCString; out AResult: TDateTime): boolean;
begin
  raise Exception.Create('Not implemented: TryXmlStrToDateTime');
end;

function XmlTryStrToTime(const AValue: AxUCString; out AResult: TDateTime): boolean;
begin
  raise Exception.Create('Not implemented: TryXmlStrToTime');
end;

function XmlTryStrToDate(const AValue: AxUCString; out AResult: TDateTime): boolean;
begin
  raise Exception.Create('Not implemented: TryXmlStrToDate');
end;

function XmlIntToStr(const AValue: integer): AxUCString;
begin
  Result := IntToStr(AValue);
end;

function XmlIntToHexStr(const AValue: integer): AxUCString;
begin
  Result := IntToHex(AValue,6);
end;

function XmlIntToHexStr(const AValue: integer; ADigits: integer): AxUCString; overload;
begin
  Result := IntToHex(AValue,ADigits);
end;

function XmlFloatToStr(const AValue: extended): AxUCString;
var
  TempDS: AxUCChar;
begin
  TempDS := FormatSettings.DecimalSeparator;
  FormatSettings.DecimalSeparator := '.';
  try
    Result := FloatToStr(AValue);
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

// 2011-01-27T14:42:50Z
function XmlDateTimeToStr(const AValue: TDateTime): AxUCString;
var
  YY,MM,DD,HH,NN,SS,MS: word;
begin
  DecodeDate(AValue,YY,MM,DD);
  DecodeTime(AValue,HH,NN,SS,MS);
  Result := Format('%d-%.2d-%.2dT%.2d:%.2d:%.2dZ',[YY,MM,DD,HH,NN,SS]);
//  Result := IntToStr(YY) + '-' + IntToStr(MM) + '-' + IntToStr(DD) + 'T' + IntToStr(HH) + ':' + IntToStr(NN) + ':' + IntToStr(SS) + 'Z';
end;

function XmlBoolToStr(const AValue: boolean): AxUCString;
begin
  if AValue then
    Result := '1'
  else
    Result := '0';
end;

function XmlStrToBoolToStr(const AValue: AxUCString): AxUCString;
var
  S: AxUCString;
begin
  S := Lowercase(AValue);
  if (S = '1') or (S = 'true') then
    Result := 'True'
  else
    Result := 'False';
end;

procedure HtmlCheckText(var AText: AxUCString);
var
  i,j: integer;
  SkipNext: boolean;
  S: AxUCString;
begin
  i := 1;
  while (i <= Length(AText)) and ((AText[i] = #10) or (AText[i] = #13)) do
    Inc(i);
  j := Length(AText);
  while (j >= 1) and ((AText[i] = #10) or (AText[i] = #13)) do
    Dec(j);

  AText := Copy(AText,i,j - i + 1);
  SetLength(S,Length(AText));

  j := 0;
  SkipNext := False;
  for i := 1 to Length(AText) do begin
    if SkipNext then begin
      SkipNext := False;
      Continue;
    end;

    case AText[i] of
      #9 : ;
      #10: begin
        Inc(j);
        S[j] := ' ';
        SkipNext := (i < Length(AText)) and (AText[i + 1] = #13);
      end;
      #13: begin
        Inc(j);
        S[j] := ' ';
        SkipNext := (i < Length(AText)) and (AText[i + 1] = #10);
      end
      else begin
        Inc(j);
        S[j] := AText[i];
      end;
    end;

  end;
  SetLength(S,j);
  AText := S;
end;

function NameOfQName(AQName: AxUCString): AxUCString;
var
  p: integer;
begin
  p := CPOs(':',AQName);
  if p > 0 then
    Result := Copy(AQName,p + 1,MAXINT)
  else
    Result := AQName;
end;

function XPGDecodeBase64(const Value: AnsiString): AnsiString;
begin
  Result := Decode4to3Ex(Value, ReTableBase64);
end;

// Don't works when optimization is on. Seems to corrupt memory.
function _CalcHash_A(const Buffer; Count: Integer): longword; assembler;
asm
        CMP     EDX,0
        JNE     @@2
        MOV     EAX,0
        JMP     @@3
@@2:
        MOV     ECX,EDX
        MOV     EDX,EAX
        XOR     EAX,EAX
        XOR     EBX,EBX
@@1:    MOV     BL,[EDX]
        ADD     EAX,EBX;
        INC     EDX
        DEC     ECX
        JNE     @@1
@@3:
end;

function CalcHash_A(const Str : AxUCString): longword;
var
  i: integer;
  P: PByteArray;
begin
  Result := 0;
  P := @Str[1];
  for i := 0 to Length(Str) * 2 - 1 do
    Result := Result + P[i];

//  Result := _CalcHash_A(Pointer(Str)^,Length(Str) * 2);
end;

function CalcHash_A_P(const Str : PAnsiChar): longword;
var
  i: integer;
begin
  Result := 0;
  i := 0;
  while Str[i] <> #0 do begin
    Result := Result + Byte(Str[i]);
    Inc(i);
  end;
end;

// RSHash
function CalcHash_B(const Str : AxUCString) : longword;
const b = 378551;
var
  a : Cardinal;
  i : Integer;
begin
  a      := 63689;
  Result := 0;
  for i := 1 to Length(Str) do
  begin
    Result := Result * a + Ord(Str[i]);
    a      := a * b;
  end;
end;

// JSHash
function CalcHash_C(const Str : AxUCString) : longword;
var
  i : Integer;
begin
  Result := 1315423911;
  for i := 1 to Length(Str) do
  begin
    Result := Result xor ((Result shl 5) + Ord(Str[i]) + (Result shr 2));
  end;
end;

// APHash
function CalcHash_D(const Str : AxUCString) : longword;
var
  i : Cardinal;
begin
  Result := $AAAAAAAA;
  for i := 1 to Length(Str) do
  begin
    if ((i - 1) and 1) = 0 then
      Result := Result xor ((Result shl 7) xor Ord(Str[i]) * (Result shr 3))
    else
      Result := Result xor (not((Result shl 11) + Ord(Str[i]) xor (Result shr 5)));
  end;
end;

function CalcHash_CRC32(const Str : AxUCString): longword;
var
  i: integer;
  b: byte;
begin
  Result:=$FFFFFFFF;
  for i := 1 to Length(Str) do begin
    if (Word(Str[i]) and $FF00) <> 0 then begin
      b := Byte(Word(Str[i]) shr 8);
      Result:= (((Result shr 8) and $00FFFFFF) xor (Ccitt32Table[(Result xor b) and $FF]));
    end;
    b := Byte(Word(Str[i]) and $00FF);
    Result:= (((Result shr 8) and $00FFFFFF) xor (Ccitt32Table[(Result xor b) and $FF]));
  end;
end;

{ TXpgPErrors }

function TXpgPErrors.Count: integer;
begin
  Result := 0;
end;

constructor TXpgPErrors.Create;
begin

end;

procedure TXpgPErrors.Error(Msg: TXpgErrorMsg; const Args: AxUCString; ARow: integer);
begin
  case Msg of
    xemUnknownAttribute: FErrorText := Format('Unknown Attribute %s',[Args]);
    xemUnknownElement  : FErrorText := Format('Unknown Element %s',[Args]);
  end;
  FErrorText := Format('Row %d: ',[ARow]) + FErrorText;

  if Assigned(FErrorEvent) then
    FErrorEvent(Self);
//  else if Msg = xemUnknownElement then
//    raise Exception.Create(FErrorText);
end;

procedure TXpgPErrors.GetDuplicates(AStrings: TStrings);
begin

end;

procedure TXpgPErrors.SetErrorEvent(const Value: TNotifyEvent);
begin
  FErrorEvent := Value;
end;

initialization
  BuildCRCTable;

end.

