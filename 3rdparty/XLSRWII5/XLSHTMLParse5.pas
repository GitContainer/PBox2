unit XLSHTMLParse5;

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

uses Classes, SysUtils,
{$ifndef BABOON}
     Graphics,
{$endif}
     XLSUtils5, XLSHTMLTypes5;

type THTMLParser = class(TObject)
private
//     ElementNamesHash: array[0..High(ElementNames)] of word;
//     AttributeNamesHash: array[0..High(AttributeNames)] of word;
     FStream: TStream;
     Buffer: XLS8PChar;
     BufferPos,BytesInBuffer: integer;
     ElementCount: integer;
     Elements: PHTMLElementArray;
     LineCount: integer;
     FSpecialDecodeText: boolean;
     FCSSText: AxUCString;

     procedure CreateHash;
     procedure NewElement(ID: THTMLElementID);
     function  ReadStream: boolean;
     function  GetNextChar(var C: WideChar; PreserveCR: boolean = False): boolean;
     function  GetNextTag: boolean;
     function  GetElementID(S: AxUCString): THTMLElementID;
     function  GetAttributeID(S: AxUCString): TElementAttributeID;
     procedure AddAttribute(Index: integer; ID: TElementAttributeID; Data: AxUCString);
     procedure ScanAttributes(Index: integer; S: AxUCString);
     procedure DoComment;
//     procedure DoCSSComment;
     procedure DoCSS;
     procedure DoScript;
public
     constructor Create(var E: THTMLElementArray);
     destructor Destroy; override;
     procedure Clear;
     procedure LoadFromStream(Stream: TStream);
     procedure LoadFromFile(const Filename: AxUCString);
     function  GetAttribute(E: THTMLElement; Attr: TElementAttributeID): PElementAttribute;
     function  GetStrAttribute(E: THTMLElement; Attr: TElementAttributeID; var S: AxUCString): boolean;

     property SpecialDecodeText: boolean read FSpecialDecodeText write FSpecialDecodeText;
     property CSSText: AxUCString read FCSSText;
     end;

function DecodeQuotedPrintable(const S: AxUCString): AxUCString;

implementation

const
  BUFFER_SIZE = 8192;

type
  TQPTableEntry = record
    CData : AxUCString;
    CChar : byte;
  end;
const
QPTable : array[0..99] of TQPTableEntry =
(
  (CData:'&nbsp;'; CChar: 32),(CData:'&quot;'; CChar: 34),(CData:'&amp;'; CChar: 38),
  (CData:'&lt;'; CChar:   60),(CData:'&gt;'; CChar:   62),

  (* ISO Latin-1 (cnt=95) *)
  (CData:'&iexcl;'; CChar:  161),(CData:'&cent;'; CChar:   162),
  (CData:'&pound;'; CChar:  163),(CData:'&curren;'; CChar: 164),(CData:'&yen;'; CChar:    165),
  (CData:'&brvbar;'; CChar: 166),(CData:'&sect;'; CChar:   167),(CData:'&uml;'; CChar:    168),
  (CData:'&copy;'; CChar:   169),(CData:'&ordf;'; CChar:   170),(CData:'&laquo;'; CChar:  171),
  (CData:'&not;'; CChar:    172),(CData:'&shy;'; CChar:    173),(CData:'&reg;'; CChar:    174),
  (CData:'&macr;'; CChar:   175),(CData:'&deg;'; CChar:    176),(CData:'&plusmn;'; CChar: 177),
  (CData:'&sup2;'; CChar:   178),(CData:'&sup3;'; CChar:   179),(CData:'&acute;'; CChar:  180),
  (CData:'&micro;'; CChar:  181),(CData:'&para;'; CChar:   182),(CData:'&middot;'; CChar: 183),
  (CData:'&cedil;'; CChar:  184),(CData:'&sup1;'; CChar:   185),(CData:'&ordm;'; CChar:   186),
  (CData:'&raquo;'; CChar:  187),(CData:'&frac14;'; CChar: 188),(CData:'&frac12;'; CChar: 189),
  (CData:'&frac34;'; CChar: 190),(CData:'&iquest;'; CChar: 191),(CData:'&Agrave;'; CChar: 192),
  (CData:'&Aacute;'; CChar: 193),(CData:'&Acirc;'; CChar:  194),(CData:'&Atilde;'; CChar: 195),
  (CData:'&Auml;'; CChar:   196),(CData:'&Aring;'; CChar:  197),(CData:'&AElig;'; CChar:  198),
  (CData:'&Ccedil;'; CChar: 199),(CData:'&Egrave;'; CChar: 200),(CData:'&Eacute;'; CChar: 201),
  (CData:'&Ecirc;'; CChar:  202),(CData:'&Euml;'; CChar:   203),(CData:'&Igrave;'; CChar: 204),
  (CData:'&Iacute;'; CChar: 205),(CData:'&Icirc;'; CChar:  206),(CData:'&Iuml;'; CChar:   207),
  (CData:'&ETH;'; CChar:    208),(CData:'&Ntilde;'; CChar: 209),(CData:'&Ograve;'; CChar: 210),
  (CData:'&Oacute;'; CChar: 211),(CData:'&Ocirc;'; CChar:  212),(CData:'&Otilde;'; CChar: 213),
  (CData:'&Ouml;'; CChar:   214),(CData:'&times;'; CChar:  215),(CData:'&Oslash;'; CChar: 216),
  (CData:'&Ugrave;'; CChar: 217),(CData:'&Uacute;'; CChar: 218),(CData:'&Ucircv;'; CChar: 219),
  (CData:'&Uuml;'; CChar:   220),(CData:'&Yacute;'; CChar: 221),(CData:'&THORN;'; CChar:  222),
  (CData:'&szlig;'; CChar:  223),(CData:'&agrave;'; CChar: 224),(CData:'&aacute;'; CChar: 225),
  (CData:'&acirc;'; CChar:  226),(CData:'&atilde;'; CChar: 227),(CData:'&auml;'; CChar:   228),
  (CData:'&aring;'; CChar:  229),(CData:'&aelig;'; CChar:  230),(CData:'&ccedil;'; CChar: 231),
  (CData:'&egrave;'; CChar: 232),(CData:'&eacute;'; CChar: 233),(CData:'&ecirc;'; CChar:  234),
  (CData:'&euml;'; CChar:   235),(CData:'&igrave;'; CChar: 236),(CData:'&iacute;'; CChar: 237),
  (CData:'&icirc;'; CChar:  238),(CData:'&iuml;'; CChar:   239),(CData:'&eth;'; CChar:    240),
  (CData:'&ntilde;'; CChar: 241),(CData:'&ograve;'; CChar: 242),(CData:'&oacute;'; CChar: 243),
  (CData:'&ocirc;'; CChar:  244),(CData:'&otilde;'; CChar: 245),(CData:'&ouml;'; CChar:   246),
  (CData:'&divide;'; CChar: 247),(CData:'&oslash;'; CChar: 248),(CData:'&ugrave;'; CChar: 249),
  (CData:'&uacute;'; CChar: 250),(CData:'&ucirc;'; CChar:  251),(CData:'&uuml;'; CChar:   252),
  (CData:'&yacute;'; CChar: 253),(CData:'&thorn;'; CChar:  254),(CData:'&yuml;'; CChar:   255)
);

function CPos(C: WideChar; S: AxUCString): integer;
begin
  for Result := 1 to Length(S) do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function DecodeQuotedPrintable(const S: AxUCString): AxUCString;
var
  i: integer;
  SC: AxUCString;
begin
  Result := S;
  for i := 1 to Length(Result) - 1 do begin
//    if (Result[i] = '&') and (Result[i + 1] = '#') and ((Length(Result) - i) >= 3) then begin
    if ((Length(Result) - i) >= 3) and (Result[i] = '&') and (Result[i + 1] = '#') then begin 
      Result[i] := WideChar(StrToIntDef(Copy(Result,i + 2,3),32));
      Result := Copy(Result,1,i) + Copy(Result,i + 6,MAXINT);
    end;
  end;
  SetLength(SC,1);
  for i := 0 to High(QPTable) do begin
    SC[1] := WideChar(QPTable[i].CChar);
    Result := _XLSReplaceStr(Result,QPTable[i].CData,SC);
  end;
end;

constructor THTMLParser.Create(var E: THTMLElementArray);
begin
  Elements := @E;
  GetMem(Buffer,BUFFER_SIZE);
  CreateHash;
end;

destructor THTMLParser.Destroy;
begin
  FreeMem(Buffer);
  Clear;
  inherited Destroy;
end;

procedure THTMLParser.Clear;
var
  i,j: integer;
begin
  for i := 0 to High(Elements^) do begin
    for j := 0 to High(Elements^[i].Attributes) do begin
      if Elements^[i].Attributes[j].AttType = atString then
        StrDispose(Elements^[i].Attributes[j].StrVal);
    end;
    SetLength(Elements^[i].Attributes,0);
  end;
  SetLength(Elements^,0);
end;

procedure THTMLParser.CreateHash;
{
var
  i: integer;
}
begin
{
  for i := 0 to High(ElementNames) do
    ElementNamesHash[i] := XLSGetHashCode(ElementNames[i][1],Length(ElementNames[i]));
  for i := 0 to High(AttributeNames) do
    AttributeNamesHash[i] := XLSGetHashCode(AttributeNames[i][1],Length(AttributeNames[i]));
}
end;

procedure THTMLParser.NewElement(ID: THTMLElementID);
begin
  if ElementCount >= Length(Elements^) then
    SetLength(Elements^,Length(Elements^) + 2000);
  FillChar(Elements^[ElementCount],SizeOf(THTMLElement),0);
  Elements^[ElementCount].ID := ID;
  Inc(ElementCount);
end;

function THTMLParser.GetElementID(S: AxUCString): THTMLElementID;
var
  i: integer;
//  Hash: word;
begin
  S := Lowercase(S);
  i := HTMLElements.IndexOf(S);
  if i >= 0 then
    Result := THTMLElementID(i)
  else
    Result := heUNKNOWN;
  {
  Hash := XLSGetHashCode(Pointer(S)^,Length(S));
  for i := 0 to High(ElementNames) do begin
    if (Hash = ElementNamesHash[i]) and (ElementNames[i] = S) then begin
      Result := THTMLElementID(i);
      Exit;
    end;
  end;
  Result := heUNKNOWN;
}
end;

function THTMLParser.GetAttributeID(S: AxUCString): TElementAttributeID;
var
  i: integer;
//  Hash: word;
begin
  S := Lowercase(S);
  i := HTMLAttributes.IndexOf(S);
  if i >= 0 then
    Result := TElementAttributeID(i)
  else
  Result := eaUNKNOWN;

{
  Hash := XLSGetHashCode(Pointer(S)^,Length(S));
  for i := 0 to High(AttributeNames) do begin
    if (Hash = AttributeNamesHash[i]) and (AttributeNames[i] = S) then begin
      Result := TElementAttributeID(i);
      Exit;
    end;
  end;
  Result := eaUNKNOWN;
}
end;

function THTMLParser.ReadStream: boolean;
begin
  BytesInBuffer := FStream.Read(Buffer^,BUFFER_SIZE);
  BufferPos := 0;
  Result := BytesInBuffer > 0;
end;

procedure THTMLParser.DoComment;
var
  C,PrevC: WideChar;
begin
  PrevC := #0;
  while GetNextChar(C) do begin
    if (C = '>') and (PrevC = '-') then
      Exit;
    PrevC := C;
  end;
end;

procedure THTMLParser.DoCSS;
var
  i: integer;
  InQuoted: boolean;
  InValue: boolean;
  C,PrevC: WideChar;
begin
  i := 0;
  PrevC := #0;
  InQuoted := False;
  InValue := False;
  while GetNextChar(C) do begin
    if (C = '>') and (PrevC = '-') then
      Break;
{
    if (C = '/') and (PrevC = '*') then begin
      DoCSSComment;
      continue;
    end;
}
    if CharInSet(C,[#9,#10,#13]) then
      C := ' '
    else if C = '"' then
      InQuoted := not InQuoted;
    if not InQuoted then begin
      if C = ':' then
        InValue := True
      else if C = ';' then
        InValue := False;
    end;
    if InQuoted or InValue or not (C = ' ') then begin
      Inc(i);
      if i > Length(FCSSText) then
        SetLength(FCSSText,Length(FCSSText) + 1024);
      FCSSText[i] := C;
    end;

    PrevC := C;
  end;
  SetLength(FCSSText,i);
  FCSSText := Copy(FCSSText,5,Length(FCSSText) - 8);
end;

//procedure THTMLParser.DoCSSComment;
//var
//  C,PrevC: WideChar;
//begin
//  PrevC := #0;
//  while GetNextChar(C) do begin
//    if (C = '*') and (PrevC = '/') then
//      Exit;
//    PrevC := C;
//  end;
//end;

procedure THTMLParser.DoScript;
var
  S: AxUCString;
  C: WideChar;
  p: integer;
begin
  S := '</script>';
  while GetNextChar(C) do begin
    if C = '<' then begin
      p := 2;
      while GetNextChar(C) and (Lowercase(C) = S[p]) and (p < Length(S)) do
        Inc(p);
      if (Lowercase(C) = S[p]) and (p = Length(S)) then begin
        NewElement(heEndScript);
        Exit;
      end;
    end;
  end;
end;

function THTMLParser.GetNextChar(var C: WideChar; PreserveCR: boolean = False): boolean;

function RStripBlank: boolean;
var
  C: WideChar;
begin
  Result := True;
  if BufferPos >= BytesInBuffer then
    Result := ReadStream;
  if Result then begin
    C := WideChar(Buffer[BufferPos]);
    if C = ' ' then begin
      Inc(BufferPos);
      Result := RStripBlank;
    end;
  end;
end;

begin
  Result := True;
  if BufferPos >= BytesInBuffer then
    Result := ReadStream;
  if Result then begin
    C := WideChar(Buffer[BufferPos]);
    Inc(BufferPos);
    if not PreserveCR and CharInSet(C,[#10,#13]) then begin
      if C = #13 then
        Inc(LineCount);
      Result := RStripBlank;
      if Result then
        Result := GetNextChar(C);
    end;
  end;
end;

procedure THTMLParser.AddAttribute(Index: integer; ID: TElementAttributeID; Data: AxUCString);
var
  i: integer;
  S: AxUCString;

function GetHtmlColor(s: AxUCString): TColor;
var
  Col: integer;
begin
  if (CompareText(s, 'white') = 0)   then Result := clWhite     else
  if (CompareText(s, 'black') = 0)   then Result := clBlack     else
  if (CompareText(s, 'red') = 0)     then Result := clRed       else
  if (CompareText(s, 'green') = 0)   then Result := clGreen     else
  if (CompareText(s, 'blue') = 0)    then Result := clBlue      else
  if (CompareText(s, 'fuchsia') = 0) then Result := clFuchsia   else
  if (CompareText(s, 'purple') = 0)  then Result := clPurple    else
  if (CompareText(s, 'yellow') = 0)  then Result := clYellow    else
{$ifdef BABOON}
{$else}
  if (CompareText(s, 'aqua') = 0)    then Result := clAqua      else
  if (CompareText(s, 'gray') = 0)    then Result := clDkGray	  else
  if (CompareText(s, 'lime') = 0)    then Result := clLime      else
  if (CompareText(s, 'maroon') = 0)  then Result := clMaroon    else
  if (CompareText(s, 'navy') = 0)    then Result := clNavy      else
  if (CompareText(s, 'olive') = 0)   then Result := clOlive     else
  if (CompareText(s, 'silver') = 0)  then Result := clGray      else
  if (CompareText(s, 'teal') = 0)    then Result := clTeal      else
{$endif}
  begin
     if (s <> '') and (s[1] = '#') then
       Delete(s, 1, 1);
     Col := StrToIntDef('$'+s, -1);
     if Col <> -1 then
       Result := ((Col and $FF) shl 16) or (Col and $FF00) or ((Col and $FF0000) shr 16)
     else
       Result := 0;
  end;
end;

procedure SpecialDecode(var Data: AxUCString);
var
  i: integer;
  Found: boolean;
begin
  for i := 1 to Length(Data) - 1 do begin
    if Byte(Data[i]) = $C3 then begin
      Found := True;
      case Byte(Data[i + 1]) of
        $A4: Data[i] := 'ä';
        $B6: Data[i] := 'ö';
        $A5: Data[i] := 'å';
        $96: Data[i] := 'Ö';
        else Found := False;
      end;
      if Found then
        Data := Copy(Data,1,i) + Copy(Data,i + 2,MAXINT);
    end;
  end;
end;

begin
  if ID = eaText then begin
    Data := DecodeQuotedPrintable(Data);
    if FSpecialDecodeText then
      SpecialDecode(Data);
  end;
  i := Length(Elements^[Index].Attributes);
  SetLength(Elements^[Index].Attributes,i + 1);
  Elements^[Index].Attributes[i].ID := ID;
  if ID in [eaSize] then begin
    Elements^[Index].Attributes[i].AttType := atNumber;
    Elements^[Index].Attributes[i].NumVal := StrToIntDef(Data,0);
    if (Data <> '') and (Data[1] = '+') then
      Elements^[Index].Attributes[i].NumVal := -(Elements^[Index].Attributes[i].NumVal + 100000);
  end
  else if ID in [eaAlign] then begin
    S := Lowercase(Data);
    Elements^[Index].Attributes[i].AttType := atAlign;
    with Elements^[Index].Attributes[i] do begin
           if S = 'left'    then AlignVal := aaLeft
      else if S = 'center'  then AlignVal := aaCenter
      else if S = 'right'   then AlignVal := aaRight
      else if S = 'justify' then AlignVal := aaJustify
      else if S = 'char'    then AlignVal := aaChar
      else if S = 'top'     then AlignVal := aaTop
      else if S = 'middle'  then AlignVal := aaMiddle
      else if S = 'bottom'  then AlignVal := aaBottom
      else AlignVal := aaLeft;
    end;
  end
  else if ID in [eaColor,eaBgcolor] then begin
    Elements^[Index].Attributes[i].AttType := atColor;
    Elements^[Index].Attributes[i].ColorVal := GetHtmlColor(Data);
  end
  else if ID in [eaHeight,eaWidth] then begin
    if CPos('%',Data) > 0 then begin
      Elements^[Index].Attributes[i].AttType := atPercent;
      Elements^[Index].Attributes[i].PercentVal := StrToIntDef(Copy(Data,1,Length(Data) - 1),100);
    end
    else begin
      Elements^[Index].Attributes[i].AttType := atNumber;
      Elements^[Index].Attributes[i].NumVal := StrToIntDef(Data,0);
    end;
  end
  else begin
    Elements^[Index].Attributes[i].AttType := atString;
    Elements^[Index].Attributes[i].StrVal := StrAlloc(Length(Data) + 1);
    StrPCopy(Elements^[Index].Attributes[i].StrVal,Data);
  end;
end;

procedure THTMLParser.ScanAttributes(Index: integer; S: AxUCString);
var
  p,i: integer;
  Val: AxUCString;
  ID: TElementAttributeID;
begin
  repeat
    S := Trim(S);
    p := CPos('=',S);
    if p > 0 then begin
      ID := GetAttributeID(Copy(S,1,p - 1));
      if (Length(S) >= p + 1) and (S[p + 1] = '"') then begin
        i := p + 2;
        while (i < Length(S)) and (S[i] <> '"') do
          Inc(i);
        Val := Copy(S,P + 2,i - (P + 2));
        S := Copy(S,i + 1,1024);
      end
      else begin
        i := CPos(' ',S);
        if i > 0 then begin
          Val := Copy(S,p + 1,i - (p + 1));
          S := Copy(S,i + 1,1024);
        end
        else begin
          Val := Copy(S,p + 1,MAXINT);
          S := '';
        end;
      end;
      AddAttribute(Index,ID,Val);
    end;
  until (p <= 0);
end;

function THTMLParser.GetNextTag: boolean;
var
  C: WideChar;
  i,p: integer;
  OK: boolean;
  S,Tag: AxUCString;
begin
  Result := True;
  NewElement(heUNKNOWN);
  i := ElementCount - 1;
  Tag := '';
  S := '';
  C := #0;
  while GetNextChar(C,True) and (C <> '<') do
    S := S + C;
  if S <> '' then begin
    Elements^[i].ID := heText;
    AddAttribute(i,eaText,S);
//    AddAttribute(i,eaText,TrimLeft(S));
    NewElement(heUNKNOWN);
    i := ElementCount - 1;
  end;
  OK := GetNextChar(C);
  if OK then begin
    if C = '!' then begin
      if GetNextChar(C) and (C = '-') then
        DoComment
    end
    else if C = '/' then begin
      while GetNextChar(C) and (C <> '>') do
        Tag := Tag + C;
      if C = '>' then
        Elements^[i].ID := THTMLElementID(Integer(GetElementID(Tag)) + Integer(heBeginEndtags) + 1)
      else
        Elements^[i].ID := heEndUNKNOWN;
      Exit;
    end;
  end;
  if CharInSet(C,['A'..'Z']) or CharInSet(C,['a'..'z']) then begin
    Tag := C;
    while GetNextChar(C) and (C <> '>') do
      Tag := Tag + C;
    if C = '>' then begin
      p := CPos(' ',Tag);
      if p > 0 then begin
        S := Copy(Tag,1,p - 1);
        ScanAttributes(i,Copy(Tag,p + 1,1024));
      end
      else
        S := Tag;
      Elements^[i].ID := GetElementID(S);
      if Elements^[i].ID = heScript then
        DoScript
      else if Elements^[i].ID = heStyle then
        DoCSS;
    end;
  end
  else if not OK then begin
    Elements^[i].ID := heEOF;
    Result := False;
  end;
end;

procedure THTMLParser.LoadFromStream(Stream: TStream);
begin
  FStream := Stream;
  LineCount := 0;
  try
    while GetNextTag do ;
  except
    raise XLSRWException.Create('HTML Error on line #' + IntToStr(LineCount));
  end;
  SetLength(Elements^,ElementCount);
end;

procedure THTMLParser.LoadFromFile(const Filename: AxUCString);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead + fmShareDenyNone);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

function THTMLParser.GetStrAttribute(E: THTMLElement; Attr: TElementAttributeID; var S: AxUCString): boolean;
var
  i: integer;
begin
  for i := 0 to High(E.Attributes) do begin
    if (E.Attributes[i].ID = Attr) and (E.Attributes[i].AttType = atString) then begin
      S := E.Attributes[i].StrVal;
      Result := True;
      Exit;
    end;
  end;
  S := '';
  Result := False;
end;

function THTMLParser.GetAttribute(E: THTMLElement; Attr: TElementAttributeID): PElementAttribute;
var
  i: integer;
begin
  for i := 0 to High(E.Attributes) do begin
    if E.Attributes[i].ID = Attr then begin
      Result := @E.Attributes[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

end.
