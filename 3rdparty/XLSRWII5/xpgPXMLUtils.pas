unit xpgPXMLUtils;

interface

uses Classes, SysUtils,
{$ifndef BABOON}
     Windows,
{$endif}
     xpgPUtils;

// TODO Fix overflow check
const AX_XML_MAX_ATTRIBUTES    = $FF;

// TODO Move to other unit
type Ax8Char = AnsiChar;
type Ax8String = AnsiString;

type TXpgAssigned =  (xaElements,xaAttributes,xaContent,xaRead);
type TXpgAssigneds = set of TXpgAssigned;

type TXpgXMLString = record
     pStart: PAnsiChar;
     pEnd: PAnsiChar;
     end;

type TXpgXMLUCString = record
     pStart: AxPUCChar;
     pEnd: AxPUCChar;
     end;

type PXpgXmlAttribute = ^TXpgXmlAttribute;
     TXpgXmlAttribute = record
     Attribute: TXpgXMLString;
     Value: TXpgXMLString;
     sAttr,sVal: AxUCString;
     Flag: integer;
     end;

type TXpgXmlAttributeArray = array of TXpgXmlAttribute;

type TXpgXMLAttributeList = class(TObject)
private
     function GetAttributes(Index: integer): AxUCString;
     function GetValues(Index: integer): AxUCString;
     function GetHashA(Index: integer): longword;
     function GetHashB(Index: integer): longword;
     function GetHashC(Index: integer): longword;
     function GetHashD(Index: integer): longword;
protected
     FSortList: TList;
     FSorted: boolean;
     FCount: integer;
     FSnatching: boolean;
     FBadEnumValue: string;

     function  GetAsString(const Str: TXpgXMLString): AxUCString;
     function  GetAsStringUTF8(const Str: TXpgXMLString): AxUCString;
     procedure MakeStrings;
     procedure Sort;
public
     FAttributes: TXpgXmlAttributeArray;

     constructor Create;
     destructor Destroy; override;

     procedure Clear;
     procedure ClearNamespace;

     procedure Assign(AAttributes: TXpgXmlAttributeArray; ACount: integer);
     function  Count: integer;
     function  Find(const AName: AxUCString): integer; overload;
     function  Find(const AName: AxUCString; out AValue: AxUCString): boolean; overload;
     procedure CollectPrefixed(APrefix: string; AList: TStrings);
     procedure BeginSnatch;
     procedure EndSnatch(BadAttributes: TStrings);

     function  AsString(const AName: AxUCString): AxUCString;
     // Error in AST, these names can't be overloaded. Probably because there is no parameter in the first.
     function  AsXmlText: AxUCString;
     function  AsXmlText2(AIndex: integer): AxUCString;

     function  AsBoolDef(const AName: AxUCString; ADefault: boolean): boolean;
     function  AsIntegerDef(const AName: AxUCString; ADefault: integer): integer;
     function  AsStringDef(const AName: AxUCString; const ADefault: AxUCString): AxUCString;
     function  AsEnumDef(const AName: AxUCString; AEnums: array of AxUCString; ADefault: integer): integer;
     function  EnumError: AxUCString;

     procedure Add(AAttributes: TXpgXmlAttributeArray; ACount: integer); overload;
     procedure Add(const AName,AValue: AxUCString); overload;

     property Attributes[Index: integer]: AxUCString read GetAttributes; default;
     property Values[Index: integer]: AxUCString read GetValues;
     property HashA[Index: integer]: longword read GetHashA;
     property HashB[Index: integer]: longword read GetHashB;
     property HashC[Index: integer]: longword read GetHashC;
     property HashD[Index: integer]: longword read GetHashD;
     end;

implementation

// uses Windows;

{ TAxXMLAttributeList }

procedure TXpgXMLAttributeList.Add(AAttributes: TXpgXmlAttributeArray; ACount: integer);
begin
  Clear;
  FAttributes := AAttributes;
  FCount := ACount;
  MakeStrings;
end;

procedure TXpgXMLAttributeList.Add(const AName, AValue: AxUCString);
begin
  if Length(FAttributes) <= FCount then
    SetLength(FAttributes,Length(FAttributes) + AX_XML_MAX_ATTRIBUTES);
  FAttributes[FCount].sAttr := AName;
  FAttributes[FCount].sVal := AValue;
  FAttributes[FCount].Flag := 0;
  Inc(FCount);
end;

function TXpgXMLAttributeList.AsBoolDef(const AName: AxUCString; ADefault: boolean): boolean;
var
  S: AxUCString;
begin
  if Find(AName,S) then begin
    S := Lowercase(S);
    if S = 'true' then
      Result := True
    else if S = 'false' then
      Result := False
    else
      Result := ADefault;
  end
  else
    Result := ADefault;
end;

function TXpgXMLAttributeList.AsEnumDef(const AName: AxUCString; AEnums: array of AxUCString; ADefault: integer): integer;
var
  S: AxUCString;
begin
  if Find(AName,S) then begin
    if FSnatching then
      FBadEnumValue := '';
    for Result := 0 to High(AEnums) do begin
      if S = AEnums[Result] then
        Exit;
    end;
    if FSnatching then
      FBadEnumValue := S;
    Result := ADefault;
  end
  else
    Result := ADefault;
end;

function TXpgXMLAttributeList.AsIntegerDef(const AName: AxUCString; ADefault: integer): integer;
var
  S: AxUCString;
begin
  if not Find(AName,S) then
    Result := ADefault
  else begin
    if S = 'unbounded' then
      Result := MAXINT
    else
      Result := StrToIntDef(S,ADefault);
  end;
end;

procedure TXpgXMLAttributeList.Assign(AAttributes: TXpgXmlAttributeArray; ACount: integer);
var
  i: integer;
begin
  for i := 0 to ACount - 1 do
    Add(AAttributes[i].sAttr,AAttributes[i].sVal);
end;

function TXpgXMLAttributeList.AsString(const AName: AxUCString): AxUCString;
begin
  if not Find(AName,Result) then
    Result := '';
end;

function TXpgXMLAttributeList.AsStringDef(const AName, ADefault: AxUCString): AxUCString;
begin
  if not Find(AName,Result) then
    Result := ADefault;
end;

function TXpgXMLAttributeList.AsXmlText2(AIndex: integer): AxUCString;
begin
  Result := FAttributes[AIndex].sAttr + '="' + FAttributes[AIndex].sVal + '"'
end;

function TXpgXMLAttributeList.AsXmlText: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FCount - 1 do
    Result := Result + AsXmlText2(i) + ' ';
  Result := Trim(Result);
end;

procedure TXpgXMLAttributeList.BeginSnatch;
begin
  FSnatching := True;
end;

procedure TXpgXMLAttributeList.Clear;
begin
  FCount := 0;
  FSorted := False;
end;

procedure TXpgXMLAttributeList.ClearNamespace;
var
  i,j: integer;
  S: AxUCString;
  P1,P2: PAnsiChar;
begin
  for i := 0 to FCount - 1 do begin
    S := FAttributes[i].sAttr;
    j := CPos(':',S);
    if j > 1 then
      FAttributes[i].sAttr := Copy(S,j + 1,MAXINT);

    P1 := FAttributes[i].Attribute.pStart;
    P2 := FAttributes[i].Attribute.pEnd;
    while (P1 < P2) and (P1^ <> ':') do
      Inc(P1);
    FAttributes[i].Attribute.pStart := P1 + 1;
  end;
end;

procedure TXpgXMLAttributeList.CollectPrefixed(APrefix: string; AList: TStrings);
var
  i: integer;
  S,S2: AxUCString;
begin
  for i := 0 to Count - 1 do begin
    S := FAttributes[i].sAttr;
    S2 := SplitAtChar(':',S);
    if Copy(S2,1,Length(APrefix)) = APrefix then begin
      AList.Add(FAttributes[i].sAttr + '=' + FAttributes[i].sVal);
      if FSnatching then
        FAttributes[i].Flag := 1;
    end;
  end;
end;


function TXpgXMLAttributeList.Count: integer;
begin
  Result := FCount;
end;

constructor TXpgXMLAttributeList.Create;
begin
  FSortList := TList.Create;
end;

destructor TXpgXMLAttributeList.Destroy;
begin
  FSortList.Free;
  inherited;
end;

procedure TXpgXMLAttributeList.EndSnatch(BadAttributes: TStrings);
var
  i: integer;
begin
  FSnatching := False;
  BadAttributes.Clear;
  for i := 0 to FCount - 1 do begin
    if FAttributes[i].Flag = 0 then
      BadAttributes.Add(FAttributes[i].sAttr);
  end;
end;

function TXpgXMLAttributeList.EnumError: AxUCString;
begin
  Result := FBadEnumValue;
  FBadEnumValue := '';
end;

function TXpgXMLAttributeList.Find(const AName: AxUCString; out AValue: AxUCString): boolean;
var
  i: integer;
begin
  i := Find(AName);
  Result := i >= 0;
  if Result then
    // TODO are spaces permitted at beginning/end of values?
    AValue := Trim(PXpgXmlAttribute(FSortList[i]).sVal);
end;

function TXpgXMLAttributeList.Find(const AName: AxUCString): integer;

function BinSearch: integer;
var
  L, H: Integer;
  mid, cmp: Integer;
begin
  Result := -1;
  L := 0;
  H := FCount - 1;
  while L <= H do
  begin
    mid := L + (H - L) shr 1;
    cmp := CompareStr(PXpgXmlAttribute(FSortList[mid]).sAttr,AName);
    if cmp < 0 then
      L := mid + 1
    else
    begin
      H := mid - 1;
      if cmp = 0 then begin
        Result := mid;
        Exit;
      end;
    end;
  end;
end;

begin
  if not FSorted then
    Sort;
  Result := BinSearch;
  if (Result >= 0) and FSnatching then
    PXpgXmlAttribute(FSortList[Result]).Flag := 1;
end;

function TXpgXMLAttributeList.GetAsString(const Str: TXpgXMLString): AxUCString;
var
  i,L: integer;
begin
  L := Integer(Str.pEnd) - Integer(Str.pStart) + 1;
  SetLength(Result,L);
  if L = 1 then
    Result[1] := AxUCChar(Str.pStart[0])
  else begin
    for i := 1 to L do
      Result[i] := AxUCChar(Str.pStart[i - 1]);
  end;
end;

function TXpgXMLAttributeList.GetAsStringUTF8(const Str: TXpgXMLString): AxUCString;
var
  L: integer;
  Hash: longword;
  P,pDest: PAnsiChar;
  pCheck: PAnsiChar;
begin
  // Entity conversion also in xpgPXML
  P := Str.pStart;
  pDest := P;
  while P <= Str.pEnd do begin
    if P^ = '&' then begin
      Hash := 0;
      pCheck := P;
      Inc(pCheck);
      while pCheck <= Str.pEnd do begin
        if pCheck^ = ';' then begin
          case Hash of
            318: begin // &amp; "
              pDest^ := '&';
              Inc(pDest);
              Inc(P,5);
            end;
            435: begin // &apos; '
              pDest^ := '''';
              Inc(pDest);
              Inc(P,6);
            end;
            219: begin // &gt; >
              pDest^ := '>';
              Inc(pDest);
              Inc(P,4);
           end;
            224: begin // &lt; <
              pDest^ := '<';
              Inc(pDest);
              Inc(P,4);
            end;
            457: begin // &quot; "
              pDest^ := '"';
              Inc(pDest);
              Inc(P,6);
            end;
            else begin
              P := pCheck;
              Inc(P);
            end;
          end;
          if P^ = '&' then begin
            Hash := 0;
            pCheck := P;
            Inc(pCheck);
            Continue;
          end
          else
            Break;
        end
        else begin
          Inc(Hash,Byte(pCheck^));
          Inc(pCheck);
        end;
      end;
    end;
    pDest^ := P^;
    Inc(P);
    Inc(pDest);
  end;
  L := Integer(Str.pEnd) - Integer(Str.pStart) + 1 - (Integer(P) - Integer(pDest));
//  L := Integer(Str.pEnd) - Integer(Str.pStart) + 1;
  SetLength(Result,L * 2);
{$ifdef BABOON}
  L := Utf8ToUnicode(PWideChar(Result),Length(Result),Str.pStart,L) - 1;
{$else}
  L := MultiByteToWideChar(CP_UTF8,0,Str.pStart,L,PWideChar(Result),L * 2);
{$endif}
  SetLength(Result,L);
end;

function TXpgXMLAttributeList.GetAttributes(Index: integer): AxUCString;
begin
  Result := FAttributes[Index].sAttr;
end;

function TXpgXMLAttributeList.GetHashA(Index: integer): longword;
var
  P: PAnsiChar;
  pEnd: PAnsiChar;
begin
  Result := 0;

  P := FAttributes[Index].Attribute.pStart;
  pEnd := FAttributes[Index].Attribute.pEnd;
  while P <= pEnd do begin
    Inc(Result,Ord(P^));
    Inc(P);
  end;
end;

function TXpgXMLAttributeList.GetHashB(Index: integer): longword;
var
  a : cardinal;
  P: PAnsiChar;
  pEnd: PAnsiChar;
begin
  Result := 0;

  a := 63689;

  P := FAttributes[Index].Attribute.pStart;
  pEnd := FAttributes[Index].Attribute.pEnd;
  while P <= pEnd do begin
    Result := Result * a + Ord(P^);
    a := a * 378551;
    Inc(P);
  end;
end;

function TXpgXMLAttributeList.GetHashC(Index: integer): longword;
var
  P: PAnsiChar;
  pEnd: PAnsiChar;
begin
  Result := 1315423911;

  P := FAttributes[Index].Attribute.pStart;
  pEnd := FAttributes[Index].Attribute.pEnd;
  while P <= pEnd do begin
    Result := Result xor ((Result shl 5) + Ord(P^) + (Result shr 2));
    Inc(P);
  end;
end;

function TXpgXMLAttributeList.GetHashD(Index: integer): longword;
var
  i: integer;
  P: PAnsiChar;
  pEnd: PAnsiChar;
begin
  Result := $AAAAAAAA;

  P := FAttributes[Index].Attribute.pStart;
  pEnd := FAttributes[Index].Attribute.pEnd;
  i := 1;
  while P <= pEnd do begin
    if ((i - 1) and 1) = 0 then
      Result := Result xor ((Result shl 7) xor Ord(P^) * (Result shr 3))
    else
      Result := Result xor (not((Result shl 11) + Ord(P^) xor (Result shr 5)));
    Inc(P);
    Inc(i);
  end;
end;

function TXpgXMLAttributeList.GetValues(Index: integer): AxUCString;
begin
  // TODO are spaces permitted at beginning/end of values?
  Result := FAttributes[Index].sVal;
end;

procedure TXpgXMLAttributeList.MakeStrings;
var
  i: integer;
begin
  for i := 0 to FCount - 1 do begin
    FAttributes[i].sAttr := GetAsString(FAttributes[i].Attribute);
    FAttributes[i].sVal := GetAsStringUTF8(FAttributes[i].Value);
  end;
end;

function ListSortCompare(Item1, Item2: Pointer): Integer;
begin
  Result := CompareStr(PXpgXmlAttribute(Item1).sAttr,PXpgXmlAttribute(Item2).sAttr);
end;

procedure TXpgXMLAttributeList.Sort;
var
  i: integer;
begin
  FSortList.Clear;
  for i := 0 to FCount - 1 do
    FSortList.Add(@FAttributes[i]);
  FSortList.Sort(ListSortCompare);
  FSorted := True;
end;

end.
