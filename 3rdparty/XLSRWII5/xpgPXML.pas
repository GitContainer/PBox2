unit xpgPXML;

interface

uses Classes, SysUtils, Contnrs,
{$ifndef BABOON}
     Windows,
{$endif}
     xpgPUtils, xpgPXMLUtils;

{$B-}
{$H+}
{$R-}
{$Q-}

// ********* TODO ********
// An attribute name can only occure once in AllocSz tag
// ********* TODO ********

// Can't be less than $7FFF
const AX_STREAM_READ_BUF_SZ     = $10000;
const AX_STREAM_READ_BUF_SZ_EX  = AX_STREAM_READ_BUF_SZ div 2;

const AX_STREAM_READ_MAX_ALLOC  = $FFFFF;

const AX_STREAM_WRITE_BUF_SZ    = $7FFF;
// Including space for converting to UTF8
const AX_STREAM_WRITE_BUF_SZ_EX = $FFFF;

const AX_MAXINDENT = 256;

type TAxBOM = (axbomNone,axbomUTF8,axbomUTF16BE,axbomUTF16LE,axbomUTF32BE,axbomUTF32LE);

type TXpgStrStack = class(TObject)
protected
     FValues  : TStringList;
public
     constructor Create;
     destructor Destroy; override;

     procedure Push(AValue: string);
     function  Pop: string;
     function  Count: integer;
     end;

type TXpgWriteXML = class(TObject)
private
     procedure SetAttributes(const Value: AxUCString);
{$ifdef AXW}
     procedure SetNamespace(const Value: AxUCString);
{$endif}
protected
     FStream      : TStream;
     FOwnsStream  : boolean;
     FAborted     : boolean;

     FText        : AxUCString;
     FAttributes  : TXpgXMLString;
     FStack       : TXpgStrStack;
     FSpaces      : Ax8String;

     FWriteIndent : boolean;

{$ifdef AXW}
     FNamespace   : AxUCString;
     FUseAttrNS   : boolean;
     FNSStack     : TStringList;
{$endif}

     procedure AddAttr(const AValue: AxUCString); {$ifdef DELPHI_7_OR_LATER} inline; {$endif}
     procedure CopyToBufUTF8(AValue: AxUCString); {$ifdef DELPHI_7_OR_LATER} inline; {$endif}
     procedure WriteString(const S: AxUCString); overload; {$ifdef DELPHI_7_OR_LATER} inline; {$endif}
     procedure WriteString(const S: PAnsiChar; const Len: integer); overload; {$ifdef DELPHI_7_OR_LATER} inline; {$endif}
{$ifdef AXW}
     function  CheckNS(const AName: AxUCString): AxUCString; inline;
{$endif}
public
     constructor Create;
     destructor Destroy; override;

     procedure SaveToStream(AStream: TStream);
     procedure SaveToFile(AFilename: AxUCString);

     procedure Declaration;
     procedure BeginTag(const AName: AxUCString); overload;
     procedure SimpleTag(const AName: AxUCString); overload;
     procedure SimpleTextTag(const AName,AText: AxUCString);
     procedure EmptyTag(const AName: AxUCString);
     procedure EndTag;
     procedure AddAttribute(const AName,AValue: AxUCString);

{$ifdef AXW}
     procedure PushNS(ANS: AxUCString);
     procedure PopNS;
{$endif}

     property Text: AxUCString read FText write FText;
     property Attributes: AxUCString write SetAttributes;

     property WriteIndent: boolean read FWriteIndent write FWriteIndent;
{$ifdef AXW}
     property Namespace  : AxUCString read FNamespace write SetNamespace;
     property UseAttrNS  : boolean read FUseAttrNS write FUseAttrNS;
{$endif}
     end;

type TXpgReadXML = class(TObject)
protected
     FStream      : TStream;
     FAborted     : boolean;
     FBOM         : TAxBOM;
     FNamespace   : TXpgXMLString;
     FHasNamespace: boolean;
     FTag         : TXpgXMLString;
     FElementVal  : TXpgXMLString;
     FAttributes  : TXpgXmlAttributeArray;
     FAttrCount   : integer;
     FText        : TXpgXMLString;
     FHasText     : boolean;
     FComment     : TXpgXMLString;
     FRow         : integer;
     FBuf         : array[0..$FF] of byte;
     // TODO
     FPreserveSpaces: boolean;

     FIncludeNamespace: boolean;

     // Preserve all chars (CR/LF etc.)
     FHtmlRawMode : boolean;

     FAttrList    : TXpgXMLAttributeList;

     procedure CheckBOM;
     function  GetAsString(const Str: TXpgXMLString): AxUCString; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  GetAsStringUTF8(const Str: TXpgXMLString): AxUCString;
     procedure MakeLowercase(const Str: TXpgXMLString);
public
     constructor Create;
     destructor Destroy; override;

     procedure Clear; virtual;
     procedure ClearNamespace;

     function  LoadFromStream(AStream: TStream): integer; virtual;
     procedure LoadFromFile(AFilename: AxUCString);

     function  LoadFromStreamHtml(AStream: TStream): integer; virtual;

     procedure Declaration; virtual;

     procedure BeginTag; virtual;
     procedure EndTag; virtual;
     procedure Comment; virtual;
     function  CmpTag(const S: AxUCString): boolean;
     function  Namespace:  AxUCString;
     function  Tag:  AxUCString;
     function  Text: AxUCString;

     function  QName: AxUCString;
     function  QNameP: PAnsiChar;
     function  QNameT: PAnsiChar; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  QNameIsC: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}
     function  QNameIsV: boolean; {$ifdef DELPHI_2006_OR_LATER} inline; {$endif}

     function  NameHashA(NS: AxUCString): longword;

     function  QNameHashA: longword;
     function  QNameHashB: longword;
     function  QNameHashC: longword;
     function  QNameHashD: longword;

     function FindAttribute(const Attr: AxUCString): AxUCString;

     function Col: integer;
     function Row: integer;

     procedure Abort;

     property Attributes: TXpgXMLAttributeList read FAttrList;

     property HasText: boolean read FHasText;
     property PreserveSpaces: boolean read FPreserveSpaces write FPreserveSpaces;
     property IncludeNamespace: boolean read FIncludeNamespace write FIncludeNamespace;
     end;

implementation

function  DecodeNumEncoding(S: AxUCString): AxUCString;
var
  p1,p2: integer;
  n    : integer;
begin
  while True do begin
    p1 := Pos('&#',S);
    p2 := CPos(';',S);

    if (p2 > p1) and (p1 > 0) then begin
      n := StrToIntDef(Copy(S,p1 + 2,p2 - p1 - 2),Ord('_'));

      S[p1] := AxUCChar(n);

      System.Delete(S,p1 + 1,p2 - p1);
    end
    else
      Break;
  end;

  Result := S;
end;

// uses Windows;

function CPos(C: char; S: string): integer;
begin
  for Result := 1 to Length(S) do begin
    if S[Result] = C then
      Exit;
  end;
  Result := -1;
end;

function SplitAtChar(C: char; var S: string): string;
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

const
  AxSpaceMap: array[0..255] of byte = (
  0,0,0,0,0,0,0,0,0,1,1,0,0,1,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
  0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);

{
function IsXMLChar(C: Ax8Char): boolean; inline;
begin
  Result := C in ['a'..'z','A'..'Z','0'..'9','_'];
end;

function IsXMLFirstChar(C: Ax8Char): boolean; inline;
begin
  Result := C in ['a'..'z','A'..'Z'];
end;
}

{ AxXMLRead }

procedure TXpgReadXML.Abort;
begin
  FAborted := True;
end;

procedure TXpgReadXML.BeginTag;
begin

end;

// Can't use as zip files don't permitts seek backwards.
procedure TXpgReadXML.CheckBOM;
//var
//  B1,B2,B3,B4: byte;
begin
//  FStream.Read(B1,1);
//  FStream.Read(B2,1);
//  if (B1 = $FE) and (B2 = $FF) then
//    FBOM := axbomUTF16BE
//  else if (B1 = $FF) and (B2 = $FE) then
//    FBOM := axbomUTF16BE
//  else begin
//    FStream.Read(B3,1);
//    if (B1 = $EF) and (B2 = $BB) and (B3 = $BF) then
//      FBOM := axbomUTF8
//    else begin
//      FStream.Read(B4,1);
//      if (B1 = $00) and (B2 = $00) and (B3 = $FE) and (B4 = $FF) then
//        FBOM := axbomUTF32BE
//      else if (B1 = $FF) and (B2 = $FE) and (B3 = $00) and (B4 = $00) then
//        FBOM := axbomUTF32LE
//      else begin
//        FBOM := axbomNone;
//        FStream.Seek(0,soFromBeginning);
//      end;
//    end;
//  end;
end;

procedure TXpgReadXML.Clear;
begin
  FAborted := False;
end;

procedure TXpgReadXML.ClearNamespace;
begin
  FNamespace.pEnd := FNamespace.pStart - 1;
end;

function TXpgReadXML.CmpTag(const S: AxUCString): boolean;
var
  i: integer;
  p: PAnsiChar;
begin
  Result := Length(S) = (FTag.pEnd - FTag.pStart + 1);
  if not Result then
    Exit;

  p := FTag.pStart;
  for i := 1 to Length(S) do begin
    if Ax8Char(S[i]) <> p^ then begin
      Result := False;
      Exit;
    end;
    Inc(p);
  end;
  Result := True;
end;

function TXpgReadXML.Col: integer;
begin
  Result := 0;
end;

procedure TXpgReadXML.Comment;
begin

end;

constructor TXpgReadXML.Create;
begin
  SetLength(FAttributes,AX_XML_MAX_ATTRIBUTES);
  FAttrList := TXpgXMLAttributeList.Create;
end;

procedure TXpgReadXML.EndTag;
begin

end;

function TXpgReadXML.FindAttribute(const Attr: AxUCString): AxUCString;
var
  i: integer;
begin
  for i := 0 to FAttrCount - 1 do begin
    if GetAsString(FAttributes[i].Attribute) = Attr then begin
      Result := GetAsString(FAttributes[i].Value);
      Exit;
    end;
  end;
end;

function TXpgReadXML.QName: AxUCString;
begin
  Result := GetAsString(FNamespace);
  if Result <> '' then
    Result := Result + ':' + GetAsString(FTag)
  else
    Result := GetAsString(FTag);
end;

function TXpgReadXML.QNameHashA: longword;
var
  P: PAnsiChar;
begin
  Result := 0;

  if FNamespace.pEnd >= FNamespace.pStart then begin
    P := FNamespace.pStart;
    while P <= FNamespace.pEnd do begin
      Inc(Result,Ord(P^));
      Inc(P);
    end;
    Inc(Result,Ord(':'));
  end;

  P := FTag.pStart;
  while P <= FTag.pEnd do begin
    Inc(Result,Ord(P^));
    Inc(P);
  end;
end;

function TXpgReadXML.QNameHashB: longword;
var
  AllocSz : longword;
  P: PAnsiChar;
begin
  Result := 0;

  AllocSz := 63689;
  if FNamespace.pEnd >= FNamespace.pStart then begin
    P := FNamespace.pStart;
    while P <= FNamespace.pEnd do begin
      Result := Result * AllocSz + Ord(P^);
      AllocSz := AllocSz * 378551;
      Inc(P);
    end;
    Result := Result * AllocSz + Ord(':');
    AllocSz := AllocSz * 378551;
  end;

  P := FTag.pStart;
  while P <= FTag.pEnd do begin
    Result := Result * AllocSz + Ord(P^);
    AllocSz := AllocSz * 378551;
    Inc(P);
  end;
end;

function TXpgReadXML.QNameHashC: longword;
var
  P: PAnsiChar;
begin
  Result := 1315423911;

  if FNamespace.pEnd >= FNamespace.pStart then begin
    P := FNamespace.pStart;
    while P <= FNamespace.pEnd do begin
      Result := Result xor ((Result shl 5) + Ord(P^) + (Result shr 2));
      Inc(P);
    end;
    Result := Result xor ((Result shl 5) + Ord(':') + (Result shr 2));
  end;

  P := FTag.pStart;
  while P <= FTag.pEnd do begin
    Result := Result xor ((Result shl 5) + Ord(P^) + (Result shr 2));
    Inc(P);
  end;
end;

function TXpgReadXML.QNameHashD: longword;
var
  i: integer;
  P: PAnsiChar;
begin
  Result := $AAAAAAAA;

  i := 1;
  if FNamespace.pEnd >= FNamespace.pStart then begin
    P := FNamespace.pStart;
    while P <= FNamespace.pEnd do begin
      if ((i - 1) and 1) = 0 then
        Result := Result xor ((Result shl 7) xor Ord(P^) * (Result shr 3))
      else
        Result := Result xor (not((Result shl 11) + Ord(P^) xor (Result shr 5)));
      Inc(P);
      Inc(i);
    end;
    if ((i - 1) and 1) = 0 then
      Result := Result xor ((Result shl 7) xor Ord(':') * (Result shr 3))
    else
      Result := Result xor (not((Result shl 11) + Ord(':') xor (Result shr 5)));
    Inc(i);
  end;

  P := @FTag.pStart;
  while P <= FTag.pEnd do begin
    if ((i - 1) and 1) = 0 then
      Result := Result xor ((Result shl 7) xor Ord(P^) * (Result shr 3))
    else
      Result := Result xor (not((Result shl 11) + Ord(P^) xor (Result shr 5)));
    Inc(P);
    Inc(i);
  end;
end;

function TXpgReadXML.QNameIsC: boolean;
begin
  Result := FTag.pStart[0] = 'c';
end;

function TXpgReadXML.QNameIsV: boolean;
begin
  Result := FTag.pStart[0] = 'v';
end;

function TXpgReadXML.QNameP: PAnsiChar;
var
  L,L2: integer;
begin
  Result := @FBuf;
  L := Integer(FNamespace.pEnd) - Integer(FNamespace.pStart) + 1;
  if L > 0 then begin
    Move(FNamespace.pStart[0],FBuf[0],L);
    FBuf[L] := Byte(':');
    L2 := Integer(FTag.pEnd) - Integer(FTag.pStart) + 1;
    Move(FTag.pStart[0],FBuf[L + 1],L2);
    FBuf[L + 1 + L2] := 0;
  end
  else begin
    L := Integer(FTag.pEnd) - Integer(FTag.pStart) + 1;
    Move(FTag.pStart[0],FBuf[0],L);
    FBuf[L] := 0;
  end;
end;

function TXpgReadXML.QNameT: PAnsiChar;
begin
  Result := @FBuf;
  FBuf[0] := Byte(FTag.pStart[0]);
  FBuf[1] := 0;
end;

function TXpgReadXML.GetAsString(const Str: TXpgXMLString): AxUCString;
var
  i,L: integer;
begin
  L := Integer(Str.pEnd) - Integer(Str.pStart) + 1;
  SetLength(Result,L);
  for i := 1 to L do
    Result[i] := AxUCChar(Str.pStart[i - 1]);
end;

function TXpgReadXML.GetAsStringUTF8(const Str: TXpgXMLString): AxUCString;
var
  L      : integer;
  Hash   : longword;
  P,pDest: PAnsiChar;
  pCheck : PAnsiChar;
  DoNum  : boolean;
begin
  // Entity conversion also in xpgPXMLUtils->GetAsStringUTF8
  DoNum := False;
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
            132: begin // &#10;
              pDest^ := #10;;
              Inc(pDest);
              Inc(P,5);
            end;
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
        else if pCheck^ = '#' then begin
          Inc(pCheck);
          while (pCheck^ <> ';') and (pCheck <= Str.pEnd) do
            Inc(pCheck);

          DoNum := True;

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

  SetLength(Result,L * 2);
{$ifdef BABOON}
  L := Utf8ToUnicode(PWideChar(Result),Length(Result),Str.pStart,L) - 1;
{$else}
  L := MultiByteToWideChar(CP_UTF8,0,Str.pStart,L,PWideChar(Result),L * 2);
{$endif}
  SetLength(Result,L);

  if DoNum then
    Result := DecodeNumEncoding(Result);
end;

function TXpgReadXML.NameHashA(NS: AxUCString): longword;
var
  i: integer;
  P: PAnsiChar;
begin
  Result := 0;

  for i := 1 to Length(NS) do
    Inc(Result,Ord(NS[i]));

  P := FTag.pStart;
  while P <= FTag.pEnd do begin
    Inc(Result,Ord(P^));
    Inc(P);
  end;
end;

function TXpgReadXML.Namespace: AxUCString;
begin
  Result := GetAsString(FNamespace);
end;

function TXpgReadXML.Row: integer;
begin
  Result := FRow;
end;

procedure TXpgReadXML.Declaration;
var
  Val: AxUCString;
begin
  if CmpTag('xml') then begin
    Val := FindAttribute('encoding');
    if (Val <> '') and (Lowercase(Val) <> 'utf-8') then
      raise Exception.Create('Unsupported encoding of xml: "' + Val + '". Must be utf-8');
  end;
end;

destructor TXpgReadXML.Destroy;
begin
  FAttrList.Free;
  inherited;
end;

procedure TXpgReadXML.LoadFromFile(AFilename: AxUCString);
var
  Stream: TFileStream;
//  Stream: TMappedFile;
begin
//  Stream := TMappedFile.Create(AFilename);
  Stream := TFileStream.Create(AFilename,fmOpenRead);
  try
    LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

function TXpgReadXML.LoadFromStream(AStream: TStream): integer;
var
  Sz   : integer;
  Ok   : boolean;
  Buf  : PByteArray;
  BufSz: integer;
  AllocSz: integer;
  ppP1,ppP2,ppLT,ppGT,ppEnd,ppMax,ppColon,ppSlash1,ppSlash2,ppEq,ppQuestion,ppQuote,ppText: PAnsiChar;
begin
  FStream := AStream;

//  CheckBOM;

  Result := 0;
  FRow := 0;
//  FStream.Seek(0,0);
  Sz := FStream.Size;
//  GetMem(Buf,Sz);
  AllocSz := AX_STREAM_READ_BUF_SZ;
  GetMem(Buf,AllocSz);
  try
    while not FAborted do begin
      ppLT := @Buf[0];
      ppMax := ppLT + AllocSz;
//      BufSz := FStream.Read(Buf[0],AX_STREAM_READ_BUF_SZ_EX);
      BufSz := FStream.Read(Buf[0],AllocSz div 2);
      if BufSz < 3 then begin
        Exit;
      end;

      Inc(Result,BufSz);

      ppEnd := ppLT + BufSz - 1;

      if Result <> Sz then begin
        OK := False;
        while not Ok do begin
          Ok := True;
          while ppEnd < ppMax do begin
            if (ppEnd^ = '<') and (ppEnd <= ppMax) then begin
              Inc(ppEnd);
              if not FStream.Read(ppEnd^,1) = 1 then
                raise Exception.Create('Unexpected end of file in XML');
              Inc(Result);
              if (ppEnd^ = '/') and (ppEnd <= ppMax) then begin
                while (ppEnd^ <> '>') and (ppEnd <= ppMax) do begin
                  Inc(ppEnd);
                  if not FStream.Read(ppEnd^,1) = 1 then
                    raise Exception.Create('Unexpected end of file in XML');
                  Inc(Result);
                end;
                Break;
              end;
            end
            else if (ppEnd^ = '/') and (ppEnd <= ppMax) then begin
              Inc(ppEnd);
              if not FStream.Read(ppEnd^,1) = 1 then
                raise Exception.Create('Unexpected end of file in XML');
              Inc(Result);

              if ppEnd^ = '>' then
                Break;
            end;

            Inc(ppEnd);
            if not FStream.Read(ppEnd^,1) = 1 then
              raise Exception.Create('Unexpected end of file in XML');
            Inc(Result);
          end;

          if ppEnd^ <> '>' then begin
            Ok := False;
            if AllocSz < AX_STREAM_READ_MAX_ALLOC then begin
              Inc(AllocSz,AX_STREAM_READ_BUF_SZ);

              ReAllocMem(Buf,AllocSz);

              FStream.Seek(0,soFromBeginning);
              BufSz := FStream.Read(Buf[0],AllocSz div 2);

              ppLT := @Buf[0];
              ppMax := ppLT + AllocSz;

              Result := 0;
              ppEnd := ppLT + BufSz - 1;
            end
            else
              raise Exception.Create('Unexpected end of file in XML');
          end;
        end;
      end;

      ppGT := ppLT + 1;
      while (ppGT < ppEnd) and not FAborted do begin
        while (ppLT^ <> '<') and (ppLT < ppEnd) do begin
          Inc(ppLT);
        end;

        if ppLT^ <> '<' then
          raise Exception.Create('Syntax error in XML: can not find "<"');

        if (PLongword(ppLT)^ and $00FF00FF) = $003E003C then begin // <x>
          Inc(ppLT);
          FTag.pStart := ppLT;
          FTag.pEnd := ppLT;
          Inc(ppLT,2);
          if (ppLT < ppEnd) and (ppLT^ <> '<') then begin
            FText.pStart := ppLT;
                  // ZZZ 120104
            while (ppLT <= ppEnd) and (ppLT^ <> '<') do begin
              if ppLT^ = #10 then
                Inc(FRow);
              Inc(ppLT);
            end;
            FText.pEnd := ppLT - 1;
            FHasText := True;
          end
          else
            FHasText := False;
          ppGT := ppLT + 1;
          FAttrList.Add(FAttributes,FAttrCount);
          FAttrCount := 0;
          BeginTag;
          Continue;
        end
        else if (PLongword(ppLT)^ and $FF00FFFF) = $3E002F3C then begin // </x>
          Inc(ppLT,2);
          FTag.pStart := ppLT;
          FTag.pEnd := ppLT;
          Inc(ppLT,2);
          ppGT := ppLT + 1;
          EndTag;
          Continue;
        end;

        ppColon := Nil;
        ppSlash1 := Nil;
        ppSlash2 := Nil;
        ppEq := Nil;
        ppQuestion := Nil;
        ppText := Nil;
        FAttrCount := 0;

        ppGT := ppLT + 1;
        if ppGT^  = '/' then begin
          ppSlash1 := ppGT;
          Inc(ppGT);
        end
        else if ppGT^ = '!' then begin
          ppP1 := Nil;
          Inc(ppGT,2);
          if ppGT^ = '-' then
            ppP1 := ppGT + 1;
          while ppGT < ppEnd do begin
            if ppGT^ = #10 then
              Inc(FRow)
            else if ppGt^ = '-' then begin
              Inc(ppGT);
              if ppGt^ = '-' then begin
                Inc(ppGT);
                if ppGT^ = '>' then
                  Break
                // <!---->
                else if ppGT^ = '-' then begin
                  Inc(ppGT);
                  if ppGT^ = '>' then
                    Break
                end;
              end;
            end;
            Inc(ppGT);
          end;
{
          while (ppGT^ <> '>') and (ppGT < ppEnd) do begin
            if ppGT^ = #10 then
              Inc(FRow);
            Inc (ppGT);
          end;
}
          if ppP1 <> Nil then begin
            FComment.pStart := ppP1;
            FComment.pEnd := ppGT - 3;
            Comment;
          end;

          while (ppGT^ <> '<') and (ppGT < ppEnd) do begin
            if ppGT^ = #10 then
              Inc(FRow);
            Inc (ppGT);
          end;

          ppLT := ppGT;
          ppGT := ppLT + 1;
          Continue;
        end;

  //      while (ppGT^ in [#09,#10,#13,' ']) and (ppGT < ppEnd) do
        while (AxSpaceMap[Byte(ppGT^)] = 1) and (ppGT < ppEnd) do begin
          if ppGT^ = #10 then
            Inc(FRow);
          Inc(ppGT);
        end;
        FTag.pStart := ppGT;
  //      while not (ppGT^ in [#09,#10,#13,' ']) and (ppGT < ppEnd) do begin
        while not (AxSpaceMap[Byte(ppGT^)] = 1) and (ppGT < ppEnd) do begin
          case ppGT^ of
            ':': ppColon := ppGT;
            '/': ppSlash2 := ppGT;
            '>': Break;
          end;
          Inc(ppGT);
        end;
        FTag.pEnd := ppGT - 1;
        if FTag.pEnd^ = '/' then
          Dec(FTag.pEnd);

  //      while (ppGT^ in [#09,#10,#13,' ']) and (ppGT < ppEnd) do
        while (AxSpaceMap[Byte(ppGT^)] = 1) and (ppGT < ppEnd) do begin
          if ppGT^ = #10 then
            Inc(FRow);
          Inc(ppGT);
        end;
        ppP1 := ppGT;

        while ppGT <= ppEnd do begin
          case ppGT^ of
            #10: begin
              ppP1 := ppGT + 1;
              Inc(FRow);
            end;
            #09,#13,' ': ppP1 := ppGT + 1;
            '>': begin
              ppP2 := ppGT + 1;
////              while (ppP2^ in [#10,#13]) and (ppP2 <= ppEnd) do begin
//              while (AxSpaceMap[Byte(ppP2^)] = 1) and (ppP2 <= ppEnd) do begin
//                if ppP2^ = #10 then
//                  Inc(FRow);
//                Inc(ppP2);
//              end;
              if (ppP2^ <> '<') and (ppP2 <= ppEnd) then begin
                while (ppP2^ <> '<') and (ppP2 <= ppEnd) do begin
                  if ppP2^ = #10 then
                    Inc(FRow);
                  Inc(ppP2);
                end;
                ppText := ppGT + 1;
                ppGT := ppP2 - 1;
                // Comments may have CR/LF
//                while (ppText < ppEnd) and (ppText^ in [#10,#13]) do
//                  Inc(ppText);
                FHasText := ppText < ppP2;
              end
              else
                FHasText := False;

              FHasNamespace := ppColon <> Nil;
              if FHasNamespace then begin
                FNamespace.pStart := FTag.pStart;
                FNamespace.pEnd := ppColon - 1;
                FTag.pStart := ppColon + 1;
              end
              else begin
                FNamespace.pStart := FTag.pStart;
                FNamespace.pEnd := FTag.pStart - 1;
              end;
              if ppQuestion <> Nil then begin
                FTag.pStart := FTag.pStart + 1;
                Declaration;
                ppLT := ppGT + 1;
                ppGT := ppLT + 1;
              end
              else begin
                if ppSlash1 <> Nil then
                  EndTag
                else begin
                  if ppText <> Nil then begin
                    FText.pStart := ppText;
                    FText.pEnd := ppGT;
                  end;
                  FAttrList.Add(FAttributes,FAttrCount);
                  FAttrCount := 0;
                  BeginTag;
                  if ppSlash2 <> Nil then
                    EndTag;
                end;
                if ppGT >= ppEnd then
                  Break;
                ppLT := ppGT + 1;
                ppGT := ppLT + 1;
              end;
              Break;
            end;
            '''',
            '"': begin  // TODO Scan entire quoted string, as they may contain characters like slash "http://www.w3.org/2001/XMLSchema"
              if ppGT^ = '''' then begin
                ppQuote := ppGT;
                Inc(ppGT);
                // TODO Check for closing ">"
                while (ppGT^ <> '''') and (ppGT <= ppEnd) do begin
                  if ppGT^ = #10 then
                    Inc(FRow);
                  Inc(ppGT);
                end;
              end
              else begin
                ppQuote := ppGT;
                Inc(ppGT);
                // TODO Check for closing ">"
                while (ppGT^ <> '"') and (ppGT <= ppEnd) do begin
                  if ppGT^ = #10 then
                    Inc(FRow);
                  Inc(ppGT);
                end;
              end;
              if (ppGT^ in ['"','''']) and (ppEq <> Nil) then begin
                // Not checked: spaces between end of attribute name and equal.
                FAttributes[FAttrCount].Attribute.pStart := ppP1;
                FAttributes[FAttrCount].Attribute.pEnd := ppEq - 1;
                FAttributes[FAttrCount].Value.pStart := ppQuote + 1;
                FAttributes[FAttrCount].Value.pEnd := ppGT - 1;
                Inc(FAttrCount);
                if FAttrCount >= Length(FAttributes) then
                  SetLength(FAttributes,Length(FAttributes) + AX_XML_MAX_ATTRIBUTES);
                ppEq := Nil;
              end;
            end;
            '=': begin
              ppEq := ppGT;
            end;
            '?': ppQuestion := ppGT;
            '/': ppSlash2 := ppGT;
          end;
          Inc(ppGT);
        end
      end;
//      Exit;
    end;
  finally
    FreeMem(Buf);
  end;
end;

function TXpgReadXML.LoadFromStreamHtml(AStream: TStream): integer;
var
  Sz: integer;
  Buf: PByteArray;
  BufSz: integer;
  ppP1,ppP2,ppLT,ppGT,ppEnd,ppMax,ppColon,ppSlash1,ppSlash2,ppEq,ppQuestion,ppText: PAnsiChar;
begin
  FStream := AStream;

//  CheckBOM;

  Result := 0;
  FRow := 0;
//  FStream.Seek(0,0);
  Sz := FStream.Size;
//  GetMem(Buf,Sz);
  GetMem(Buf,AX_STREAM_READ_BUF_SZ);
  try
    while not FAborted do begin
      ppLT := @Buf[0];
      ppMax := ppLT + AX_STREAM_READ_BUF_SZ;
      BufSz := FStream.Read(Buf[0],AX_STREAM_READ_BUF_SZ - AX_STREAM_READ_BUF_SZ_EX);
      if BufSz < 3 then begin
        Exit;
      end;

      Inc(Result,BufSz);

      ppEnd := ppLT + BufSz - 1;

      if Result <> Sz then begin
        while ppEnd < ppMax do begin
          if (ppEnd^ = '<') and (ppEnd <= ppMax) then begin
            Inc(ppEnd);
            if not FStream.Read(ppEnd^,1) = 1 then
              raise Exception.Create('Unexpected end of file in XML');
            Inc(Result);
            if (ppEnd^ = '/') and (ppEnd <= ppMax) then begin
              while (ppEnd^ <> '>') and (ppEnd <= ppMax) do begin
                Inc(ppEnd);
                if not FStream.Read(ppEnd^,1) = 1 then
                  raise Exception.Create('Unexpected end of file in XML');
                Inc(Result);
              end;
              Break;
            end;
          end
          else if (ppEnd^ = '/') and (ppEnd <= ppMax) then begin
            Inc(ppEnd);
            if not FStream.Read(ppEnd^,1) = 1 then
              raise Exception.Create('Unexpected end of file in XML');
            Inc(Result);

            if ppEnd^ = '>' then
              Break;
          end;

          Inc(ppEnd);
          if not FStream.Read(ppEnd^,1) = 1 then
            raise Exception.Create('Unexpected end of file in XML');
          Inc(Result);
        end;

        if ppEnd^ <> '>' then
          raise Exception.Create('Unexpected end of file in XML');
      end;

      ppGT := ppLT + 1;
      while (ppGT < ppEnd) and not FAborted do begin
        while (ppLT^ <> '<') and (ppLT < ppEnd) do begin
          Inc(ppLT);
        end;

        if ppLT^ <> '<' then
          raise Exception.Create('Syntax error in XML: can not find "<"');

        ppColon := Nil;
        ppSlash1 := Nil;
        ppSlash2 := Nil;
        ppQuestion := Nil;
        ppText := Nil;
        FAttrCount := 0;

        ppGT := ppLT + 1;
        if ppGT^  = '/' then begin
          ppSlash1 := ppGT;
          Inc(ppGT);
        end
        else if ppGT^ = '!' then begin
          ppP1 := Nil;
          Inc(ppGT,2);
          if ppGT^ = '-' then
            ppP1 := ppGT + 1;
          while ppGT < ppEnd do begin
            if ppGT^ = #10 then
              Inc(FRow)
            else if ppGt^ = '-' then begin
              Inc(ppGT);
              if ppGt^ = '-' then begin
                Inc(ppGT);
                if ppGT^ = '>' then
                  Break
                // <!---->
                else if ppGT^ = '-' then begin
                  Inc(ppGT);
                  if ppGT^ = '>' then
                    Break
                end;
              end;
            end
            else if ppGT^ = '>' then
              Break;
            Inc(ppGT);
          end;
          if ppP1 <> Nil then begin
            FComment.pStart := ppP1;
            FComment.pEnd := ppGT - 3;
            Comment;
          end;

          while (ppGT^ <> '<') and (ppGT < ppEnd) do begin
            if ppGT^ = #10 then
              Inc(FRow);
            Inc (ppGT);
          end;

          ppLT := ppGT;
          ppGT := ppLT + 1;
          Continue;
        end;

        while (AxSpaceMap[Byte(ppGT^)] = 1) and (ppGT < ppEnd) do begin
          if ppGT^ = #10 then
            Inc(FRow);
          Inc(ppGT);
        end;
        FTag.pStart := ppGT;
        while not (AxSpaceMap[Byte(ppGT^)] = 1) and (ppGT < ppEnd) do begin
          case ppGT^ of
            ':': ppColon := ppGT;
            '/': ppSlash2 := ppGT;
            '>': Break;
          end;
          Inc(ppGT);
        end;
        FTag.pEnd := ppGT - 1;
        if FTag.pEnd^ = '/' then
          Dec(FTag.pEnd);

        while ppGT <= ppEnd do begin
          case ppGT^ of
            '>': begin
              ppP2 := ppGT + 1;
              if (ppP2^ <> '<') and (ppP2 <= ppEnd) then begin
                while (ppP2^ <> '<') and (ppP2 <= ppEnd) do begin
                  if ppP2^ = #10 then
                    Inc(FRow);
                  Inc(ppP2);
                end;
                ppText := ppGT + 1;
                ppGT := ppP2 - 1;
                if not FHtmlRawMode then begin
                  while (ppText^ in [#10,#13]) and (ppText <= ppEnd) do
                    Inc(ppText);
                end;
                FHasText := ppText <= ppGT;
              end
              else
                FHasText := False;

              FHasNamespace := ppColon <> Nil;
              if FHasNamespace then begin
                FNamespace.pStart := FTag.pStart;
                FNamespace.pEnd := ppColon - 1;
                FTag.pStart := ppColon + 1;
              end
              else begin
                FNamespace.pStart := FTag.pStart;
                FNamespace.pEnd := FTag.pStart - 1;
              end;
              if ppQuestion <> Nil then begin
                FTag.pStart := FTag.pStart + 1;
                Declaration;
                ppLT := ppGT + 1;
                ppGT := ppLT + 1;
              end
              else begin
                if ppText <> Nil then begin
                  FText.pStart := ppText;
                  FText.pEnd := ppGT;
                end;
                if ppSlash1 <> Nil then
                  EndTag
                else begin
                  FAttrList.Add(FAttributes,FAttrCount);
                  FAttrCount := 0;
                  BeginTag;
                  if ppSlash2 <> Nil then
                    EndTag;
                end;
                if ppGT >= ppEnd then
                  Break;
                ppLT := ppGT + 1;
                ppGT := ppLT + 1;
              end;
              Break;
            end;
            #10,
            #13,
            ' ': begin
              while (ppGT^ <> '>') and (ppGT <= ppEnd) do begin
                while (ppGT^ in [' ',#10,#13]) and (ppGT <= ppEnd) do
                  Inc(ppGT);
                ppP1 := ppGT;

                while not (ppGT^ in [' ','=','>']) and (ppGT <= ppEnd) do
                  Inc(ppGT);
                ppP2 := ppGT - 1;

                while not (ppGT^ in ['=','>']) and (ppGT <= ppEnd) do
                  Inc(ppGT);

                if ppGT^ = '=' then begin
                  Inc(ppGT);
                  while (ppGT^ in [#10,#13,' ']) and (ppGT <= ppEnd) do
                    Inc(ppGT);
                  ppEQ := ppGT;
                  if ppGT^ = '''' then begin
                    Inc(ppEQ);
                    Inc(ppGT);
                    while (ppGT^ <> '''') and (ppGT <= ppEnd) do
                      Inc(ppGT);
                  end
                  else if ppGT^ = '"' then begin
                    Inc(ppEQ);
                    Inc(ppGT);
                    while (ppGT^ <> '"') and (ppGT <= ppEnd) do
                      Inc(ppGT);
                  end
                  else begin
                    Inc(ppGT);
                    while not (ppGT^ in [#10,#13,' ','>']) and (ppGT <= ppEnd) do
                      Inc(ppGT);
                  end;
                  if (ppGT^ in [#10,#13,' ','"','''','>']) and (ppEq <> Nil) then begin
                    FAttributes[FAttrCount].Attribute.pStart := ppP1;
                    FAttributes[FAttrCount].Attribute.pEnd := ppP2;
                    FAttributes[FAttrCount].Value.pStart := ppEq;
                    FAttributes[FAttrCount].Value.pEnd := ppGT - 1;
                    Inc(FAttrCount);
                    if FAttrCount >= Length(FAttributes) then
                      SetLength(FAttributes,Length(FAttributes) + AX_XML_MAX_ATTRIBUTES);
                  end;
                end;
                if ppGT^ <> '>' then
                  Inc(ppGT);
              end;
              Dec(ppGT);
            end;
            '?': ppQuestion := ppGT;
            '/': ppSlash2 := ppGT;
          end;
          Inc(ppGT);
        end
      end;
//      Exit;
    end;
  finally
    FreeMem(Buf);
  end;
end;

procedure TXpgReadXML.MakeLowercase(const Str: TXpgXMLString);
var
  i,L: integer;
begin
  L := Integer(Str.pEnd) - Integer(Str.pStart) + 1;
  for i := 0 to L do begin
    if Str.pStart[i] in ['A'..'Z'] then
      Str.pStart[i] := AnsiChar(Ord('a') + (Byte(Str.pStart[i]) - Ord('A')));
  end;
end;

function TXpgReadXML.Tag: AxUCString;
begin
  Result := GetAsString(FTag);
end;

function TXpgReadXML.Text: AxUCString;
begin
  if FHasText then
    Result := GetAsStringUTF8(FText)
  else
    Result := '';
end;

{ TXpgWriteXML }

procedure TXpgWriteXML.BeginTag(const AName: AxUCString);
begin
  if FAttributes.pEnd = FAttributes.pStart then
{$ifdef AXW}
    WriteString('<' + CheckNS(AName) + '>' + #13)
{$else}
    WriteString('<' + AName + '>' + #13)
{$endif}
  else begin
{$ifdef AXW}
    WriteString('<' + CheckNS(AName) + ' ');
{$else}
    WriteString('<' + AName + ' ');
{$endif}
    WriteString(FAttributes.pStart,Integer(FAttributes.pEnd) - Integer(FAttributes.pStart));
    WriteString('>' + #13);

    FAttributes.pEnd := FAttributes.pStart;
  end;

  FStack.Push(AName);
end;

procedure TXpgWriteXML.AddAttr(const AValue: AxUCString);
var
  i: integer;
  L,L2: integer;
begin
  L := Length(AValue);
  L2 := (Integer(FAttributes.pEnd) - Integer(FAttributes.pStart) + 1);
  if L2 > AX_STREAM_WRITE_BUF_SZ then
    raise Exception.Create('Write XML buffer owerflow');

  for i := 1 to L do begin
    FAttributes.pEnd^ := AnsiChar(Byte(AValue[i]) and $00FF);
    Inc(FAttributes.pEnd);
  end;
end;

procedure TXpgWriteXML.AddAttribute(const AName, AValue: AxUCString);
begin
{$ifdef AXW}
  if FUseAttrNS then
    AddAttr(CheckNS(AName))
  else
    AddAttr(AName);
{$else}
  AddAttr(AName);
{$endif}
  AddAttr('="');
  CopyToBufUTF8(AValue);
  AddAttr('" ');
end;

{$ifdef AXW}
function TXpgWriteXML.CheckNS(const AName: AxUCString): AxUCString;
begin
  if FNamespace <> '' then
    Result := FNamespace + ':' + AName
  else
    Result := AName;
end;
{$endif}

procedure TXpgWriteXML.CopyToBufUTF8(AValue: AxUCString);
var
  i,j,L,Sz: integer;
begin
  // The purpose with this is to avoid reallocating the string for each entity as this is expensive.
  Sz := 0;
  for i := 1 to Length(AValue) do begin
    case AValue[i] of
      '&' : Inc(Sz,4); // &amp; &
      '''': Inc(Sz,5); // &apos; '
      '>' : Inc(Sz,3); // &gt; >
      '<' : Inc(Sz,3); // &lt; <
      '"' : Inc(Sz,5); // &quot; "
    end;
  end;

  if Sz > 0 then begin
    L := Length(AValue);
    SetLength(AValue,Length(AValue) + Sz);
    j := Length(AValue);

    for i := L downto 1 do begin
      case AValue[i] of
        '&' : begin
          AValue[j] := ';';
          Dec(j);
          AValue[j] := 'p';
          Dec(j);
          AValue[j] := 'm';
          Dec(j);
          AValue[j] := 'a';
          Dec(j);
          AValue[j] := '&';
          Dec(j);
        end;
        '''': begin
          AValue[j] := ';';
          Dec(j);
          AValue[j] := 's';
          Dec(j);
          AValue[j] := 'o';
          Dec(j);
          AValue[j] := 'p';
          Dec(j);
          AValue[j] := 'a';
          Dec(j);
          AValue[j] := '&';
          Dec(j);
        end;
        '>' : begin
          AValue[j] := ';';
          Dec(j);
          AValue[j] := 't';
          Dec(j);
          AValue[j] := 'g';
          Dec(j);
          AValue[j] := '&';
          Dec(j);
        end;
        '<' : begin
          AValue[j] := ';';
          Dec(j);
          AValue[j] := 't';
          Dec(j);
          AValue[j] := 'l';
          Dec(j);
          AValue[j] := '&';
          Dec(j);
        end;
        '"' : begin
          AValue[j] := ';';
          Dec(j);
          AValue[j] := 't';
          Dec(j);
          AValue[j] := 'o';
          Dec(j);
          AValue[j] := 'u';
          Dec(j);
          AValue[j] := 'q';
          Dec(j);
          AValue[j] := '&';
          Dec(j);
        end;
        else begin
          AValue[j] := AValue[i];
          Dec(j);
        end;
      end;
    end;
  end;

{$ifdef BABOON}
  L := UnicodeToUtf8(FAttributes.pEnd,AX_STREAM_WRITE_BUF_SZ,PWideChar(AValue),Length(AValue)) - 1;
{$else}
  L := WideCharToMultiByte(CP_UTF8,0,PWideChar(AValue),Length(AValue),FAttributes.pEnd,AX_STREAM_WRITE_BUF_SZ,Nil,Nil);
{$endif}
  Inc(FAttributes.pEnd,L);

  if ((Integer(FAttributes.pEnd) - Integer(FAttributes.pStart) + 1)) > AX_STREAM_WRITE_BUF_SZ then
    raise Exception.Create('Write XML buffer owerflow');
end;

constructor TXpgWriteXML.Create;
begin
  FStack := TXpgStrStack.Create;
  FWriteIndent := True;

  SetLength(FSpaces,AX_MAXINDENT * 2);
  FillChar(Pointer(FSpaces)^,AX_MAXINDENT * 2,' ');

  GetMem(FAttributes.pStart,AX_STREAM_WRITE_BUF_SZ_EX);
  FAttributes.pEnd := FAttributes.pStart;

{$ifdef AXW}
  FNSStack := TStringList.Create;
{$endif}
end;

procedure TXpgWriteXML.Declaration;
begin
  WriteString('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + #13);
end;

destructor TXpgWriteXML.Destroy;
begin
  FStack.Free;
  if FOwnsStream and (FStream <> Nil) then
    FStream.Free;
  FreeMem(FAttributes.pStart);

{$ifdef AXW}
  FNSStack.Free;
{$endif}

  inherited;
end;

procedure TXpgWriteXML.EmptyTag(const AName: AxUCString);
begin
{$ifdef AXW}
  WriteString('<' + CheckNS(AName) + '/>' + #13);
{$else}
  WriteString('<' + AName + '/>' + #13);
{$endif}
end;

procedure TXpgWriteXML.EndTag;
begin
  if FText <> '' then begin
    CopyToBufUTF8(FText);
    WriteString(FAttributes.pStart,Integer(FAttributes.pEnd) - Integer(FAttributes.pStart));
    FText := '';
    FAttributes.pEnd := FAttributes.pStart;
  end;
{$ifdef AXW}
  WriteString('</' + CheckNS(FStack.Pop) + '>' + #13);
{$else}
  WriteString('</' + FStack.Pop + '>' + #13);
{$endif}
end;

{$ifdef AXW}
procedure TXpgWriteXML.PopNS;
begin
  FNamespace := FNSStack[FNSStack.Count - 1];
  FNSStack.Delete(FNSStack.Count - 1);
end;

procedure TXpgWriteXML.PushNS(ANS: AxUCString);
begin
  FNSStack.Add(FNamespace);
  FNamespace := ANS;
end;
{$endif}

procedure TXpgWriteXML.SaveToFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  if FStream = Nil then begin
    Stream := TFileStream.Create(AFilename,fmCreate);
    FOwnsStream := True;
    SaveToStream(Stream);
  end;
end;

procedure TXpgWriteXML.SaveToStream(AStream: TStream);
begin
  FStream := AStream;
  Declaration;
end;

procedure TXpgWriteXML.SetAttributes(const Value: AxUCString);
begin
  AddAttr(Value);
end;

{$ifdef AXW}
procedure TXpgWriteXML.SetNamespace(const Value: AxUCString);
begin
  FNamespace := Value;
  FUseAttrNS := TRue;
end;
{$endif}

procedure TXpgWriteXML.SimpleTextTag(const AName, AText: AxUCString);
begin
  if FAttributes.pEnd = FAttributes.pStart then begin
{$ifdef AXW}
    WriteString('<' + CheckNS(AName) + '>');
{$else}
    WriteString('<' + AName + '>');
{$endif}
    CopyToBufUTF8(AText);
    WriteString(FAttributes.pStart,Integer(FAttributes.pEnd) - Integer(FAttributes.pStart));
{$ifdef AXW}
    WriteString('</' + CheckNS(AName) + '>' + #13);
{$else}
    WriteString('</' + AName + '>' + #13);
{$endif}
  end
  else begin
{$ifdef AXW}
    WriteString('<' + CheckNS(AName) + ' ');
{$else}
    WriteString('<' + AName + ' ');
{$endif}
    WriteString(FAttributes.pStart,Integer(FAttributes.pEnd) - Integer(FAttributes.pStart));
    WriteString('>');

    FAttributes.pEnd := FAttributes.pStart;
    CopyToBufUTF8(AText);
    WriteString(FAttributes.pStart,Integer(FAttributes.pEnd) - Integer(FAttributes.pStart));

{$ifdef AXW}
    WriteString('</' + CheckNS(AName) + '>' + #13);
{$else}
    WriteString('</' + AName + '>' + #13);
{$endif}
  end;
  FAttributes.pEnd := FAttributes.pStart;
end;

procedure TXpgWriteXML.WriteString(const S: PAnsiChar; const Len: integer);
begin
//  if FWriteIndent then begin
//    if FStack.Count <= AX_MAXINDENT then
//      FStream.Write(Pointer(FSpaces)^,FStack.Count * 2)
//    else
//      FStream.Write(Pointer(FSpaces)^,AX_MAXINDENT * 2)
//  end;
//
  FStream.Write(Pointer(S)^,Len);
end;

procedure TXpgWriteXML.SimpleTag(const AName: AxUCString);
begin
  if FText <> '' then begin
    SimpleTextTag(AName,FText);
    FText := '';
  end
  else begin
    if FAttributes.pStart = FAttributes.pEnd then
{$ifdef AXW}
      WriteString('<' + CheckNS(AName) + '/>' + #13)
{$else}
      WriteString('<' + AName + '/>' + #13)
{$endif}
    else begin
{$ifdef AXW}
      WriteString('<' + CheckNS(AName) + ' ');
{$else}
      WriteString('<' + AName + ' ');
{$endif}
      WriteString(FAttributes.pStart,Integer(FAttributes.pEnd) - Integer(FAttributes.pStart));
      WriteString('/>' + #13);

      FAttributes.pEnd := FAttributes.pStart;
    end;
  end;
end;

procedure TXpgWriteXML.WriteString(const S: AxUCString);
var
  S8: Ax8String;
begin
//  if FWriteIndent then begin
//    if FStack.Count <= AX_MAXINDENT then
//      FStream.Write(Pointer(FSpaces)^,FStack.Count * 2)
//    else
//      FStream.Write(Pointer(FSpaces)^,AX_MAXINDENT * 2)
//  end;

  S8 := Ax8String(S);
  FStream.Write(Pointer(S8)^,Length(S8));
end;

{ TXpgStrStack }

function TXpgStrStack.Count: integer;
begin
  Result := FValues.Count;
end;

constructor TXpgStrStack.Create;
begin
  FValues := TStringList.Create;
end;

destructor TXpgStrStack.Destroy;
begin
  FValues.Free;
  inherited;
end;

function TXpgStrStack.Pop: string;
begin
  if FValues.Count < 1 then
    raise Exception.Create('Empty stack');
  Result := FValues[FValues.Count - 1];
  FValues.Delete(FValues.Count - 1);
end;

procedure TXpgStrStack.Push(AValue: string);
begin
  FValues.Add(AValue);
end;

end.
