unit Xc12DataSST5;

{
********************************************************************************
******* XLSReadWriteII V5.00                                             *******
*******                                                                  *******
******* Copyright(C) 1999,2012 Lars Arvidsson, Axolot Data               *******
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

{$I AxCompilers.inc}

{$B-}
{$R-}
{$H+}

interface

uses Classes, SysUtils, Contnrs,
     XLSUtils5,
     Xc12Common5, Xc12DataStyleSheet5;

type PXLSString = ^TXLSString;
     TXLSString = record
     RefCount: integer;
     Len: word;
     Options: byte;
     Data: record end;
     end;

const STRID_COMPRESSED      = $00;
const STRID_UNICODE         = $01;
const STRID_RICH            = $08;
const STRID_RICH_UNICODE    = STRID_RICH + STRID_UNICODE;
const STRID_FAREAST         = $04;
const STRID_FAREAST_RICH    = STRID_FAREAST + STRID_RICH;
const STRID_FAREAST_UC      = STRID_FAREAST + STRID_UNICODE;
const STRID_FAREAST_RICH_UC = STRID_FAREAST + STRID_UNICODE + STRID_RICH;

const SST_HASHSIZE          = $FFFF;

  {$IFDEF CPUX64}
type XLSPointerArray = array [0..256*1024*1024 - 2] of Pointer;
  {$ELSE !CPUX64}
type XLSPointerArray = array [0..512*1024*1024 - 2] of Pointer;
  {$ENDIF !CPUX64}
type XLSPPointerArray = ^XLSPointerArray;

// Unicode
type PXLSStringUC = ^TXLSStringUC;
     TXLSStringUC = record
     RefCount: integer;
     Len: word;
     Options: byte;
     Data: record end;
     end;

// Compressed with formatting
type PXLSStringRich = ^TXLSStringRich;
     TXLSStringRich = record
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     Data: record end;
     end;

// Unicode with formatting
type PXLSStringRichUC = ^TXLSStringRichUC;
     TXLSStringRichUC = record
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     Data: record end;
     end;

// Compressed with Far East data
type PXLSStringFarEast = ^TXLSStringFarEast;
     TXLSStringFarEast = record
     RefCount: integer;
     Len: word;
     Options: byte;
     FarEastDataSize: longword;
     Data: record end;
     end;

// Unicode with Far East data
type PXLSStringFarEastUC = ^TXLSStringFarEastUC;
     TXLSStringFarEastUC = record
     RefCount: integer;
     Len: word;
     Options: byte;
     FarEastDataSize: longword;
     Data: record end;
     end;

// Compressed with Far East data and formatting
type PXLSStringFarEastRich = ^TXLSStringFarEastRich;
     TXLSStringFarEastRich = record
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     FarEastDataSize: longword;
     Data: record end;
     end;

// Unicode with Far East data and formatting
type PXLSStringFarEastRichUC = ^TXLSStringFarEastRichUC;
     TXLSStringFarEastRichUC = record
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     FarEastDataSize: longword;
     Data: record end;
     end;

type PPHashItem = ^PHashItem;
     PHashItem = ^THashItem;
     THashItem = record
     Next: PHashItem;
     Key: string;
     Value: Integer;
     end;

type TStringHash = class
private
     Buckets: array of PHashItem;
protected
     function Find(const Key: AxUCString): PPHashItem;
     function HashOf(const Key: AxUCString): Cardinal; virtual;
public
     constructor Create(Size: Cardinal = 256);
     destructor Destroy; override;

     procedure Add(const Key: AxUCString; Value: Integer);
     procedure Clear;
     procedure Remove(const Key: AxUCString);
     function Modify(const Key: AxUCString; Value: Integer): Boolean;
     function ValueOf(const Key: AxUCString): Integer;
end;

type TXc12DataSST = class(TXc12Data)
private
     function GetItems(Index: integer): PXLSString;
     function GetIsFormatted(Index: integer): boolean;
     function GetItemText(Index: integer): AxUCString;
protected
     FHashTable : TStringHash;
     FHashValid : boolean;

     FStyles    : TXc12DataStyleSheet;
     FFonts     : TXc12Fonts;
     FSST       : XLSPPointerArray;
     FTotalCount: integer;
     FAllocCount: integer;
     FFreeCount : integer;

     FCompress  : boolean;

     FIsExcel97 : boolean;

     procedure BuildHash;
     procedure Grow;
     function  IsUnicode(const S: AxUCString): boolean;
     function  GetDataPointer(Value: PXLSString): PByteArray;

     function  MakeString(const S: AxUCString): PXLSString;
     function  MakeCompressedString(const WS: AxUCString): PXLSString;

     procedure EncodeString(var AText: AxUCString);
public
     constructor Create(AStyles: TXc12DataStyleSheet);
     destructor Destroy; override;

     procedure AfterRead;

     procedure Clear;

     function  AddRawString(const ALen: integer; const AData: PByteArray): integer; overload;
     function  AddRawString(const ALen: integer; const AData: PWordArray): integer; overload;

     // Don't tries to find a matching string.
     function  FileAddString(const AText: AxUCString): integer;
     function  AddString(AText: AxUCString): integer;
     function  AddFormattedString(const AText: AxUCString; AFontRuns: TXc12DynFontRunArray): integer;

     function  GetText(XString: PXLSString): AxUCString;

     procedure UsesString(AIndex: integer);
     procedure ReleaseString(AIndex: integer);

     function  GetFontRuns(Value: PXLSString): PXc12FontRunArray;
     function  GetFontRunsCount(Value: PXLSString): integer;

     property CompressStrings: boolean read FCompress write FCompress;
     property TotalCount: integer read FTotalCount;
     property Items[Index: integer]: PXLSString read GetItems; default;
     property IsFormatted[Index: integer]: boolean read GetIsFormatted;
     property ItemText[Index: integer]: AxUCString read GetItemText;

     property IsExcel97: boolean read FIsExcel97 write FIsExcel97;
     end;

function  SSTStringIsRichStr(const ASSTString: PXLSString): boolean;

implementation

function  SSTStringIsRichStr(const ASSTString: PXLSString): boolean;
begin
  Result := ASSTString.Options in [STRID_RICH,STRID_RICH_UNICODE,STRID_FAREAST_RICH,STRID_FAREAST_RICH_UC];
end;

const SST_ALLOC_CHUNK = $FF;

{ TStringHash }

procedure TStringHash.Add(const Key: AxUCString; Value: Integer);
var
  Hash: Integer;
  Bucket: PHashItem;
begin
  Hash := HashOf(Key) mod Cardinal(Length(Buckets));
  New(Bucket);
  Bucket^.Key := Key;
  Bucket^.Value := Value;
  Bucket^.Next := Buckets[Hash];
  Buckets[Hash] := Bucket;
end;

procedure TStringHash.Clear;
var
  i: Integer;
  P, N: PHashItem;
begin
  for i := 0 to Length(Buckets) - 1 do begin
    P := Buckets[I];
    while P <> nil do begin
      N := P^.Next;
      Dispose(P);
      P := N;
    end;
    Buckets[I] := nil;
  end;
end;

constructor TStringHash.Create(Size: Cardinal);
begin
  inherited Create;
  SetLength(Buckets, Size);
end;

destructor TStringHash.Destroy;
begin
  Clear;

  inherited Destroy;
end;

function TStringHash.Find(const Key: AxUCString): PPHashItem;
var
  Hash: Integer;
begin
  Hash := HashOf(Key) mod Cardinal(Length(Buckets));
  Result := @Buckets[Hash];
  while Result^ <> nil do begin
    if Result^.Key = Key then
      Exit
    else
      Result := @Result^.Next;
  end;
end;

function TStringHash.HashOf(const Key: AxUCString): Cardinal;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(Key) do
    Result := ((Result shl 2) or (Result shr (SizeOf(Result) * 8 - 2))) xor Ord(Key[i]);
end;

function TStringHash.Modify(const Key: AxUCString; Value: Integer): Boolean;
var
  P: PHashItem;
begin
  P := Find(Key)^;
  if P <> nil then begin
    Result := True;
    P^.Value := Value;
  end
  else
    Result := False;
end;

procedure TStringHash.Remove(const Key: AxUCString);
var
  P: PHashItem;
  Prev: PPHashItem;
begin
  Prev := Find(Key);
  P := Prev^;
  if P <> nil then begin
    Prev^ := P^.Next;
    Dispose(P);
  end;
end;

function TStringHash.ValueOf(const Key: AxUCString): Integer;
var
  P: PHashItem;
begin
  P := Find(Key)^;
  if P <> nil then
    Result := P^.Value
  else
    Result := -1;
end;

{ TXc12DataSST }

function TXc12DataSST.AddFormattedString(const AText: AxUCString; AFontRuns: TXc12DynFontRunArray): integer;
var
  i: integer;
  XLSStr: PXLSString;
  PRuns: PXc12FontRunArray;

function RichMakeString(const S: AxUCString): PXLSString;
begin
  Result := AllocMem(SizeOf(TXLSStringRichUC) + (Length(S) * 2) + Length(AFontRuns) * SizeOf(TXc12FontRun));
  Result.Options := STRID_UNICODE + STRID_RICH;
  PXLSStringRichUC(Result).FormatCount := Length(AFontRuns);

  Result.RefCount := 0;
  Result.Len := Length(S);
  Move(Pointer(S)^,GetDataPointer(Result)^,Length(S) * 2);
end;

function RichMakeCompressedString(const WS: AxUCString): PXLSString;
var
  i: integer;
  S: XLS8String;
begin
  SetLength(S,Length(WS));
  for i := 1 to Length(WS) do
    S[i] := AnsiChar(WS[i]);

  Result := AllocMem(SizeOf(TXLSStringRich) + Length(S) + Length(AFontRuns) * SizeOf(TXc12FontRun));
  Result.Options := STRID_COMPRESSED + STRID_RICH;
  PXLSStringRich(Result).FormatCount := Length(AFontRuns);

  Result.RefCount := 0;
  Result.Len := Length(S);
  Move(Pointer(S)^,GetDataPointer(Result)^,Length(S));
end;

begin
  if FTotalCount >= FAllocCount then
    Grow;

  if IsUnicode(AText) then
    XLSStr := RichMakeString(AText)
  else
    XLSStr := RichMakeCompressedString(AText);

  FSST[FTotalCount] := XLSStr;

  PRuns := GetFontRuns(XLSStr);
  for i := 0 to High(AFontRuns) do begin
    PRuns[i].Index := AFontRuns[i].Index;
    PRuns[i].Font := AFontRuns[i].Font;
  end;

  Result := FTotalCount;
  Inc(FTotalCount);
end;

function TXc12DataSST.AddRawString(const ALen: integer; const AData: PByteArray): integer;
var
  i: integer;
  j: integer;
  P: PXLSString;
  PData: PByteArray;
begin
  if FTotalCount >= FAllocCount then
    Grow;

  GetMem(P,SizeOf(TXLSString) + ALen);
  P.Options := STRID_COMPRESSED;
  P.RefCount := 0;
  P.Len := ALen;


  PData := GetDataPointer(P);
  j := 0;
  for i := 0 to ALen - 1 do begin
    if (AData[i] >= 32) or (AData[i] = 10) then begin
      PData[j] := AData[i];
      Inc(j);
    end;
  end;
  ReallocMem(P,SizeOf(TXLSString) + j);
  P.Len := j;


//  Move(AData^,GetDataPointer(P)^,ALen);

  FSST[FTotalCount] := P;

  Result := FTotalCount;
  Inc(FTotalCount);
end;

function TXc12DataSST.AddRawString(const ALen: integer; const AData: PWordArray): integer;
var
  i: integer;
  j: integer;
  P: PXLSString;
  PData: PWordArray;
begin
  if FTotalCount >= FAllocCount then
    Grow;

  GetMem(P,SizeOf(TXLSString) + ALen * 2);
  P.Options := STRID_UNICODE;
  P.RefCount := 0;
  P.Len := ALen;


  PData := PWordArray(GetDataPointer(P));
  j := 0;
  for i := 0 to ALen - 1 do begin
    if (AData[i] >= 32) or (AData[i] in [10,13]) then begin
      PData[j] := AData[i];
      Inc(j);
    end;
  end;
  ReallocMem(P,SizeOf(TXLSString) + j * 2);
  P.Len := j;


//  Move(AData^,GetDataPointer(P)^,ALen * 2);

  FSST[FTotalCount] := P;

  Result := FTotalCount;
  Inc(FTotalCount);
end;

function TXc12DataSST.AddString(AText: AxUCString): integer;
var
  P: PXLSString;
  SlotFound: boolean;
begin
  if not FHashValid then
    BuildHash;

  if not FIsExcel97 then
    EncodeString(AText);

  Result := FHashTable.ValueOf(AText);

  if Result < 0 then begin
    if FTotalCount >= FAllocCount then
      Grow;

    if IsUnicode(AText) then
      P := MakeString(AText)
    else
      P := MakeCompressedString(AText);

    Result := FTotalCount;

    SlotFound := False;
    if FFreeCount > 0 then begin
      Dec(FFreeCount);

      for Result := 0 to FTotalCount - 1 do begin
        if FSST[Result] = Nil then begin
          SlotFound := True;
          Break;
        end;
      end;
    end;

    FHashTable.Add(AText,Result);
    FSST[Result] := P;

    if not SlotFound then
      Inc(FTotalCount);
  end;
end;

procedure TXc12DataSST.AfterRead;
begin
  BuildHash;
end;

procedure TXc12DataSST.BuildHash;
var
  i: integer;
  S: AxUCString;
begin
  FHashTable.Clear;
  for i := 0 to FTotalCount - 1 do begin
    // Don't hash rich strings.
    if (FSST[i] <> Nil) and ((PXLSString(FSST[i]).Options and STRID_RICH) = 0) then begin
      S := ItemText[i];
      FHashTable.Add(S,i);
    end;
  end;
  FHashValid := True;
end;

procedure TXc12DataSST.Clear;
var
  i: integer;
  PRuns: PXc12FontRunArray;
begin
  for i := 0 to FTotalCount - 1 do begin
    if FSST[i] <> Nil then begin
      if (PXLSString(FSST[i]).Options and STRID_RICH) <> 0 then begin
        PRuns := GetFontRuns(PXLSString(FSST[i]));
        FStyles.XFEditor.FreeFonts(PRuns,GetFontRunsCount(PXLSString(FSST[i])));
      end;
      FreeMem(FSST[i]);
      FSST[i] := Nil;
    end;
  end;
  FTotalCount := 0;
  FAllocCount := 0;
  FFreeCount := 0;
  FreeMem(FSST);
  FSST := Nil;

  FHashTable.Clear;
  FHashValid := False;
end;

constructor TXc12DataSST.Create(AStyles: TXc12DataStyleSheet);
begin
  FStyles := AStyles;
  FFonts := FStyles.Fonts;

  FHashTable := TStringHash.Create($FFFF);

  FCompress := True;
end;

procedure TXc12DataSST.ReleaseString(AIndex: integer);
var
  PRuns: PXc12FontRunArray;
begin
  if FSST[AIndex] <> Nil then begin
    Dec(PXLSString(FSST[AIndex]).RefCount);
    if PXLSString(FSST[AIndex]).RefCount <= 0 then begin
      if (PXLSString(FSST[AIndex]).Options and STRID_RICH) <> 0 then begin
        PRuns := GetFontRuns(PXLSString(FSST[AIndex]));
        FStyles.XFEditor.FreeFonts(PRuns,GetFontRunsCount(PXLSString(FSST[AIndex])));
      end
      else
        // Rich strings are never hashed.
        FHashTable.Remove(ItemText[AIndex]);
      FreeMem(FSST[AIndex]);
      FSST[AIndex] := Nil;

      if AIndex = (FTotalCount - 1) then
        Dec(FTotalCount)
      else
        Inc(FFreeCount);
    end;
  end;
end;

destructor TXc12DataSST.Destroy;
begin
  Clear;
  FHashTable.Free;
end;

procedure TXc12DataSST.EncodeString(var AText: AxUCString);
var
  i: integer;
  L: integer;
begin
  // Check for _xnnnn_

  L := Length(AText) - 6;
  i := 1;
  while i <= L do begin
    if (AText[i] = '_') and (AText[i + 6] = '_') and (AText[i + 1] = 'x') and (StrToIntDef('$' + Copy(AText,i + 2,4),MAXINT) <> MAXINT) then begin
      System.Insert('_x005F',AText,i);
      Inc(i,13);
      Inc(L,6);
    end
    else
      Inc(i);
  end;
end;

function TXc12DataSST.FileAddString(const AText: AxUCString): integer;
var
  P: PXLSString;
begin
  if FTotalCount >= FAllocCount then
    Grow;

  if IsUnicode(AText) then
    P := MakeString(AText)
  else
    P := MakeCompressedString(AText);

  Result := FTotalCount;

  FSST[Result] := P;

  Inc(FTotalCount);
end;

function TXc12DataSST.GetDataPointer(Value: PXLSString): PByteArray;
begin
  case Value.Options of
    STRID_COMPRESSED     : Result := @PXLSString(Value).Data;
    STRID_UNICODE        : Result := @PXLSStringUC(Value).Data;
    STRID_RICH           : Result := @PXLSStringRich(Value).Data;
    STRID_RICH_UNICODE   : Result := @PXLSStringRichUC(Value).Data;
    STRID_FAREAST        : Result := @PXLSStringFarEast(Value).Data;
    STRID_FAREAST_UC     : Result := @PXLSStringFarEastUC(Value).Data;
    STRID_FAREAST_RICH   : Result := @PXLSStringFarEastRich(Value).Data;
    STRID_FAREAST_RICH_UC: Result := @PXLSStringFarEastRichUC(Value).Data;
    else
      raise XLSRWException.Create('SST: Unhandled string type.');
  end;
end;

function TXc12DataSST.GetFontRuns(Value: PXLSString): PXc12FontRunArray;
begin
  case Value.Options of
    STRID_RICH:            Result := PXc12FontRunArray(NativeInt(@PXLSStringRich(Value).Data) +
                                     PXLSStringRich(Value).Len);
    STRID_RICH_UNICODE:    Result := PXc12FontRunArray(NativeInt(@PXLSStringRichUC(Value).Data) +
                                     PXLSStringRichUC(Value).Len * 2);
    STRID_FAREAST_RICH:    Result := PXc12FontRunArray(NativeInt(@PXLSStringFarEastRich(Value).Data) +
                                     PXLSStringFarEastRich(Value).Len);
    STRID_FAREAST_RICH_UC: Result := PXc12FontRunArray(NativeInt(@PXLSStringFarEastRichUC(Value).Data) +
                                     PXLSStringFarEastRichUC(Value).Len * 2);
    else                   Result := Nil;
  end;
end;

function TXc12DataSST.GetFontRunsCount(Value: PXLSString): integer;
begin
  if (Value.Options and STRID_RICH) <> 0 then
    Result := PXLSStringRich(Value).FormatCount
  else
    Result := 0;
end;

function TXc12DataSST.GetIsFormatted(Index: integer): boolean;
var
  XStr: PXLSString;
begin
  XStr := Items[Index];
  Result := (XStr <> Nil) and (XStr.Options in [STRID_RICH,STRID_RICH_UNICODE,STRID_FAREAST_RICH,STRID_FAREAST_RICH_UC]);
end;

function TXc12DataSST.GetItems(Index: integer): PXLSString;
begin
  if (Index < 0) or (Index >= FTotalCount) then
    raise XLSRWException.Create('Index out of range');
  Result := PXLSString(FSST[Index]);
end;

function TXc12DataSST.GetItemText(Index: integer): AxUCString;
var
  XStr: PXLSString;
begin
  if (Index < 0) or (Index >= FTotalCount) then
    raise XLSRWException.Create('Index out of range');
  XStr := PXLSString(FSST[Index]);
  Result := GetText(XStr);
end;

function TXc12DataSST.GetText(XString: PXLSString): AxUCString;
var
  S8: XLS8String;
begin
  case XString.Options of
    STRID_COMPRESSED: begin
//      SetLength(Result,XString.Len);
//      for i := 1 to XString.Len do
//        Result[i] := Char(PByteArray(@XString.Data)[i - 1]);

      SetLength(S8,XString.Len);
      Move(PByteArray(@XString.Data)^,Pointer(S8)^,XString.Len);
      Result := AxUCString(S8);
    end;
    STRID_UNICODE: begin
      SetLength(Result,XString.Len);
      Move(PByteArray(@XString.Data)^,Pointer(Result)^,XString.Len * 2);
    end;
    STRID_RICH: begin
//      SetLength(Result,XString.Len);
//      for i := 1 to XString.Len do
//        Result[i] := Char(PByteArray(@PXLSStringRich(XString).Data)[i - 1]);

      SetLength(S8,XString.Len);
      with PXLSStringRich(XString)^ do
        Move(PByteArray(@Data)^,Pointer(S8)^,XString.Len);
      Result := AxUCString(S8);
    end;
    STRID_RICH_UNICODE: begin
      SetLength(Result,XString.Len);
      with PXLSStringRichUC(XString)^ do
        Move(PByteArray(@Data)^,Pointer(Result)^,XString.Len * 2);
    end;
    STRID_FAREAST: begin
//      SetLength(Result,XString.Len);
//      for i := 1 to XString.Len do
//        Result[i] := Char(PByteArray(@PXLSStringFarEast(XString).Data)[i - 1]);

      SetLength(S8,XString.Len);
      with PXLSStringFarEast(XString)^ do
        Move(PByteArray(@Data)^,Pointer(S8)^,XString.Len);
      Result := AxUCString(S8);
    end;
    STRID_FAREAST_RICH: begin
//      SetLength(Result,XString.Len);
//      for i := 1 to XString.Len do
//        Result[i] := Char(PByteArray(@PXLSStringFarEastRich(XString).Data)[i - 1]);

      SetLength(S8,XString.Len);
      with PXLSStringFarEastRich(XString)^ do
        Move(PByteArray(@Data)^,Pointer(S8)^,XString.Len);
      Result := AxUCString(S8);
    end;
    STRID_FAREAST_UC: begin
      SetLength(Result,XString.Len);
      with PXLSStringFarEastUC(XString)^ do
        Move(PByteArray(@Data)^,Pointer(Result)^,XString.Len * 2);
    end;
    STRID_FAREAST_RICH_UC: begin
      SetLength(Result,XString.Len);
      with PXLSStringFarEastRichUC(XString)^ do
        Move(PByteArray(@Data)^,Pointer(Result)^,XString.Len * 2);
    end;
  end;
end;

procedure TXc12DataSST.Grow;
begin
  Inc(FAllocCount,SST_ALLOC_CHUNK);
  ReAllocMem(FSST,FAllocCount * SizeOf(Pointer));
end;

function TXc12DataSST.IsUnicode(const S: AxUCString): boolean;
var
  i: integer;
  W: word;
begin
  if not FCompress then
    Result := True
  else begin
    for i := 1 to Length(S) do begin
      W := Word(S[i]);       // This is not tested. Need a multibyte char.
      if ((W and $FF00) <> 0){ or (ByteType(S[i], 1) <> mbSingleByte)} then begin
        Result := True;
        Exit;
      end;
    end;
    Result := False;
  end;
end;

function TXc12DataSST.MakeCompressedString(const WS: AxUCString): PXLSString;
var
  i: integer;
  S: XLS8String;
begin
  SetLength(S,Length(WS));
  for i := 1 to Length(WS) do
    S[i] := AnsiChar(WS[i]);
  GetMem(Result,SizeOf(TXLSString) + Length(S));
  Result.Options := STRID_COMPRESSED;
  Result.RefCount := 0;
  Result.Len := Length(S);
  Move(Pointer(S)^,GetDataPointer(Result)^,Length(S));
end;

function TXc12DataSST.MakeString(const S: AxUCString): PXLSString;
begin
  GetMem(Result,SizeOf(TXLSString) + Length(S) * 2);
  Result.Options := STRID_UNICODE;
  Result.RefCount := 0;
  Result.Len := Length(S);
  Move(Pointer(S)^,GetDataPointer(Result)^,Length(S) * 2);
end;

procedure TXc12DataSST.UsesString(AIndex: integer);
begin
  if FSST[AIndex] <> Nil then
    Inc(PXLSString(FSST[AIndex]).RefCount)
  else
    raise XLSRWException.Create('SST String is Nil');
end;

end.
