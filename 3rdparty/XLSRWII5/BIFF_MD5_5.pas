unit BIFF_MD5_5;

// **********************************************************************
// ** Copyright (C) 1990, RSA Data Security, Inc. All rights reserved. **
// **                                                                  **
// ** License to copy and use this software is granted provided that   **
// ** it is identified as the "RSA Data Security, Inc. MD5 Message     **
// ** Digest Algorithm" in all material mentioning or referencing this **
// ** software or this function.                                       **
// **                                                                  **
// ** License is also granted to make and use derivative works         **
// ** provided that such works are identified as "derived from the RSA **
// ** Data Security, Inc. MD5 Message Digest Algorithm" in all         **
// ** material mentioning or referencing the derived work.             **
// **                                                                  **
// ** RSA Data Security, Inc. makes no representations concerning      **
// ** either the merchantability of this software or the suitability   **
// ** of this software for any particular purpose.  It is provided "as **
// ** is" without express or implied warranty of any kind.             **
// **                                                                  **
// ** These notices must be retained in any copies of any part of this **
// ** documentation and/or software.                                   **
// **********************************************************************

interface

uses Classes, SysUtils,
     XLSUtils5;

type TLongwordArray = array[0..15] of longword;
type PLongwordArray = ^TLongwordArray;

type TMD5Context = record
    I: array[0..1] of longword;
    Buf: array[0..3] of longword;
    InBuf: array[0..63] of byte;
    Digest: array [0..15] of byte;
    end;

type TMD5= class(TObject)
protected
{
    FInitialized: boolean;
    LenHi, LenLo: longword;
    Index: longword;
    CurrentHash: array[0..3] of longword;
    HashBuffer: array[0..63] of byte;
}
    procedure Compress(Buf,Data: PLongwordArray);
public
    procedure Init(var Ctx: TMD5Context);
    procedure Update(var Ctx: TMD5Context; InBuf: PByteArray; InLen: longword);
    procedure Store(var Ctx: TMD5Context);
    end;

function XLSMD5String(const S: AnsiString): AxUCString;
function XLSMD5Buffer(const Buffer; Size: Integer): AxUCString;


{******************************************************************************}
{******************************************************************************}
implementation
{$R-}{$Q-}

function LRot32(a, b: longword): longword;
begin
  Result:= (a shl b) or (a shr (32-b));
end;

procedure TMD5.Compress(Buf,Data: PLongwordArray);
var
  A, B, C, D: longword;
begin
  A:= Buf[0];
  B:= Buf[1];
  C:= Buf[2];
  D:= Buf[3];

  A:= B + LRot32(A + (D xor (B and (C xor D))) + Data[ 0] + $d76aa478,7);
  D:= A + LRot32(D + (C xor (A and (B xor C))) + Data[ 1] + $e8c7b756,12);
  C:= D + LRot32(C + (B xor (D and (A xor B))) + Data[ 2] + $242070db,17);
  B:= C + LRot32(B + (A xor (C and (D xor A))) + Data[ 3] + $c1bdceee,22);
  A:= B + LRot32(A + (D xor (B and (C xor D))) + Data[ 4] + $f57c0faf,7);
  D:= A + LRot32(D + (C xor (A and (B xor C))) + Data[ 5] + $4787c62a,12);
  C:= D + LRot32(C + (B xor (D and (A xor B))) + Data[ 6] + $a8304613,17);
  B:= C + LRot32(B + (A xor (C and (D xor A))) + Data[ 7] + $fd469501,22);
  A:= B + LRot32(A + (D xor (B and (C xor D))) + Data[ 8] + $698098d8,7);
  D:= A + LRot32(D + (C xor (A and (B xor C))) + Data[ 9] + $8b44f7af,12);
  C:= D + LRot32(C + (B xor (D and (A xor B))) + Data[10] + $ffff5bb1,17);
  B:= C + LRot32(B + (A xor (C and (D xor A))) + Data[11] + $895cd7be,22);
  A:= B + LRot32(A + (D xor (B and (C xor D))) + Data[12] + $6b901122,7);
  D:= A + LRot32(D + (C xor (A and (B xor C))) + Data[13] + $fd987193,12);
  C:= D + LRot32(C + (B xor (D and (A xor B))) + Data[14] + $a679438e,17);
  B:= C + LRot32(B + (A xor (C and (D xor A))) + Data[15] + $49b40821,22);

  A:= B + LRot32(A + (C xor (D and (B xor C))) + Data[ 1] + $f61e2562,5);
  D:= A + LRot32(D + (B xor (C and (A xor B))) + Data[ 6] + $c040b340,9);
  C:= D + LRot32(C + (A xor (B and (D xor A))) + Data[11] + $265e5a51,14);
  B:= C + LRot32(B + (D xor (A and (C xor D))) + Data[ 0] + $e9b6c7aa,20);
  A:= B + LRot32(A + (C xor (D and (B xor C))) + Data[ 5] + $d62f105d,5);
  D:= A + LRot32(D + (B xor (C and (A xor B))) + Data[10] + $02441453,9);
  C:= D + LRot32(C + (A xor (B and (D xor A))) + Data[15] + $d8a1e681,14);
  B:= C + LRot32(B + (D xor (A and (C xor D))) + Data[ 4] + $e7d3fbc8,20);
  A:= B + LRot32(A + (C xor (D and (B xor C))) + Data[ 9] + $21e1cde6,5);
  D:= A + LRot32(D + (B xor (C and (A xor B))) + Data[14] + $c33707d6,9);
  C:= D + LRot32(C + (A xor (B and (D xor A))) + Data[ 3] + $f4d50d87,14);
  B:= C + LRot32(B + (D xor (A and (C xor D))) + Data[ 8] + $455a14ed,20);
  A:= B + LRot32(A + (C xor (D and (B xor C))) + Data[13] + $a9e3e905,5);
  D:= A + LRot32(D + (B xor (C and (A xor B))) + Data[ 2] + $fcefa3f8,9);
  C:= D + LRot32(C + (A xor (B and (D xor A))) + Data[ 7] + $676f02d9,14);
  B:= C + LRot32(B + (D xor (A and (C xor D))) + Data[12] + $8d2a4c8a,20);

  A:= B + LRot32(A + (B xor C xor D) + Data[ 5] + $fffa3942,4);
  D:= A + LRot32(D + (A xor B xor C) + Data[ 8] + $8771f681,11);
  C:= D + LRot32(C + (D xor A xor B) + Data[11] + $6d9d6122,16);
  B:= C + LRot32(B + (C xor D xor A) + Data[14] + $fde5380c,23);
  A:= B + LRot32(A + (B xor C xor D) + Data[ 1] + $a4beea44,4);
  D:= A + LRot32(D + (A xor B xor C) + Data[ 4] + $4bdecfa9,11);
  C:= D + LRot32(C + (D xor A xor B) + Data[ 7] + $f6bb4b60,16);
  B:= C + LRot32(B + (C xor D xor A) + Data[10] + $bebfbc70,23);
  A:= B + LRot32(A + (B xor C xor D) + Data[13] + $289b7ec6,4);
  D:= A + LRot32(D + (A xor B xor C) + Data[ 0] + $eaa127fa,11);
  C:= D + LRot32(C + (D xor A xor B) + Data[ 3] + $d4ef3085,16);
  B:= C + LRot32(B + (C xor D xor A) + Data[ 6] + $04881d05,23);
  A:= B + LRot32(A + (B xor C xor D) + Data[ 9] + $d9d4d039,4);
  D:= A + LRot32(D + (A xor B xor C) + Data[12] + $e6db99e5,11);
  C:= D + LRot32(C + (D xor A xor B) + Data[15] + $1fa27cf8,16);
  B:= C + LRot32(B + (C xor D xor A) + Data[ 2] + $c4ac5665,23);

  A:= B + LRot32(A + (C xor (B or (not D))) + Data[ 0] + $f4292244,6);
  D:= A + LRot32(D + (B xor (A or (not C))) + Data[ 7] + $432aff97,10);
  C:= D + LRot32(C + (A xor (D or (not B))) + Data[14] + $ab9423a7,15);
  B:= C + LRot32(B + (D xor (C or (not A))) + Data[ 5] + $fc93a039,21);
  A:= B + LRot32(A + (C xor (B or (not D))) + Data[12] + $655b59c3,6);
  D:= A + LRot32(D + (B xor (A or (not C))) + Data[ 3] + $8f0ccc92,10);
  C:= D + LRot32(C + (A xor (D or (not B))) + Data[10] + $ffeff47d,15);
  B:= C + LRot32(B + (D xor (C or (not A))) + Data[ 1] + $85845dd1,21);
  A:= B + LRot32(A + (C xor (B or (not D))) + Data[ 8] + $6fa87e4f,6);
  D:= A + LRot32(D + (B xor (A or (not C))) + Data[15] + $fe2ce6e0,10);
  C:= D + LRot32(C + (A xor (D or (not B))) + Data[ 6] + $a3014314,15);
  B:= C + LRot32(B + (D xor (C or (not A))) + Data[13] + $4e0811a1,21);
  A:= B + LRot32(A + (C xor (B or (not D))) + Data[ 4] + $f7537e82,6);
  D:= A + LRot32(D + (B xor (A or (not C))) + Data[11] + $bd3af235,10);
  C:= D + LRot32(C + (A xor (D or (not B))) + Data[ 2] + $2ad7d2bb,15);
  B:= C + LRot32(B + (D xor (C or (not A))) + Data[ 9] + $eb86d391,21);

  Inc(Buf[0],A);
  Inc(Buf[1],B);
  Inc(Buf[2],C);
  Inc(Buf[3],D);
end;

procedure TMD5.Init(var Ctx: TMD5Context);
begin
  Ctx.I[0] := 0;
  Ctx.I[1] := 0;
  Ctx.Buf[0]:= $67452301;
  Ctx.Buf[1]:= $efcdab89;
  Ctx.Buf[2]:= $98badcfe;
  Ctx.Buf[3]:= $10325476;
end;

procedure TMD5.Store(var Ctx: TMD5Context);
var
  i,j: integer;
begin
  j := 0;
  for i := 0 to 3 do begin
    Ctx.Digest[j] := Ctx.Buf[i] and $FF;
    Ctx.Digest[j + 1] := (Ctx.Buf[i] shr 8) and $FF;
    Ctx.Digest[j + 2] := (Ctx.Buf[i] shr 16) and $FF;
    Ctx.Digest[j + 3] := (Ctx.Buf[i] shr 24) and $FF;
    Inc(j,4);
  end;
end;

procedure TMD5.Update(var Ctx: TMD5Context; InBuf: PByteArray; InLen: longword);
var
  TmpBuf: array[0..15] of longword;
  mdi,i,j,k: integer;
begin
  mdi := ((Ctx.I[0] shr 3) and $3F);

  if (Ctx.I[0] + (InLen shl 3)) < Ctx.I[0] then
    Inc(Ctx.I[1]);
  Inc(Ctx.I[0],InLen shl 3);
  Inc(Ctx.I[1],InLen shr 29);

  for k := 0 to InLen - 1 do begin
    Ctx.InBuf[mdi] := InBuf[k];
    Inc(mdi);
    if mdi = $40 then begin
      j := 0;
      for i := 0 to 15 do begin
        TmpBuf[i] := ((Ctx.InBuf[j + 3]) shl 24) or ((Ctx.InBuf[j + 2]) shl 16) or ((Ctx.InBuf[j + 1]) shl 8) or (Ctx.InBuf[j]);
        Inc(j,4);
      end;
      Compress (@Ctx.Buf,@TmpBuf[0]);
      mdi := 0;
    end;
  end;
end;

function XLSMD5String(const S: AnsiString): AxUCString;
begin
  Result := XLSMD5Buffer(PAnsiChar(S)^,Length(S));
end;

function XLSMD5Buffer(const Buffer; Size: Integer): AxUCString;
var
  MD5: TMD5;
  Context: TMD5Context;
begin
  MD5 := TMD5.Create;
  try
    MD5.Init(Context);
    MD5.Update(Context, PByteArray(@Buffer), Size);
    MD5.Store(Context);
  finally
    MD5.Free;
  end;
end;

end.
