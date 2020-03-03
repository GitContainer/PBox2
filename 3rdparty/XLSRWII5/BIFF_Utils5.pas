unit BIFF_Utils5;

interface

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

uses SysUtils, Math,
{$ifdef BABOON}
{$else}
     vcl.Graphics,
{$endif}
     BIFF_RecsII5,
     XLSUtils5,
     Xc12Utils5;

type TFormulaValType = (fvFloat,fvBoolean,fvError,fvString,fvRef,fvArea,fvExtRef,fvExtArea,fvNull);
const TFormulaValTypeRef = [fvRef,fvArea,fvExtRef,fvExtArea];

type PRecPTGS = ^TRecPTGS;
     TRecPTGS = packed record
     Size: word;
     PTGS: PByteArray;
     end;

type TFormulaValue = record
     vString:  AxUCString;
     case ValType: TFormulaValType of
       fvFloat   : (vFloat:   double);
       fvBoolean : (vBoolean: boolean);
       fvError   : (vError:   TXc12CellError);
       fvString  : (xvString:  boolean);
       fvRef     : (vRef:     array[0..1] of word);
       fvArea    : (vArea:    array[0..3] of word);
       fvExtRef  : (vExtRef:  array[0..2] of word);
       fvExtArea : (vExtArea: array[0..4] of word);
     end;

type TDynByteArray = array of byte;

type PLongWordArray = ^TLongWordArray;
     TLongWordArray = array[0..8191] of Longword;

type TByte8Array = array[0..7] of byte;

type TFormulaErrorEvent = procedure(Sender: TObject; ErrorId: integer; ErrorStr: AxUCString) of object;
type TFunctionEvent     = procedure(Sender: TObject; FuncName: AxUCString; Args: array of TFormulaValue; var Result: TFormulaValue) of object;

type TFormulaNameType = (ntName,ntExternName,ntExternSheet,ntCurrBook,ntCellValue);

function  ClipAreaToSheet(var C1,R1,C2,R2: integer): boolean;
procedure FVClear(var FV: TFormulaValue);
procedure FVSetNull(var FV: TFormulaValue);
procedure FVSetFloat(var FV: TFormulaValue; Value: double);
procedure FVSetBoolean(var FV: TFormulaValue; Value: boolean);
procedure FVSetError(var FV: TFormulaValue; Value: TXc12CellError);
procedure FVSetString(var FV: TFormulaValue; Value: AxUCString);
procedure FVSetRef(var FV: TFormulaValue; Col,Row: word);
procedure FVSetArea(var FV: TFormulaValue; Col1,Row1,Col2,Row2: word);
procedure FVSetXRef(var FV: TFormulaValue; Col,Row,Sheet: word);
procedure FVSetXArea(var FV: TFormulaValue; Col1,Row1,Col2,Row2,Sheet: word);
function  FVGetFloat(FV: TFormulaValue): double;
function  FVGetString(FV: TFormulaValue): AxUCString;
function  FVGetBoolean(FV: TFormulaValue): boolean;
function  FVGetVariant(FV: TFormulaValue): Variant;
function  FVCompare(FV1,FV2: TFormulaValue; var Res: double): boolean;
function  FVSize(FV: TFormulaValue): integer;

function  RGBToClosestIndexColor(Color: longword): TXc12IndexColor;
function  ValidArea(C1,R1,C2,R2: integer): boolean;
function  ValidRef(C,R: integer): boolean;
function  CellErrorToErrorCode(Error: TXc12CellError): byte;
function  DecodeRK(Value: longint): double;
function  Bit8StrToWideString(P: PByteArray; CharCount: integer): AxUCString;
function  ErrorCodeToCellError(Code: integer): TXc12CellError;
function  XColorToTColor(XC: TXc12IndexColor): TColor;
function  XColorToRGB(XC: TXc12IndexColor): longword;
function  IntToXColor(Value: word): TXc12IndexColor;
function  BufUnicodeZToWS(Buf: PByteArray; Len: integer): AxUCString;
function  UCStringLenFile(S: AxUCString): integer;
function  CompressedStringLen(S: XLS8String): integer;
function  ByteToWideString(P: PByteArray; CharCount: integer): AxUCString;
function  CompressedStrToWideString(S: XLS8String): AxUCString;
function  ByteToCompressedStr(P: PByteArray; CharCount: integer): XLS8String;
function  ErrorCodeToText(Code: integer): AxUCString;
function  MyWideUppercase(S: AxUCString): AxUCString;
procedure WideStringToByteStr(S: AxUCString; P: PByteArray);
function  ByteStrToWideString(P: PByteArray; CharCount: integer): AxUCString;

implementation

function RGBToClosestIndexColor(Color: longword): TXc12IndexColor;
var
  i,j: integer;
  C: integer;
  R1,G1,B1: byte;
  R2,G2,B2: byte;
  V1,V2: double;
begin
  j := 8;
  R1 := Color and $FF;
  G1 := (Color and $FF00) shr 8;
  B1 := (Color and $FF0000) shr 16;
  V1 := $FFFFFF;
  for i := 8 to 63 do begin
    C := Xc12IndexColorPalette[i];
    R2 := C and $FF;
    G2 := (C and $FF00) shr 8;
    B2 := (C and $FF0000) shr 16;
    V2 := Abs(R1 - R2) + Abs(G1 - G2) + Abs(B1 - B2);
    if Abs(V2) < Abs(V1) then begin
      V1 := V2;
      j := i;
    end;
  end;
  Result := TXc12IndexColor(j);
end;

function  ValidRef(C,R: integer): boolean;
begin
  Result := (C >= 0) and (C <= MAXCOL) and (R >= 0) and (R <= MAXROW);
end;

function  ValidArea(C1,R1,C2,R2: integer): boolean;
begin
  Result := ValidRef(C1,R1) and ValidRef(C2,R2);
end;

function CellErrorToErrorCode(Error: TXc12CellError): byte;
begin
  case Error of
    errUnknown: Result := $2A;
    errNull:  Result := $00;
    errDiv0:  Result := $07;
    errValue: Result := $0F;
    errRef:   Result := $17;
    errName:  Result := $1D;
    errNum:   Result := $24;
    errNA:    Result := $2A;
    else
      Result := $2A;
  end;
end;

procedure WideStringToByteStr(S: AxUCString; P: PByteArray);
begin
  P[0] := 1;
  Move(Pointer(S)^,P[1],Length(S) * 2);
end;

function DecodeRK(Value: longint): double;
var
  RK: TRK;
begin
  RK.DW[0] := 0;
//  RK.DW[1] := Value and $FFFFFFFC;
  RK.DW[1] := Value and LongInt($FFFFFFFC);
  case (Value and $3) of
    0: Result := RK.V;
    1: Result := RK.V / 100;
    2: Result := Integer(RK.DW[1]) / 4;
    3: Result := Integer(RK.DW[1]) / 400;
    else
      Result := RK.V;
  end;
end;

function Bit8StrToWideString(P: PByteArray; CharCount: integer): AxUCString;
var
  i: integer;
begin
  SetLength(Result,CharCount);
  for i := 0 to CharCount - 1 do
    Result[i + 1] := AxUCChar(P[i]);
end;

function ErrorCodeToCellError(Code: integer): TXc12CellError;
var
  V: byte;
begin
  case Code of
    $00: V := 1;
    $07: V := 2;
    $0F: V := 3;
    $17: V := 4;
    $1D: V := 5;
    $24: V := 6;
    $2A: V := 7;
    else V := 0;
  end;
  Result := TXc12CellError(V);
end;

function XColorToTColor(XC: TXc12IndexColor): TColor;
begin
  Result := Xc12IndexColorPalette[Integer(XC)];
end;

function XColorToRGB(XC: TXc12IndexColor): longword;
var
  tmp: longword;
begin
  Result := XColorToTColor(XC);
  tmp := Result and $00FF0000;
  Result := Result + (((Result and $000000FF) shl 16) or tmp);
end;

function IntToXColor(Value: word): TXc12IndexColor;
begin
  if Value <= Word(High(TXc12IndexColor)) then
    Result := TXc12IndexColor(Value)
  else
    Result := xcAutomatic;
end;

function BufUnicodeZToWS(Buf: PByteArray; Len: integer): AxUCString;
var
  p: integer;
begin
  if Len > 0 then begin
    SetLength(Result,(Len div 2) - 1);
    Move(Buf^,Pointer(Result)^,Len - 2);
    p := CPos(#0,Result);
    if (p > 0) and (p < Len) then
      Result := Copy(Result,1,p - 1);
  end
  else
    Result := '';
end;

function ClipAreaToSheet(var C1,R1,C2,R2: integer): boolean;
begin
  if (C1 > MAXCOL) or (R1 > MAXROW) or (C2 < 0) or (R2 < 0) then
    Result := False
  else begin
    C1 := Max(C1,0);
    R1 := Max(R1,0);
    C2 := Min(C2,MAXCOL);
    R2 := Min(R2,MAXROW);
    Result := True;
  end;
end;

function UCStringLenFile(S: AxUCString): integer;
begin
  if S = '' then
    Result := 0
  else
    Result := Length(S) * 2 + 1;
end;

function CompressedStringLen(S: XLS8String): integer;
begin
  if S = '' then
    Result := 0
  else if S[1] = #0 then
    Result := Length(S) - 1
  else
    Result := (Length(S) - 1) div 2;
end;

function ByteToWideString(P: PByteArray; CharCount: integer): AxUCString;
var
  S: XLS8String;
begin
  if CharCount <= 0 then
    Result := ''
  else begin
    if P[0] = 1 then begin
      SetLength(S,((CharCount * 2) and $00FF) + 1);
      Move(P[0],Pointer(S)^,((CharCount * 2) and $00FF) + 1)
    end
    else begin
      SetLength(S,(CharCount and $00FF) + 1);
      Move(P[0],Pointer(S)^,(CharCount and $00FF) + 1);
    end;
    Result := CompressedStrToWideString(S);
  end;
end;

function CompressedStrToWideString(S: XLS8String): AxUCString;
begin
  if Length(S) <= 0 then
    Result := ''
  else begin
    if S[1] = #0 then begin
      Result := Copy(String(S),2,MAXINT);
    end
    else if S[1] = #1 then begin
      SetLength(Result,(Length(S) - 1) div 2);
      S := Copy(S,2,MAXINT);
      Move(Pointer(S)^,Pointer(Result)^,Length(S));
    end
    else
      Result := String(S);
//      raise XLSRWException.Create('Bad excel string id.');
  end;
end;

function ByteToCompressedStr(P: PByteArray; CharCount: integer): XLS8String;
begin
  if P[0] = 0 then begin
    SetLength(Result,CharCount + 1);
    Move(P[1],Result[2],CharCount);
    Result[1] := #0;
  end
  else begin
    SetLength(Result,CharCount * 2 + 1);
    Move(P[1],Result[2],CharCount * 2);
    Result[1] := #1;
  end;
end;


procedure FVClear(var FV: TFormulaValue);
begin
  FV.ValType := fvError;
  FV.vError := errUnknown;
end;

procedure FVSetNull(var FV: TFormulaValue);
begin
  FV.ValType := fvNull;
end;

procedure FVSetFloat(var FV: TFormulaValue; Value: double);
begin
  FV.ValType := fvFloat;
  FV.vFloat := Value;
end;

procedure FVSetBoolean(var FV: TFormulaValue; Value: boolean);
begin
  FV.ValType := fvBoolean;
  FV.vBoolean := Value;
end;

procedure FVSetError(var FV: TFormulaValue; Value: TXc12CellError);
begin
  FV.ValType := fvError;
  FV.vError := Value;
end;

procedure FVSetString(var FV: TFormulaValue; Value: AxUCString);
begin
  FV.ValType := fvString;
  FV.vString := Value;
end;

procedure FVSetRef(var FV: TFormulaValue; Col,Row: word);
begin
  FV.ValType := fvRef;
  FV.vRef[0] := Col;
  FV.vRef[1] := Row;
end;

procedure FVSetArea(var FV: TFormulaValue; Col1,Row1,Col2,Row2: word);
begin
  FV.ValType := fvArea;
  FV.vArea[0] := Col1;
  FV.vArea[1] := Row1;
  FV.vArea[2] := Col2;
  FV.vArea[3] := Row2;
end;

procedure FVSetXRef(var FV: TFormulaValue; Col,Row,Sheet: word);
begin
  FV.ValType := fvExtRef;
  FV.vExtRef[0] := Col;
  FV.vExtRef[1] := Row;
  FV.vExtRef[2] := Sheet;
end;

procedure FVSetXArea(var FV: TFormulaValue; Col1,Row1,Col2,Row2,Sheet: word);
begin
  FV.ValType := fvExtArea;
  FV.vExtArea[0] := Col1;
  FV.vExtArea[1] := Row1;
  FV.vExtArea[2] := Col2;
  FV.vExtArea[3] := Row2;
  FV.vExtArea[4] := Sheet;
end;

function FVGetFloat(FV: TFormulaValue): double;
begin
  if FV.ValType = fvFloat then
    Result := FV.vFloat
  else
    Result := 0;
end;

function  FVGetString(FV: TFormulaValue): AxUCString;
begin
  if FV.ValType = fvString then
    Result := FV.vString
  else
    Result := '';
end;

function FVGetBoolean(FV: TFormulaValue): boolean;
begin
  if FV.ValType = fvBoolean then
    Result := FV.vBoolean
  else
    Result := False;
end;

function  FVGetVariant(FV: TFormulaValue): Variant;
begin
  case FV.ValType of
    fvFloat    : Result := FV.vFloat;
    fvBoolean  : Result := FV.vBoolean;
    fvError    : Result := Xc12CellErrorNames[FV.vError];
    fvString   : Result := FV.vString;
    fvRef      : Result := ColRowToRefStr(FV.vRef[0],FV.vRef[1],False,False);
    fvArea     : Result := AreaToRefStr(FV.vArea[0],FV.vArea[1],FV.vArea[2],FV.vArea[3],False,False,False,False);
    // Sheet name is not included
    fvExtRef   : Result := ColRowToRefStr(FV.vExtRef[0],FV.vExtRef[1],False,False);
    fvExtArea  : Result := AreaToRefStr(FV.vExtArea[0],FV.vExtArea[1],FV.vExtArea[2],FV.vExtArea[3],False,False,False,False);
  end;
end;

function FVCompare(FV1,FV2: TFormulaValue; var Res: double): boolean;
begin
  Result := (FV1.ValType <> fvError) and (FV2.ValType <> fvError) and (FV1.ValType = FV2.ValType);
  if not Result then
    Exit;
  if FV1.ValType in TFormulaValTypeRef then
    raise XLSRWException.Create('Illegal value in comparision');
  case FV1.ValType of
    fvFloat   : Res := FV1.vFloat - FV2.vFloat;
    fvBoolean : Res := Integer(FV1.vBoolean = FV2.vBoolean);
{$ifdef DELPHI_5}
    fvString  : Res := CompareStr(MyWideUppercase(FV1.vString),MyWideUppercase(FV2.vString));
{$else}
    fvString  : Res := WideCompareStr(MyWideUppercase(FV1.vString),MyWideUppercase(FV2.vString));
{$endif}
  end;
end;

function FVSize(FV: TFormulaValue): integer;
begin
  case FV.ValType of
    fvArea,fvExtArea:
      Result := (FV.vArea[2] - FV.vArea[0] + 1) * (FV.vArea[3] - FV.vArea[1] + 1);
    else
      Result := 1;
  end;
end;

function ErrorCodeToText(Code: integer): AxUCString;
begin
  case Code of
    $00: Result := Xc12CellErrorNames[errNull];
    $07: Result := Xc12CellErrorNames[errDiv0];
    $0F: Result := Xc12CellErrorNames[errValue];
    $17: Result := Xc12CellErrorNames[errRef];
    $1D: Result := Xc12CellErrorNames[errName];
    $24: Result := Xc12CellErrorNames[errNum];
    $2A: Result := Xc12CellErrorNames[errNA];
    else Result := '#???';
  end;
end;

function MyWideUppercase(S: AxUCString): AxUCString;
begin
{$ifdef DELPHI_5}
  Result := Uppercase(S);
{$else}
  Result := WideUppercase(S);
{$endif}
end;

function ByteStrToWideString(P: PByteArray; CharCount: integer): AxUCString;
var
  i: integer;
begin
  SetLength(Result,CharCount);
  if CharCount <= 0  then
    Exit
  else if P[0] = 1 then begin
    P := PByteArray(NativeInt(P) + 1);
    Move(P^,Pointer(Result)^,CharCount * 2);
  end
  else begin
    for i := 1 to CharCount do
      Result[i] := AxUCChar(P[i]);
  end;
end;

end.
