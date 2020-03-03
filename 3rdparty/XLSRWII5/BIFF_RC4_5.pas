unit BIFF_RC4_5;

interface

uses Classes, SysUtils;

type TRC4Key = record
     State: array[0..255] of byte;
     x,y: byte;
     end;

procedure RC4PrepareKey(KeyData: PByteArray; DataLen: integer; var Key: TRC4Key);
procedure RC4(Buffer: PByteArray; BuffLen: integer; var Key: TRC4Key);

implementation

procedure SwapByte(var A,B: byte);
var
  T: byte;
begin
  T := A;
  A := B;
  B := T;
end;

procedure RC4PrepareKey(KeyData: PByteArray; DataLen: integer; var Key: TRC4Key);
var
  Index1: byte;
  Index2: byte;
  State: PByteArray;
  Counter: longword;
begin
  State := @Key.state;
  for Counter := 0 to 255 do
    State[Counter] := Counter;
  Key.X := 0;
  Key.Y := 0;
  Index1 := 0;
  Index2 := 0;
  for Counter := 0 to 255 do begin
    Index2 := (KeyData[Index1] + State[counter] + Index2) mod 256;
    SwapByte(State[Counter],State[Index2]);
    Index1 := (Index1 + 1) mod DataLen;
  end;
end;

procedure RC4(Buffer: PByteArray; BuffLen: integer; var Key: TRC4Key);
var
  X: byte;
  Y: byte;
  State: PByteArray;
  XorIndex: byte;
  Counter: longword;
begin
  if BuffLen <= 0 then
    Exit;
  X := Key.X;
  Y := Key.Y;

  State := @Key.state;
  for Counter := 0 to BuffLen - 1 do begin
    X := (X + 1) mod 256;
    Y := (State[X] + Y) mod 256;
    SwapByte(State[X],State[Y]);

    XorIndex := (State[X] + State[Y]) mod 256;

    Buffer[Counter] := Buffer[Counter] xor State[XorIndex];
  end;
  Key.X := X;
  Key.Y := Y;
end;

end.
