library EncPass;
{$IF CompilerVersion >= 21.0}
{$WEAKLINKRTTI ON}
{$RTTI EXPLICIT METHODS([]) PROPERTIES([]) FIELDS([])}
{$IFEND}

uses
  IdHashMessageDigest,
  System.SysUtils,
  System.Classes;

{$R *.res}

function GetMemoryMD5(const MemStream: TMemoryStream): String;
var
  MyMD5: TIdHashMessageDigest5;
begin
  MyMD5 := TIdHashMessageDigest5.Create;
  Try
    Result := MyMD5.HashStreamAsHex(MemStream);
  Finally
    MyMD5.Free;
  End;
end;

function GetString_MD5(const strValue: String): String; // »ñÈ¡×Ö·û´® MD5
var
  MemSteam: TMemoryStream;
  B       : TBytes;
  n       : Integer;
begin
  B := BytesOf(strValue);
  n := Length(B);

  MemSteam := TMemoryStream.Create;
  Try
    MemSteam.SetSize(n);
    MemSteam.Position := 0;
    MemSteam.Write(B[0], n);
    MemSteam.Position := 0;
    Result            := GetMemoryMD5(MemSteam);
  Finally
    MemSteam.Free;
    SetLength(B, 0);
  End;
end;

function EncPassword(const strPassword: PAnsiChar): PAnsiChar; stdcall;
begin
  Result := PAnsiChar(PChar(GetString_MD5(string(PChar(strPassword)))));
end;

exports
  EncPassword;

begin

end.
