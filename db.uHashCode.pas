unit db.uHashCode;

interface

uses IdHashMessageDigest, Classes, System.SysUtils, IdHashCRC, IdHashSHA, IdSSLOpenSSLHeaders;

function GetFile_MD5(const iFileName: String): String;  // 获取文件 MD5
function GetString_MD5(const strValue: String): String; // 获取字符串 MD5

function GetFile_CRC(const iFileName: String): String;  // 获取文件 CRC
function GetString_CRC(const strValue: String): String; // 获取字符串 CRC

function GetFile_SHA1(const iFileName: String): String;  // 获取文件 SHA1
function GetString_SHA1(const strValue: String): String; // 获取字符串 SHA1

function GetFile_SHA256(const iFileName: String): String;  // 获取文件 SHA256
function GetString_SHA256(const strValue: String): String; // 获取字符串 SHA256

function GetFile_SHA384(const iFileName: String): String;  // 获取文件 SHA384
function GetString_SHA384(const strValue: String): String; // 获取字符串 SHA384

function GetFile_SHA512(const iFileName: String): String;  // 获取文件 SHA512
function GetString_SHA512(const strValue: String): String; // 获取字符串 SHA512

implementation

var
  FOpenSSLLoad: Boolean = False;

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

{ 获取文件 MD5 }
function GetFile_MD5(const iFileName: String): String;
var
  MemSteam: TMemoryStream;
begin
  MemSteam := TMemoryStream.Create;
  Try
    MemSteam.LoadFromFile(iFileName);
    Result := GetMemoryMD5(MemSteam);
  Finally
    MemSteam.Free;
  End;
end;

{ 获取字符串 MD5 }
function GetString_MD5(const strValue: String): String;
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

function GetMemoryCRC(const MemStream: TMemoryStream): String;
var
  crc: TIdHashCRC32;
begin
  crc := TIdHashCRC32.Create;
  try
    Result := crc.HashStreamAsHex(MemStream);
  finally
    crc.Free;
  end;
end;

{ 获取文件 CRC }
function GetFile_CRC(const iFileName: String): String;
var
  MemStream: TMemoryStream;
begin
  MemStream := TMemoryStream.Create;
  try
    MemStream.LoadFromFile(iFileName);
    Result := GetMemoryCRC(MemStream);
  finally
    MemStream.Free;
  end;
end;

{ 获取字符串 CRC }
function GetString_CRC(const strValue: String): String;
var
  MemStream: TMemoryStream;
  B        : TBytes;
  n        : Integer;
begin
  B := BytesOf(strValue);
  n := Length(B);

  MemStream := TMemoryStream.Create;
  try
    MemStream.SetSize(n);
    MemStream.Position := 0;
    MemStream.Write(B[0], n);
    MemStream.Position := 0;
    Result             := GetMemoryCRC(MemStream);
  finally
    MemStream.Free;
    SetLength(B, 0);
  end;
end;

function GetMemorySHA1(const MemStream: TMemoryStream): String;
var
  SHA1: TIdHashSHA1;
begin
  SHA1 := TIdHashSHA1.Create;
  try
    Result := SHA1.HashStreamAsHex(MemStream);
  finally
    SHA1.Free;
  end;
end;

{ 获取文件 SHA1 }
function GetFile_SHA1(const iFileName: String): String;
var
  MemStream: TMemoryStream;
begin
  MemStream := TMemoryStream.Create;
  try
    MemStream.LoadFromFile(iFileName);
    Result := GetMemorySHA1(MemStream);
  finally
    MemStream.Free;
  end;
end;

{ 获取字符串 SHA1 }
function GetString_SHA1(const strValue: String): String;
var
  MemStream: TMemoryStream;
  B        : TBytes;
  n        : Integer;
begin
  B := BytesOf(strValue);
  n := Length(B);

  MemStream := TMemoryStream.Create;
  try
    MemStream.SetSize(n);
    MemStream.Position := 0;
    MemStream.Write(B[0], n);
    MemStream.Position := 0;
    Result             := GetMemorySHA1(MemStream);
  finally
    MemStream.Free;
    SetLength(B, 0);
  end;
end;

function GetMemorySHA256(const MemStream: TMemoryStream): String;
var
  SHA256: TIdHashSHA256;
begin
  if FOpenSSLLoad then
  begin
    SHA256 := TIdHashSHA256.Create;
    try
      Result := SHA256.HashStreamAsHex(MemStream);
    finally
      SHA256.Free;
    end;
  end
  else
  begin
    Result := 'Error';
  end;
end;

{ 获取文件 SHA256 }
function GetFile_SHA256(const iFileName: String): String;
var
  MemStream: TMemoryStream;
begin
  MemStream := TMemoryStream.Create;
  try
    MemStream.LoadFromFile(iFileName);
    Result := GetMemorySHA256(MemStream);
  finally
    MemStream.Free;
  end;
end;

{ 获取字符串 SHA256 }
function GetString_SHA256(const strValue: String): String;
var
  MemStream: TMemoryStream;
  B        : TBytes;
  n        : Integer;
begin
  B := BytesOf(strValue);
  n := Length(B);

  MemStream := TMemoryStream.Create;
  try
    MemStream.SetSize(n);
    MemStream.Position := 0;
    MemStream.Write(B[0], n);
    MemStream.Position := 0;
    Result             := GetMemorySHA256(MemStream);
  finally
    MemStream.Free;
    SetLength(B, 0);
  end;
end;

function GetMemorySHA384(const MemStream: TMemoryStream): String;
var
  SHA384: TIdHashSHA384;
begin
  if FOpenSSLLoad then
  begin
    SHA384 := TIdHashSHA384.Create;
    try
      Result := SHA384.HashStreamAsHex(MemStream);
    finally
      SHA384.Free;
    end;
  end
  else
  begin
    Result := 'Error';
  end;
end;

{ 获取文件 SHA384 }
function GetFile_SHA384(const iFileName: String): String;
var
  MemStream: TMemoryStream;
begin
  MemStream := TMemoryStream.Create;
  try
    MemStream.LoadFromFile(iFileName);
    Result := GetMemorySHA384(MemStream);
  finally
    MemStream.Free;
  end;
end;

{ 获取字符串 SHA384 }
function GetString_SHA384(const strValue: String): String;
var
  MemStream: TMemoryStream;
  B        : TBytes;
  n        : Integer;
begin
  B := BytesOf(strValue);
  n := Length(B);

  MemStream := TMemoryStream.Create;
  try
    MemStream.SetSize(n);
    MemStream.Position := 0;
    MemStream.Write(B[0], n);
    MemStream.Position := 0;
    Result             := GetMemorySHA384(MemStream);
  finally
    MemStream.Free;
    SetLength(B, 0);
  end;
end;

function GetMemorySHA512(const MemStream: TMemoryStream): String;
var
  SHA512: TIdHashSHA512;
begin
  if FOpenSSLLoad then
  begin
    SHA512 := TIdHashSHA512.Create;
    try
      Result := SHA512.HashStreamAsHex(MemStream);
    finally
      SHA512.Free;
    end;
  end
  else
  begin
    Result := 'Error';
  end;
end;

{ 获取文件 SHA512 }
function GetFile_SHA512(const iFileName: String): String;
var
  MemStream: TMemoryStream;
begin
  MemStream := TMemoryStream.Create;
  try
    MemStream.LoadFromFile(iFileName);
    Result := GetMemorySHA512(MemStream);
  finally
    MemStream.Free;
  end;
end;

{ 获取字符串 SHA512 }
function GetString_SHA512(const strValue: String): String;
var
  MemStream: TMemoryStream;
  B        : TBytes;
  n        : Integer;
begin
  B := BytesOf(strValue);
  n := Length(B);

  MemStream := TMemoryStream.Create;
  try
    MemStream.SetSize(n);
    MemStream.Position := 0;
    MemStream.Write(B[0], n);
    MemStream.Position := 0;
    Result             := GetMemorySHA512(MemStream);
  finally
    MemStream.Free;
    SetLength(B, 0);
  end;
end;

initialization
  FOpenSSLLoad := IdSSLOpenSSLHeaders.Load();

end.
