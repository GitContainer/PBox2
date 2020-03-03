unit XLSEncryption5;

{
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

// https://www.lyquidity.com/devblog/?p=35
// https://github.com/magnumripper/JohnTheRipper

uses Classes, SysUtils, Windows, ActiveX, ComObj, Math,

     xpgPSimpleDOM,
     XLSUtils5,
     XLSDCPsha1, XLSDCPsha512,
     XLSSynCrypto;

type TXLSCryptoResult = (xcrUnknown,xcrOk,xcrMissingPassword,xcrWrongPassword,xcrUnsupportedEncryption);

type PXLSRecEncVersion = ^TXLSRecEncVersion;
     TXLSRecEncVersion = packed record
     Major: word;
     Minor: word;
     end;

type TXLSCryptoHash = (xchSHA1,xchSHA512);
type TXLSCryptoCiper = (xccAESECB,xccAESCBC);

type PXLSRecEncHeader = ^TXLSRecEncHeader;
     TXLSRecEncHeader = packed record
     Flags       : longword;
     SizeExtra   : longword;
     AlgId       : longword;
     AlgIDHash   : longword;
     KeySize     : longword;
     ProviderType: longword;
     Reserved1   : longword;
     Reserved2   : longword;
//     CSPName     : string variable
     end;

type PXLSRecEncVerifier = ^TXLSRecEncVerifier;
     TXLSRecEncVerifier = packed record
     SaltSize             : longword;
     Salt                 : array[0..15] of byte;
     EncryptedVerifier    : array[0..15] of byte;
     VerifierHashSize     : longword;
     EncryptedVerifierHash: array[0..31] of byte;
     end;

type TXLSVector = class(TObject)
private
     function  GetAsInteger(AIndex: integer): integer;
     procedure SetAsInteger(AIndex: integer; const Value: integer);
     function  GetAsString: AxUCString;
protected
     FBytes: PByteArray;
     FSize : integer;
public
     constructor Create(ASize,AValue: integer); overload;
     constructor Create(AVector: TXLSVector); overload;

     destructor Destroy; override;

     procedure Clear;

     procedure Resize(ASize: integer); overload;
     procedure Resize(ASize,AValue: integer); overload;

     procedure Assign(AArray: array of byte); overload;
     procedure Assign(AStr: AnsiString); overload;
     procedure Assign(AVector: TXLSVector); overload;

     procedure CopyFrom(ABytes: PByteArray; ACount,AOffset: integer); overload;
     procedure CopyFrom(AVector: TXLSVector; AOffset: integer); overload;
     procedure CopyFrom(AVector: TXLSVector;ACount, AOffset: integer); overload;

     function  Equal(AVector: TXLSVector; ACompareSize: integer = 0): boolean;

     property Bytes: PByteArray read FBytes;
     property Size : integer read FSize;

     property AsInteger[AIndex: integer]: integer read GetAsInteger write SetAsInteger;
     property AsString: AxUCString read GetAsString;
     end;

type TXLSCrypto = class(TObject)
protected
     FSaltSize : integer;
     FSpinCount: integer;
     FKeyBits  : integer;
     FHashSize : integer;
     FPassword : AxUCString;

     procedure FillRandom(AArray: PByteArray; ASize: integer);
public
     constructor Create(APassword: AxUCString);
     destructor Destroy; override;

     function  InitiateRead(ABuf: PByteArray; ABufSize: integer): TXLSCryptoResult; virtual; abstract;
     function  InitiateWrite(AStream : IStream): TXLSCryptoResult; virtual; abstract;
     function  DecryptDocument(AInStream : IStream; AOutStream: TStream): boolean; virtual; abstract;
     function  EncryptDocument(AInStream : TStream; AOutStream: IStream): boolean; virtual; abstract;
     end;

type TXLSCryptoStandard = class(TXLSCrypto)
protected
     FHeader  : TXLSRecEncHeader;
     FVerifier: TXLSRecEncVerifier;
     FKey     : array[0..15] of byte;

     procedure CalcKey;
     function  VerifyPassword: boolean;
     procedure EncryptVerifier(AVeriVal: PByteArray);
public
     constructor Create(APassword: AxUCString);

     function  InitiateRead(ABuf: PByteArray; ABufSize: integer): TXLSCryptoResult; override;
     function  InitiateWrite(AStream : IStream): TXLSCryptoResult; override;
     function  DecryptDocument(AInStream : IStream; AOutStream: TStream): boolean; override;
     function  EncryptDocument(AInStream : TStream; AOutStream: IStream): boolean; override;
     function  EncryptDocumentFile(AInStream,AOutStream: TStream): boolean;
     end;

type TXLSCryptoAgile = class(TXLSCrypto)
protected
     FSpinCount           : integer;
     FBlockSize           : integer;
     FKeyBits             : integer;
     FHashSize            : integer;
     FCipherAlgorithm     : TXLSCryptoCiper;
     FHashAlgorithm       : TXLSCryptoHash;

     FKeyDataSalt               : TXLSVector;

     FSaltValue                 : TXLSVector;
     FEncryptedVerifierHashInput: TXLSVector;
     FEncryptedVerifierHashValue: TXLSVector;
     FEncryptedKeyValue         : TXLSVector;

     FKey                       : TXLSVector;

     function  ReadXML(ABuf: PByteArray; ASize: integer): boolean;

     procedure SetupSHA512;

     procedure SetTestSHA512;

     function  HashCalc(output,input: TXLSVector; Algorithm: TXLSCryptoHash): boolean;

	   function  CalculateBlock(rBlock: PByteArray; aBlockSize: integer; rHashFinal: TXLSVector; rInput: TXLSVector; rOutput: TXLSVector): boolean;
     function  CalculateHashFinal(rPassword: AxUCString; aHashFinal: TXLSVector): boolean;
     function  GenerateEncryptionKey(rPassword: AxUCString): boolean;
public
     constructor Create(APassword: AxUCString);
     destructor Destroy; override;

     function  InitiateRead(ABuf: PByteArray; ABufSize: integer): TXLSCryptoResult; override;
     function  InitiateWrite(AStream : IStream): TXLSCryptoResult; override;
     function  DecryptDocument(AInStream : IStream; AOutStream: TStream): boolean; override;
     function  EncryptDocument(AInStream : TStream; AOutStream: IStream): boolean; override;

     function  Test: boolean;
     end;

type TXLSEncryption = class(TObject)
private
     procedure SetPassword(const Value: AxUCString);
protected
     FOleStorage: IStorage;
     FLockBytes : ILockBytes;
     FPassword  : AxUCString;
     FPasswEvent: TXLSPasswordEvent;
     FHeader    : TXLSRecEncHeader;

     FCrypto    : TXLSCrypto;

     FResult    : TXLSCryptoResult;

     FVerifier  : TXLSRecEncVerifier;
     FKey       : array[0..15] of byte;
     FWrongPassw: boolean;
     FInStream  : TStream;
     FOutStream : TMemoryStream;

     function  DoEncryptionInfo(var ABuf: PByteArray; ASize: integer): TXLSCrypto;
     procedure WriteStorageToDestStream(AStream: TStream);
public
     constructor Create;
     destructor Destroy; override;

     procedure Test;

     function LoadFromFile(const AFilename: string): boolean;
     function LoadFromStream(AStream: TStream): boolean;
     function SaveToStream(AStream: TStream): boolean;
     function SaveToStreamTest(AStream: TStream; ACrypto: TXLSCrypto): boolean;

     property Password     : AxUCString read FPassword write SetPassword;
     property CryptoResult : TXLSCryptoResult read FResult;
     property InStream     : TStream write FInStream;
     property OutStream    : TMemoryStream read FOutStream;

     property OnPassword   : TXLSPasswordEvent read FPasswEvent write FPasswEvent;
     end;

implementation

const EncryptedVerifierHashInputData: array[0..7] of byte = ($FE, $A7, $D2, $76, $3B, $4B, $9E, $79);
const EncryptedVerifierHashValue    : array[0..7] of byte = ($D7, $AA, $0F, $6D, $30, $61, $34, $4E);

const EncryptedKeyValueData         : array[0..7] of byte = ($14, $6e, $0b, $e7, $ab, $ac, $d0, $d6);


{ TXLSEncryption }

function FileGetTempName(const Prefix: string): string;
var
  TempPath, TempFile: string;
  R: Cardinal;
begin
  Result := '';
  R := GetTempPath(0, nil);
  SetLength(TempPath, R);
  R := GetTempPath(R, PChar(TempPath));
  if R <> 0 then
  begin
    SetLength(TempPath, StrLen(PChar(TempPath)));
    SetLength(TempFile, MAX_PATH);
    R := GetTempFileName(PChar(TempPath), PChar(Prefix), 0, PChar(TempFile));
    if R <> 0 then
    begin
      SetLength(TempFile, StrLen(PChar(TempFile)));
      Result := TempFile;
    end;
  end;
end;

constructor TXLSEncryption.Create;
begin
  FOutStream := TMemoryStream.Create;
end;

destructor TXLSEncryption.Destroy;
begin
  FOutStream.Free;

  if FCrypto <> Nil then
    FCrypto.Free;

  inherited;
end;

function TXLSEncryption.DoEncryptionInfo(var ABuf: PByteArray; ASize: integer): TXLSCrypto;
var
  Ver: PXLSRecEncVersion;
begin
  Ver := PXLSRecEncVersion(ABuf);

  ABuf := PByteArray(NativeInt(ABuf) + SizeOf(TXLSRecEncVersion));

  if ((Ver.Major in [$0003,$0004]) and (Ver.Minor = $0002)) then
    Result := TXLSCryptoStandard.Create(FPassword)
  else if ((Ver.Major = $0004) and (Ver.Minor = $0004)) then
    Result := TXLSCryptoAgile.Create(FPassword)
  else
    Result := Nil;
end;

function TXLSEncryption.LoadFromFile(const AFilename: string): boolean;
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmOpenRead + fmShareDenyNone);
  try
    Result := LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

function TXLSEncryption.LoadFromStream(AStream: TStream): boolean;
var
  DataHandle: HGLOBAL;
  Buffer,B  : PByteArray;
  EStatStg  : IEnumStatStg;
  Stat      : array[0..255] of TSTATSTG;
  StatStg   : TSTATSTG;
  Count     : LongInt;
  OleStream : IStream;
begin
  if (FPassword = '') and Assigned(FPasswEvent) then
    FPasswEvent(Self,FPassword);

  if FPassword = '' then begin
    FResult := xcrMissingPassword;
    Result := False;
    Exit;
  end;

  DataHandle := GlobalAlloc(GMEM_MOVEABLE,AStream.Size);
  Buffer := GlobalLock(DataHandle);
  try
    AStream.ReadBuffer(Buffer^, AStream.Size);
    AStream.Seek(0,soFromBeginning);

    OleCheck(CreateILockBytesOnHGlobal(DataHandle, True, FLockBytes));
    OleCheck(StgOpenStorageOnILockBytes(FLockBytes, Nil, STGM_READ or STGM_TRANSACTED or STGM_SHARE_DENY_NONE , Nil, 0, FOleStorage));
  finally
    GlobalUnlock(DataHandle);
  end;

  OleCheck(FOleStorage.EnumElements(0,Nil,0,EStatStg));
  EStatStg.Next(High(Stat),Stat,@Count);

  Result := FOleStorage.OpenStream('EncryptionInfo',Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,OleStream) = S_OK;
  if Result then begin
    OleStream.Stat(StatStg,0);
    GetMem(Buffer,StatStg.cbSize);
    try
      B := Buffer;

      OleStream.Read(B,StatStg.cbSize,@Count);
      FCrypto := DoEncryptionInfo(B,Count);
      if FCrypto <> Nil then begin
        B := PByteArray(NativeInt(B) + (NativeInt(B) - NativeInt(Buffer)));
        FResult := FCrypto.InitiateRead(B,Count - (NativeInt(B) - NativeInt(Buffer)));
        if FResult <> xcrOk then
          Exit;

        Result := FOleStorage.OpenStream('EncryptedPackage',Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,OleStream) = S_OK;
        if Result then begin
          Result := FCrypto.DecryptDocument(OleStream,FOutStream);

          FResult := xcrOk;
        end
        else
          FResult := xcrUnknown;
      end
      else
        FResult := xcrUnsupportedEncryption;
    finally
      FreeMem(Buffer);
    end;
  end;
end;

function TXLSEncryption.SaveToStream(AStream: TStream): boolean;
var
  HdrStream : IStream;
  DocStream : IStream;
  Crypto    : TXLSCrypto;
begin
  OleCheck(CreateILockBytesOnHGlobal(0, True, FLockBytes));
  OleCheck(StgCreateDocfileOnILockBytes(FLockBytes, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE, 0,FOleStorage));

  OleCheck(FOleStorage.CreateStream('EncryptionInfo',  STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,HdrStream));
  OleCheck(FOleStorage.CreateStream('EncryptedPackage',STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,DocStream));

  Crypto := TXLSCryptoStandard.Create(FPassword);
  try
    Crypto.InitiateWrite(HdrStream);

    Crypto.EncryptDocument(FInStream,DocStream)
  finally
    Crypto.Free;
  end;

  WriteStorageToDestStream(AStream);

  Result := True;
end;

function TXLSEncryption.SaveToStreamTest(AStream: TStream; ACrypto: TXLSCrypto): boolean;
var
  HdrStream : IStream;
  DocStream : IStream;
begin
  OleCheck(CreateILockBytesOnHGlobal(0, True, FLockBytes));
  OleCheck(StgCreateDocfileOnILockBytes(FLockBytes, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE, 0,FOleStorage));

  OleCheck(FOleStorage.CreateStream('EncryptionInfo',  STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,HdrStream));
  OleCheck(FOleStorage.CreateStream('EncryptedPackage',STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,DocStream));

  ACrypto.InitiateWrite(HdrStream);

  ACrypto.EncryptDocument(FInStream,DocStream);

  WriteStorageToDestStream(AStream);

  Result := True;
end;

procedure TXLSEncryption.SetPassword(const Value: AxUCString);
begin
  FPassword := Value;
end;

procedure TXLSEncryption.Test;
const
  Key: array[0..15] of byte = (161, 122, 23, 97, 23, 35, 245, 131, 202, 70, 181, 42, 226, 143, 131, 0);
var
  StorageIn : IStorage;
  StorageOut: IStorage;
  EStatStg  : IEnumStatStg;
  Stat      : array[0..255] of TSTATSTG;
  Count     : LongInt;
  Stream    : TFileStream;
  StreamIIn : IStream;
  StreamIOut: IStream;
  Crypto    : TXLSCryptoStandard;
begin
  OleCheck(StgOpenStorage('d:\xtemp\t1in.xlsx',Nil,STGM_READ or STGM_TRANSACTED or STGM_SHARE_DENY_NONE ,Nil,0,StorageIn));

  OleCheck(StorageIn.EnumElements(0,Nil,0,EStatStg));
  EStatStg.Next(High(Stat),Stat,@Count);

  OleCheck(StorageIn.OpenStream('EncryptedPackage',Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,StreamIIn));

  OleCheck(StgCreateDocfile('d:\xtemp\t1out.xlsx', STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE, 0, StorageOut));
  OleCheck(StorageIn.CopyTo(0,Nil,Nil,StorageOut));

  OleCheck(StorageOut.DestroyElement('EncryptedPackage'));

  OleCheck(StorageOut.CreateStream('EncryptedPackage',STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,StreamIOut));

  Crypto := TXLSCryptoStandard.Create('storsugga');
//  Stream := TMemoryStream.Create;
  Stream := TFileStream.Create('d:\xtemp\Deceypted.xlsx',fmCreate);
  try
    Move(Key,Crypto.FKey,16);

    Crypto.DecryptDocument(StreamIIn,Stream);
//    Stream.Seek(0,soFromBeginning);
//    Crypto.EncryptDocument(Stream,StreamIOut);
  finally
    Crypto.Free;
    Stream.Free;
  end;
end;

procedure TXLSEncryption.WriteStorageToDestStream(AStream: TStream);
var
  DataHandle: HGLOBAL;
  Buffer: Pointer;
begin
  OleCheck(FLockBytes.Flush);
  OleCheck(FOleStorage.Commit(STGC_DEFAULT));

  OleCheck(FLockBytes.Flush);
  OleCheck(GetHGlobalFromILockBytes(FLockBytes, DataHandle));
  Buffer := GlobalLock(DataHandle);
  try
    AStream.WriteBuffer(Buffer^, GlobalSize(DataHandle));
  finally
    GlobalUnlock(DataHandle);
  end;
end;

{ TXLSVector }

procedure TXLSVector.Assign(AArray: array of byte);
begin
  Resize(Length(AArray));

  Move(AArray[0],FBytes[0],Length(AArray));
end;

procedure TXLSVector.Assign(AVector: TXLSVector);
begin
  Resize(AVector.Size);

  Move(AVector.FBytes[0],FBytes[0],FSize);
end;

procedure TXLSVector.CopyFrom(AVector: TXLSVector; AOffset: integer);
begin
  if (AVector.Size + AOffset) > FSize then
    raise Exception.Create('Vector to long');

  Move(AVector.FBytes[0],FBytes[AOffset],AVector.Size);
end;

procedure TXLSVector.Assign(AStr: AnsiString);
begin
  Resize(Length(AStr));

  Move(Pointer(AStr)^,FBytes[0],FSize);
end;

procedure TXLSVector.Clear;
begin
  Resize(0,0);
end;

procedure TXLSVector.CopyFrom(ABytes: PByteArray; ACount,AOffset: integer);
begin
  if (ACount + AOffset) > FSize then
    raise Exception.Create('Vector to long');

  Move(ABytes[0],FBytes[AOffset],ACount);
end;

constructor TXLSVector.Create(AVector: TXLSVector);
begin
  Assign(AVector);
end;

constructor TXLSVector.Create(ASize, AValue: integer);
begin
  Resize(ASize,AValue);
end;

destructor TXLSVector.Destroy;
begin
  FreeMem(FBytes);

  inherited;
end;

function TXLSVector.Equal(AVector: TXLSVector; ACompareSize: integer = 0): boolean;
begin
  if ACompareSize > 0 then begin
    Result := (FSize >= ACompareSize) and (AVector.Size >= ACompareSize);

    if Result then
      Result := CompareMem(@FBytes[0],@AVector.FBytes[0],ACompareSize);
  end
  else begin
    Result := FSize = AVector.Size;

    if Result then
      Result := CompareMem(@FBytes[0],@AVector.FBytes[0],FSize);
  end;
end;

function TXLSVector.GetAsInteger(AIndex: integer): integer;
begin
  Result := PIntegerArray(FBytes)[AIndex];
end;

function TXLSVector.GetAsString: AxUCString;
var
  i: integer;
begin
  Result := '';
  for i := 0 to FSize - 1 do
    Result := Result + IntToStr(Bytes[i]) + ', ';

  Result := Copy(Result,1,Length(Result) - 2);
end;

procedure TXLSVector.Resize(ASize, AValue: integer);
begin
  FSize := ASize;

  ReAllocMem(FBytes,FSize);

  FillChar(FBytes[0],FSize,AValue);
end;

procedure TXLSVector.SetAsInteger(Aindex: integer; const Value: integer);
begin
  PIntegerArray(FBytes)[AIndex] := Value;
end;

procedure TXLSVector.Resize(ASize: integer);
begin
  Resize(ASize,0);
end;

procedure TXLSVector.CopyFrom(AVector: TXLSVector; ACount, AOffset: integer);
begin
  if (ACount + AOffset) > FSize then
    raise Exception.Create('Vector to long');

  Move(AVector.FBytes[0],FBytes[AOffset],ACount);
end;

{ TXLSCryptoAgile }

function TXLSCryptoAgile.CalculateBlock(rBlock: PByteArray; aBlockSize: integer; rHashFinal, rInput, rOutput: TXLSVector): boolean;
var
  keySize  : integer;
  hash     : TXLSVector;
  dataFinal: TXLSVector;
  key      : TXLSVector;
  AES      : TAESCBC;
  IV       : TAESBlock;
begin
  hash := TXLSVector.Create(FHashSize, 0);
	dataFinal := TXLSVector.Create(FHashSize + aBlockSize, 0);

  dataFinal.CopyFrom(rHashFinal,0);
  dataFinal.CopyFrom(rBlock,aBlockSize,FHashSize);

	hashCalc(hash, dataFinal, FHashAlgorithm);

	keySize := FKeyBits div 8;
	key := TXLSVector.Create(keySize, 0);

  Key.CopyFrom(hash,keySize,0);

  AES := TAESCBC.Create(Key.Bytes[0],Key.Size * 8);
  try
    Move(FSaltValue.Bytes[0],IV,FSaltValue.Size);

    AES.IV := IV;

    AES.Decrypt(rInput.Bytes,rOutput.Bytes,rInput.Size);
  finally
    AES.Free;
  end;

  hash.Free;
  dataFinal.Free;
  key.Free;

	Result := True;
end;

function TXLSCryptoAgile.CalculateHashFinal(rPassword: AxUCString; aHashFinal: TXLSVector): boolean;
var
  i                 : integer;
  passwordByteLength: integer;
  passwordByteArray : PByteArray;
  initialData       : TXLSVector;
  hash              : TXLSVector;
  data              : TXLSVector;
begin
	passwordByteLength := Length(rPassword) * 2;

	initialData := TXLSVector.Create(FSaltSize + passwordByteLength,0);
  initialData.CopyFrom(FSaltValue,0);

	passwordByteArray := @rPassword[1];

  initialData.CopyFrom(passwordByteArray,passwordByteLength,FSaltSize);

	hash := TXLSVector.Create(FHashSize, 0);

	hashCalc(hash, initialData, FHashAlgorithm);

	data := TXLSVector.Create(FHashSize + 4, 0);

	for i := 0 to FSpinCOunt - 1 do begin
    data.AsInteger[0] := i;

    data.CopyFrom(hash,4);
		hashCalc(hash, data, FHashAlgorithm);
  end;

  aHashFinal.Assign(hash);

  initialData.Free;
  hash.Free;
  data.Free;

	Result := True;
end;

constructor TXLSCryptoAgile.Create(APassword: AxUCString);
begin
  inherited Create(APassword);

  FKeyDataSalt := TXLSVector.Create(0,0);

  FSaltValue := TXLSVector.Create(0,0);
  FEncryptedVerifierHashInput := TXLSVector.Create(0,0);
  FEncryptedVerifierHashValue := TXLSVector.Create(0,0);
  FEncryptedKeyValue := TXLSVector.Create(0,0);

  FKey := TXLSVector.Create(0,0);

  SetupSHA512;
end;

// 648A19
function TXLSCryptoAgile.DecryptDocument(AInStream: IStream; AOutStream: TStream): boolean;
var
  totalSize       : int64;
  segment         : integer;
{$ifdef DELPHI_XE8_OR_LATER}
  P               : largeuint;
{$else}
  P               : largeint;
{$endif}
  Ptr             : Pointer;

  inBuf           : array[0..4095] of byte;
  outBuf          : array[0..4095] of byte;

  saltWithBlockKey: TXLSVector;
  hash            : TXLSVector;
  IV              : TAESBlock;

  AES             : TAESCBC;
begin
  AInStream.Read(@totalSize,4,@P);
  AInStream.Seek(4,STREAM_SEEK_CUR,P); // Reserved 4 Bytes

  segment := 0;

  saltWithBlockKey := TXLSVector.Create(FSaltSize + sizeof(segment), 0);
  saltWithBlockKey.CopyFrom(FKeyDataSalt,0);

  hash := TXLSVector.Create(FHashSize, 0);

  while (AInStream.Read(@inBuf[0],4096,@P) = S_OK) and (P > 0) do begin
    saltWithBlockKey.AsInteger[FSaltSize div 4] := segment;

    hashCalc(hash, saltWithBlockKey, xchSHA512);

    AES := TAESCBC.Create(FKey.Bytes[0],FKey.Size * 8);
    try
      Move(hash.Bytes[0],IV,FSaltSize);

      AES.IV := IV;

      AES.Decrypt(@inBuf[0],@outBuf[0],P);

      Ptr := @outBuf[0];
      AOutStream.Write(Ptr^,P);
    finally
      AES.Free;
    end;

    Inc(segment);
  end;

  AOutStream.Seek(0,soBeginning);

  saltWithBlockKey.Free;
  hash.Free;

  Result := True;
end;

destructor TXLSCryptoAgile.Destroy;
begin
  FKeyDataSalt.Free;

  FSaltValue.Free;
  FEncryptedVerifierHashInput.Free;
  FEncryptedVerifierHashValue.Free;
  FEncryptedKeyValue.Free;

  FKey.Free;

  inherited;
end;

function TXLSCryptoAgile.EncryptDocument(AInStream: TStream; AOutStream: IStream): boolean;
begin
  Result := False;
end;

function TXLSCryptoAgile.GenerateEncryptionKey(rPassword: AxUCString): boolean;
var
  hashFinal         : TXLSVector;
  hashInput         : TXLSVector;
  hashValue         : TXLSVector;
  hash              : TXLSVector;
begin
	FKey.Clear;
	FKey.Resize(FKeyBits div 8, 0);

  hashFinal := TXLSVector.Create(FHashSize, 0);
	calculateHashFinal(rPassword, hashFinal);

	hashInput := TXLSVector.Create(FSaltSize, 0);
	calculateBlock(@EncryptedVerifierHashInputData[0], sizeof(EncryptedVerifierHashInputData), hashFinal, FEncryptedVerifierHashInput, hashInput);

	hashValue := TXLSVector.Create(FEncryptedVerifierHashValue.Size, 0);
	calculateBlock(@EncryptedVerifierHashValue[0], sizeof(EncryptedVerifierHashValue), hashFinal, FEncryptedVerifierHashValue, hashValue);

	hash := TXLSVector.Create(FHashSize, 0);
	hashCalc(hash, hashInput, xchSHA512);

	if hash.Equal(hashValue,16) then begin
		calculateBlock(@EncryptedKeyValueData[0],SizeOf(EncryptedKeyValueData), hashFinal, FEncryptedKeyValue, FKey);

    Result := True;
  end
  else
    Result := False;

  hashFinal.Free;
  hashInput.Free;
  hashValue.Free;
  hash.Free;
end;

function TXLSCryptoAgile.HashCalc(output, input: TXLSVector; Algorithm: TXLSCryptoHash): boolean;
var
  SHA1  : TDCP_sha1;
  SHA512: TDCP_sha512;
  P     : PByteArray;
begin
  output.Clear;
  output.Resize(64, 0);


  case Algorithm of
    xchSHA1  : begin
      SHA1 := TDCP_sha1.Create;
      try
        SHA1.Init;
        SHA1.Update(input.Bytes[0],input.Size);
        P := output.Bytes;
        SHA1.Final(P[0]);
      finally
        SHA1.Free;
      end;
    end;
    xchSHA512: begin
      SHA512 := TDCP_sha512.Create;
      try
        SHA512.Init;
        SHA512.Update(input.Bytes[0],input.Size);
        P := output.Bytes;
        SHA512.Final(P[0]);
      finally
        SHA512.Free;
      end;
    end;
  end;

  Result := True;
end;

function TXLSCryptoAgile.InitiateRead(ABuf: PByteArray; ABufSize: integer): TXLSCryptoResult;
begin
  if not ReadXML(ABuf,ABufSize) then begin
    Result := xcrUnsupportedEncryption;

    Exit;
  end;

  if not GenerateEncryptionKey(FPassword) then
    Result := xcrWrongPassword
  else
    Result := xcrOk;
end;

function TXLSCryptoAgile.InitiateWrite(AStream: IStream): TXLSCryptoResult;
begin
  Result := xcrUnknown;
end;

// <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
// <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
//   <keyData saltSize="16"
//            blockSize="16"
//            keyBits="256"
//            hashSize="64"
//            cipherAlgorithm="AES"
//            cipherChaining="ChainingModeCBC"
//            hashAlgorithm="SHA512"
//            saltValue="toVT80MHZcBENQjcEla7+Q=="/>
//   <dataIntegrity encryptedHmacKey="al0y67uJS2e899wcyLW/86bC62JZseTc6BCpzScGMbMfx5CxVISNXI/jP/qLPtxTKeHSWU4bCCH67Gc+YjJ8vQ=="
//                  encryptedHmacValue="WniuYEOfQveSMS+1978NzWDkEpRGM3ZacnIE61ilICnpPpHgKVioguyyThjtwgq7HxsblsKgGQtOnn4F+HALOg=="/>
//   <keyEncryptors>
//     <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
//       <p:encryptedKey spinCount="100000"
//                       saltSize="16"
//                       blockSize="16"
//                       keyBits="256"
//                       hashSize="64"
//                       cipherAlgorithm="AES"
//                       cipherChaining="ChainingModeCBC"
//                       hashAlgorithm="SHA512"
//                       saltValue="mlyS6emPfVIo6ajHaIRcNQ=="
//                       encryptedVerifierHashInput="ZK9BSfhxMGJjg4Vmuccd+g=="
//                       encryptedVerifierHashValue="64EenNxAbeAeCNqq2U2/j/jAQBvqnDteeyFaGlYs3A/QPt+U0BGF1+jwZyKPmoX+KQgjuZiEXzcUg9UZMFx/jw=="
//                       encryptedKeyValue="NDzZT19PF7P1D5fRZ89XyTH45Gm9FJQD8hInUt47Foc="/>
//     </keyEncryptor>
//   </keyEncryptors>
// </encryption>

function TXLSCryptoAgile.ReadXML(ABuf: PByteArray; ASize: integer): boolean;
var
  DOM : TXpgSimpleDOM;
  Node: TXpgDOMNode;
  Attr: TXpgDOMAttribute;
  S   : AnsiString;
  S2  : AxUCString;
  Sz  : integer;
begin
  Result := False;

  DOM := TXpgSimpleDOM.Create;
  try
    DOM.LoadFromBuffer(ABuf,ASize);

    Node := DOM.Root.Find('encryption/keyData');
    if Node <> Nil then begin
      Attr := Node.Attributes.Find('saltValue');
      if Attr <> Nil then begin
        S := Attr.AsBase64;

        Sz := Node.Attributes.FindValueIntDef('saltSize',0);
        if Length(S) <> Sz then
          Exit;

        FKeyDataSalt.Assign(S);
      end
      else
        Exit;
    end
    else
      Exit;


//    Node := DOM.Root.Find('encryption/keyData');
    Node := DOM.Root.Find('encryption/keyEncryptors/keyEncryptor/p:encryptedKey');

    if Node <> Nil then begin
      Attr := Node.Attributes.Find('saltValue');
      if Attr <> Nil then begin
        S := Attr.AsBase64;

        FSaltSize := Node.Attributes.FindValueIntDef('saltSize',0);
        if Length(S) <> Integer(FSaltSize) then
          Exit;

        FSaltValue.Assign(S);
      end
      else
        Exit;

      FSpinCount := Node.Attributes.FindValueIntDef('spinCount',0);
      FBlockSize := Node.Attributes.FindValueIntDef('blockSize',0);
      FKeyBits := Node.Attributes.FindValueIntDef('keyBits',0);
      FHashSize := Node.Attributes.FindValueIntDef('hashSize',0);

      S2 := Node.Attributes.FindValue('cipherAlgorithm');
      if S2 = 'AES' then begin
        S2 := Node.Attributes.FindValue('cipherChaining');
        if S2 = 'ChainingModeCBC' then
          FCipherAlgorithm := xccAESCBC;
      end
      else
        Exit;

      S2 := Node.Attributes.FindValue('hashAlgorithm');
      if S2 = 'SHA1' then
        FHashAlgorithm := xchSHA1
      else if S2 = 'SHA512' then
        FHashAlgorithm := xchSHA512;

      Attr := Node.Attributes.Find('encryptedVerifierHashInput');
      if Attr <> Nil then begin
        S := Attr.AsBase64;
        FEncryptedVerifierHashInput.Assign(S);
      end
      else
        Exit;

      Attr := Node.Attributes.Find('encryptedVerifierHashValue');
      if Attr <> Nil then begin
        S := Attr.AsBase64;
        FEncryptedVerifierHashValue.Assign(S);
      end
      else
        Exit;

      Attr := Node.Attributes.Find('encryptedKeyValue');
      if Attr <> Nil then begin
        S := Attr.AsBase64;
        FEncryptedKeyValue.Assign(S);
      end
      else
        Exit;

    end
    else
      Exit;
  finally
    DOM.Free;
  end;

  Result := True;
end;

procedure TXLSCryptoAgile.SetTestSHA512;
begin
  FPassword := 'password';

  FSaltValue.Assign([$59,$B4,$9C,$64,$C0,$D2,$9D,$E7,$33,$F0,$02,$58,$37,$32,$7D,$50]);
  FEncryptedVerifierHashInput.Assign([$70,$AC,$C7,$94,$66,$46,$EA,$30,$0F,$C1,$3C,$FE,$3B,$D7,$51,$E2]);
  FEncryptedVerifierHashValue.Assign([$62,$7C,$8B,$DB,$7D,$98,$46,$22,$8A,$AE,$A8,$1E,$EE,$D4,$34,$D0,$22,$BB,$93,$BB,$5F,$4D,$A1,$46,$CB,$3A,$D9,$D8,$47,$DE,$9E,$C9]);
end;

procedure TXLSCryptoAgile.SetupSHA512;
begin
  FSaltSize := 16;
  FSpinCount := 100000;
  FKeyBits  := 256;
  FHashSize := 64;

  FSaltValue.Resize(16,0);
  FEncryptedVerifierHashInput.Resize(16,0);
  FEncryptedVerifierHashValue.Resize(32,0);
  FEncryptedKeyValue.Resize(32,0);
end;

function TXLSCryptoAgile.Test: boolean;
begin
  SetTestSHA512;

  Result := GenerateEncryptionKey(FPassword);
end;

{ TXLSCrypto }

constructor TXLSCrypto.Create(APassword: AxUCString);
begin
  FPassword := APassword;
end;

destructor TXLSCrypto.Destroy;
begin

  inherited;
end;

procedure TXLSCrypto.FillRandom(AArray: PByteArray; ASize: integer);
var
  i: integer;
begin
  Randomize;

  for i := 0 to ASize - 1 do
    AArray[i] := Byte(Random(256));
end;

{ TXLSCryptoStandard }

procedure TXLSCryptoStandard.CalcKey;
var
  i   : integer;
  Sz  : integer;
  PwSz: longword;
  SHA1: TSHA1;
  Buf : PByteArray;
  Hash: TSHA1Digest;
begin
  PwSz := Length(FPassword) * 2;
  Sz := PwSz + FVerifier.SaltSize;
  GetMem(Buf,Max(Sz,SizeOf(TSHA1Digest) + 4));
  try
    Move(FVerifier.Salt[0],Buf[0],FVerifier.SaltSize);
    Move(Pointer(FPassword)^,Buf[FVerifier.SaltSize],PwSz);

    SHA1.Full(Buf,Sz,Hash);

    for i := 0 to 49999 do begin
      Move(i,Buf[0],4);
      Move(Hash[0],Buf[4],SizeOf(TSHA1Digest));
      SHA1.Full(Buf,SizeOf(TSHA1Digest) + 4,Hash);
    end;

    i := 0;
    Move(Hash[0],Buf[0],SizeOf(TSHA1Digest));
    Move(i,Buf[SizeOf(TSHA1Digest)],4);

    SHA1.Full(Buf,SizeOf(TSHA1Digest) + 4,Hash);

    ReAllocMem(Buf,64);
    FillChar(Buf^,64,$36);

    for i := 0 to SizeOf(TSHA1Digest) - 1 do
      Buf[i] := Buf[i] xor Hash[i];

    SHA1.Full(Buf,64,Hash);

    Move(Hash[0],FKey[0],16);
  finally
    FreeMem(Buf);
  end;
end;

constructor TXLSCryptoStandard.Create(APassword: AxUCString);
begin
  inherited Create(APassword);
end;

function TXLSCryptoStandard.DecryptDocument(AInStream: IStream; AOutStream: TStream): boolean;
var
  Sz     : integer;
{$ifdef DELPHI_XE8_OR_LATER}
  P      : largeuint;
{$else}
  P      : largeint;
{$endif}
  AES    : TAESECB;
  iBuf   : array[0..4095] of byte;
  oBuf   : array[0..4095] of byte;
//  Ptr    : Pointer;
begin
  AInStream.Read(@Sz,4,@P);
  AInStream.Read(@Sz,4,@P);

  AES := TAESECB.Create(FKey,16 * 8);
  try
    while True do begin
      AInStream.Read(@iBuf[0],SizeOf(iBuf),@P);
      Sz := P;
      if Sz <= 0 then
        Break;

      AES.Decrypt(@iBuf[0],@oBuf[0],Sz);

      // This works with DDX 10.2
      AOutStream.Write(oBuf,Sz);

      // Works with previous version2 ?
//      Ptr := @oBuf[0];
//      AOutStream.Write(Ptr,Sz);
    end;

    AOutStream.Seek(0,soBeginning);
  finally
    AES.Free;
  end;

  AOutStream.Seek(0,soFromBeginning);

  Result := True;
end;

procedure TXLSCryptoStandard.EncryptVerifier(AVeriVal: PByteArray);
var
  AES : TAESECB;
  SHA1: TSHA1;
  Hash: TSHA1Digest;
  HashEnc: array[0..31] of byte;
begin
  AES := TAESECB.Create(FKey,16 * 8);
  try
    AES.Encrypt(AVeriVal,@FVerifier.EncryptedVerifier,16);
  finally
    AES.Free;
  end;

  SHA1.Full(AVeriVal,16,Hash);

  FillChar(HashEnc,32,#0);
  Move(Hash,HashEnc,20);

  AES := TAESECB.Create(FKey,16 * 8);
  try
    AES.Encrypt(@HashEnc,@FVerifier.EncryptedVerifierHash,32);
  finally
    AES.Free;
  end;
end;

function TXLSCryptoStandard.EncryptDocument(AInStream: TStream; AOutStream: IStream): boolean;
var
  Sz     : integer;
  Tmp    : integer;
  n      : integer;
{$ifdef DELPHI_XE8_OR_LATER}
  P      : largeuint;
{$else}
  P      : largeint;
{$endif}
  AES    : TAESECB;
  IV     : TAESBlock;
  iBuf   : array[0..4095] of byte;
  oBuf   : array[0..4095] of byte;
begin
  Sz := AInStream.Size;
  AOutStream.Write(@Sz,4,@P);
  Tmp := 0;
  AOutStream.Write(@Tmp,4,@P);

  Sz := 0;
  AES := TAESECB.Create(FKey,16 * 8);
  try
    FillChar(IV,16,#0);
    AES.IV := IV;
    while True do begin
      n := AInStream.Read(iBuf,SizeOf(iBuf));
      if n = 0 then
        Break
      else if (n mod 16) <> 0 then
        n := ((n div 16) * 16) + 16;

      AES.Encrypt(@iBuf[0],@oBuf[0],n);
      AOutStream.Write(@oBuf[0],n,@P);

      Inc(Sz,n);
    end;
  finally
    AES.Free;
  end;

  Result := True;
end;

function TXLSCryptoStandard.EncryptDocumentFile(AInStream,AOutStream: TStream): boolean;
var
  Sz     : integer;
  Tmp    : integer;
  n      : integer;
  AES    : TAESECB;
  iBuf   : array[0..4095] of byte;
  oBuf   : array[0..4095] of byte;
  Ptr    : Pointer;
begin
  Sz := AInStream.Size;
  
  Ptr := @Sz;
  AOutStream.Write(Ptr,4);

  Tmp := 0;
  
  Ptr := @Tmp;
  AOutStream.Write(Ptr,4);

  Sz := 0;
  AES := TAESECB.Create(FKey,16 * 8);
  try
    while True do begin
      n := AInStream.Read(iBuf,SizeOf(iBuf));
      if n = 0 then
        Break
      else if (n mod 16) <> 0 then
        n := ((n div 16) * 16) + 16;

      AES.Encrypt(@iBuf[0],@oBuf[0],n);
      
      Ptr := @oBuf[0];
      AOutStream.Write(Ptr,n);

      Inc(Sz,n);
    end;
  finally
    AES.Free;
  end;

  Result := True;
end;

function TXLSCryptoStandard.InitiateRead(ABuf: PByteArray; ABufSize: integer): TXLSCryptoResult;
var
  HdrSz   : longword;
//  Flags   : longword;
  Sz      : integer;
  Name    : WideString;
begin
  // Skip flags
//  ABuf := PByteArray(NativeInt(ABuf) + 4);

  HdrSz := PLongword(ABuf)^;
  ABuf := PByteArray(NativeInt(ABuf) + 4);

  Move(ABuf^,FHeader,SizeOf(TXLSRecEncHeader));

  ABuf := PByteArray(NativeInt(ABuf) + SizeOf(TXLSRecEncHeader));

  Sz := HdrSz - SizeOf(TXLSRecEncHeader);
  SetLength(Name,Sz div 2);
  Move(ABuf^,Pointer(Name)^,Sz);

  ABuf := PByteArray(NativeInt(ABuf) + Sz);

  Move(ABuf^,FVerifier,SizeOf(TXLSRecEncVerifier));

  CalcKey;

  if VerifyPassword then
    Result := xcrOk
  else
    Result := xcrWrongPassword;
end;

function TXLSCryptoStandard.InitiateWrite(AStream: IStream): TXLSCryptoResult;
var
  Ver    : TXLSRecEncVersion;
  Hdr    : TXLSRecEncHeader;
  Flags  : longword;
  V      : longword;
{$ifdef DELPHI_XE8_OR_LATER}
  P      : largeuint;
{$else}
  P      : largeint;
{$endif}
  Name   : AxUCString;
  VeriVal: array[0..15] of byte;
begin
  Name := 'Microsoft Enhanced RSA and AES Cryptographic Provider'#0;

  Ver.Major := $0003;
  Ver.Minor := $0002;

  AStream.Write(@Ver,SizeOf(TXLSRecEncVersion),@P);

  Flags := $00000024;
  AStream.Write(@Flags,4,@P);

  V := SizeOf(TXLSRecEncHeader) + Length(Name) * 2;
  AStream.Write(@V,4,@P);

  Hdr.Flags := Flags;
  Hdr.SizeExtra := 0;
  Hdr.AlgId := $0000660E;
  Hdr.AlgIdHash := $00008004;
  Hdr.KeySize := $00000080;
  Hdr.ProviderType := $00000018;
  Hdr.Reserved1 := 0;
  Hdr.Reserved2 := 0;
  AStream.Write(@Hdr,SizeOf(TXLSRecEncHeader),@P);

  AStream.Write(@Name[1],Length(Name) * 2,@P);

  FillRandom(@VeriVal[0],SizeOf(VeriVal));

  FVerifier.SaltSize := 16;
  FillRandom(@FVerifier.Salt[0],SizeOf(FVerifier.Salt));
  FVerifier.VerifierHashSize := 20;

  CalcKey;

  EncryptVerifier(@VeriVal);

  AStream.Write(@FVerifier,SizeOf(TXLSRecEncVerifier),@P);

  Result := xcrOk;
end;

function TXLSCryptoStandard.VerifyPassword: boolean;
var
  AES    : TAESECB;
  SHA1   : TSHA1;
  Hash   : TSHA1Digest;
  VHash  : array[0..31] of byte;
  VeriVal: array[0..15] of byte;
begin
  AES := TAESECB.Create(FKey,16 * 8);
  try
    AES.Decrypt(@FVerifier.EncryptedVerifier,@VeriVal,16);
  finally
    AES.Free;
  end;

  AES := TAESECB.Create(FKey,16 * 8);
  try
    AES.Decrypt(@FVerifier.EncryptedVerifierHash,@VHash,32);
  finally
    AES.Free;
  end;

  SHA1.Full(@VeriVal,16,Hash);

  Result := CompareMem(@VHash,@Hash,16);
end;

end.
