unit BIFF_Stream5;

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

uses SysUtils, Classes,
{$ifdef BABOON}
     BIFF_ICompoundStream5,
{$else}
     Windows, ActiveX,
{$ifdef DELPHI_XE2_OR_LATER}
     System.Win.ComObj,
{$else}
     ComObj,
{$endif}
{$endif}
     Math,
     BIFF_RecsII5, BIFF_EscherTypes5, BIFF_VBA5, BIFF_MD5_5, BIFF_RC4_5,
     Xc12Utils5, XLSUtils5;

type TExtraObjects = class(TObject)
private
     FSavedLockBytes: ILockBytes;
     FSavedStorage: IStorage;
public
     procedure Clear;

     property SavedLockBytes: ILockBytes read FSavedLockBytes;
     property SavedStorage: IStorage read FSavedStorage;
     end;

type TXLSStream = class(TObject)
private
     OleStorage: IStorage;
     OleStream: IStream;
     FTargetStream: TStream;
     FSourceStream: TStream;
     FLockBytes: ILockBytes;
     FExtraObjects: TExtraObjects;
     FVBA: TXLSVBA;
     FFileVersion: TExcelVersion;
     FStream: TStream;
     FStreamSize: integer;
     FWrittenBytes: integer;
     FWriteCONTINUEPos: integer;
     FReadCONTINUEPos: integer;
     FReadCONTINUEActive: boolean;
     FMaxBytesWrite: integer;
     FSaveVBA: boolean;
     FISAPIRead: boolean;
     FIsEncrypted: boolean;
     FMD5Ctx: TMD5Context;
     FRC4Key: TRC4Key;
     FNextReKeyBlock: integer;

     function  OpenFileStreamRead(const Filename: string): TExcelVersion;
     function  IntWrite(const Buffer; Count: Longint): longint;
     procedure WriteStorageToDestStream;
     function  OpenOleStreamRead(Filename: PWideChar): TExcelVersion;
     procedure MakeKey(Block: longword; ctxVal: TMD5Context; var Key: TRC4Key);
     procedure EncryptSkipBytes(Start,Count: integer);
     function  VerifyPassword(FILEPASS: PRecFILEPASS; Password: AxUCString): boolean;
public
     constructor Create(VBA: TXLSVBA);
     destructor Destroy; override;

     procedure CreatePassword(FILEPASS: PRecFILEPASS; Password: AxUCString);
     function  OpenStorageRead(const Filename: AxUCString): TExcelVersion;
     procedure OpenStorageWrite(const Filename: AxUCString; Version: TExcelVersion);
     procedure OpenRawFileStreamWrite(const Filename: string);
     procedure OpenRawFileStreamRead(const Filename: string);
     procedure OpenRawMemStreamWrite;
     procedure OpenRawMemStreamRead;
     function  Read(var Buffer; Count: Longint): longint;
     function  ReadUnencrypted(var Buffer; Count: Longint): longint;
     function  ReadUnencryptedSync(var Buffer; Count: Longint): longint;
     function  ReadHeader(var Header: TBIFFHeader): integer;
     // Do not affect continue
     function  ReadOtherHeader(var Header: TBIFFHeader): integer;
     function  PeekHeader: word;
     function  PeekMSOHeader: word;
     function  Write(const Buffer; Count: Longint): longint;
     procedure WriteCONTINUE(RecId: word; const Buffer; Count: Longint);
     procedure WriteHeader(RecId: word; Length: word);
     procedure WriteMSOHeader(Version: byte; Instance: word; FBT: word; Length: longword);
     procedure WriteBufHeader(const Buffer; RecId: word; Length: word);
     procedure WriteWord(RecId: word; Value: word);
     procedure WLWord(Value: longword);
     procedure WWord(Value: word);
     procedure WByte(Value: byte);
     procedure WriteUnicodeStr16(Value: XLS8String);
     procedure WriteUnicodeStr8(Value: XLS8String);
     procedure WriteWideString(Value: AxUCString; WriteLen: boolean = True);
     procedure WriteCellArea(Area: TRec32CellArea);
     procedure WriteCellAreaI(Area: TRecCellAreaI);
     function  ReadWideString(Length: integer; var Str: AxUCString): integer;
     function  Seek(Offset: Longint; Origin: Word): longint;
     function  Pos: longint;
     procedure BeginCONTINUEWrite(MaxRecSize: integer);
     procedure CheckCONTINUEWrite(const Buffer; Count: Longint);
     procedure EndCONTINUEWrite;
     procedure BeginCONTINUERead;
     procedure ResetCONTINUERead;
     procedure EndCONTINUERead;
     procedure WriteVBA;
     function  SetReadDecrypt(FILEPASS: PRecFILEPASS; Password: AxUCString): boolean;
     procedure Decrypt(Buffer: PByteArray; Count: integer);
     procedure CloseStorage;
     procedure CloseStream;
     procedure EncryptFile;
     // For debugging
     procedure WriteFile(RecId,Length: word; Filename: string);

     property  Size: integer read FStreamSize;
     property  WrittenBytes: integer read FWrittenBytes write FWrittenBytes;
     property  TargetStream: TStream read FTargetStream write FTargetStream;
     property  SourceStream: TStream read FSourceStream write FSourceStream;
     property  ExtraObjects: TExtraObjects read FExtraObjects write FExtraObjects;
     property  SaveVBA: boolean read FSaveVBA write FSaveVBA;
     property  ISAPIRead: boolean read FISAPIRead write FISAPIRead;

     property  IsEncrypted: boolean read FIsEncrypted;
     property FileVersion: TExcelVersion read FFileVersion;

     property ReadCONTINUEActive: boolean read FReadCONTINUEActive;
     end;

type TXLSStreamObj = class(TObject)
protected
public
     procedure Read(Stream: TXLSStream); virtual; abstract;
     procedure Write(Stream: TXLSStream); virtual; abstract;
     end;

implementation

const EncryptReKeyBlockSz = $0400;

{ TExtraObjects }

procedure TExtraObjects.Clear;
begin
  FSavedStorage := Nil;
  FSavedLockBytes := Nil;
end;


{ TXLSStream }

constructor TXLSStream.Create(VBA: TXLSVBA);
begin
  FVBA := VBA;
  FWriteCONTINUEPos := -1;
  FSaveVBA := False;
end;

destructor TXLSStream.Destroy;
begin
  CloseStorage;
  inherited Destroy;
end;

procedure TXLSStream.CloseStorage;
begin
  if FStream <> Nil then
    FStream.Free;
  FStream := Nil;
  if OleStream <> Nil then begin
    try
      if (FTargetStream <> Nil) and (FLockBytes <> nil) then begin
        OleCheck(FLockBytes.Flush);
        OleCheck(OleStorage.Commit(STGC_DEFAULT));
        WriteStorageToDestStream;
      end;
    finally
      OleStorage := nil;
      OleStream := nil;
      FLockBytes := Nil;
    end;
  end;
end;

procedure TXLSStream.WriteVBA;
var
  Storage: IStorage;
begin
  if (FVBA <> Nil) and FVBA.HasData then begin
    OleCheck(OleStorage.CreateStorage(VBA_ROOTDIR_NAME,STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE,0,0,Storage));
    FVBA.Write(Storage);
    Storage := Nil;
  end;
end;

function TXLSStream.OpenOleStreamRead(Filename: PWideChar): TExcelVersion;
var
  Stat: TSTATSTG;
  FileMode: integer;
  DataHandle: HGLOBAL;
  Buffer: PByteArray;
  Hdr: TBIFFHeader;
begin
// NOTE! Do not use OleStorage.EnumElements as this will cause a memory leak.

  CloseStorage;
  if FISAPIRead then
    FileMode := STGM_READ or STGM_SHARE_DENY_WRITE
  else
    FileMode := STGM_READ or STGM_TRANSACTED or STGM_SHARE_DENY_NONE;

  if FSourceStream <> Nil then begin
    DataHandle := GlobalAlloc(GMEM_MOVEABLE,FSourceStream.Size);
    Buffer := GlobalLock(DataHandle);
    try
      FSourceStream.ReadBuffer(Buffer^, FSourceStream.Size);
      FSourceStream.Seek(0,soFromBeginning);

      OleCheck(CreateILockBytesOnHGlobal(DataHandle, True, FLockBytes));
      OleCheck(StgOpenStorageOnILockBytes(FLockBytes, Nil, STGM_READ or STGM_TRANSACTED or STGM_SHARE_DENY_NONE , Nil, 0, OleStorage));
    finally
      GlobalUnlock(DataHandle);
    end;
  end
  else
    OleCheck(StgOpenStorage(Filename,Nil,FileMode ,Nil,0,OleStorage));

  if FExtraObjects <> Nil then begin
    OleCheck(CreateILockBytesOnHGlobal(0,True,FExtraObjects.FSavedLockBytes));
    OleCheck(StgCreateDocfileOnILockBytes(FExtraObjects.FSavedLockBytes,STGM_CREATE + STGM_READWRITE + STGM_SHARE_EXCLUSIVE,0,FExtraObjects.FSavedStorage));

    OleCheck(OleStorage.CopyTo(0,Nil,Nil,FExtraObjects.FSavedStorage));
    FExtraObjects.FSavedStorage.DestroyElement('Workbook');
    FExtraObjects.FSavedStorage.DestroyElement('Book');
  end;

  FileMode := STGM_READ;
  if OleStorage.OpenStream('Workbook',Nil,STGM_DIRECT or FileMode or STGM_SHARE_EXCLUSIVE,0,OleStream) = S_OK then
    Result := xvExcel97
  else if OleStorage.OpenStream('Book',Nil,STGM_DIRECT or FileMode or STGM_SHARE_EXCLUSIVE,0,OleStream) = S_OK then
    Result := xvExcel50
  else if OleStorage.OpenStream('EncryptionInfo',Nil,STGM_DIRECT or FileMode or STGM_SHARE_EXCLUSIVE,0,OleStream) = S_OK then
    Result := xvOffice2007Encrypted
  else
    raise XLSRWException.Create('The file is not an Excel workbook');

  ReadHeader(Hdr);

  OleStream.Stat(Stat,0);
  FStreamSize := Stat.cbSize;

  if FVBA <> Nil then
    FVBA.Read(OleStorage);
end;

function TXLSStream.OpenFileStreamRead(const Filename: string): TExcelVersion;
var
  Header: TBIFFHeader;
begin
  Result := xvNone;
  CloseStorage;
  FStream := TFileStream.Create(Filename,fmOpenRead);
  FStreamSize := FStream.Size;
  FStream.Read(Header,SizeOf(TBIFFHeader));
  if (Header.RecID and $FF) = BIFFRECID_BOF then begin
    case (Header.RecID and $FF00) of
      $0000: Result := xvExcel21;
      $0200: Result := xvExcel30;
      $0400: Result := xvExcel40;
    end;
  end;
  FStream.Seek(0,soFromBeginning);
end;

procedure TXLSStream.OpenRawFileStreamRead(const Filename: string);
begin
  CloseStorage;
  FStream := TFileStream.Create(Filename,fmOpenRead);
  FStreamSize := 0;
  FFileVersion := xvExcel97;
end;

procedure TXLSStream.OpenRawFileStreamWrite(const Filename: string);
begin
  CloseStorage;
  FStream := TFileStream.Create(Filename,fmCreate);
  FStreamSize := 0;
  FFileVersion := xvExcel97;
end;

procedure TXLSStream.OpenRawMemStreamRead;
begin
  CloseStorage;
  FStream := TMemoryStream.Create;
  FStreamSize := 0;
  FFileVersion := xvExcel97;
end;

procedure TXLSStream.OpenRawMemStreamWrite;
begin
  CloseStorage;
  FStream := TMemoryStream.Create;
  FStreamSize := 0;
  FFileVersion := xvExcel97;
end;

function TXLSStream.Read(var Buffer; Count: Longint): integer;
var
  R,Sz: integer;
  P,P2: Pointer;
  Header: TBIFFHeader;
begin
  if FStream = Nil then begin
    if FReadCONTINUEActive then begin
      if (FReadCONTINUEPos + Count) > MAXRECSZ_97 then begin
        Result := 0;
        P := @Buffer;
        repeat
          Sz := Min(MAXRECSZ_97 - FReadCONTINUEPos,Count);
          if Sz > 0 then begin
            if Header.RecID = BIFFRECID_MSO_0866 then
              Inc(Sz,14);
            OleCheck(OleStream.Read(P,Sz,@R));
            if FIsEncrypted and (Sz > 0) then
              Decrypt(P,Sz);
            if Header.RecID = BIFFRECID_MSO_0866 then begin
              P2 := Pointer(NativeInt(Buffer) + 14);
              System.Move(P2,Buffer,Sz - 14);
            end;
            Inc(Result,R);
            Dec(Count,R);
            P := Pointer(NativeInt(P) + R);
          end
          else
            R := Count;
          FReadCONTINUEPos := 0;
          if Count > 0 then begin
            ReadHeader(Header);
            // Bug in Excel. MSODRAWINGGROUP may be continued with MSODRAWINGGROUP instead of CONTINUE.
            // Bug in Excel. GELFRAME may be continued with GELFRAME instead of CONTINUE.
            if not ((Header.RecID in [BIFFRECID_CONTINUE,BIFFRECID_MSODRAWINGGROUP]) or (Header.RecID = CHARTRECID_GELFRAME) or (Header.RecID = BIFFRECID_MSO_0866)) then
              raise XLSRWException.Create('CONTINUE record is missing');
          end;
        until ((Count <= 0) or (R <= 0));
        FReadCONTINUEPos := R;
      end
      else begin
        OleCheck(OleStream.Read(@Buffer,Count,@Result));
        Inc(FReadCONTINUEPos,Count);
      end;
    end
    else begin
      OleCheck(OleStream.Read(@Buffer,Count,@Result));
    end;
  end
  else
    Result := FStream.Read(Buffer,Count);

  if FIsEncrypted and (Count > 0) then
    Decrypt(@Buffer,Count);
end;

function TXLSStream.ReadUnencrypted(var Buffer; Count: Integer): longint;
var
  TempEnc: boolean;
begin
  TempEnc := FIsEncrypted;
  try
    FIsEncrypted := False;
    Result := Read(Buffer,Count);
  finally
    FIsEncrypted := TempEnc;
  end;
end;

function TXLSStream.ReadUnencryptedSync(var Buffer; Count: Integer): longint;
var
  Buf: array[0..EncryptReKeyBlockSz - 1] of byte;
begin
  Result := ReadUnencrypted(Buffer,Count);

  if FIsEncrypted then
//    When reading unencrypted records (BOF), it seems that the RC4 not shall
//    be re-keyed, if there is a re-key in that record.
    RC4(@Buf,Count,FRC4Key);
//    EncryptSkipBytes(Pos,Count);

end;

function TXLSStream.ReadHeader(var Header: TBIFFHeader): integer;
var
  TempEnc,TempCont: boolean;
begin
  TempEnc := FIsEncrypted;
  TempCont := FReadCONTINUEActive;
  try
    FIsEncrypted := False;
    FReadCONTINUEActive := False;
    Result := Read(Header,SizeOf(TBIFFHeader));
  finally
    FIsEncrypted := TempEnc;
    FReadCONTINUEActive := TempCont;
  end;
  if FIsEncrypted then
    EncryptSkipBytes(Pos - SizeOf(TBIFFHeader),SizeOf(TBIFFHeader));
end;

function TXLSStream.ReadOtherHeader(var Header: TBIFFHeader): integer;
var
  TempEnc: boolean;
begin
  TempEnc := FIsEncrypted;
  try
    FIsEncrypted := False;
    Result := Read(Header,SizeOf(TBIFFHeader));
  finally
    FIsEncrypted := TempEnc;
  end;
  if FIsEncrypted then
    EncryptSkipBytes(Pos - SizeOf(TBIFFHeader),SizeOf(TBIFFHeader));
end;

procedure TXLSStream.WriteHeader(RecId: word; Length: word);
var
  Header: TBIFFHeader;
begin
  Header.RecId := RecId;
  Header.Length := Length;
  IntWrite(Header,SizeOf(TBIFFHeader));
end;

procedure TXLSStream.WriteBufHeader(const Buffer; RecId: word; Length: word);
var
  Header: TBIFFHeader;
begin
  Header.RecId := RecId;
  Header.Length := Length;
  IntWrite(Header,SizeOf(TBIFFHeader));
  IntWrite(Buffer,Length);
end;

procedure TXLSStream.WriteWord(RecId, Value: word);
begin
  WriteHeader(RecId,2);
  IntWrite(Value,2);
end;

function TXLSStream.IntWrite(const Buffer; Count: Integer): longint;
begin
  if FStream = Nil then
    OleCheck(OleStream.Write(@Buffer,Count,@Result))
  else
    Result := FStream.Write(Buffer,Count);
end;

function TXLSStream.Write(const Buffer; Count: Longint): longint;
begin
  Result := 0;
  if FMaxBytesWrite > 0 then
    CheckCONTINUEWrite(Buffer,Count)
  else begin
    Result := IntWrite(Buffer,Count);
    Inc(FWrittenBytes,Result);
  end;
end;

procedure TXLSStream.CheckCONTINUEWrite(const Buffer; Count: Longint);
var
  WBytes: integer;
  P: Pointer;
begin
  P := @Buffer;
  if (Count + FWrittenBytes) > FMaxBytesWrite then begin
    WBytes := FMaxBytesWrite - FWrittenBytes;
    Seek(FWriteCONTINUEPos + 2,soFromBeginning);
    IntWrite(FMaxBytesWrite,SizeOf(word));
    Seek(0,soFromEnd);
    IntWrite(P^,WBytes);
    FWriteCONTINUEPos := Pos;
    WriteHeader(BIFFRECID_CONTINUE,0);
    FWrittenBytes := 0;
    Dec(Count,WBytes);
    P := Pointer(NativeInt(P) + WBytes);
    CheckCONTINUEWrite(P^,Count);
  end
  else begin
    IntWrite(Buffer,Count);
    Inc(FWrittenBytes,Count);
  end;
end;

function TXLSStream.Seek(Offset: Longint; Origin: Word): integer;
var
{$ifdef BABOON}
  Pos: largeint;
{$else}
{$ifdef DELPHI_XE8_OR_LATER}
  Pos: largeuint;
{$else}
  Pos: largeint;
{$endif}
{$endif}
begin
  if FStream = Nil then begin
    case Origin of
      soFromBeginning:
        OleCheck(OleStream.Seek(Offset,STREAM_SEEK_SET,Pos));
      soFromCurrent:
        OleCheck(OleStream.Seek(Offset,STREAM_SEEK_CUR,Pos));
      soFromEnd:
        OleCheck(OleStream.Seek(Offset,STREAM_SEEK_END,Pos));
    end;
    Result := Pos;
  end
  else
    Result := FStream.Seek(Offset,Origin);
end;

function TXLSStream.Pos: longint;                   
begin
  Result := Seek(0,1);
end;

function TXLSStream.OpenStorageRead(const Filename: AxUCString): TExcelVersion;
begin
  if (FSourceStream <> Nil) or (StgIsStorageFile(PWideChar(Filename)) = S_OK) then
    Result := OpenOLEStreamRead(PWideChar(Filename))
  else
    Result := OpenFileStreamRead(Filename);
  FFileVersion := Result;
end;

procedure TXLSStream.OpenStorageWrite(const Filename: AxUCString; Version: TExcelVersion);
begin
  CloseStorage;
  FFileVersion := Version;
  if Version >= xvExcel50 then begin
    if Assigned(FTargetStream) then begin
      OleCheck(CreateILockBytesOnHGlobal(0, True, FLockBytes));
      OleCheck(StgCreateDocfileOnILockBytes(FLockBytes, STGM_READWRITE or STGM_SHARE_EXCLUSIVE or STGM_CREATE, 0, OleStorage));
    end
    else begin
      if Filename = '' then
        raise XLSRWException.Create('Filename is missing');
      OleCheck(StgCreateDocfile(PWideChar(Filename), STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE, 0, OleStorage));
    end;
{
    if FExtraObjects <> Nil then
      FExtraObjects.WriteStreams(OleStorage);
}
    if (FExtraObjects <> Nil) and (FExtraObjects.FSavedStorage <> Nil) then
      OleCheck(FExtraObjects.FSavedStorage.CopyTo(0,Nil,Nil,OleStorage));

    if Version = xvExcel50 then
      OleStorage.CreateStream('Book',STGM_DIRECT or STGM_WRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,OleStream)
//      OleCheck(OleStorage.CreateStream('Book',STGM_DIRECT or STGM_WRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,OleStream))
    else
      // STGM_READWRITE for encrypted files
      OleCheck(OleStorage.CreateStream('Workbook',STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,OleStream));
//      OleCheck(OleStorage.CreateStream('Workbook',STGM_DIRECT or STGM_WRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,OleStream));
  end
  else begin
    if Assigned(FTargetStream) then
      FStream := FTargetStream
    else
      FStream := TFileStream.Create(Filename,fmCreate);
  end;
end;

procedure TXLSStream.WriteStorageToDestStream;
var
  DataHandle: HGLOBAL;
  Buffer: Pointer;
begin
  OleCheck(FLockBytes.Flush);
  OleCheck(GetHGlobalFromILockBytes(FLockBytes, DataHandle));
  Buffer := GlobalLock(DataHandle);
  try
    FTargetStream.WriteBuffer(Buffer^, GlobalSize(DataHandle));
  finally
    GlobalUnlock(DataHandle);
  end;
end;

procedure TXLSStream.BeginCONTINUEWrite(MaxRecSize: integer);
begin
  FMaxBytesWrite := MaxRecSize;
  FWrittenBytes := 0;
  FWriteCONTINUEPos := Pos;
end;

procedure TXLSStream.EndCONTINUEWrite;
var
  V: word;
begin
  if FWrittenBytes > 0 then begin
    Seek(FWriteCONTINUEPos + 2,soFromBeginning);
    V := FWrittenBytes;
    IntWrite(V,SizeOf(word));
    Seek(0,soFromEnd);
  end;
  FMaxBytesWrite := -1;

end;

procedure TXLSStream.WLWord(Value: longword);
begin
  IntWrite(Value,SizeOf(longword));
end;

procedure TXLSStream.WWord(Value: word);
begin
  IntWrite(Value,SizeOf(word));
end;

procedure TXLSStream.WByte(Value: byte);
begin
  IntWrite(Value,SizeOf(byte));
end;

procedure TXLSStream.WriteUnicodeStr16(Value: XLS8String);
var
  L: word;
begin
  if Length(Value) <= 0 then
    WWord(0)
  else begin
    if Value[1] = #1 then
      L := (Length(Value) - 1) div 2
    else
      L := Length(Value) - 1;
    WWord(L);
    Write(Pointer(Value)^,Length(Value));
  end;
end;

procedure TXLSStream.WriteUnicodeStr8(Value: XLS8String);
var
  L: byte;
begin
  if Length(Value) <= 0 then
    WByte(0)
  else begin
    if Value[1] = #1 then
      L := (Length(Value) - 1) div 2
    else
      L := Length(Value) - 1;
    WByte(L);
    Write(Pointer(Value)^,Length(Value));
  end;
end;

procedure TXLSStream.WriteWideString(Value: AxUCString; WriteLen: boolean = True);
begin
  if WriteLen then
    WWord(Length(Value));
  if Value <> '' then begin
    WByte(1);
    Write(Pointer(Value)^,Length(Value) * 2);
  end
end;

function TXLSStream.ReadWideString(Length: integer; var Str: AxUCString): integer;
var
  B: byte;
  S: string;
begin
  Read(B,1);
  if B = 1 then begin
    SetLength(Str,Length);
    Read(Pointer(Str)^,Length * 2);
    Result := Length * 2 + 1;
  end
  else begin
    SetLength(S,Length);
    Read(Pointer(S)^,Length);
    Str := S;
    Result := Length+ 1;
  end;
end;

procedure TXLSStream.ResetCONTINUERead;
begin
  FReadCONTINUEPos := 0;
end;

procedure TXLSStream.BeginCONTINUERead;
begin
  if FReadCONTINUEActive then
    raise XLSRWException.Create('CONTINUE read active');
  FReadCONTINUEActive := True;
  FReadCONTINUEPos := 0;
end;

procedure TXLSStream.EndCONTINUERead;
begin
  FReadCONTINUEActive := False;
  FReadCONTINUEPos := 0;
end;

function TXLSStream.SetReadDecrypt(FILEPASS: PRecFILEPASS; Password: AxUCString): boolean;
begin
  Result := VerifyPassword(FILEPASS,Password);
  if Result then begin
    FIsEncrypted := True;
    FNextReKeyBlock := -1;
    EncryptSkipBytes(0,Pos);
  end;
end;

procedure TXLSStream.EncryptSkipBytes(Start, Count: integer);
var
  Buf: array[0..EncryptReKeyBlockSz - 1] of byte;
  Block: integer;
begin
  Block := (Start + Count) div EncryptReKeyBlockSz;
  if Block <> FNextReKeyBlock then begin
    FNextReKeyBlock := Block;
    MakeKey(FNextReKeyBlock,FMD5Ctx,FRC4Key);
    Count := (Start + Count) mod EncryptReKeyBlockSz;
  end;
  RC4(@Buf,Count,FRC4Key);
end;

procedure TXLSStream.Decrypt(Buffer: PByteArray; Count: integer);
var
  p,Step: integer;
begin
  p := Pos - Count;
  while FNextReKeyBlock <> ((p + Count) div EncryptReKeyBlockSz) do begin
    Step := EncryptReKeyBlockSz - (p mod EncryptReKeyBlockSz);
    RC4(Buffer,Step,FRC4Key);
    Buffer := PByteArray(NativeInt(Buffer) + Step);
    Inc(p,Step);
    Dec(Count,Step);
    Inc(FNextReKeyBlock);
    MakeKey(FNextReKeyBlock,FMD5Ctx,FRC4Key);
  end;
  RC4(Buffer,Count,FRC4Key);
end;

procedure TXLSStream.MakeKey(Block: longword; ctxVal: TMD5Context; var Key: TRC4Key);
var
  Ctx: TMD5Context;
  PWArray: array[0..63] of byte;
  MD5: TMD5;
begin
  FillChar(PWarray,SizeOf(PWArray),0);

  Move(ctxVal.Digest[0],PWArray[0],5);

  PWArray[5] := Byte(Block and $FF);
  PWArray[6] := Byte((Block shr 8) and $FF);
  PWArray[7] := Byte((Block shr 16) and $FF);
  PWArray[8] := Byte((Block shr 24) and $FF);
  PWArray[9] := $80;
  PWArray[56] := $48;

  MD5 := TMD5.Create;
  try
    MD5.Init(Ctx);
    MD5.Update(Ctx,@PWArray,64);
    MD5.Store(Ctx);
  finally
    MD5.Free;
  end;
  RC4PrepareKey(@Ctx.Digest,16,Key);
end;

function TXLSStream.VerifyPassword(FILEPASS: PRecFILEPASS; Password: AxUCString): boolean;
var
  i: integer;
  PWArray: array[0..63] of byte;
  Ctx1,Ctx2: TMD5Context;
  Salt: array[0..63] of byte;
  HashedSalt: array[0..63] of byte;
  MD5: TMD5;
  Offset,KeyOffset: integer;
  ToCopy: longword;
begin
  // What is the max password length? How does Excel deal with passwords
  // exceeding max length?
  Password := Copy(Password,1,(Length(PWArray) div 2) - 1);
  FillChar(PWarray,SizeOf(PWArray),0);
  for i := 1 to Length(Password) do
    PWordArray(@PWarray)[i - 1] := Word(Password[i]);
  PWarray[Length(Password) * 2] := $80;
  PWarray[56] := Length(Password) shl 4;

  MD5 := TMD5.Create;
  try
    MD5.Init(Ctx1);
    MD5.Update(Ctx1,@PWArray,64);
    MD5.Store(Ctx1);

    Offset := 0;
    KeyOffset := 0;
    ToCopy := 5;

    MD5.Init(FMD5Ctx);
    while Offset <> 16 do begin
      if (64 - Offset) < 5 then
        ToCopy := 64 - Offset;

      Move(Ctx1.Digest[KeyOffset],PWArray[Offset],ToCopy);
      Inc(Offset,ToCopy);

      if Offset = 64 then begin
        MD5.Update(FMD5Ctx,@PWArray,64);
        KeyOffset := ToCopy;
        ToCopy := 5 - ToCopy;
        Offset := 0;
        Continue;
      end;

      KeyOffset := 0;
      ToCopy := 5;
      Move(FILEPASS.DocId,PWArray[Offset],16);
      Inc(Offset,16);
    end;
    PWArray[16] := $80;
    FillChar(PWArray[17],47,0);
    PWArray[56] := $80;
    PWArray[57] := $0A;
                                                 
    MD5.Update(FMD5Ctx,@PWArray, 64);
    MD5.Store(FMD5Ctx);

//    WriteBuf(@ctxVal.Digest,16);

    MakeKey (0,FMD5Ctx,FRC4Key);

    Move(FILEPASS.Salt,Salt,SizeOf(FILEPASS.Salt));
    RC4(@Salt,16,FRC4Key);
    Move(FILEPASS.HashedSalt,HashedSalt,16);
    RC4(@HashedSalt,16,FRC4Key);

    Salt[16] := $80;
    FillChar(Salt[17],47,0);
    Salt[56] := $80;

    MD5.Init(Ctx2);
    MD5.Update(Ctx2,@Salt,64);
    MD5.Store(Ctx2);

    Result := CompareMem(@Ctx2.Digest,@HashedSalt,16);
  finally
    MD5.Free;
  end;
end;

procedure TXLSStream.WriteFile(RecId, Length: word; Filename: string);
var
  S: AnsiString;
  B: byte;
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead);
  try
    WriteHeader(RecId,Length);
    SetLength(S,3);
    while (Length > 0) and (Stream.Read(Pointer(S)^,3) = 3) do begin
      B := 0;
      if S[1] in ['0'..'9'] then
        B := (Byte(S[1]) - Byte('0')) shl 4
      else if S[1] in ['A'..'F'] then
        B := (Byte(S[1]) - Byte('A') + 10) shl 4;
      if S[2] in ['0'..'9'] then
        B := B + (Byte(S[2]) - Byte('0'))
      else if S[2] in ['A'..'F'] then
        B := B + (Byte(S[2]) - Byte('A') + 10);
      WByte(B);
      Dec(Length);
    end;
  finally
    Stream.Free;
  end;
end;

function TXLSStream.PeekHeader: word;
var
  Header: TBIFFHeader;
  Sz: integer;
begin
  if FStream <> Nil then
    Sz := FStream.Read(Header,SizeOf(TBIFFHeader))
  else
    OleCheck(OleStream.Read(@Header,SizeOf(TBIFFHeader),@Sz));
  if Sz = SizeOf(TBIFFHeader) then begin
//  if ReadUnencrypted(Header,SizeOf(TBIFFHeader)) = SizeOf(TBIFFHeader) then begin
    Result := Header.RecID;
    Seek(-SizeOf(TBIFFHeader),soFromCurrent);
  end
  else
    Result := 0;
end;

function TXLSStream.PeekMSOHeader: word;
var
  Sz: integer;
begin
  OleCheck(OleStream.Read(@Result,2,@Sz));
  OleCheck(OleStream.Read(@Result,2,@Sz));
  if Sz = 2 then
    Seek(-4,soFromCurrent)
  else
    Result := 0;
end;

procedure TXLSStream.WriteCONTINUE(RecId: word; const Buffer; Count: Integer);
var
  P: PByteArray;
  L,L2: integer;
begin
  WWord(RecId);
  if Count > MAXRECSZ_97 then begin
    WWord(MAXRECSZ_97);
    Write(Buffer,MAXRECSZ_97);
    P := PByteArray(NativeInt(@Buffer) + MAXRECSZ_97);
    L := Count - MAXRECSZ_97;
    while L > 0 do begin
      WWord(BIFFRECID_CONTINUE);
      if L > MAXRECSZ_97 then
        L2 := MAXRECSZ_97
      else
        L2 := L;
      Write(L2,2);
      Write(P^,L2);
      P := PByteArray(NativeInt(P) + L2);
      Dec(L,MAXRECSZ_97);
    end;
  end
  else begin
    WWord(Count);
    if Count > 0 then
      Write(Buffer,Count);
  end;
end;

procedure TXLSStream.WriteCellArea(Area: TRec32CellArea);
begin
  if FFileVersion <= xvExcel97 then begin
    WWord(Area.Row1);
    WWord(Area.Row2);
    WWord(Area.Col1);
    WWord(Area.Col2);
  end
  else
    raise XLSRWException.Create('Unhandled Excel version');
end;

procedure TXLSStream.CreatePassword(FILEPASS: PRecFILEPASS; Password: AxUCString);
const
  TestDocId: array[0..15] of byte = ($22,$F7,$B5,$FC,$9F,$4A,$B2,$4F,$41,$A6,$47,$E2,$36,$67,$21,$89);
  TestSalt: array[0..63] of byte = ($B5, $DB, $ED, $4E, $64, $76, $37, $43, $34, $43, $64, $66, $76, $A7, $DA, $FD, $80, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $0, $80, $0, $0, $0, $0, $0, $0, $0);
var
  i: integer;
  PWArray: array[0..63] of byte;
  Ctx1,Ctx2: TMD5Context;
  Salt: array[0..63] of byte;
  HashedSalt: array[0..63] of byte;
  MD5: TMD5;
  Offset,KeyOffset: integer;
  ToCopy: longword;
begin
  Randomize;
  for i := 0 to High(FILEPASS.DocId) do
//    FILEPASS.DocId[i] := Random($FF);
    FILEPASS.DocId[i] := TestDocId[i];
  Password := Copy(Password,1,(Length(PWArray) div 2) - 1);
  FillChar(PWarray,SizeOf(PWArray),0);
  for i := 1 to Length(Password) do
    PWordArray(@PWarray)[i - 1] := Word(Password[i]);
  PWarray[Length(Password) * 2] := $80;
  PWarray[56] := Length(Password) shl 4;

  MD5 := TMD5.Create;
  try
    MD5.Init(Ctx1);
    MD5.Update(Ctx1,@PWArray,64);
    MD5.Store(Ctx1);

    Offset := 0;
    KeyOffset := 0;
    ToCopy := 5;

    MD5.Init(FMD5Ctx);
    while Offset <> 16 do begin
      if (64 - Offset) < 5 then
        ToCopy := 64 - Offset;

      Move(Ctx1.Digest[KeyOffset],PWArray[Offset],ToCopy);
      Inc(Offset,ToCopy);

      if Offset = 64 then begin
        MD5.Update(FMD5Ctx,@PWArray,64);
        KeyOffset := ToCopy;
        ToCopy := 5 - ToCopy;
        Offset := 0;
        Continue;
      end;

      KeyOffset := 0;
      ToCopy := 5;
      Move(FILEPASS.DocId,PWArray[Offset],16);
      Inc(Offset,16);
    end;
    PWArray[16] := $80;
    FillChar(PWArray[17],47,0);
    PWArray[56] := $80;
    PWArray[57] := $0A;

    MD5.Update(FMD5Ctx,@PWArray, 64);
    MD5.Store(FMD5Ctx);

    MakeKey (0,FMD5Ctx,FRC4Key);

    for i := 0 to High(Salt) do
//      Salt[i] := Random($FF);
      Salt[i] := TestSalt[i];

    Salt[16] := $80;
    FillChar(Salt[17],47,0);
    Salt[56] := $80;
    MD5.Init(Ctx2);
    MD5.Update(Ctx2,@Salt,64);
    MD5.Store(Ctx2);

    RC4(@Salt,16,FRC4Key);
    Move(Salt,FILEPASS.Salt,SizeOf(FILEPASS.Salt));

    Move(Ctx2.Digest,HashedSalt,16);
    RC4(@HashedSalt,16,FRC4Key);
    Move(HashedSalt,FILEPASS.HashedSalt,SizeOf(FILEPASS.Salt));
//    Result := CompareMem(@Ctx2.Digest,@HashedSalt,16);
  finally
    MD5.Free;
  end;
end;

procedure TXLSStream.EncryptFile;
var
  Header: TBIFFHeader;
  PBuf: PByteArray;
  FirstRec: boolean;
  LW: longword;

procedure SkipRec;
begin
  ReadHeader(Header);
  Seek(Header.Length,soFromCurrent);
end;

begin
  Seek(0,soFromBeginning);
  FirstRec := True;
  GetMem(PBuf,MAXRECSZ_97);
  try
    SkipRec; // BOF
    SkipRec; // FILEPASS
    FNextReKeyBlock := -1;
    EncryptSkipBytes(0,Pos);
    while ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
      EncryptSkipBytes(Pos - SizeOf(TBIFFHeader),SizeOf(TBIFFHeader));
      if (Header.RecID and $FF) = BIFFRECID_BOF then begin
        Read(PBuf^,Header.Length);
        RC4(PBuf,Header.Length,FRC4Key);
      end
      else begin
        Read(PBuf^,Header.Length);
        if Header.Length > 0 then begin
          if Header.RecID = BIFFRECID_BOUNDSHEET then begin
            LW := PRecBOUNDSHEET8(PBuf).BOFPos;
            Decrypt(PBuf,Header.Length);
            PRecBOUNDSHEET8(PBuf).BOFPos := LW;
          end
          else
            Decrypt(PBuf,Header.Length);
          if not FirstRec then begin
            Seek(-Header.Length,soFromCurrent);
            Write(PBuf^,Header.Length);
          end;
          FirstRec := False;
        end;
      end;
    end;
  finally
    FreeMem(PBuf);
  end;
end;

procedure TXLSStream.WriteCellAreaI(Area: TRecCellAreaI);
begin
  if FFileVersion <= xvExcel97 then begin
    WWord(Area.Row1);
    WWord(Area.Row2);
    WWord(Area.Col1);
    WWord(Area.Col2);
  end
  else
    raise XLSRWException.Create('Unhandled Excel version');
end;

procedure TXLSStream.CloseStream;
begin
end;

procedure TXLSStream.WriteMSOHeader(Version: byte; Instance: word; FBT: word; Length: longword);
var
  Header: TMSOHeader;
begin
  Header.VerInst := Version + (Instance shl 4);
  Header.FBT := FBT;
  Header.Length := Length;
  Write(Header,SizeOf(TMSOHeader));
end;

end.
