unit XLSZip5;

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils, Contnrs, Math,
     XLSUtils5, XLSZlibPas5;

const ZL_DEF_COMPRESSIONMETHOD  = $8;  { Deflate }
const ZL_ENCH_COMPRESSIONMETHOD = $9;  { Enchanced Deflate }
const ZL_DEF_COMPRESSIONINFO    = $7;  { 32k window for Deflate }
const ZL_PRESET_DICT            = $20;

const ZL_FASTEST_COMPRESSION    = $0;
const ZL_FAST_COMPRESSION       = $1;
const ZL_DEFAULT_COMPRESSION    = $2;
const ZL_MAXIMUM_COMPRESSION    = $3;

const ZL_FCHECK_MASK            = $1F;
const ZL_CINFO_MASK             = $F0; { mask out leftmost 4 bits }
const ZL_FLEVEL_MASK            = $C0; { mask out leftmost 2 bits }
const ZL_CM_MASK                = $0F; { mask out rightmost 4 bits }

const ZL_MULTIPLE_DISK_SIG      = $08074b50; // 'PK'#7#8
const ZL_DATA_DESCRIPT_SIG      = $08074b50; // 'PK'#7#8
const ZL_LOCAL_HEADERSIG        = $04034b50; // 'PK'#3#4
const ZL_CENTRAL_HEADERSIG      = $02014b50; // 'PK'#1#2
const ZL_EOC_HEADERSIG          = $06054b50; // 'PK'#5#6

type TZipCompressionType   = (ctNormal, ctMaximum, ctFast, ctSuperFast, ctNone, ctUnknown);
type TZipCompressionMethod = (cmStored, cmShrunk, cmReduced1, cmReduced2, cmReduced3, cmReduced4, cmImploded, cmTokenizingReserved, cmDeflated, cmDeflated64, cmDCLImploding, cmPKWAREReserved);

type TEndOfCentralDir = packed record
     EndOfCentralDirSignature        : Cardinal;  //    4 bytes  (0x06054b50)
     NumberOfThisDisk                : WORD;      //    2 bytes
     NumberOfTheDiskWithTheStart     : WORD;      //    2 bytes
     TotalNumberOfEntriesOnThisDisk  : WORD;      //    2 bytes
     TotalNumberOfEntries            : WORD;      //    2 bytes
     SizeOfTheCentralDirectory       : Cardinal;  //    4 bytes
     OffsetOfStartOfCentralDirectory : Cardinal;  //    4 bytes
     ZipfileCommentLength            : WORD;      //    2 bytes
     end;

type TCentralDirectoryFile = packed record
     CentralFileHeaderSignature     : Cardinal;   //    4 bytes  (0x02014b50)
     VersionMadeBy                  : WORD;       //    2 bytes
     VersionNeededToExtract         : WORD;       //    2 bytes
     GeneralPurposeBitFlag          : WORD;       //    2 bytes
     CompressionMethod              : WORD;       //    2 bytes
     LastModFileTimeDate            : Cardinal;   //    4 bytes
     Crc32                          : Cardinal;   //    4 bytes
     CompressedSize                 : Cardinal;   //    4 bytes
     UncompressedSize               : Cardinal;   //    4 bytes
     FilenameLength                 : WORD;       //    2 bytes
     ExtraFieldLength               : WORD;       //    2 bytes
     FileCommentLength              : WORD;       //    2 bytes
     DiskNumberStart                : WORD;       //    2 bytes
     InternalFileAttributes         : WORD;       //    2 bytes
     ExternalFileAttributes         : Cardinal;   //    4 bytes
     RelativeOffsetOfLocalHeader    : Cardinal;   //    4 bytes
     FileName                       : AnsiString; //    variable size
     ExtraField                     : AnsiString; //    variable size
     FileComment                    : AnsiString; //    variable size
     end;

type TLocalFile = packed record
     LocalFileHeaderSignature       : Cardinal;   //    4 bytes  (0x04034b50)
     VersionNeededToExtract         : WORD;       //    2 bytes
     GeneralPurposeBitFlag          : WORD;       //    2 bytes
     CompressionMethod              : WORD;       //    2 bytes
     LastModFileTimeDate            : Cardinal;   //    4 bytes
     Crc32                          : Cardinal;   //    4 bytes
     CompressedSize                 : Cardinal;   //    4 bytes
     UncompressedSize               : Cardinal;   //    4 bytes
     FilenameLength                 : WORD;       //    2 bytes
     ExtraFieldLength               : WORD;       //    2 bytes
     FileName                       : AnsiString; //    variable size
     ExtraField                     : AnsiString; //    variable size
     CompressedDataPos              : integer;    //
     end;

type TDataDescriptor = packed record
     DescriptorSignature            : Cardinal;   //    4 bytes UNDOCUMENTED
     Crc32                          : Cardinal;   //    4 bytes
     CompressedSize                 : Cardinal;   //    4 bytes
     UncompressedSize               : Cardinal;   //    4 bytes
     end;

type TXLSZipOpenMode = (xzomClosed,xzomRead,xzomWrite);

type TXLSZipArchive = class;

     TXLSZipFile = class(TStream)
private
     function GetName: AxUCString;
protected
     FParent              : TXLSZipArchive;
     FDate                : TDateTime;
     FIsEncrypted         : boolean;
     FIsFolder            : boolean;
     FCompressionType     : TZipCompressionType;
     FCentralDirectoryFile: TCentralDirectoryFile;
     FLocalFile           : TLocalFile;
     FDStream             : TStream;

     FSeekCount           : integer;

     function  GetCentralEntrySize: Cardinal;
// Not used ?
//{$ifndef DELPHI_6}
//     function  GetSize: int64; override;
//{$endif}
public
     constructor Create(AParent: TXLSZipArchive);
     destructor Destroy; override;

     procedure OpenRead;
     function  Read(var Buffer; Count: Longint): Longint; override;
     // Close the stream by calling Seek with Offset = MAXINT
{$ifdef DELPHI_5}
     function  Seek(Offset: Longint; Origin: Word): Longint; override;
{$else}
     function  Seek(const Offset: Int64; Origin: TSeekOrigin): Int64; override;
{$endif}
     procedure _SaveToStream(AStream: TStream);
     function  Write(const Buffer; Count: Longint): Longint; override;
     procedure Close;

     property Name: AxUCString read GetName;
     property CompressedSize: cardinal read FLocalFile.CompressedSize;
     property UncompressedSize: cardinal read FLocalFile.UncompressedSize;
     end;

     TXLSZipArchive = class(TObjectList)
private
     function GetItems(Index: integer): TXLSZipFile;
protected
     FOpenMode           : TXLSZipOpenMode;
     FStream             : TStream;
     FOwnsStream         : boolean;
     FComment            : AxUCString;
     FEndOfCentralDirPos : cardinal;
     FEndOfCentralDir    : TEndOfCentralDir;
     FHasBadEntries      : boolean;
     FLocalHeaderNumFiles: integer;
     FStoreFolders       : boolean;
     FStoreRelativePath  : boolean;

     FCurrZip            : TXLSZipFile;
     FCurrStreamStart    : integer;

     function  Add: TXLSZipFile;

     function  ParseLocalHeaders: Boolean;
     function  GetLocalEntry(Offset,NextOffset : Integer): TLocalFile;
     procedure LoadLocalHeaders;
     function  FindCentralDirectory: Boolean;
     function  ParseCentralHeaders: Boolean;
     procedure ReadFile;
     procedure ClearEndOfCentralDir;
     procedure WriteCentralDirectory;
     function  AddFolderChain(ItemName: AxUCString): Boolean;
     function  AddStreamFast(ItemName: AxUCString; Stream:TStream): TXLSZipFile;
public
     constructor Create;
     destructor Destroy; override;

     function  Find(AFilename: AxUCString): integer;

     procedure OpenRead(AFilename: AxUCString); overload;
     procedure OpenRead(AStream: TStream); overload;
     procedure OpenWrite(AFilename: AxUCString); overload;
     procedure OpenWrite(AStream: TStream); overload;
     procedure AddStream(AFilename: AxUCString; AStream: TStream);

     function  CreateStream(const AFilename: AxUCString): TCompressionStream;
     procedure CloseStream(AStream: TCompressionStream);

     procedure Close;

     property Items[Index: integer]: TXLSZipFile read GetItems; default;
     end;

implementation

function ZipExtractFilePath(AFilename: AxUCString): AxUCString;
var
  p1,p2: integer;
begin
  p1 := RCPos('.',AFilename);
  if p1 > 1 then begin
    p2 := RCPos('/',AFilename);
    if (p2 > 0) and (p2 < p1) then
      Result := Copy(AFilename,1,p2)
    else
      Result := AFilename;
  end
  else
    Result := AFilename;
end;

function ToZipName(FileName: AxUCString): AxUCString;
var
  P : Integer;
Begin
  Result := StringReplace(FileName,'\','/',[rfReplaceAll]);
  P := Pos(':/',Result);
  if P > 0 then
    System.Delete(Result,1,P + 1);
  P := Pos('//',Result);
  if P > 0 then begin
    System.Delete(Result,1,P + 1);
    P := CPos('/',Result);
    if P > 0 then begin
      System.Delete(Result,1,P);
      P := CPos('/',Result);
      if P > 0 then
        System.Delete(Result,1,P);
    end;
  end;
end;

{ TXLSZipArchive }

function TXLSZipArchive.Add: TXLSZipFile;
begin
  Result := TXLSZipFile.Create(Self);
  inherited Add(Result);
end;

function TXLSZipArchive.AddFolderChain(ItemName: AxUCString): Boolean;
Var
 FN     : AxUCString;
 TN     : AxUCString;
 INCN   : AxUCString;
 P      : Integer;
 MS     : TMemoryStream;
 NoMore : Boolean;
begin
  FN   := ZipExtractFilePath(ToZipName(ItemName));
  TN   := FN;
  INCN := '';
  MS   := TMemoryStream.Create;
  try
    repeat
      NoMore := True;
      P := Pos('\',TN);
      if P > 0 then begin
        INCN := INCN + Copy(TN,1,P);
        System.Delete(TN,1,P);
        MS.Position := 0;
        MS.Size := 0;
        if Find(INCN) < 0 then begin
          AddStreamFast(INCN,MS);
          NoMore := False;
        end
        else
          raise XLSRWException.Create('Duplicate filename in Zip');
      end;
    until NoMore;
    Result := True;
  finally
    MS.Free;
  end;
end;

procedure TXLSZipArchive.AddStream(AFilename: AxUCString; AStream: TStream);
begin
  if FStoreFolders and FStoreRelativePath then
    AddFolderChain(AFileName);
  AddStreamFast(AFileName,AStream);
end;

function TXLSZipArchive.AddStreamFast(ItemName: AxUCString; Stream: TStream): TXLSZipFile;
Var
  Compressor   : TCompressionStream;
  CM           : WORD;
  p            : integer;
  UncompSz     : Integer;
  FCRC32       : Cardinal;
  Level        : TCompressionLevel;
begin
  //*********************************** COMPRESS DATA
  if not FStoreRelativePath then
    ItemName := ExtractFileName(ItemName);

  ItemName := ToZipName(ItemName);

  if Find(ItemName) >= 0 then
    raise XLSRWException.CreateFmt('Duplicate file "%s" in Zip',[ItemName]);

  UncompSz  := Stream.Size - Stream.Position;
  if UncompSz > 0 then
    CM := 8
  else
    CM := 0;

  Level := clFastest;

  //***********************************
  //*********************************** FILL RECORDS
  Result := Add;
  with Result.FLocalFile do begin
    LocalFileHeaderSignature := $04034B50;
    VersionNeededToExtract   := 20;
    GeneralPurposeBitFlag    := 0;
    CompressionMethod        := CM;
    LastModFileTimeDate      := DateTimeToFileDate(Now);
    Crc32                    := 0;
    CompressedSize           := 0;
    UncompressedSize         := UncompSz;
    FilenameLength           := Length(ItemName);
    ExtraFieldLength         := 0;
    FileName                 := AnsiString(ItemName);
    ExtraField               := '';
    CompressedDataPos        := 0;
  end;

  with Result.FCentralDirectoryFile do begin
    CentralFileHeaderSignature     := $02014B50;
    VersionMadeBy                  := 20;
    VersionNeededToExtract         := 20;
    GeneralPurposeBitFlag          := 0;
    CompressionMethod              := CM;
    LastModFileTimeDate            := DateTimeToFileDate(Now);
    Crc32                          := 0;
    CompressedSize                 := 0;
    UncompressedSize               := UncompSz;
    FilenameLength                 := Length(ItemName);
    ExtraFieldLength               := 0;
    FileCommentLength              := 0;
    DiskNumberStart                := 0;
    InternalFileAttributes         := 0;
    ExternalFileAttributes         := $00000020; // faArchive;
    RelativeOffsetOfLocalHeader    := FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    FileName                       := AnsiString(ItemName);
    ExtraField                     := '';
    FileComment                    := '';
  end;

  //************************************ SAVE LOCAL HEADER AND COMPRESSED DATA
  FStream.Position := Result.FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
  FStream.Write(Result.FLocalFile,SizeOf(Result.FLocalFile) - 2 * SizeOf(AnsiString) - SizeOf(Integer));
  if Result.FLocalFile.FilenameLength > 0 then
    FStream.Write(Result.FLocalFile.FileName[1],Result.FLocalFile.FilenameLength);
  if UncompSz > 0 then begin
    p := FStream.Position;

    Compressor := TCompressionStream.Create(Level,FStream,True);
    try
      Compressor.WriteStream(Stream);
      FCRC32 := Compressor.CRC32Val;
    finally
      Compressor.Free;
    end;

    Result.FLocalFile.CompressedSize := FStream.Position - p;
    Result.FLocalFile.Crc32 := FCRC32;
    Result.FCentralDirectoryFile.CompressedSize := FStream.Position - p;
    Result.FCentralDirectoryFile.Crc32 := FCRC32;

{$ifdef DELPHI_5}
    FStream.Seek(Result.FCentralDirectoryFile.RelativeOffsetOfLocalHeader,0);
    FStream.Write(Result.FLocalFile,SizeOf(Result.FLocalFile) - 2 * SizeOf(AnsiString) - SizeOf(Integer));
    FStream.Seek(0,soFromEnd);
{$else}
    FStream.Seek(Result.FCentralDirectoryFile.RelativeOffsetOfLocalHeader,soBeginning);
    FStream.Write(Result.FLocalFile,SizeOf(Result.FLocalFile) - 2 * SizeOf(AnsiString) - SizeOf(Integer));
    FStream.Seek(0,soFromEnd);
{$endif}
  end;

  //************************************ MARK START OF CENTRAL DIRECTORY
  FEndOfCentralDir.OffsetOfStartOfCentralDirectory := FStream.Position;

  //************************************ SAVE END CENTRAL DIRECTORY RECORD
  FEndOfCentralDirPos := FStream.Position;
  FEndOfCentralDir.SizeOfTheCentralDirectory := FEndOfCentralDirPos - FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
  Inc(FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
  Inc(FEndOfCentralDir.TotalNumberOfEntries);

  Result.FDate := Now;

  if (Result.FCentralDirectoryFile.GeneralPurposeBitFlag and 1) > 0 then
    Result.FIsEncrypted := True
  else
    Result.FIsEncrypted := False;
  Result.FIsFolder := (Result.FCentralDirectoryFile.ExternalFileAttributes and faDirectory) > 0;
  Result.FCompressionType := ctUnknown;
  if (Result.FCentralDirectoryFile.CompressionMethod = 8) or (Result.FCentralDirectoryFile.CompressionMethod=9) then begin
    case Result.FCentralDirectoryFile.GeneralPurposeBitFlag AND 6 of
      0 : Result.FCompressionType := ctNormal;
      2 : Result.FCompressionType := ctMaximum;
      4 : Result.FCompressionType := ctFast;
      6 : Result.FCompressionType := ctSuperFast
    end;
  end;
end;

procedure TXLSZipArchive.Close;
begin
  if FStream <> Nil then begin
    if FOpenMode = xzomWrite then
      WriteCentralDirectory;
    if FOwnsStream then
      FStream.Free;
    FStream := Nil;
  end;
  FOpenMode := xzomClosed;
  Clear;
end;

procedure TXLSZipArchive.CloseStream(AStream: TCompressionStream);
var
  FCRC32: Cardinal;
  UncompressedSize: integer;
begin
  AStream.Flush;

  UncompressedSize := AStream.UncompressedSize;

  FCRC32 := AStream.CRC32Val;
  
  AStream.Free;

  FCurrZip.FLocalFile.UncompressedSize := UncompressedSize;
  FCurrZip.FLocalFile.CompressedSize := FStream.Position - FCurrStreamStart;
  FCurrZip.FLocalFile.Crc32 := FCRC32;

  FCurrZip.FCentralDirectoryFile.UncompressedSize := UncompressedSize;
  FCurrZip.FCentralDirectoryFile.CompressedSize := FStream.Position - FCurrStreamStart;
  FCurrZip.FCentralDirectoryFile.Crc32 := FCRC32;

  FStream.Seek(Integer(FCurrZip.FCentralDirectoryFile.RelativeOffsetOfLocalHeader),soFromBeginning);
  FStream.Write(FCurrZip.FLocalFile,SizeOf(FCurrZip.FLocalFile) - 2 * SizeOf(AnsiString) - SizeOf(Integer));
  FStream.Seek(0,soFromEnd);

  //************************************ MARK START OF CENTRAL DIRECTORY
  FEndOfCentralDir.OffsetOfStartOfCentralDirectory := FStream.Position;

  //************************************ SAVE END CENTRAL DIRECTORY RECORD
  FEndOfCentralDirPos := FStream.Position;
  FEndOfCentralDir.SizeOfTheCentralDirectory := FEndOfCentralDirPos - FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
  Inc(FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk);
  Inc(FEndOfCentralDir.TotalNumberOfEntries);

  FCurrZip.FDate := Now;

  if (FCurrZip.FCentralDirectoryFile.GeneralPurposeBitFlag and 1) > 0 then
    FCurrZip.FIsEncrypted := True
  else
    FCurrZip.FIsEncrypted := False;
  FCurrZip.FIsFolder := (FCurrZip.FCentralDirectoryFile.ExternalFileAttributes and faDirectory) > 0;
  FCurrZip.FCompressionType := ctUnknown;
  if (FCurrZip.FCentralDirectoryFile.CompressionMethod = 8) or (FCurrZip.FCentralDirectoryFile.CompressionMethod=9) then begin
    case FCurrZip.FCentralDirectoryFile.GeneralPurposeBitFlag AND 6 of
      0 : FCurrZip.FCompressionType := ctNormal;
      2 : FCurrZip.FCompressionType := ctMaximum;
      4 : FCurrZip.FCompressionType := ctFast;
      6 : FCurrZip.FCompressionType := ctSuperFast
    end;
  end;

  FCurrZip := Nil;
end;

constructor TXLSZipArchive.Create;
begin
  inherited Create;
  FStoreFolders := True;
  FStoreRelativePath := True;
end;

function TXLSZipArchive.CreateStream(const AFilename: AxUCString): TCompressionStream;
var
  CM           : WORD;
  Level        : TCompressionLevel;
  ItemName     : AxUCString;
begin
  if FCurrZip <> Nil then
    raise XLSRWException.Create('There is an active zip');

  if FStoreFolders and FStoreRelativePath then
    AddFolderChain(AFileName);

  ItemName := AFileName;

  //*********************************** COMPRESS DATA
  if not FStoreRelativePath then
    ItemName := ExtractFileName(ItemName);

  ItemName := ToZipName(ItemName);

  if Find(ItemName) >= 0 then
    raise XLSRWException.CreateFmt('Duplicate file "%s" in Zip',[ItemName]);

  CM := 8;

  Level := clFastest;

  //***********************************
  //*********************************** FILL RECORDS
  FCurrZip := Add;
  with FCurrZip.FLocalFile do begin
    LocalFileHeaderSignature := $04034B50;
    VersionNeededToExtract   := 20;
    GeneralPurposeBitFlag    := 0;
    CompressionMethod        := CM;
    LastModFileTimeDate      := DateTimeToFileDate(Now);
    Crc32                    := 0;
    CompressedSize           := 0;
    UncompressedSize         := 0;
    FilenameLength           := Length(ItemName);
    ExtraFieldLength         := 0;
    FileName                 := AnsiString(ItemName);
    ExtraField               := '';
    CompressedDataPos        := 0;
  end;

  with FCurrZip.FCentralDirectoryFile do begin
    CentralFileHeaderSignature     := $02014B50;
    VersionMadeBy                  := 20;
    VersionNeededToExtract         := 20;
    GeneralPurposeBitFlag          := 0;
    CompressionMethod              := CM;
    LastModFileTimeDate            := DateTimeToFileDate(Now);
    Crc32                          := 0;
    CompressedSize                 := 0;
    UncompressedSize               := 0;
    FilenameLength                 := Length(ItemName);
    ExtraFieldLength               := 0;
    FileCommentLength              := 0;
    DiskNumberStart                := 0;
    InternalFileAttributes         := 0;
    ExternalFileAttributes         := $00000020; // faArchive;
    RelativeOffsetOfLocalHeader    := FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    FileName                       := AnsiString(ItemName);
    ExtraField                     := '';
    FileComment                    := '';
  end;

  //************************************ SAVE LOCAL HEADER AND COMPRESSED DATA
  FStream.Position := FCurrZip.FCentralDirectoryFile.RelativeOffsetOfLocalHeader;
  FStream.Write(FCurrZip.FLocalFile,SizeOf(FCurrZip.FLocalFile) - 2 * SizeOf(AnsiString) - SizeOf(Integer));
  if FCurrZip.FLocalFile.FilenameLength > 0 then
    FStream.Write(FCurrZip.FLocalFile.FileName[1],FCurrZip.FLocalFile.FilenameLength);

  FCurrStreamStart := FStream.Position;

  Result := TCompressionStream.Create(Level,FStream,True);
end;

procedure TXLSZipArchive.OpenWrite(AFilename: AxUCString);
begin
  Close;
  FOpenMode := xzomWrite;
  FStream := TFileStream.Create(AFilename,fmCreate);
  FOwnsStream := True;
  ClearEndOfCentralDir;
end;

procedure TXLSZipArchive.OpenWrite(AStream: TStream);
begin
  Close;
  FOpenMode := xzomWrite;
  FStream := AStream;
  FOwnsStream := False;
  ClearEndOfCentralDir;
end;

destructor TXLSZipArchive.Destroy;
begin
  Close;
  inherited;
end;

function TXLSZipArchive.Find(AFilename: AxUCString): integer;
begin
  for Result := 0 to Count - 1 do begin
    if CompareText(AFilename,Items[Result].Name) = 0 then
      Exit;
  end;
  Result := -1;
end;

function TXLSZipArchive.FindCentralDirectory: Boolean;
var
  SeekStart : Integer;
  FilePos   : Integer;
  BR        : Integer;
  Byte_     : Array[0..3] of Byte;
begin
  Result := False;
  if FStream.Size < 22 then
    Exit;
  if FStream.Size < 256 then
    SeekStart := FStream.Size
  else
    SeekStart := 256;
  FilePos := FStream.Size-22;
  BR := SeekStart;
  repeat
    FStream.Position := FilePos;
    FStream.Read(Byte_,4);
    if Byte_[0] = $50 then begin
      if (Byte_[1] = $4B) and (Byte_[2] = $05) and (Byte_[3]=$06) then begin
        FStream.Position := FilePos;
        FEndOfCentralDirPos := FStream.Position;
        FStream.Read(FEndOfCentralDir,SizeOf(TEndOfCentralDir));
        Result  := True;
      end
      else begin
        Dec(FilePos,4);
        Dec(BR ,4);
      end;
    end
    else begin
      Dec(FilePos);
      Dec(BR)
    end;
    if BR < 0 then begin
      case SeekStart of
        256 : begin
          SeekStart := 1024;
          FilePos := FStream.Size - (256 + 22);
          BR := SeekStart;
        end;
        1024 : begin
          SeekStart := 65536;
          FilePos := FStream.Size - (1024 + 22);
          BR := SeekStart;
        end;
        65536 : SeekStart := -1;
      end;
    end;
    if BR < 0 then
      SeekStart := -1;
    if FStream.Size < SeekStart then
      SeekStart := -1;
  until (Result) or (SeekStart = -1);
end;

function TXLSZipArchive.GetItems(Index: integer): TXLSZipFile;
begin
  Result := TXLSZipFile(inherited Items[Index]);
end;

function TXLSZipArchive.GetLocalEntry(Offset, NextOffset: Integer): TLocalFile;
var
  Byte_ : Array[0..4] of Byte;
  DataDescriptor: TDataDescriptor;
begin
  FillChar(Result,SizeOf(Result),0);
  FStream.Position := Offset;
  FStream.Read(Byte_,4);
  if  (Byte_[0] = $50) and (Byte_[1] = $4B) and (Byte_[2] = $03) and (Byte_[3] = $04) then begin
    FStream.Position := Offset;
    FStream.Read(Result,SizeOf(Result) - 2 * SizeOf(AnsiString) - SizeOf(integer));
    if Result.FilenameLength > 0 then begin
      SetLength(Result.FileName,Result.FilenameLength);
      FStream.Read(Result.FileName[1],Result.FilenameLength);
    end;
    if Result.ExtraFieldLength > 0 then begin
      SetLength(Result.ExtraField,Result.ExtraFieldLength);
      FStream.Read(Result.ExtraField[1],Result.ExtraFieldLength);
    end;
    if (Result.GeneralPurposeBitFlag and (1 shl 3)) > 0 then begin
      FStream.Read(DataDescriptor,SizeOf(TDataDescriptor));
      Result.Crc32            := DataDescriptor.Crc32;
      Result.CompressedSize   := DataDescriptor.CompressedSize;
      Result.UnCompressedSize := DataDescriptor.UnCompressedSize;
    end;
    Result.CompressedDataPos := FStream.Position;
  end
  else
    raise XLSRWException.Create('Unknown zip file');
end;

procedure TXLSZipArchive.LoadLocalHeaders;
Var
  X : Integer;
  LF: TLocalFile;
begin
  FHasBadEntries := False;
  for X := 0 To Count - 1 do begin
    LF := GetLocalEntry(Items[X].FCentralDirectoryFile.RelativeOffsetOfLocalHeader,0);
    Items[X].FLocalFile := LF;
    SetLength(LF.FileName,0);
    SetLength(LF.ExtraField,0);
    LF.CompressedDataPos := 0;
    if Items[X].FLocalFile.LocalFileHeaderSignature <> $04034b50 then
      FHasBadEntries := True;
  end;
end;

procedure TXLSZipArchive.OpenRead(AFilename: AxUCString);
begin
  Close;

  FOpenMode := xzomRead;
  FStream := TFileStream.Create(AFilename,fmOpenRead{$ifndef DELPHI_5},fmShareDenyWrite{$endif});

  FOwnsStream := True;

  ReadFile;
end;

procedure TXLSZipArchive.OpenRead(AStream: TStream);
begin
  Close;

  FOpenMode := xzomRead;
  FStream := AStream;

  FOwnsStream := False;

  ReadFile;
end;

function TXLSZipArchive.ParseCentralHeaders: Boolean;
var
  i     : Integer;
  Entry : TXLSZipFile;
  CDFile: TCentralDirectoryFile;
begin
  Result := False;
  try
    FStream.Position := FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
    for i := 0 to FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk - 1 do begin
      FillChar(CDFile,SizeOf(TCentralDirectoryFile),0);
      FStream.Read(CDFile,SizeOf(TCentralDirectoryFile) - 3 * SizeOf(AnsiString));
      Entry := Add;
      Entry.FDate := FileDateToDateTime(CDFile.LastModFileTimeDate);
      if (CDFile.GeneralPurposeBitFlag and 1) > 0 then
        Entry.FIsEncrypted := True
      else
        Entry.FIsEncrypted := False;
      if CDFile.FilenameLength > 0 then begin
        SetLength(CDFile.FileName,CDFile.FilenameLength);
        FStream.Read(CDFile.FileName[1],CDFile.FilenameLength)
      end;
      if CDFile.ExtraFieldLength > 0 then begin
        SetLength(CDFile.ExtraField,CDFile.ExtraFieldLength);
        FStream.Read(CDFile.ExtraField[1], CDFile.ExtraFieldLength);
      end;
      if CDFile.FileCommentLength > 0 then begin
        SetLength(CDFile.FileComment,CDFile.FileCommentLength);
        FStream.Read(CDFile.FileComment[1],CDFile.FileCommentLength);
      end;
      Entry.FIsFolder := (CDFile.ExternalFileAttributes and faDirectory) > 0;

      Entry.FCompressionType := ctUnknown;
      if (CDFile.CompressionMethod = 8) or (CDFile.CompressionMethod = 9) then begin
        case CDFile.GeneralPurposeBitFlag and 6 of
          0 : Entry.FCompressionType := ctNormal;
          2 : Entry.FCompressionType := ctMaximum;
          4 : Entry.FCompressionType := ctFast;
          6 : Entry.FCompressionType := ctSuperFast
        end;
      end;
      Entry.FCentralDirectoryFile := CDFile;
      SetLength(CDFile.FileName,0);
      SetLength(CDFile.ExtraField,0);
      SetLength(CDFile.FileComment,0);
   end;
   except
     Exit;
   end;
   Result := Count = FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk;
end;

function TXLSZipArchive.ParseLocalHeaders: Boolean;
Var
  Poz                 : Integer;
  NLE                 : Integer;
  Byte_               : Array[0..4] of Byte;
  LocalFile           : TLocalFile;
  DataDescriptor      : TDataDescriptor;
  Entry               : TXLSZipFile;
  CDFile              : TCentralDirectoryFile;
  CDSize              : Cardinal;
  L                   : Integer;
  NoMore              : Boolean;
begin
  Result               := False;
  FLocalHeaderNumFiles := 0;
  Clear;

  Poz    := 0;
  NLE    := 0;
  CDSize := 0;
  repeat
    NoMore      := True;
    FStream.Position := Poz;
    FStream.Read(Byte_,4);
    if (Byte_[0]  = $50) and (Byte_[1]  = $4B) and (Byte_[2]  = $03) and (Byte_[3]  = $04) then begin
      Result := True;
      Inc(FLocalHeaderNumFiles);
      NoMore := False;
      FStream.Position := Poz;
      FStream.Read(LocalFile,SizeOf(TLocalFile) - 2 * SizeOf(AnsiString) - SizeOf(Integer));
      if LocalFile.FilenameLength > 0 then begin
        SetLength(LocalFile.FileName,LocalFile.FilenameLength);
        FStream.Read(LocalFile.FileName[1],LocalFile.FilenameLength);
      end;
      if LocalFile.ExtraFieldLength > 0 then begin
        SetLength(LocalFile.ExtraField,LocalFile.ExtraFieldLength);
        FStream.Read(LocalFile.ExtraField[1],LocalFile.ExtraFieldLength);
      end;
      if (LocalFile.GeneralPurposeBitFlag and (1 shl 3)) > 0 then begin
        FStream.Read(DataDescriptor,SizeOf(TDataDescriptor));
        LocalFile.Crc32            := DataDescriptor.Crc32;
        LocalFile.CompressedSize   := DataDescriptor.CompressedSize;
        LocalFile.UncompressedSize := DataDescriptor.UncompressedSize;
      end;
      FStream.Position := FStream.Position+LocalFile.CompressedSize;

      FillChar(CDFile,SizeOf(TCentralDirectoryFile),0);
      CDFile.CentralFileHeaderSignature  := $02014B50;
      CDFile.VersionMadeBy               := 20;
      CDFile.VersionNeededToExtract      := LocalFile.VersionNeededToExtract;
      CDFile.GeneralPurposeBitFlag       := LocalFile.GeneralPurposeBitFlag;
      CDFile.CompressionMethod           := LocalFile.CompressionMethod;
      CDFile.LastModFileTimeDate         := LocalFile.LastModFileTimeDate;
      CDFile.Crc32                       := LocalFile.Crc32;
      CDFile.CompressedSize              := LocalFile.CompressedSize;
      CDFile.UncompressedSize            := LocalFile.UncompressedSize;
      CDFile.FilenameLength              := LocalFile.FilenameLength;
      CDFile.ExtraFieldLength            := LocalFile.ExtraFieldLength;
      CDFile.FileCommentLength           := 0;
      CDFile.DiskNumberStart             := 0;
      CDFile.InternalFileAttributes      := LocalFile.VersionNeededToExtract;
      CDFile.ExternalFileAttributes      := $00000020; // faArchive;
      CDFile.RelativeOffsetOfLocalHeader := Poz;
      CDFile.FileName                    := LocalFile.FileName;
      L := Length(CDFile.FileName);
      if L > 0 then begin
        if CDFile.FileName[L] = '/' then
          CDFile.ExternalFileAttributes := faDirectory;
      end;
      CDFile.ExtraField                  := LocalFile.ExtraField;
      CDFile.FileComment                 := '';

      Entry                              := Add;
      Entry.FDate                        := FileDateToDateTime(CDFile.LastModFileTimeDate);
      if (CDFile.GeneralPurposeBitFlag and 1) > 0 then
        Entry.FIsEncrypted    := True
      else
         Entry.FIsEncrypted   := False;
      Entry.FIsFolder         := (CDFile.ExternalFileAttributes and faDirectory) > 0;
      Entry.FCompressionType  := ctUnknown;
      if (CDFile.CompressionMethod=8) or (CDFile.CompressionMethod=9) then begin
        case CDFile.GeneralPurposeBitFlag and 6 of
          0 : Entry.FCompressionType := ctNormal;
          2 : Entry.FCompressionType := ctMaximum;
          4 : Entry.FCompressionType := ctFast;
          6 : Entry.FCompressionType := ctSuperFast
        end;
      end;
      Entry.FCentralDirectoryFile := CDFile;
      Poz         := FStream.Position;
      Inc(NLE);
      CDSize      := CDSize + Entry.GetCentralEntrySize;
    end;
  until NoMore;

  FEndOfCentralDir.EndOfCentralDirSignature        := $06054b50;
  FEndOfCentralDir.NumberOfThisDisk                := 0;
  FEndOfCentralDir.NumberOfTheDiskWithTheStart     := 0;
  FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk  := NLE;
  FEndOfCentralDir.SizeOfTheCentralDirectory       := CDSize;
  FEndOfCentralDir.OffsetOfStartOfCentralDirectory := FStream.Position;
  FEndOfCentralDir.ZipfileCommentLength            := 0;
end;

procedure TXLSZipArchive.ReadFile;
begin
  if FindCentralDirectory then begin
    if ParseCentralHeaders then begin
      LoadLocalHeaders;
    end
    else
      raise XLSRWException.Create('File is not a valid ZIP file');
  end
  else begin
    if ParseLocalHeaders then begin
      if Count > 0 then
        FHasBadEntries := True;
    end;
  end;
end;

procedure TXLSZipArchive.WriteCentralDirectory;
var
  i: integer;
begin
  FStream.Seek(0,soFromEnd);

  for i := 0 to Count - 1 do begin
    FStream.Write(Self.Items[i].FCentralDirectoryFile,SizeOf(Self.Items[i].FCentralDirectoryFile) - 3 * SizeOf(AnsiString));
    if Self.Items[i].FCentralDirectoryFile.FilenameLength > 0 Then
       FStream.Write(Self.Items[i].FCentralDirectoryFile.FileName[1],Self.Items[i].FCentralDirectoryFile.FilenameLength);
    if Self.Items[i].FCentralDirectoryFile.ExtraFieldLength > 0 Then
       FStream.Write(Self.Items[i].FCentralDirectoryFile.ExtraField[1],Self.Items[i].FCentralDirectoryFile.ExtraFieldLength);
    if Self.Items[i].FCentralDirectoryFile.FileCommentLength > 0 Then
       FStream.Write(Self.Items[i].FCentralDirectoryFile.FileComment[1],Self.Items[i].FCentralDirectoryFile.FileCommentLength);
  end;
  FEndOfCentralDirPos := FStream.Position;
  FEndOfCentralDir.SizeOfTheCentralDirectory := FEndOfCentralDirPos - FEndOfCentralDir.OffsetOfStartOfCentralDirectory;
  FStream.Write(FEndOfCentralDir, SizeOf(TEndOfCentralDir));
end;

procedure TXLSZipArchive.ClearEndOfCentralDir;
begin
  FEndOfCentralDir.EndOfCentralDirSignature        := $06054b50;
  FEndOfCentralDir.NumberOfThisDisk                := 0;
  FEndOfCentralDir.NumberOfTheDiskWithTheStart     := 0;
  FEndOfCentralDir.TotalNumberOfEntriesOnThisDisk  := 0;
  FEndOfCentralDir.TotalNumberOfEntries            := 0;
  FEndOfCentralDir.SizeOfTheCentralDirectory       := 0;
  FEndOfCentralDir.OffsetOfStartOfCentralDirectory := 0;
  FEndOfCentralDir.ZipfileCommentLength            := 0;
end;

{ TXLSZipFile }

procedure TXLSZipFile.Close;
begin
  if FDStream <> Nil then begin
    if FDStream is TDecompressionStream then
      FDStream.Free;
    FDStream := Nil;
  end;
end;

constructor TXLSZipFile.Create(AParent: TXLSZipArchive);
begin
  FParent := AParent;
end;

destructor TXLSZipFile.Destroy;
begin
  Close;
  inherited;
end;

function TXLSZipFile.GetCentralEntrySize: Cardinal;
begin
  Result := SizeOf(TCentralDirectoryFile) - 3 * SizeOf(AnsiString) +
                   FCentralDirectoryFile.FilenameLength +
                   FCentralDirectoryFile.ExtraFieldLength +
                   FCentralDirectoryFile.FileCommentLength;
end;

function TXLSZipFile.GetName: AxUCString;
begin
  Result := AxUCString(FLocalFile.FileName);
end;

//{$ifndef DELPHI_6}
//function TXLSZipFile.GetSize: int64;
//begin
//  Result := FLocalFile.UncompressedSize;
//end;
//{$endif}

procedure TXLSZipFile.OpenRead;
begin
  FParent.FStream.Seek(FLocalFile.CompressedDataPos,soFromBeginning);
  case FCentralDirectoryFile.CompressionMethod of
    0: FDStream := FParent.FStream;
    8: FDStream := TDecompressionStream.CreateZip(FParent.FStream,FLocalFile.CompressedSize);
  end;
//  FDStream.Seek(0,0);
end;

function TXLSZipFile.Read(var Buffer; Count: Integer): Longint;
var
  MaxSz: integer;
begin
  if FDStream <> Nil then begin
    if FCentralDirectoryFile.CompressionMethod = 8 then
      Result := TDecompressionStream(FDStream).ReadZip(Buffer,Count)
    else begin
      MaxSz := FLocalFile.CompressedSize - (FParent.FStream.Position - FLocalFile.CompressedDataPos) ;
      Result := FDStream.Read(Buffer,Min(Count,MaxSz));
    end;
  end
  else
    raise XLSRWException.Create('ZIP file is not open');
end;

{$ifdef DELPHI_5}
function TXLSZipFile.Seek(Offset: Longint; Origin: Word): Longint;
{$else}
function TXLSZipFile.Seek(const Offset: Int64; Origin: TSeekOrigin): Int64;
{$endif}
begin
// In D6 is the TStream.GetSize method not virtual and can not be overloaded.
// Instead D6 uses Seek in order to get the file size. As this is a compressed
// file, Seek is only partial implemented and don't work with D5,D6.
// The code below solves the problem. The first call to Seek is to the end
// of the file in order to get the file size. The second call is to the
// beginning and is ignored here. It's not possible to FDStream.Seek
// as the file pointer is not at the beginning of the file.
{$ifdef DELPHI_5}
  if FSeekCount > 0 then begin
    Dec(FSeekCount);
    Exit;
  end;

  if (Offset = 0) and (Origin = soFromEnd) then begin
    Result := FLocalFile.UncompressedSize;
    FSeekCount := 1;
  end
  else
{$endif}
{$ifdef DELPHI_6}
  if FSeekCount > 0 then begin
    Dec(FSeekCount);
    Exit;
  end;

  if (Offset = 0) and (Origin = soEnd) then begin
    Result := FLocalFile.UncompressedSize;
    FSeekCount := 1;
  end
  else
{$endif}
    Result := FDStream.Seek(Offset,Origin);
end;

function TXLSZipFile.Write(const Buffer; Count: Integer): Longint;
begin
  raise XLSRWException.Create('Not implemented, write');
end;

procedure TXLSZipFile._SaveToStream(AStream: TStream);
const
  ZBUFSZ = $FFF;
var
  Sz,SzRead: integer;
  Buf: PAnsiChar;
begin
  GetMem(Buf,ZBufSz);
  try
    OpenRead;

    Sz := FLocalFile.UncompressedSize;
    while Sz > 0 do begin
      SzRead := Read(Buf^,ZBUFSZ);
      AStream.WriteBuffer(Buf^,SzRead);
      Dec(Sz,SzRead);
    end;

    Close;
  finally
    FreeMem(Buf);
  end;
end;

end.
