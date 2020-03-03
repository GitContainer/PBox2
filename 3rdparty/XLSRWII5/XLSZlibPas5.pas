{Portable Network Graphics Delphi ZLIB linking  (16 May 2002) }

{This unit links ZLIB to pngimage unit in order to implement  }
{the library. It's now using the new ZLIB version, 1.1.4      }
{Note: The .obj files must be located in the subdirectory \obj}

unit XLSZlibPas5;

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface


uses SysUtils, Classes, Math;

{$H+}

{$ifndef DELPHI_7_OR_LATER}
type NativeInt = integer;
{$endif}

type

  TAlloc = function (AppData: Pointer; Items, Size: Integer): Pointer;
  TFree = procedure (AppData, Block: Pointer);

  // Internal structure.  Ignore.
  // NOT packed. Will not work in x64.
  TZStreamRec = record
    next_in: PAnsiChar;   // next input byte
    avail_in: Integer;    // number of bytes available at next_in
    total_in: Integer;    // total nb of input bytes read so far

    next_out: PAnsiChar;  // next output byte should be put here
    avail_out: Integer;   // remaining free space at next_out
    total_out: Integer;   // total nb of bytes output so far

    msg: PAnsiChar;       // last error message, NULL if no error
    internal: Pointer;    // not visible by applications

    zalloc: TAlloc;       // used to allocate the internal state
    zfree: TFree;         // used to free the internal state
    AppData: Pointer;     // private data object passed to zalloc and zfree

    data_type: Integer;   //  best guess about the data type: ascii or binary
    adler: Integer;       // adler32 value of the uncompressed data
    reserved: Integer;    // reserved for future use
  end;

{$ifdef MSWINDOWS}
function inflateInit_(var strm: TZStreamRec; version: PAnsiChar; recsize: Integer): Integer; forward;
function inflate(var strm: TZStreamRec; flush: Integer): Integer; forward;
function inflateEnd(var strm: TZStreamRec): Integer; forward;
function deflateInit_(var strm: TZStreamRec; level: Integer; version: PAnsiChar; recsize: Integer): Integer; forward;
function deflate(var strm: TZStreamRec; flush: Integer): Integer; forward;
function deflateEnd(var strm: TZStreamRec): Integer; forward;
{$else}
function inflateInit_(var strm: TZStreamRec; version: PAnsiChar; recsize: Integer): Integer;
function inflate(var strm: TZStreamRec; flush: Integer): Integer;
function inflateEnd(var strm: TZStreamRec): Integer;
function deflateInit_(var strm: TZStreamRec; level: Integer; version: PAnsiChar; recsize: Integer): Integer;
function deflate(var strm: TZStreamRec; flush: Integer): Integer;
function deflateEnd(var strm: TZStreamRec): Integer;
{$endif}

const
  zlib_version = '1.2.5';


const
  Z_NO_FLUSH      = 0;
  Z_PARTIAL_FLUSH = 1;
  Z_SYNC_FLUSH    = 2;
  Z_FULL_FLUSH    = 3;
  Z_FINISH        = 4;

  Z_OK            = 0;
  Z_STREAM_END    = 1;
  Z_NEED_DICT     = 2;
  Z_ERRNO         = (-1);
  Z_STREAM_ERROR  = (-2);
  Z_DATA_ERROR    = (-3);
  Z_MEM_ERROR     = (-4);
  Z_BUF_ERROR     = (-5);
  Z_VERSION_ERROR = (-6);

  Z_NO_COMPRESSION       =   0;
  Z_BEST_SPEED           =   1;
  Z_BEST_COMPRESSION     =   9;
  Z_DEFAULT_COMPRESSION  = (-1);

  Z_FILTERED            = 1;
  Z_HUFFMAN_ONLY        = 2;
  Z_DEFAULT_STRATEGY    = 0;

  Z_BINARY   = 0;
  Z_ASCII    = 1;
  Z_UNKNOWN  = 2;

  Z_DEFLATED = 8;

{$ifdef WIN64}
  z_errmsg: array[0..9] of PAnsiChar = (
{$else}
  _z_errmsg: array[0..9] of PAnsiChar = (
{$endif}
    'need dictionary',      // Z_NEED_DICT      (2)
    'stream end',           // Z_STREAM_END     (1)
    '',                     // Z_OK             (0)
    'file error',           // Z_ERRNO          (-1)
    'stream error',         // Z_STREAM_ERROR   (-2)
    'data error',           // Z_DATA_ERROR     (-3)
    'insufficient memory',  // Z_MEM_ERROR      (-4)
    'buffer error',         // Z_BUF_ERROR      (-5)
    'incompatible version', // Z_VERSION_ERROR  (-6)
    ''
  );

  type
  EZlibError = class(Exception);
  ECompressionError = class(EZlibError);
  EDecompressionError = class(EZlibError);

  TCustomZlibStream = class(TStream)
  private
    FStrm: TStream;
    FStrmPos: Integer;
    FOnProgress: TNotifyEvent;
    FZRec: TZStreamRec;
    FBuffer: array [Word] of AnsiChar;
  protected
    constructor Create(Strm: TStream);
    procedure Progress(Sender: TObject); dynamic;
    property OnProgress: TNotifyEvent read FOnProgress write FOnProgress;
  end;


  TCompressionLevel = (clNone, clFastest, clDefault, clMax);

  TCompressionStream = class(TCustomZlibStream)
  private
    FCRC32: longword;
    FTempBuffer: Pointer;
    FTempBufSz: integer;

    function GetCompressionRate: Single;
  public
    constructor Create(CompressionLevel: TCompressionLevel; Dest: TStream; AsZip: boolean);
    destructor Destroy; override;

    function  Read(var Buffer; Count: Longint): Longint; override;
    function  Write(const Buffer; Count: Longint): Longint; override;
    function  DoWrite(const Buffer; Count: Longint): Longint;
    procedure WriteStream(AStream: TStream);
    function  Seek(Offset: Longint; Origin: Word): Longint; override;
    procedure Flush;

    function  UncompressedSize: integer;

    property CRC32Val: longword read FCRC32 write FCRC32;
    property CompressionRate: Single read GetCompressionRate;
    property OnProgress;
  end;


  TDecompressionStream = class(TCustomZlibStream)
  protected
    FSize: integer;
    FZipMode: boolean;
  public
    constructor Create(Source: TStream);
    constructor CreateZip(Source: TStream; ASize: integer);
    destructor Destroy; override;
    function  ReadZip(var Buffer; Count: Longint): Longint;
    function  Read(var Buffer; Count: Longint): Longint; override;
    function  Read_(var Buffer; Count: Longint): Longint;
    function  Write(const Buffer; Count: Longint): Longint; override;
    function  Seek(Offset: Longint; Origin: Word): Longint; override;

    property ZipMode: boolean read FZipMode write FZipMode;
    property OnProgress;
  end;

procedure CompressBuf(const InBuf: Pointer; InBytes: Integer; out OutBuf: Pointer; out OutBytes: Integer);
procedure DecompressBuf(const InBuf: Pointer; InBytes: Integer; OutEstimate: Integer; out OutBuf: Pointer; out OutBytes: Integer);

implementation

const TEMP_BUFFER_SZ = $1FFF;

{$ifdef MSWINDOWS}
{$ifdef WIN64}
{$L obj\x64\adler32.obj}
{$L obj\x64\deflate.obj}
{$L obj\x64\infback.obj}
{$L obj\x64\inffast.obj}
{$L obj\x64\inflate.obj}
{$L obj\x64\inftrees.obj}
{$L obj\x64\trees.obj}
{$L obj\x64\compress.obj}
{$L obj\x64\crc32.obj}
{$else}
{$L obj\adler32.obj}
{$L obj\deflate.obj}
{$L obj\infback.obj}
{$L obj\inffast.obj}
{$L obj\inflate.obj}
{$L obj\inftrees.obj}
{$L obj\trees.obj}
{$L obj\compress.obj}
{$L obj\crc32.obj}
{$endif}
{$endif}

{$ifdef WIN64}
procedure memset(P: Pointer; B: Byte; count: Integer);cdecl;
{$else}
procedure _memset(P: Pointer; B: Byte; count: Integer);cdecl;
{$endif}
begin
  FillChar(P^, count, B);
end;

{$ifdef WIN64}
procedure memcpy(dest, source: Pointer; count: Integer);cdecl;
{$else}
procedure _memcpy(dest, source: Pointer; count: Integer);cdecl;
{$endif}
begin
  Move(source^, dest^, count);
end;


{$ifdef MSWINDOWS}
function adler32(adler: LongInt; const buf: PAnsiChar; len: Integer): LongInt; external;

// deflate compresses data
function deflateInit_ (var strm: TZStreamRec; level: Integer; version: PAnsiChar; recsize: Integer): Integer; external;
function deflateInit2_(var strm: TZStreamRec; level: Integer; method: integer; windowBits: integer; memLevel: integer; strategy: integer; version: PAnsiChar; recsize: Integer): integer; external;
//                                                                             -15                  8                  0
function deflate(var strm: TZStreamRec; flush: Integer): Integer; external;
function deflateEnd(var strm: TZStreamRec): Integer; external;

// inflate decompresses data
function inflateInit_(var strm: TZStreamRec; version: PAnsiChar; recsize: Integer): Integer; external;
function inflateInit2_(var strm: TZStreamRec; WindowBits: integer; version: PAnsiChar; recsize: Integer): Integer; external;
function inflate(var strm: TZStreamRec; flush: Integer): Integer; external;
function inflateEnd(var strm: TZStreamRec): Integer; external;
function inflateReset(var strm: TZStreamRec): Integer; external;

function crc32(crc: longword; buf: PAnsiChar; len: longword): longword; external;
{$else}
function inflateInit_(var strm: TZStreamRec; version: PAnsiChar; recsize: Integer): Integer;
begin
  Result := 0;
end;

function inflate(var strm: TZStreamRec; flush: Integer): Integer;
begin
  Result := 0;
end;

function inflateEnd(var strm: TZStreamRec): Integer;
begin
  Result := 0;
end;

function deflateInit_(var strm: TZStreamRec; level: Integer; version: PAnsiChar; recsize: Integer): Integer;
begin
  Result := 0;
end;

function deflate(var strm: TZStreamRec; flush: Integer): Integer;
begin
  Result := 0;
end;

function deflateEnd(var strm: TZStreamRec): Integer;
begin
  Result := 0;
end;

function deflateInit2_(var strm: TZStreamRec; level: Integer; method: integer; windowBits: integer; memLevel: integer; strategy: integer; version: PAnsiChar; recsize: Integer): integer;
begin
  Result := 0;
end;

function inflateInit2_(var strm: TZStreamRec; WindowBits: integer; version: PAnsiChar; recsize: Integer): Integer;
begin
  Result := 0;
end;

function crc32(crc: longword; buf: PAnsiChar; len: longword): longword;
begin
  Result := 0;
end;

function inflateReset(var strm: TZStreamRec): Integer;
begin
  Result := 0;
end;
{$endif}

function zcalloc(AppData: Pointer; Items, Size: Integer): Pointer;
begin
  GetMem(Result, Items*Size);
end;

procedure zcfree(AppData, Block: Pointer);
begin
  FreeMem(Block);
end;

function zlibAllocMem(AppData: Pointer; Items, Size: Integer): Pointer;
{$IFDEF MSWINDOWS}
  register;
{$ENDIF}
{$IFDEF LINUX}
  cdecl;
{$ENDIF}
begin
  Result := AllocMem(Items * Size);
end;

procedure zlibFreeMem(AppData, Block: Pointer);
{$IFDEF MSWINDOWS}
  register;
{$ENDIF}
{$IFDEF LINUX}
  cdecl;
{$ENDIF}
begin
  FreeMem(Block);
end;

function DCheck(code: Integer): Integer;
begin
  Result := code;
  if code < 0 then
    raise EDecompressionError.Create('ZLib decompress error');  //!!
end;

function CCheck(code: Integer): Integer;
begin
  Result := code;
  if code < 0 then
    raise ECompressionError.Create('ZLib decompress error'); //!!
end;

procedure DecompressBuf(const InBuf: Pointer; InBytes: Integer; OutEstimate: Integer; out OutBuf: Pointer; out OutBytes: Integer);
var
  strm: TZStreamRec;
  P: Pointer;
  BufInc: Integer;
begin
  FillChar(strm, sizeof(strm), 0);
  strm.zalloc := zlibAllocMem;
  strm.zfree := zlibFreeMem;
  BufInc := (InBytes + 255) and not 255;
  if OutEstimate = 0 then
    OutBytes := BufInc
  else
    OutBytes := OutEstimate;
  GetMem(OutBuf, OutBytes);
  try
    strm.next_in := InBuf;
    strm.avail_in := InBytes;
    strm.next_out := OutBuf;
    strm.avail_out := OutBytes;
    DCheck(inflateInit_(strm, zlib_version, sizeof(strm)));
    try
      while DCheck(inflate(strm, Z_FINISH)) <> Z_STREAM_END do
      begin
        P := OutBuf;
        Inc(OutBytes, BufInc);
        ReallocMem(OutBuf, OutBytes);
        strm.next_out := PAnsiChar(NativeInt(OutBuf) + (NativeInt(strm.next_out) - NativeInt(P)));
        strm.avail_out := BufInc;
      end;
    finally
      DCheck(inflateEnd(strm));
    end;
    ReallocMem(OutBuf, strm.total_out);
    OutBytes := strm.total_out;
  except
    FreeMem(OutBuf);
    raise
  end;
end;

procedure CompressBuf(const InBuf: Pointer; InBytes: Integer; out OutBuf: Pointer; out OutBytes: Integer);
var
  strm: TZStreamRec;
  P: Pointer;
begin
  FillChar(strm, sizeof(strm), 0);
  strm.zalloc := zlibAllocMem;
  strm.zfree := zlibFreeMem;
  OutBytes := ((InBytes + (InBytes div 10) + 12) + 255) and not 255;
  GetMem(OutBuf, OutBytes);
  try
    strm.next_in := InBuf;
    strm.avail_in := InBytes;
    strm.next_out := OutBuf;
    strm.avail_out := OutBytes;
    CCheck(deflateInit_(strm, Z_BEST_COMPRESSION, zlib_version, sizeof(strm)));
    try
      while CCheck(deflate(strm, Z_FINISH)) <> Z_STREAM_END do
      begin
        P := OutBuf;
        Inc(OutBytes, 256);
        ReallocMem(OutBuf, OutBytes);
        strm.next_out := PAnsiChar(NativeInt(OutBuf) + (NativeInt(strm.next_out) - NativeInt(P)));
        strm.avail_out := 256;
      end;
    finally
      CCheck(deflateEnd(strm));
    end;
    ReallocMem(OutBuf, strm.total_out);
    OutBytes := strm.total_out;
  except
    FreeMem(OutBuf);
    raise
  end;
end;

// TCustomZlibStream

constructor TCustomZLibStream.Create(Strm: TStream);
begin
  inherited Create;
  FStrm := Strm;
  FStrmPos := Strm.Position;
  FZRec.zalloc := zlibAllocMem;
  FZRec.zfree := zlibFreeMem;
end;

procedure TCustomZLibStream.Progress(Sender: TObject);
begin
  if Assigned(FOnProgress) then FOnProgress(Sender);
end;


// TCompressionStream

constructor TCompressionStream.Create(CompressionLevel: TCompressionLevel; Dest: TStream; AsZip: boolean);
const
  Levels: array [TCompressionLevel] of ShortInt = (Z_NO_COMPRESSION, Z_BEST_SPEED, Z_DEFAULT_COMPRESSION, Z_BEST_COMPRESSION);
begin
  inherited Create(Dest);

  GetMem(FTempBuffer,TEMP_BUFFER_SZ);

  FCRC32 := 0;
  FZRec.next_out := FBuffer;
  FZRec.avail_out := sizeof(FBuffer);
  if AsZip then
    CCheck(deflateInit2_(FZRec, Levels[CompressionLevel],8, -15, 8, 0,zlib_version, sizeof(FZRec)))
  else
    CCheck(deflateInit_(FZRec, Levels[CompressionLevel], zlib_version, sizeof(FZRec)));
end;

//constructor TCompressionStream.CreateZip(CompressionLevel: TCompressionLevel; Dest: TStream);
//const
//  Levels: array [TCompressionLevel] of ShortInt = (Z_NO_COMPRESSION, Z_BEST_SPEED, Z_DEFAULT_COMPRESSION, Z_BEST_COMPRESSION);
//begin
//  inherited Create(Dest);
//  FCRC32 := 0;
//  FZRec.next_out := FBuffer;
//  FZRec.avail_out := sizeof(FBuffer);
//  CCheck(deflateInit2_(FZRec, Levels[CompressionLevel],8, -15, 8, 0,zlib_version, sizeof(FZRec)));
//end;

destructor TCompressionStream.Destroy;
begin
  Flush;

  FZRec.next_in := nil;
  FZRec.avail_in := 0;
  try
    if FStrm.Position <> FStrmPos then FStrm.Position := FStrmPos;
    while (CCheck(deflate(FZRec, Z_FINISH)) <> Z_STREAM_END) and (FZRec.avail_out = 0) do
    begin
      FStrm.WriteBuffer(FBuffer, sizeof(FBuffer));
      FZRec.next_out := FBuffer;
      FZRec.avail_out := sizeof(FBuffer);
    end;
    if FZRec.avail_out < sizeof(FBuffer) then
      FStrm.WriteBuffer(FBuffer, sizeof(FBuffer) - FZRec.avail_out);
  finally
    deflateEnd(FZRec);
  end;

  FreeMem(FTempBuffer);

  inherited Destroy;
end;

function TCompressionStream.Read(var Buffer; Count: Longint): Longint;
begin
  raise ECompressionError.Create('Invalid stream operation');
end;

function TCompressionStream.DoWrite(const Buffer; Count: Longint): Longint;
begin
  FZRec.next_in := @Buffer;
  FZRec.avail_in := Count;
  if FStrm.Position <> FStrmPos then FStrm.Position := FStrmPos;
  while (FZRec.avail_in > 0) do begin
    CCheck(deflate(FZRec, 0));
    if FZRec.avail_out = 0 then begin
      FStrm.WriteBuffer(FBuffer, sizeof(FBuffer));
      FZRec.next_out := FBuffer;
      FZRec.avail_out := sizeof(FBuffer);
      FStrmPos := FStrm.Position;
      Progress(Self);
    end;
  end;
  Result := Count;
  FCRC32 := crc32(FCRC32,@Buffer,Count);
end;

procedure TCompressionStream.Flush;
begin
  if FTempBufSz > 0 then
    DoWrite(FTempBuffer^,FTempBufSz);
  FTempBufSz := 0;
end;

function TCompressionStream.Write(const Buffer; Count: Integer): Longint;
var
  P: Pointer;
begin
  Result := Count;
  if Count > TEMP_BUFFER_SZ then begin
    if FTempBufSz > 0 then begin
      DoWrite(FTempBuffer^,FTempBufSz);
      FTempBufSz := 0;
    end;
    DoWrite(Buffer,Count);
  end
  else if (FTempBufSz + Count) > TEMP_BUFFER_SZ then begin
    DoWrite(FTempBuffer^,FTempBufSz);
    FTempBufSz := Count;
    System.Move(Buffer,FTempBuffer^,FTempBufSz);
  end
  else begin
    P := Pointer(NativeInt(FTempBuffer) + FTempBufSz);
    System.Move(Buffer,P^,Count);
    Inc(FTempBufSz,Count);
  end;
end;

procedure TCompressionStream.WriteStream(AStream: TStream);
var
  Sz,Sz2: integer;
  Buf: array [Word] of AnsiChar;
begin
  AStream.Seek(0,soFromBeginning);
  Sz := AStream.Size;
  while Sz > 0 do begin
    Sz2 := AStream.Read(Buf,SizeOf(Buf));
    DoWrite(Buf,Sz2);
    Dec(Sz,Sz2);
  end;
end;

function TCompressionStream.Seek(Offset: Longint; Origin: Word): Longint;
begin
  if (Offset = 0) and (Origin = soFromCurrent) then
    Result := FZRec.total_in
  else
    raise ECompressionError.Create('Invalid stream operation');
end;

function TCompressionStream.UncompressedSize: integer;
begin
  Result := FZRec.total_in;
end;

function TCompressionStream.GetCompressionRate: Single;
begin
  if FZRec.total_in = 0 then
    Result := 0
  else
    Result := (1.0 - (FZRec.total_out / FZRec.total_in)) * 100.0;
end;


// TDecompressionStream

constructor TDecompressionStream.Create(Source: TStream);
begin
  inherited Create(Source);
  FZRec.next_in := FBuffer;
  FZRec.avail_in := 0;
  DCheck(inflateInit_(FZRec, zlib_version, sizeof(FZRec)));
end;

constructor TDecompressionStream.CreateZip(Source: TStream; ASize: integer);
begin
  inherited Create(Source);
  FSize := ASize;
  FZRec.next_in := FBuffer;
  FZRec.avail_in := 0;
  DCheck(inflateInit2_(FZRec, -15,zlib_version, sizeof(FZRec)));
end;

destructor TDecompressionStream.Destroy;
begin
  FStrm.Seek(-FZRec.avail_in, 1);
  inflateEnd(FZRec);
  inherited Destroy;
end;

function TDecompressionStream.Read(var Buffer; Count: Longint): Longint;
begin
  if FZipMode then
    Result := ReadZip(Buffer,Count)
  else
    Result := Read_(Buffer,Count);
end;

function TDecompressionStream.ReadZip(var Buffer; Count: Integer): Longint;
begin
  FZRec.next_out := @Buffer;
  FZRec.avail_out := Count;
  if FStrm.Position <> FStrmPos then FStrm.Position := FStrmPos;
  while (FZRec.avail_out > 0) do
  begin
    if FZRec.avail_in = 0 then
    begin
      FZRec.avail_in := FStrm.Read(FBuffer, Min(SizeOf(FBuffer),FSize));
      Dec(FSize,FZRec.avail_in);
      if FZRec.avail_in = 0 then
      begin
        Result := Count - FZRec.avail_out;
        Exit;
      end;
	    FZRec.next_in := FBuffer;
      FStrmPos := FStrm.Position;
      Progress(Self);
    end;
    CCheck(inflate(FZRec, 0));
  end;
  Result := Count;
end;

function TDecompressionStream.Read_(var Buffer; Count: Integer): Longint;
begin
  FZRec.next_out := @Buffer;
  FZRec.avail_out := Count;
  if FStrm.Position <> FStrmPos then FStrm.Position := FStrmPos;
  while (FZRec.avail_out > 0) do
  begin
    if FZRec.avail_in = 0 then
    begin
      FZRec.avail_in := FStrm.Read(FBuffer, sizeof(FBuffer));
      if FZRec.avail_in = 0 then
      begin
        Result := Count - FZRec.avail_out;
        Exit;
      end;
	    FZRec.next_in := FBuffer;
      FStrmPos := FStrm.Position;
      Progress(Self);
    end;
    CCheck(inflate(FZRec, 0));
  end;
  Result := Count;
end;

function TDecompressionStream.Write(const Buffer; Count: Longint): Longint;
begin
  raise EDecompressionError.Create('Invalid stream operation');
end;

function TDecompressionStream.Seek(Offset: Longint; Origin: Word): Longint;
var
  I: Integer;
  Buf: array [0..4095] of Char;
begin
//  if Offset = $FFFFFFF then begin
//    FZipMode := True;
//    Exit;
//  end;

  if (Offset = 0) and (Origin = soFromBeginning) then
  begin
    DCheck(inflateReset(FZRec));
    FZRec.next_in := FBuffer;
    FZRec.avail_in := 0;
    FStrm.Position := 0;
    FStrmPos := 0;
  end
  else if ( (Offset >= 0) and (Origin = soFromCurrent)) or
          ( ((Offset - FZRec.total_out) > 0) and (Origin = soFromBeginning)) then
  begin
    if Origin = soFromBeginning then Dec(Offset, FZRec.total_out);
    if Offset > 0 then
    begin
      for I := 1 to Offset div sizeof(Buf) do
        ReadBuffer(Buf, sizeof(Buf));
      ReadBuffer(Buf, Offset mod sizeof(Buf));
    end;
  end
  else
    raise EDecompressionError.Create('Invalid stream operation');
  Result := FZRec.total_out;
end;


end.

