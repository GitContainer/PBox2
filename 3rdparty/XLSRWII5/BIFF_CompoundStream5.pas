unit BIFF_CompoundStream5;

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

uses Classes, SysUtils, Math, ContNrs,
     XLSUtils5;

const SID_Free       = -1;
const SID_EndOfChain = -2;
const SID_SAT        = -3;
const SID_MSAT       = -4;

const DirEntry_Empty     = 0;
const DirEntry_Storage   = 1;
const DirEntry_Stream    = 2;
const DirEntry_LockBytes = 3;
const DirEntry_Property  = 4;
const DirEntry_Root      = 5;

type TStorageOpenMode = (somClosed,somRead,somWrite);

type TCompoundDocHeader = packed record
     Id: array[0..7] of byte;
     DocId: array[0..15] of byte;
     RevisionNumber: word;
     VersionNumber: word;
     ByteOrder: word;
     SectorSize: word;
     ShortSectorSize: word;
     Unused1: array[0..9] of byte;
     TotalSectors: longword;
     DirSID: longint;
     Unused2: longword;
     MinStdStream: longword;
     ShortSecSID: longint;
     ShortSecCount: longword;
     MasterSecSID: longint;
     MasterSecCount: longword;
     FirstMasterSecSIDs: array[0..108] of longint;
     end;

type PDirEntry = ^TDirEntry;
     TDirEntry = packed record
     Name: array[0..31] of WideChar;
     NameSize: word;
     EntryType: byte;
     NodeColor: byte;
     LeftChildDID: longint;
     RightChildDID: longint;
     RootDID: longint;
     Id: array[0..15] of byte;
     UserFlags: longword;
     CreationTimestamp: array[0..1] of longword;
     LastModTimestamp: array[0..1] of longword;
     FirstSID: integer;
     StreamSize: longword;
     Unused: longword;
     end;

type TCacheFile = class(TObjectList)
private
     FName: WideString;
     FData: PByteArray;
     FDataSize: integer;
     FIsDirectory: boolean;
     FFiles: TCacheFile;

     function GetItems(Index: integer): TCacheFile;
     procedure SetData(const Value: PByteArray);
public
     function Add: TCacheFile;
     procedure Clear; override;

     property Items[Index: integer]: TCacheFile read GetItems; default;
     property Name: WideString read FName write FName;
     property Data: PByteArray read FData write SetData;
     property DataSize: integer read FDataSize write FDataSize;
     property IsDirectory: boolean read FIsDirectory write FIsDirectory;
     end;

     TCompoundStorage = class;

     TCompFile = class(TObjectList)
private
     FStorage: TCompoundStorage;
     FFiles: TCompFile;
     FName: WideString;
     FDataSize: integer;
     FIsDirectory: boolean;
     FSID: integer;
     FSize: integer;
     FCurrPos: integer;
     FWritingSmallSec: boolean;
     FDID: integer;
     FLeftDID: integer;
     FRightDID: integer;
     FRootDID: integer;
     FCreationTimestamp: array[0..1] of longword;
     FLastModTimestamp: array[0..1] of longword;

     function  GetItems(Index: integer): TCompFile;
protected
     function  DecompressVBA(var Buf: PByteArray; Offset: integer): integer;
public
     constructor Create(CompStorage: TCompoundStorage);
     destructor Destroy; override;

     function  Add(Name: WideString; Size: integer = 0): TCompFile;
     function  FindStream(Name: WideString; SearchSubDirs: boolean): TCompFile;
     procedure Cache(var Data: PByteArray);
     procedure Open;
     function  Read(const Buffer; Count: Longint): longint;
     function  Write(const Buffer; Count: Longint): longint;
     function  Seek(Offset: Longint; Origin: Word): longint;
     procedure Close;

     property Items[Index: integer]: TCompFile read GetItems; default;
     property IsDirectory: boolean read FIsDirectory write FIsDirectory;
     property Files: TCompFile read FFiles;
     property Name: WideString read FName;
     property DataSize: integer read FDataSize;
     property DID: integer read FDID;
     property LeftDID: integer read FLeftDID;
     property RightDID: integer read FRightDID;
     property RootDID: integer read FRootDID;
     property Size: integer read FSize;
     end;

     PLongintArray = ^TLongintArray;
     TLongintArray = array[0..16383] of Longint;

     TCompoundStorage = class(TObject)
private
     FOpenMode: TStorageOpenMode;
     FStream: TStream;
     FOwnsStream: boolean;
     FHeader: TCompoundDocHeader;
     FSectorSize: integer;
     FShortSectorSize: integer;
     FSAT: PLongintArray;
     FSSAT: PLongintArray;
     FFiles: TCompFile;
     FShortSectors: PByteArray;
     FShortSectorsSize: integer;
     FWBuf: PByteArray;
     FMaxSID: integer;
     FMaxSSID: integer;
     FLastSID: integer;
     FLastPos: integer;

     function  SIDToPos(SID: integer): integer;
     function  GetSID: integer;
     function  AllocSIDs(Count: integer; IsMSAT: boolean = False): integer;
     function  AllocSSIDs(Count: integer): integer;
     procedure WriteShort(const Buffer; Count: Longint);
     procedure WriteSAT;
     procedure WriteDirectories;
     procedure WriteShortSectors;

     procedure ReadSAT;
     procedure ReadShortSectors;
     procedure ReadDirectories;
     function  ReadStream(SS: boolean; SID,Position,Count: integer; Buf: PByteArray): integer;
     procedure Clear;
     procedure SetDefault;
     procedure PadSector;
     procedure PadSectorNeg;
     procedure SetFiles(const Value: TCompFile);
public
     constructor Create;
     destructor Destroy; override;

     procedure LoadFromStream(Stream: TStream);
     procedure LoadFromFile(Filename: WideString);
     procedure ReadCache(CacheFile: TCacheFile; ExceptThese: array of WideString);
     procedure WriteCache(CacheFile: TCacheFile);
     procedure SaveToStream(Stream: TStream);
     procedure SaveToFile(Filename: WideString);
     procedure Close;
     procedure CloseFile;
     function  IsStorage(Filename: WideString): boolean;

     property Files: TCompFile read FFiles write SetFiles;
     end;

implementation

const CompoundId: array[0..7] of byte = ($D0,$CF,$11,$E0,$A1,$B1,$1A,$E1);

const READ_FILEMODE  = fmShareDenyNone;
const WRITE_FILEMODE = fmShareDenyNone;

function CompareDir(Item1, Item2: TCompFile): Integer;
var
  i: integer;
begin
  Result := Length(Item1.Name) - Length(Item2.Name);
  if Result = 0 then begin
    Result := WideCompareStr(Item1.Name,Item2.Name);

    // TODO: Find a generic comparision that works.
    for i := 1 to Length(Item1.Name) do begin
      if CharInSet(Item1.Name[i],['a'..'z','A'..'Z']) and (Item2.Name[i] = '_') then
        Result := -1
      else if (Item1.Name[i] = '_') and CharInSet(Item2.Name[i],['a'..'z','A'..'Z']) then
        Result := 1
      else
//        Result := Word(Item1.Name[i]) - Word(Item2.Name[i]);
        Result := WideCompareStr(Item1.Name[i],Item2.Name[i]);
      if Result <> 0 then
        Exit;
    end;

  end;
end;

{ TCompoundStorage }

procedure TCompoundStorage.Clear;
begin
  if FFiles <> Nil then
    FFiles.Clear;
  FreeMem(FSAT);
  FSAT := Nil;
  FreeMem(FSSAT);
  FSSAT := Nil;
  if FOwnsStream then
    FStream.Free;
  FStream := Nil;
  FOwnsStream := False;
  FOpenMode := somClosed;
  FreeMem(FShortSectors);
  FShortSectors := Nil;
  FShortSectorsSize := 0;
  FMaxSID := 0;
  FMaxSSID := 0;

  SetDefault;
end;

constructor TCompoundStorage.Create;
begin
  FFiles := TCompFile.Create(Self);
  Clear;
  SetDefault;
end;

destructor TCompoundStorage.Destroy;
begin
  Clear;
  FreeMem(FWBuf);
  FFiles.Free;
  inherited;
end;

procedure TCompoundStorage.LoadFromFile(Filename: WideString);
begin
  FStream := TFileStream.Create(Filename,fmOpenRead,READ_FILEMODE);
  LoadFromStream(FStream);
  FOwnsStream := True;
end;

procedure TCompoundStorage.LoadFromStream(Stream: TStream);
begin
  if FOpenMode <> somClosed then
    raise Exception.Create('Storage is open');
  Clear;
  FOwnsStream := False;
  FStream := Stream;
  FOpenMode := somRead;
  FLastSID := -1;
  FLastPos := -1;

  Stream.Read(FHeader,SizeOf(TCompoundDocHeader));
  if not CompareMem(@FHeader.Id[0],@CompoundId[0],Length(CompoundId)) then
    raise Exception.Create('The file is not a Compound File');
  if FHeader.ByteOrder <> $FFFE then
    raise Exception.Create('File uses big-endian byte order');

  FSectorSize := 1 shl FHeader.SectorSize;
  FShortSectorSize := 1 shl FHeader.ShortSectorSize;
  ReAllocMem(FWBuf,FHeader.MinStdStream);

  ReadSAT;
  ReadDirectories;
  ReadShortSectors;
end;

procedure TCompoundStorage.ReadCache(CacheFile: TCacheFile; ExceptThese: array of WideString);

function InExcept(S: WideString): boolean;
var
  i: integer;
begin
  Result := False;
  S := AnsiUppercase(S);
  for i := 0 to High(ExceptThese) do begin
    Result := S = AnsiUppercase(ExceptThese[i]);
    if Result then
      Exit;
  end;
end;

procedure ReadFiles(Cache: TCacheFile; CFile: TCompFile);
var
  i: integer;
  D: PByteArray;
  CF: TCacheFile;
begin
  for i := 0 to CFile.Count - 1 do begin
    if not InExcept(CFile[i].FName) then begin
      CF := Cache.Add;
      D := Nil;
      if not CFile[i].FIsDirectory then
        CFile[i].Cache(D);
      CF.Data := D;
      CF.Name := CFile[i].FName;
      CF.DataSize := CFile[i].FDataSize;
      CF.IsDirectory := CFile[i].FIsDirectory;
      if CFile[i].Count > 0 then
        ReadFiles(CF,CFile[i]);
    end;
  end;
end;

begin
  CacheFile.Clear;
  // Root entry
  CacheFile.Name := FFiles.Name;
  CacheFile.IsDirectory := True;
  ReadFiles(CacheFile,FFiles);
end;

procedure TCompoundStorage.WriteCache(CacheFile: TCacheFile);

procedure WriteFile(CFile: TCompFile; Cache: TCacheFile);
var
  i: integer;
  CF: TCompFile;
begin
  for i := 0 to Cache.Count - 1 do begin
    CF := CFile.Add(Cache[i].FName);
    CF.FIsDirectory := Cache[i].FIsDirectory;
    if not CF.FIsDirectory then begin
      CF.Open;
      CF.Write(Cache[i].FData^,Cache[i].FDataSize);
      CF.Close;
    end;
    if Cache[i].Count > 0 then
      WriteFile(CF,Cache[i]);
  end;
end;

begin
  if (CacheFile.FName <> '' ) and (CacheFile.FName <> 'Root Entry') then
    raise Exception.Create('First cache file is not Root Entry');
  WriteFile(FFiles,CacheFile);
end;

procedure TCompoundStorage.ReadDirectories;
var
  i: integer;
  pDir: PDirEntry;
  SID: integer;
  Dirs: TList;

procedure ReadDir;
begin
  New(pDir);
  FStream.Read(pDir^,SizeOf(TDirEntry));

  if not (pDir.EntryType in [DirEntry_Empty,DirEntry_Storage,DirEntry_Stream,DirEntry_Root]) then
    raise Exception.Create('Unhandled entry in directory');
  Dirs.Add(pDir);
end;

function GetDirName(DID: integer): WideString;
var
  i,L: integer;
begin
  L := (PDirEntry(Dirs[DID]).NameSize div 2) - 1;
  SetLength(Result,L);
  for i := 0 to L - 1 do
    Result[i + 1] := PDirEntry(Dirs[DID]).Name[i];
end;

procedure AddDir(Parent: TCompFile; DID: integer);
var
  pDir: PDirEntry;
  CompFile: TCompFile;
begin
  pDir := PDirEntry(Dirs[DID]);
  if pDir.EntryType <> DirEntry_Empty then begin
    CompFile := Parent.Add(GetDirName(DID),pDir.StreamSize);
    CompFile.FSize := pDir.StreamSize;
    CompFile.FIsDirectory := pDir.EntryType = DirEntry_Storage;
    CompFile.FSID := pDir.FirstSID;
    CompFile.FDID := DID;
    CompFile.FLeftDID := pDir.LeftChildDID;
    CompFile.FRightDID := pDir.RightChildDID;
    CompFile.FRootDID := pDir.RootDID;
    CompFile.FCreationTimestamp[0] := pDir.CreationTimestamp[0];
    CompFile.FCreationTimestamp[1] := pDir.CreationTimestamp[1];
    CompFile.FLastModTimestamp[0] := pDir.LastModTimestamp[0];
    CompFile.FLastModTimestamp[1] := pDir.LastModTimestamp[1];
    if pDir.LeftChildDID >= 0 then
      AddDir(Parent,pDir.LeftChildDID);
    if pDir.RightChildDID >= 0 then
      AddDir(Parent,pDir.RightChildDID);
    if pDir.RootDID >= 0 then begin
      AddDir(CompFile,pDir.RootDID);
      CompFile.Sort(@CompareDir);
    end;
  end;
end;

begin
  SID := FHeader.DirSID;
  Dirs := TList.Create;
  try
    repeat
      FStream.Seek(SIDToPos(SID),soFromBeginning);
      for i := 1 to FSectorSize div SizeOf(TDirEntry) do
        ReadDir;
      SID := FSAT[SID];
    until (SID = SID_EndOfChain);
    FFiles.FName := 'Root Entry';
    FFiles.FIsDirectory := True;
    FFiles.FSID := PDirEntry(Dirs[0]).FirstSID;
    FFiles.FDataSize := 0;
    if PDirEntry(Dirs[0]).RootDID >= 0 then
      AddDir(FFiles,PDirEntry(Dirs[0]).RootDID);
    FFiles.Sort(@CompareDir);
  finally
    for i := 0 to Dirs.Count - 1 do
      FreeMem(Dirs[i]);
    Dirs.Free;
  end;
end;

procedure TCompoundStorage.ReadSAT;
var
  i,SID: integer;
begin
  if (FHeader.MasterSecCount > 0) or (FHeader.MasterSecSID <> -2) then
    raise Exception.Create('Unhandled: more than 109 SID');
  for i := 0 to 108 do begin
    if FHeader.FirstMasterSecSIDs[i] < 0 then
      Break;
    ReAllocMem(FSAT,FSectorSize * (i + 1));
    FStream.Seek(SIDToPos(FHeader.FirstMasterSecSIDs[i]),soFromBeginning);
    FStream.Read(FSAT[(i * FSectorSize) div 4],FSectorSize);
  end;
  if FHeader.ShortSecCount > 0 then begin
    SID := FHeader.ShortSecSID;
    i := 0;
    repeat
      ReAllocMem(FSSAT,FSectorSize * (i + 1));
      FStream.Seek(SIDToPos(SID),soFromBeginning);
      FStream.Read(FSSAT[(i * FSectorSize) div 4],FSectorSize);
      Inc(i);
      SID := FSAT[SID];
    until (SID = SID_EndOfChain);
  end;
end;

procedure TCompoundStorage.ReadShortSectors;
var
  Count: integer;
  P: PByteArray;
  SID: integer;
begin
  Count := 0;
  SID := FFiles.FSID;
  while SID <> SID_EndOfChain do begin
    ReAllocMem(FShortSectors,(Count + 1) * FSectorSize);
    FStream.Seek(SIDToPos(SID),soFromBeginning);
    P := PByteArray(Integer(FShortSectors) + Count * FSectorSize);
    FStream.Read(P^,FSectorSize);
    SID := FSAT[SID];
    Inc(Count);
  end;
end;

function TCompoundStorage.ReadStream(SS: boolean; SID,Position,Count: integer; Buf: PByteArray): integer;

procedure ReadShortSectors;
var
  P,pSec: PByteArray;
begin
  Result := 0;
  P := Buf;
  while (Position >= FShortSectorSize) and (SID <> SID_EndOfChain) do begin
    Dec(Position,FShortSectorSize);
    SID := FSSAT[SID];
  end;
  if SID = SID_EndOfChain then
    Exit;
  while (SID <> SID_EndOfChain) and (Count > 0) do begin
    pSec := PByteArray(Integer(FShortSectors) + Position + SID * FShortSectorSize);
    if Count < (FShortSectorSize - Position) then begin
      Move(pSec^,P^,Count);
      Inc(Result,Count);
    end
    else begin
      Move(pSec^,P^,FShortSectorSize - Position);
      Inc(Result,FShortSectorSize - Position);
    end;
    P := PByteArray(Integer(P) + (FShortSectorSize - Position));
    Dec(Count,FShortSectorSize - Position);
    Position := 0;
    SID := FSSAT[SID];
  end;
end;

procedure ReadSectors;
var
  P: PByteArray;
begin
  Result := 0;
  P := Buf;
{
  if (FLastSID >= 0) and (Position >= FLastPos) and (Position < ((FLastPos div FSectorSize) + 1) * FSectorSize) then begin
    Dec(Position,(FLastPos div FSectorSize) * FSectorSize);
    SID := FLastSID;
  end
  else begin
}
    while (Position >= FSectorSize) and (SID <> SID_EndOfChain) do begin
      Dec(Position,FSectorSize);
      SID := FSAT[SID];
    end;
//  end;
  if SID = SID_EndOfChain then
    Exit;
  FLastPos := Position + Count;
  while (SID <> SID_EndOfChain) and (Count > 0) do begin
    FStream.Seek(SIDToPos(SID) + Position,soFromBeginning);
    if Count < (FSectorSize - Position) then
      Inc(Result,FStream.Read(P^,Count))
    else
      Inc(Result,FStream.Read(P^,FSectorSize - Position));
    P := PByteArray(Integer(P) + (FSectorSize - Position));
    Dec(Count,FSectorSize - Position);
    Position := 0;
    SID := FSAT[SID];
  end;
  FLastSID := SID - 1;
end;

begin
  if SS then
    ReadShortSectors
  else
    ReadSectors;
end;

function TCompoundStorage.SIDToPos(SID: integer): integer;
begin
  Result := SizeOf(TCompoundDocHeader) + SID * FSectorSize;
end;

procedure TCompoundStorage.SetDefault;
var
  i: integer;
begin
  FillChar(FHeader,SizeOf(TCompoundDocHeader),#0);
  for i := 0 to High(FHeader.FirstMasterSecSIDs) do
    FHeader.FirstMasterSecSIDs[i] := SID_Free;
  Move(CompoundId[0],FHeader.Id[0],SizeOf(CompoundId));
  FHeader.RevisionNumber := $003E;
  FHeader.VersionNumber := $0003;
  FHeader.ByteOrder := $FFFE;
  FHeader.SectorSize := 9; // 512 bytes
  FSectorSize := 1 shl FHeader.SectorSize;
  FHeader.ShortSectorSize := 6; // 64 bytes
  FShortSectorSize := 1 shl FHeader.ShortSectorSize;
  FHeader.MinStdStream := 4096;
  FHeader.ShortSecSID := SID_EndOfChain;
  FHeader.MasterSecSID := SID_EndOfChain;
  ReAllocMem(FWBuf,FHeader.MinStdStream);
end;

procedure TCompoundStorage.Close;
var
  i: integer;
begin
  if FOpenMode = somWrite then begin
    for i := 0 to FFiles.Count - 1 do begin
      if not FFiles[i].IsDirectory then
        FFiles[i].Close;
    end;

    WriteShortSectors;
    WriteDirectories;
    WriteSAT;

    FStream.Seek(0,soFromBeginning);
    FHeader.TotalSectors := Ceil((FMaxSID * 4) / FSectorSize);
    FStream.Write(FHeader,SizeOf(TCompoundDocHeader));
  end;
  FOpenMode := somClosed;
end;

procedure TCompoundStorage.SaveToFile(Filename: WideString);
begin
  FStream := TFileStream.Create(Filename,fmCreate,WRITE_FILEMODE);
  SaveToStream(FStream);
  FOwnsStream := True;
end;

procedure TCompoundStorage.SaveToStream(Stream: TStream);
begin
  if FOpenMode <> somClosed then
    raise Exception.Create('Storage is open');
//  Clear;
  FreeMem(FShortSectors);
  FShortSectors := Nil;
  FShortSectorsSize := 0;
  FOpenMode := somWrite;
  FStream := Stream;
  FOwnsStream := False;

  FStream.Write(FHeader,SizeOf(TCompoundDocHeader));
end;

function TCompoundStorage.AllocSIDs(Count: integer; IsMSAT: boolean = False): integer;
var
  i: integer;
begin
  Result := FMaxSID;
  if Count <= 0 then
    Exit;
  ReAllocMem(FSAT,(FMaxSID + Count) * SizeOf(longint));
  for i := 0 to Count - 2 do begin
{
    if IsMSAT then
      FSAT[FMaxSID + i] := SID_SAT
    else
}
      FSAT[FMaxSID + i] := FMaxSID + i + 1;
  end;
  if IsMSAT then
    FSAT[FMaxSID + Count - 1] := SID_SAT
  else
    FSAT[FMaxSID + Count - 1] := SID_EndOfChain;
  Inc(FMaxSID,Count);
end;

function TCompoundStorage.GetSID: integer;
begin
  Result := FMaxSID;
end;

procedure TCompoundStorage.WriteShort(const Buffer; Count: Integer);
var
  P: PByteArray;
  PadSz: integer;
begin
  PadSz := FShortSectorSize - ((FShortSectorsSize + Count) mod FShortSectorSize);
  if PadSz = FShortSectorSize then
    PadSz := 0;

  ReAllocMem(FShortSectors,FShortSectorsSize + Count + PadSz);
  P := PByteArray(Integer(FShortSectors) + FShortSectorsSize);
  Move(Buffer,P^,Count);
  P := PByteArray(Integer(FShortSectors) + FShortSectorsSize + Count);
  Inc(FShortSectorsSize,Count + PadSz);
  FillChar(P^,PadSz,#0);
end;

procedure TCompoundStorage.WriteDirectories;
var
  i,Sz: integer;
  DID: integer;
  Dir: TDirEntry;

procedure AssignDID(CFile: TCompFile);
var
  i: integer;
begin
  for i := 0 to CFile.Count - 1 do begin
    Inc(DID);
    CFile[i].FRootDID := -1;
    CFile[i].FLeftDID := -1;
    CFile[i].FRightDID := -1;
    if CFile[i].IsDirectory then begin
      if CFile[i].Count > 0 then
        CFile[i].FRootDID := DID;
    end
    else begin
      if i < (CFile.Count - 1) then
        CFile[i].FRightDID := DID;
    end;
    if CFile[i].Count > 0 then begin
      AssignDID(CFile[i]);
      if i < (CFile.Count - 1) then
        CFile[i].FRightDID := DID;
    end;
  end;
end;

procedure WriteDir(CFile: TCompFile);
var
  i: integer;
begin
  for i := 0 to CFile.Count - 1 do begin
    FillChar(Dir,SizeOf(TDirEntry),#0);
    Dir.RootDID := CFile[i].FRootDID;
    Dir.LeftChildDID := CFile[i].FLeftDID;
    Dir.RightChildDID := CFile[i].FRightDID;
    Dir.NodeColor := 1;
    Dir.CreationTimestamp[0] := CFile[i].FCreationTimestamp[0];
    Dir.CreationTimestamp[1] := CFile[i].FCreationTimestamp[1];
    Dir.LastModTimestamp[0] := CFile[i].FLastModTimestamp[0];
    Dir.LastModTimestamp[1] := CFile[i].FLastModTimestamp[1];
    if CFile[i].IsDirectory then begin
      Dir.EntryType := DirEntry_Storage;
      Dir.FirstSID := 0;
      Dir.StreamSize := 0;
    end
    else begin
      Dir.EntryType := DirEntry_Stream;
//      Dir.StreamSize := CFile[i].FDataSize;
      Dir.StreamSize := CFile[i].FSize;
      Dir.FirstSID := CFile[i].FSID;
    end;
    Move(Pointer(CFile[i].FName)^,Dir.Name,Length(CFile[i].FName) * 2);
    Dir.Name[Length(CFile[i].FName)] := #0;
    Dir.NameSize := Length(CFile[i].FName) * 2 + 2;
    FStream.Write(Dir,SizeOf(TDirEntry));

    if CFile[i].Count > 0 then
      WriteDir(CFile[i]);
  end;
end;

procedure SortDirs(CFile: TCompFile);
var
  i: integer;
begin
  CFile.Sort(@CompareDir);
  for i := 0 to CFile.Count - 1 do
    SortDirs(CFile[i]);
end;

begin
  DID := 0;
  FHeader.DirSID := GetSID;
  FFiles.FName := 'Root Entry';

  SortDirs(FFiles);

  FillChar(Dir,SizeOf(TDirEntry),#0);
  Move(Pointer(FFiles.FName)^,Dir.Name,Length(FFiles.FName) * 2);
  Dir.Name[Length(FFiles.FName)] := #0;
  Dir.NameSize := Length(FFiles.FName) * 2 + 2;
  Dir.EntryType := DirEntry_Root;
  Dir.NodeColor := 1;
  if FShortSectorsSize > 0 then begin
    Dir.FirstSID := FFiles.FSID;
    Dir.StreamSize := FFiles.FDataSize;
  end
  else
    Dir.FirstSID := SID_EndOfChain;
  Dir.RootDID := -1;
  Dir.LeftChildDID := -1;
  Dir.RightChildDID := -1;
  Inc(DID);
  if FFiles.Count > 0 then
    Dir.RootDID := DID;

  AssignDID(FFiles);

  FStream.Write(Dir,SizeOf(TDirEntry));

  WriteDir(FFiles);
  AllocSIDs(Ceil((DID * SizeOf(TDirEntry)) / FSectorSize));
  Sz := FSectorSize - (FStream.Position mod FSectorSize);
  if Sz <> FSectorSize then begin
    FillChar(Dir,SizeOf(TDirEntry),#0);
    for i := 0 to (Sz div SizeOf(TDirEntry)) - 1 do
      FStream.Write(Dir,SizeOf(TDirEntry));
  end;
end;

procedure TCompoundStorage.WriteSAT;
var
  i,Count,SID: integer;
begin
  Count := Ceil((FMaxSID * 4) / FSectorSize);
  if Count > 109 then
    raise Exception.Create('Can only write 109 MSAT');
  FHeader.MasterSecSID := SID_EndOfChain;
  FHeader.MasterSecCount := 0;
  SID := AllocSIDs(Count,True);
  if FStream.Position <> SIDToPos(SID) then
    raise Exception.CreateFmt('SID position error: Pos=%d SID=%d',[FStream.Position,SIDToPos(SID)]);
  FStream.Write(FSAT^,FMaxSID * 4);
  PadSectorNeg;

  for i := 0 to Count - 1 do
    FHeader.FirstMasterSecSIDs[i] := SID + i;
  for i := Count to High(FHeader.FirstMasterSecSIDs) do
    FHeader.FirstMasterSecSIDs[i] := SID_Free;
end;

procedure TCompoundStorage.WriteShortSectors;
var
  PadSz: integer;
begin
  if FShortSectorsSize > 0 then begin
    PadSz := FShortSectorSize - (FStream.Position mod FShortSectorSize);
    if PadSz = FShortSectorSize then
      PadSz := 0;
    FFiles.FDataSize := FShortSectorsSize + PadSz;
    FFiles.FSID := AllocSIDs(Ceil(FShortSectorsSize / FSectorSize));
    if FStream.Position <> SIDToPos(FFiles.FSID) then
      raise Exception.CreateFmt('SID position error: Pos=%d SID=%d',[FStream.Position,SIDToPos(FFiles.FSID)]);
    FStream.Write(FShortSectors^,FShortSectorsSize);
    PadSector;

    FHeader.ShortSecCount := Ceil((FMaxSSID * 4) / FSectorSize);
    FHeader.ShortSecSID := AllocSIDs(FHeader.ShortSecCount);
    if FStream.Position <> SIDToPos(FHeader.ShortSecSID) then
      raise Exception.CreateFmt('SID position error: Pos=%d SID=%d',[FStream.Position,SIDToPos(FHeader.ShortSecSID)]);
    FStream.Write(FSSAT^,FMaxSSID * 4);
    PadSectorNeg;
  end
  else begin
    FFiles.FSID := SID_EndOfChain;
    FHeader.ShortSecCount := 0;
    FHeader.ShortSecSID := SID_EndOfChain;;
  end;
end;

function TCompoundStorage.AllocSSIDs(Count: integer): integer;
var
  i: integer;
begin
  Result := FMaxSSID;
  if Count <= 0 then
    Exit;
  ReAllocMem(FSSAT,(FMaxSSID + Count) * SizeOf(longint));
  for i := 0 to Count - 2 do
    FSSAT[FMaxSSID + i] := FMaxSSID + i + 1;
  FSSAT[FMaxSSID + Count - 1] := SID_EndOfChain;
  Inc(FMaxSSID,Count);
end;

procedure TCompoundStorage.PadSector;
var
  b: byte;
  Sz: integer;
begin
  Sz := FSectorSize - (FStream.Position mod FSectorSize);
  if Sz = FSectorSize then
    Exit;
  b := 0;
  while Sz > 0 do begin
    FStream.Write(b,1);
    Dec(Sz);
  end;
end;

procedure TCompoundStorage.PadSectorNeg;
var
  i: integer;
  Sz: integer;
begin
  Sz := FSectorSize - (FStream.Position mod FSectorSize);
  if Sz = FSectorSize then
    Exit;
  i := -1;
  while Sz > 0 do begin
    FStream.Write(i,4);
    Dec(Sz,4);
  end;
end;

procedure TCompoundStorage.CloseFile;
begin
  if FOwnsStream then begin
    FStream.Free;
    FStream := Nil;
  end;
end;

procedure TCompoundStorage.SetFiles(const Value: TCompFile);

procedure AssignStorage(CFile: TCompFile);
var
  i: integer;
begin
  for i := 0 to CFile.Count - 1 do begin
    CFile[i].FStorage := Self;
    if CFile[i].Count > 0 then
      AssignStorage(CFile[i]);
  end;
end;

begin
  if Value <> Nil then
    FFiles.Free;
  FFiles := Value;
  if Value <> Nil then begin
    FFiles.FStorage := Self;
    AssignStorage(FFiles);
  end;
end;

function TCompoundStorage.IsStorage(Filename: WideString): boolean;
var
  Buf: array[0..7] of byte;
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(Filename,fmOpenRead,READ_FILEMODE);
  try
    Stream.Read(Buf,Length(Buf));
    Result := CompareMem(@Buf[0],@CompoundId[0],Length(CompoundId));
  finally
    Stream.Free;
  end;
end;

{ TCompFile }

function TCompFile.Add(Name: WideString; Size: integer = 0): TCompFile;
begin
  if FFiles = Nil then
    FFiles := TCompFile.Create(FStorage);
  Result := TCompFile.Create(FStorage);
  Result.FName := Name;
  Result.FDataSize := Size;
  inherited Add(Result);
end;

procedure TCompFile.Cache(var Data: PByteArray);
begin
  if FIsDirectory then
    raise Exception.Create('Can not read a directory');
  GetMem(Data,FDataSize);
  Read(Data^,FDataSize);
end;

procedure TCompFile.Close;
var
  Sz,Count: integer;
begin
  if FStorage.FOpenMode = somWrite then begin
    FSize := FCurrPos;
    if FWritingSmallSec then begin
      Count := Ceil(FCurrPos / FStorage.FShortSectorSize);
      FSID := FStorage.AllocSSIDs(Count);
      FStorage.WriteShort(FStorage.FWBuf^,FCurrPos);
    end
    else begin
      FSID := FStorage.GetSID;
      Sz := FStorage.FStream.Position - FStorage.SIDToPos(FSID);
      if (FStorage.FStream.Position - Sz) <> FStorage.SIDToPos(FSID) then
        raise Exception.CreateFmt('SID position error: Pos=%d SID=%d',[FStorage.FStream.Position - Sz,FStorage.SIDToPos(FSID)]);
      Count := Ceil(Sz / FStorage.FSectorSize);
      FStorage.AllocSIDs(Count);
      FStorage.PadSector;
    end;
  end;
  FCurrPos := 0;
  FWritingSmallSec := False;
end;

constructor TCompFile.Create(CompStorage: TCompoundStorage);
begin
  inherited Create;

  FStorage := CompStorage;
end;

function TCompFile.DecompressVBA(var Buf: PByteArray; Offset: integer): integer;
const VBA_WINDOW_SZ = $1000;
var
  i,Index: integer;
  Data: PByteArray;
  c,Flag: byte;
  Mask,Shift: integer;
  Pos,WinPos,Distance,SrcPos: integer;
  Token,Len: word;
  Clean: boolean;
  WinBuf: array[0..VBA_WINDOW_SZ - 1] of byte;

function GetByte: byte;
begin
  Result := Data[Index];
  Inc(Index);
end;

function GetWord: word;
var
  P: PWordArray;
begin
  P := PWordArray(Integer(Data) + Index);
  Result := P[0];
  Inc(Index,2);
end;

begin
  Result := 0;
  if FDataSize <= 0 then
    Exit;

  GetMem(Data,FDataSize);
  try
    Read(Data^,FDataSize);

    Clean := True;
    Pos := 0;
    Index := Offset + 3;
    while Index < FDataSize do begin
      Flag := GetByte;
      Mask := 1;
      while Mask < $0100 do begin
        if (Flag and Mask) <> 0 then begin
          WinPos := Pos mod VBA_WINDOW_SZ;
          if WinPos <= $80 then begin
            if WinPos <= $20 then begin
              if WinPos <= $10 then Shift := 12 else Shift := 11;
            end
            else begin
              if WinPos <= $40 then Shift := 10 else Shift := 9;
            end;
          end
          else begin
            if WinPos <= $200 then begin
              if WinPos <= $100 then Shift := 8 else Shift := 7;
            end
            else if WinPos <= $800 then begin
              if WinPos <= $400 then Shift := 6 else Shift := 5;
            end
            else
              Shift := 4;
          end;
          Token := GetWord;
          Len := (Token and ((1 shl Shift) - 1)) + 3;
          Distance := Token shr Shift;
          Clean := True;
          for i := 0 to Len - 1 do begin
            SrcPos := (Pos - Distance - 1) mod VBA_WINDOW_SZ;
            c := WinBuf[SrcPos];
            WinBuf[Pos mod VBA_WINDOW_SZ] := c;
            Inc(Pos);
          end;
        end
        else begin
          if (Pos <> 0) and ((Pos mod VBA_WINDOW_SZ) = 0) and Clean then begin
            GetWord;
            Clean := False;
            ReAllocMem(Buf,Result + VBA_WINDOW_SZ);
            System.Move(WinBuf,Buf[Result],VBA_WINDOW_SZ);
            Inc(Result,VBA_WINDOW_SZ);
            Break;
          end;
          WinBuf[Pos mod VBA_WINDOW_SZ] := GetByte;
          Inc(Pos);
          Clean := True;
        end;

        Mask := Mask shl 1;
      end;
    end;
    if (Pos mod VBA_WINDOW_SZ) <> 0 then begin
      ReAllocMem(Buf,Result + (Pos mod VBA_WINDOW_SZ));
      System.Move(WinBuf,Buf[Result],Pos mod VBA_WINDOW_SZ);
      Inc(Result,Pos mod VBA_WINDOW_SZ);
    end;
  finally
    FreeMem(Data);
  end;
end;

destructor TCompFile.Destroy;
begin
  FFiles.Free;
  inherited;
end;

function TCompFile.FindStream(Name: WideString; SearchSubDirs: boolean): TCompFile;
var
  i: integer;
begin
  Name := WideUppercase(Name);
  for i := 0 to Count - 1 do begin
    if not Items[i].IsDirectory and (WideUppercase(Items[i].FName) = Name) then begin
      Result := Items[i];
      Exit;
    end
    else if Items[i].IsDirectory and SearchSubDirs then begin
      Result := Items[i].FindStream(Name,True);
      if Result <> Nil then
        Exit;
    end;
  end;
  Result := Nil;
end;

function TCompFile.GetItems(Index: integer): TCompFile;
begin
  Result := TCompFile(inherited Items[Index]);
end;

procedure TCompFile.Open;
begin
  if FStorage.FOpenMode = somWrite then begin
    FCurrPos := 0;
    FSize := 0;
  end;
end;

function TCompFile.Read(const Buffer; Count: integer): longint;
begin
  if FIsDirectory then
    raise Exception.Create('Can not read a directory');
  if FCurrPos >= FSize then
    Result := 0
  else begin
    Result := FStorage.ReadStream(FSize < Integer(FStorage.FHeader.MinStdStream),FSID,FCurrPos,Count,@Buffer);
    Inc(FCurrPos,Result);
  end;
end;

function TCompFile.Seek(Offset: Integer; Origin: Word): longint;
begin
  if (Offset = 0) and (Origin = soFromCurrent) then
    Result := FCurrPos
  else if FStorage.FOpenMode = somRead then begin
    case Origin of
      soFromBeginning: FCurrPos := Offset;
      soFromCurrent  : Inc(FCurrPos,Offset);
      soFromEnd      : FCurrPos := FSize + Offset;
      else
        raise Exception.Create('Invalid Seek origin')
    end;
    if FCurrPos < 0 then
      FCurrPos := 0
    else if FCurrPos > FSize then
      FCurrPos := FSize;
    Result := FCurrPos;
  end
  else if FStorage.FOpenMode = somWrite then begin
    if FSize < Integer(FStorage.FHeader.MinStdStream) then begin
      case Origin of
        soFromBeginning: FCurrPos := Offset;
        soFromCurrent  : Inc(FCurrPos,Offset);
        soFromEnd      : FCurrPos := FSize + Offset;
        else
          raise Exception.Create('Invalid Seek origin')
      end;
      if FCurrPos < 0 then
        FCurrPos := 0
      else if FCurrPos > FSize then
        FCurrPos := FSize;
      Result := FCurrPos;
    end
    else begin
      case Origin of
        soFromBeginning: FCurrPos := FStorage.FStream.Seek(Offset + SizeOf(TCompoundDocHeader),Origin);
        soFromCurrent  : FCurrPos := FStorage.FStream.Seek(Offset,Origin) - SizeOf(TCompoundDocHeader);
        soFromEnd      : FCurrPos := FStorage.FStream.Seek(Offset,Origin) - SizeOf(TCompoundDocHeader);
        else
          raise Exception.Create('Invalid Seek origin')
      end;
      Result := FCurrPos;
      if Result < 0 then
        raise Exception.Create('Invalid position after seek');
    end;
  end
  else
    raise Exception.Create('File is not opened');
end;

function TCompFile.Write(const Buffer; Count: Integer): longint;

function WriteSmall(const Buffer; Count: Integer): longint;
var
  P: PByteArray;
begin
  Result := Count;
  P := PByteArray(Integer(FStorage.FWBuf) + FCurrPos);
  System.Move(Buffer,P^,Count);
  FWritingSmallSec := True;
  Inc(FCurrPos,Count);
end;

function WriteStd(const Buffer; Count: Integer): longint;
begin
  Result := Count;
  if FWritingSmallSec then begin
    FWritingSmallSec := False;
    WriteStd(FStorage.FWBuf^,FSize);
    FCurrPos := FSize;
  end;
  FStorage.FStream.Write(Buffer,Count);
  Inc(FCurrPos,Count);
end;

begin
  if FIsDirectory then
    raise Exception.Create('Can not write a directory');
  if (FSize + Count) < Integer(FStorage.FHeader.MinStdStream) then
    Result := WriteSmall(Buffer,Count)
  else
    Result := WriteStd(Buffer,Count);
  if FCurrPos > FSize then
    FSize := FCurrPos;
end;

{ TCacheFile }

function TCacheFile.Add: TCacheFile;
begin
  if FFiles = Nil then
    FFiles := TCacheFile.Create;
  Result := TCacheFile.Create;
  inherited Add(Result);
end;

procedure TCacheFile.Clear;
begin
  inherited;

  FFiles.Free;
  FFiles := Nil;
  FreeMem(FData);
  FData := Nil;
  FName := '';
  FDataSize := 0;
  FIsDirectory := False;
end;

function TCacheFile.GetItems(Index: integer): TCacheFile;
begin
  Result := TCacheFile(inherited Items[Index]);
end;

procedure TCacheFile.SetData(const Value: PByteArray);
begin
  FData := Value;
end;

end.
