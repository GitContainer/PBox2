unit BIFF_ICompoundStream5;

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

uses Classes, SysUtils,
     XLSUtils5, BIFF_CompoundStream5;

const STGM_CREATE           = $00001000;
const STGM_READ             = $00000000;
const STGM_WRITE            = $00000001;
const STGM_READWRITE        = $00000002;
const STGM_SHARE_EXCLUSIVE  = $00000010;
const STGM_DIRECT           = $00000000;
const STGM_SHARE_DENY_NONE  = $00000040;
const STGM_SHARE_DENY_READ  = $00000030;
const STGM_SHARE_DENY_WRITE = $00000020;

const STGM_TRANSACTED       = $10001;

const STGC_DEFAULT          = $10002;

const STATFLAG_NONAME       = $10003;

const STREAM_SEEK_SET       = 0;
const STREAM_SEEK_CUR       = 1;
const STREAM_SEEK_END       = 2;

const GMEM_MOVEABLE         = 1;

type LargeInt = int64;
type POleStr = PWideChar;
type TCLSID = TGUID;
type TSNB = ^POleStr;
type PIID = PGUID;

type TFileTime = TDateTime;
type BOOL = boolean;

type PStatStg = ^TStatStg;
  tagSTATSTG = record
    pwcsName: POleStr;
    dwType: Longint;
    cbSize: Largeint;
    mtime: TFileTime;
    ctime: TFileTime;
    atime: TFileTime;
    grfMode: Longint;
    grfLocksSupported: Longint;
    clsid: TCLSID;
    grfStateBits: Longint;
    reserved: Longint;
  end;
  TStatStg = tagSTATSTG;
  STATSTG = TStatStg;

type PXLSGlobalMem = ^TXLSGlobalMem;
     TXLSGlobalMem = record
     Data: Pointer;
     Size: NativeInt;
     Stream: TStream;
     end;

type IEnumStatStg = interface(IUnknown)
    ['{B4C201E9-C970-49E7-AAEE-5BF9EFFB389C}']
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
    function Skip(celt: Longint): HResult;
    function Reset: HResult;
    function Clone(out enm: IEnumStatStg): HResult;
  end;

type IStream = interface(IUnknown)
    ['{6464605F-B967-41CD-ABC0-D3B7DCCD3936}']
    function Read(pv: Pointer; cb: Longint; pcbRead: PLongint): HResult;
    function Write(pv: Pointer; cb: Longint; pcbWritten: PLongint): HResult;
    function Seek(dlibMove: Largeint; dwOrigin: Longint; out libNewPosition: Largeint): HResult;
    function SetSize(libNewSize: Largeint): HResult;
    function CopyTo(stm: IStream; cb: Largeint; out cbRead: Largeint; out cbWritten: Largeint): HResult;
    function Commit(grfCommitFlags: Longint): HResult;
    function Revert: HResult;
    function LockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
    function UnlockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
    function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
    function Clone(out stm: IStream): HResult;
  end;

type IStorage = interface(IUnknown)
    ['{12732E38-3E8C-4D79-84F7-0CFB14014082}']
    function CreateStream(pwcsName: POleStr; grfMode: Longint; reserved1: Longint; reserved2: Longint; out stm: IStream): HResult;
    function OpenStream(pwcsName: POleStr; reserved1: Pointer; grfMode: Longint; reserved2: Longint; out stm: IStream): HResult;
    function CreateStorage(pwcsName: POleStr; grfMode: Longint; dwStgFmt: Longint; reserved2: Longint; out stg: IStorage): HResult;
    function OpenStorage(pwcsName: POleStr; const stgPriority: IStorage; grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stg: IStorage): HResult;
    function CopyTo(ciidExclude: Longint; rgiidExclude: PIID; snbExclude: TSNB; const stgDest: IStorage): HResult;
    function MoveElementTo(pwcsName: POleStr; const stgDest: IStorage; pwcsNewName: POleStr; grfFlags: Longint): HResult;
    function Commit(grfCommitFlags: Longint): HResult;
    function Revert: HResult;
    function EnumElements(reserved1: Longint; reserved2: Pointer; reserved3: Longint; out enm: IEnumStatStg): HResult;
    function DestroyElement(pwcsName: POleStr): HResult;
    function RenameElement(pwcsOldName: POleStr; pwcsNewName: POleStr): HResult;
    function SetElementTimes(pwcsName: POleStr; const ctime: TFileTime; const atime: TFileTime; const mtime: TFileTime): HResult;
    function SetClass(const clsid: TCLSID): HResult; stdcall; function SetStateBits(grfStateBits: Longint; grfMask: Longint): HResult;
    function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
    // Not original.
    function Compound : TCompoundStorage;
    function Cache: TCacheFile;
  end;

type ILockBytes = interface(IUnknown)
    ['{CDCE42EE-0381-420B-9D82-9F2F8ECFC312}']
    function ReadAt(ulOffset: Largeint; pv: Pointer; cb: Longint; pcbRead: PLongint): HResult;
    function WriteAt(ulOffset: Largeint; pv: Pointer; cb: Longint; pcbWritten: PLongint): HResult;
    function Flush: HResult;
    function SetSize(cb: Largeint): HResult;
    function LockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
    function UnlockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
    function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
    // Not original.
    function Data: PXLSGlobalMem;
  end;

// *****************************************************************************
// ******* This is by no way a correct or complete implementation if the *******
// ******* IStorage and IStream interfaces but it will do the job for    *******
// ******* the component.                                                *******
// *****************************************************************************

type TCompLockBytes = class(TInterfacedObject,ILockBytes)
protected
     FData: PXLSGlobalMem;
public
     destructor Destroy; override;

     function ReadAt(ulOffset: Largeint; pv: Pointer; cb: Longint; pcbRead: PLongint): HResult;
     function WriteAt(ulOffset: Largeint; pv: Pointer; cb: Longint; pcbWritten: PLongint): HResult;
     function Flush: HResult;
     function SetSize(cb: Largeint): HResult;
     function LockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
     function UnlockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
     function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
    // Not original.
     function Data: PXLSGlobalMem;
     end;

type TCompStorage = class(TInterfacedObject,IStorage)
protected
    FCompound : TCompoundStorage;
    FCache    : TCacheFile;
    FFilename : WideString;
public
    constructor Create;
    destructor Destroy; override;

    function CreateStream(pwcsName: POleStr; grfMode: Longint; reserved1: Longint; reserved2: Longint; out stm: IStream): HResult;
    function OpenStream(pwcsName: POleStr; reserved1: Pointer; grfMode: Longint; reserved2: Longint; out stm: IStream): HResult;
    function CreateStorage(pwcsName: POleStr; grfMode: Longint; dwStgFmt: Longint; reserved2: Longint; out stg: IStorage): HResult;
    function OpenStorage(pwcsName: POleStr; const stgPriority: IStorage; grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stg: IStorage): HResult;
    function CopyTo(ciidExclude: Longint; rgiidExclude: PIID; snbExclude: TSNB; const stgDest: IStorage): HResult;
    function MoveElementTo(pwcsName: POleStr; const stgDest: IStorage; pwcsNewName: POleStr; grfFlags: Longint): HResult;
    function Commit(grfCommitFlags: Longint): HResult;
    function Revert: HResult;
    function EnumElements(reserved1: Longint; reserved2: Pointer; reserved3: Longint; out enm: IEnumStatStg): HResult;
    function DestroyElement(pwcsName: POleStr): HResult;
    function RenameElement(pwcsOldName: POleStr; pwcsNewName: POleStr): HResult;
    function SetElementTimes(pwcsName: POleStr; const ctime: TFileTime; const atime: TFileTime; const mtime: TFileTime): HResult;
    function SetClass(const clsid: TCLSID): HResult; stdcall; function SetStateBits(grfStateBits: Longint; grfMask: Longint): HResult;
    function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
    // Not original.
    function Compound: TCompoundStorage;
    procedure AddCache;
    function Cache: TCacheFile;
    end;

type TCompStream = class(TInterfacedObject,IStream)
protected
     FCompFile: TCompFile;
public
     constructor Create(ACompFile: TCompFile);
     destructor Destroy; override;

     function Read(pv: Pointer; cb: Longint; pcbRead: PLongint): HResult;
     function Write(pv: Pointer; cb: Longint; pcbWritten: PLongint): HResult;
     function Seek(dlibMove: Largeint; dwOrigin: Longint; out libNewPosition: Largeint): HResult;
     function SetSize(libNewSize: Largeint): HResult;
     function CopyTo(stm: IStream; cb: Largeint; out cbRead: Largeint; out cbWritten: Largeint): HResult;
     function Commit(grfCommitFlags: Longint): HResult;
     function Revert: HResult;
     function LockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
     function UnlockRegion(libOffset: Largeint; cb: Largeint; dwLockType: Longint): HResult;
     function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
     function Clone(out stm: IStream): HResult;
     end;

function StgOpenStorage(pwcsName: POleStr; stgPriority: IStorage; grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stgOpen: IStorage): HResult;
function CreateILockBytesOnHGlobal(hglob: HGlobal; fDeleteOnRelease: BOOL; out lkbyt: ILockBytes): HResult;
function StgCreateDocfileOnILockBytes(lkbyt: ILockBytes; grfMode: Longint; reserved: Longint; out stgOpen: IStorage): HResult;
function StgCreateDocfile(pwcsName: POleStr; grfMode: Longint; reserved: Longint; out stgOpen: IStorage): HResult;
function StgOpenStorageOnILockBytes(lkbyt: ILockBytes; stgPriority: IStorage; grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stgOpen: IStorage): HResult;
function StgIsStorageFile(pwcsName: POleStr): HResult;
function GetHGlobalFromILockBytes(lkbyt: ILockBytes; out hglob: HGlobal): HResult;

procedure OleCheck(Result: HResult);

function  GlobalAlloc(AFlags: longword; ASize: integer): NativeInt;
function  GlobalLock(AHandle: NativeInt): Pointer;
procedure GlobalUnlock(AHandle: NativeInt);
function  GlobalSize(AHandle: longword): integer;

implementation

function StgOpenStorage(pwcsName: POleStr; stgPriority: IStorage; grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stgOpen: IStorage): HResult;
var
  Compound: TCompStorage;
begin
  Compound := TCompStorage.Create;

  Compound.FFilename := pwcsName;
  Compound.FCompound.LoadFromFile(pwcsName);

  stgOpen := Compound;

  Result := S_OK;
end;

function CreateILockBytesOnHGlobal(hglob: HGlobal; fDeleteOnRelease: BOOL; out lkbyt: ILockBytes): HResult;
var
  LBytes: TCompLockBytes;
begin
  LBytes := TCompLockBytes.Create;
  LBytes.FData := PXLSGlobalMem(hglob);

  lkbyt := LBytes;

  Result := S_OK;
end;

function StgCreateDocfileOnILockBytes(lkbyt: ILockBytes; grfMode: Longint; reserved: Longint; out stgOpen: IStorage): HResult;
var
  Compound: TCompStorage;
begin
  Compound := TCompStorage.Create;
  Compound.AddCache;

  stgOpen := Compound;

  Result := S_OK;
end;

function StgCreateDocfile(pwcsName: POleStr; grfMode: Longint; reserved: Longint; out stgOpen: IStorage): HResult;
var
  Compound: TCompStorage;
begin
  Compound := TCompStorage.Create;

  Compound.FFilename := pwcsName;
  Compound.FCompound.SaveToFile(pwcsName);

  stgOpen := Compound;

  Result := S_OK;
end;

type TMyMemStream = class(TMemoryStream);

function StgOpenStorageOnILockBytes(lkbyt: ILockBytes; stgPriority: IStorage; grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stgOpen: IStorage): HResult;
var
  Compound: TCompStorage;
begin
  if lkbyt.Data <> Nil then begin
    Compound := TCompStorage.Create;
    lkbyt.Data.Stream := TMemoryStream.Create;

    TMyMemStream(lkbyt.Data.Stream).SetPointer(lkbyt.Data.Data,lkbyt.Data.Size);

    Compound.FCompound.LoadFromStream(lkbyt.Data.Stream);

    stgOpen := Compound;

    Result := S_OK;
  end
  else
    Result := S_FALSE;
end;

function StgIsStorageFile(pwcsName: POleStr): HResult;
var
  Compound: TCompStorage;
begin
  Result := S_FALSE;
  Compound := TCompStorage.Create;
  try
    if Compound.FCompound.IsStorage(pwcsName) then
      Result := S_OK;
  finally
    Compound.Free;
  end;
end;

function GetHGlobalFromILockBytes(lkbyt: ILockBytes; out hglob: HGlobal): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;
end;

procedure OleCheck(Result: HResult);
begin
  if Result <> S_OK then
    raise Exception.Create('Gnyffa Pomperipossa!');
end;

function GlobalAlloc(AFlags: longword; ASize: integer): NativeInt;
var
  M: PXLSGlobalMem;
begin
  GetMem(M,SizeOf(TXLSGlobalMem));
  GetMem(M.Data,ASize);
  M.Size := ASize;
  Result := NativeInt(M);
end;

function  GlobalLock(AHandle: NativeInt): Pointer;
begin
  Result := PXLSGlobalMem(AHandle).Data;
end;

procedure GlobalUnlock(AHandle: NativeInt);
begin
// Do nothing. Memory is owned by TCompLockBytes.
//  FreeMem(PXLSGlobalMem(AHandle).Data);
//  FreeMem(PXLSGlobalMem(AHandle));
end;

function  GlobalSize(AHandle: longword): integer;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;
end;

{ TStorage }

procedure TCompStorage.AddCache;
begin
  FCache := TCacheFile.Create;
end;

function TCompStorage.Cache: TCacheFile;
begin
  Result := FCache;
end;

function TCompStorage.Commit(grfCommitFlags: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;
end;

function TCompStorage.Compound: TCompoundStorage;
begin
  Result := FCompound;
end;

function TCompStorage.CopyTo(ciidExclude: Integer; rgiidExclude: PIID; snbExclude: TSNB; const stgDest: IStorage): HResult;
begin
  if Cache <> Nil then
    stgDest.Compound.WriteCache(Cache)
  else
    Compound.ReadCache(stgDest.Cache,['Book','Workbook']);

  Result := S_OK;
end;

constructor TCompStorage.Create;
begin
  FCompound := TCompoundStorage.Create;
end;

function TCompStorage.CreateStorage(pwcsName: POleStr; grfMode, dwStgFmt, reserved2: Integer; out stg: IStorage): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;
end;

function TCompStorage.CreateStream(pwcsName: POleStr; grfMode, reserved1, reserved2: Integer; out stm: IStream): HResult;
var
  CFile: TCompFile;
begin
  Result := S_FALSE;

  if FCompound <> Nil then begin
    CFile := FCompound.Files.Add(pwcsName);
    CFile.Open;
    stm := TCompStream.Create(CFile);

    Result := S_OK;
  end;
end;

destructor TCompStorage.Destroy;
begin
  FCompound.Close;
  FCompound.CloseFile;

  FCompound.Free;

  if FCache <> Nil then
    FCache.Free;

  inherited;
end;

function TCompStorage.DestroyElement(pwcsName: POleStr): HResult;
begin
  Result := S_OK;
end;

function TCompStorage.EnumElements(reserved1: Integer; reserved2: Pointer; reserved3: Integer; out enm: IEnumStatStg): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;
end;

function TCompStorage.MoveElementTo(pwcsName: POleStr; const stgDest: IStorage; pwcsNewName: POleStr; grfFlags: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;
end;

function TCompStorage.OpenStorage(pwcsName: POleStr; const stgPriority: IStorage; grfMode: Integer; snbExclude: TSNB; reserved: Integer; out stg: IStorage): HResult;
begin
  FCompound.LoadFromFile(pwcsName);
  Result := S_OK;
end;

function TCompStorage.OpenStream(pwcsName: POleStr; reserved1: Pointer; grfMode, reserved2: Integer; out stm: IStream): HResult;
var
  CFile: TCompFile;
begin
  Result := S_FALSE;

  CFile := FCompound.Files.FindStream(pwcsName,False);
  if CFile <> Nil then begin
    stm := TCompStream.Create(CFile);
    Result := S_OK;
  end;
end;

function TCompStorage.RenameElement(pwcsOldName, pwcsNewName: POleStr): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStorage.Revert: HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStorage.SetClass(const clsid: TCLSID): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStorage.SetElementTimes(pwcsName: POleStr; const ctime, atime, mtime: TFileTime): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStorage.SetStateBits(grfStateBits, grfMask: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStorage.Stat(out statstg: TStatStg; grfStatFlag: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

{ TCompStream }

function TCompStream.Clone(out stm: IStream): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStream.Commit(grfCommitFlags: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStream.CopyTo(stm: IStream; cb: Largeint; out cbRead, cbWritten: Largeint): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

constructor TCompStream.Create(ACompFile: TCompFile);
begin
  FCompFile := ACompFile;
end;

destructor TCompStream.Destroy;
begin
  if FCompFile <> Nil then
    FCompFile.Close;

  inherited;
end;

function TCompStream.LockRegion(libOffset, cb: Largeint; dwLockType: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStream.Read(pv: Pointer; cb: Integer; pcbRead: PLongint): HResult;
begin
  Result := S_FALSE;
  if FCompFile <> Nil then begin
    pcbRead^ := FCompFile.Read(pv^,cb);
    Result := S_OK;
  end;
end;

function TCompStream.Revert: HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStream.Seek(dlibMove: Largeint; dwOrigin: Integer; out libNewPosition: Largeint): HResult;
begin
  Result := S_FALSE;
  if FCompFile <> Nil then begin
    case dwOrigin of
      STREAM_SEEK_SET: libNewPosition := FCompFile.Seek(dlibMove,soFromBeginning);
      STREAM_SEEK_END: libNewPosition := FCompFile.Seek(dlibMove,soFromEnd);
      STREAM_SEEK_CUR: libNewPosition := FCompFile.Seek(dlibMove,soFromCurrent);
      else             Exit;
    end;
    Result := S_OK;
  end;
end;

function TCompStream.SetSize(libNewSize: Largeint): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStream.Stat(out statstg: TStatStg; grfStatFlag: Integer): HResult;
begin
  Result := S_FALSE;
  if FCompFile <> Nil then begin
    FillChar(statstg,SizeOf(TStatStg),#0);
    statstg.cbSize := FCompFile.Size;
    Result := S_OK;
  end;
end;

function TCompStream.UnlockRegion(libOffset, cb: Largeint; dwLockType: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompStream.Write(pv: Pointer; cb: Integer; pcbWritten: PLongint): HResult;
begin
  Result := S_FALSE;
  if FCompFile <> Nil then begin
    pcbWritten^ := FCompFile.Write(pv^,cb);
    Result := S_OK;
  end;
end;

{ TCompLockBytes }

function TCompLockBytes.Data: PXLSGlobalMem;
begin
  Result := FData;
end;

destructor TCompLockBytes.Destroy;
begin
  if FData <> Nil then begin
    if FData.Data <> Nil then
      FreeMem(FData.Data);
    if FData.Stream <> Nil then
      FData.Stream.Free;
    FreeMem(FData);
  end;

  inherited;
end;

function TCompLockBytes.Flush: HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompLockBytes.LockRegion(libOffset, cb: Largeint; dwLockType: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompLockBytes.ReadAt(ulOffset: Largeint; pv: Pointer; cb: Integer; pcbRead: PLongint): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompLockBytes.SetSize(cb: Largeint): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompLockBytes.Stat(out statstg: TStatStg; grfStatFlag: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompLockBytes.UnlockRegion(libOffset, cb: Largeint; dwLockType: Integer): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

function TCompLockBytes.WriteAt(ulOffset: Largeint; pv: Pointer; cb: Integer; pcbWritten: PLongint): HResult;
begin
  raise Exception.Create('TODO');
  Result := S_FALSE;;
end;

end.
