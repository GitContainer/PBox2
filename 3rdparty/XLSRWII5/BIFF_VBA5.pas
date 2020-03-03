unit BIFF_VBA5;

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

uses SysUtils, Classes, Contnrs,
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
     BIFF_WideStrList5,
     XLSUtils5;

//* ~exclude
const VBA_ROOTDIR_NAME = '_VBA_PROJECT_CUR';
//* ~exclude
const VBA_SUBDIR_NAME  = 'VBA';
//* ~exclude
const VBA_PROJDIR_NAME = '_VBA_PROJECT';

//* ~exclude
type TVBAReadResult = (vbarrFailure,vbarrReadOnly,vbarrOk);

//* ~exclude
type TVBARecHeader = packed record
     RecId: word;
     Length: longword;
     end;

//* ~exclude
type PVBARecordData = ^TVBARecordData;
     TVBARecordData = record
     RecId: word;
     Length: longword;
     Data: record end;
     end;

//* ~exclude
type TRecVBAProjVersion = packed record
     ExcelVBA: word;
     Mac: word;
     Unknown: word;
     end;

//* ~exclude
type TRecVBALanguageId = packed record
     Id1: longword;
     Id2: longword;
     CharSet: word;
     Len1: word;
     Unknown1: longword;
     Unknown2: longword;
     Len2: word;
     Len3: word;
     Len4: word;
     end;

//* ~exclude
type TModuleData = record
     Name: WideString;
     Offset: longword;
     end;

{$ifdef BIFFSCAN}
type  TVbaRecordEvent = procedure(Id: word; Name: AxUCString; Data: TDynByteArray; Description: AxUCString) of object;
{$endif}

//* Module type
type TVbaModuleType = (vmtMacro,    //* Macro module.
                       vmtDocument  //* Document module.
                       );

//* ~exclude
type TRecordStorageVBA = class(TList)
private
     function  GetAsString(Index: integer): XLS8String;
     function  GetAsWideString(Index: integer): AxUCString;
     procedure SetAsString(Index: integer; const Value: XLS8String);
     procedure SetAsWideString(Index: integer; const Value: AxUCString);
     function  GetItems(Index: integer): PVBARecordData;
public
     destructor Destroy; override;
     procedure Clear; override;
     function  AddRec(Header: TVBARecHeader; Data: PByteArray): PVBARecordData;
     procedure Delete(Index: integer);
     procedure RenameModule(Name: AxUCString);
     procedure Write(Stream: TStream);

     property Items[Index: integer]: PVBARecordData read GetItems;
     property AsString[Index: integer]: XLS8String read GetAsString write SetAsString;
     property AsWideString[Index: integer]: AxUCString read GetAsWideString write SetAsWideString;
     end;

//* ~exclude
type TXLSVBAProject = class(TObject)
private
     FProjVersion: TRecVBAProjVersion;
     FLangId: TRecVBALanguageId;
     FLinkStrings: array of AxUCString;
     FModules: array of TModuleData;
{$ifdef BIFFSCAN}
     FRecordEvent: TVbaRecordEvent;
{$endif}
public
     function Read(Stream: IStream): boolean;
     procedure Clear;

{$ifdef BIFFSCAN}
     property OnRecord: TVbaRecordEvent read FRecordEvent write FRecordEvent;
{$endif}
     end;

//* Base class for VBA modules.
type TVBAModule = class(TObject)
private
     procedure SetName(const Value: AxUCString);
protected
     FName: AxUCString;
     FFileOffset: integer;
     FModuleType: TVbaModuleType;
     FAttributes: TXLSValueWideStringList;
     FDir: TRecordStorageVBA;
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* The name of the module.
     property Name: AxUCString read FName write SetName;
//     property FileOffset: integer read FFileOffset write FFileOffset;
     //* Module type.
     property ModuleType: TVbaModuleType read FModuleType write FModuleType;
{$ifdef BIFFSCAN}
     property Attributes: TXLSValueWideStringList read FAttributes;
{$endif}
     end;

//* Source code VBA module.
type TVBAModuleSource = class(TVBAModule)
private
     FSource: TStringList;
protected
     procedure AddSrcExtractAttr(Src: XLS8String);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* The source code in the module.
     property Source: TStringList read FSource;
     end;

//* VBA macros object. This object stores the VBA source code modules for the workbook.
//* ~[br]
//* ~[br]
//* ~[b Example:]
//* ~[br]
//* ~[sample
//*  // Button with macro example.
//*  // IMPORTANT! As XLSReadWriteII not can check that the syntax of the macro
//*  // code is correct, try to test run it in Excel first.
//*  // If there are errors in the code, Excel will open the file with an
//*  // error message.
//*  EditVBA must be set to True in order to edit macros.
//*  XLS.VBA.EditVBA := True;
//*  // Add a VBA module and a procedure.
//*  with XLS.VBA.AddModule('Module1',vmtMacro) do begin
//*    Source.Text :=
//*    'Sub MyTestMacro()' + #13#10 +
//*    'MsgBox "MyTestMacro executed"' + #13#10 +
//*    'End Sub';
//*  end;
//*
//*  // Add a button control.
//*  with XLS.Sheet[0~[].ControlsObjects.Buttons.Add do begin
//*    Text := 'Macro test';
//*    Col1 := 2;
//*    Row1 := 4;
//*    Col2 := 3;
//*    Row2 := 5;
//*    // The name of the macro must match the previous added macro procedure.
//*    MacroName := 'MyTestMacro';
//*  end;
//* ]
type TXLSVBA = class(TObject)
private
     FModules: TObjectList;
     FVBAProject: TXLSVBAProject;
     FProject: TXLSWideStringList;
     FProjectwm: array of byte;
     FEditVBA: boolean;
     FDisableVBA: boolean;
     FEditStateLocked: boolean;
     FIsEditing: boolean;
     FDir: TRecordStorageVBA;
     FSavedLockBytes: ILockBytes;
     FSavedStorage: IStorage;
     FHighCompression: boolean;
{$ifdef BIFFSCAN}
     FRecordEvent: TVbaRecordEvent;
{$endif}

     procedure SetEditVBA(const Value: boolean);
     function  StreamSize(Stream: IStream): integer;
     function  DecompressVBA(BufIn: PByteArray; var BufOut: PByteArray; BufSize, Offset: integer): integer;
     procedure CompressVBA(Buf: PByteArray; BufSize: integer; Stream: IStream);
     procedure ReadDir(Storage: IStorage; Buf: PByteArray; BufSz: integer);
     function  GetModules(Index: integer): TVBAModuleSource;
     function  ReadStrmToBuf(Stream: IStream; var Buf: PByteArray): integer;
     procedure Strip(stgVBA: IStorage);
     procedure ReadPROJECT(Storage: IStorage);
     procedure ReadPROJECTwm(Storage: IStorage);
     procedure WritePROJECT(Storage: IStorage);
     procedure WritePROJECTwm(Storage: IStorage);
     procedure WriteModules(Storage: IStorage);
     procedure WriteDir(Storage: IStorage);
     procedure WriteEmptyVBA_PROJECT(Stream: IStream);
     procedure Int_AddModule(var Buf: PByteArray; Storage: IStorage);
     function  GetCompressShift(Index: integer): integer;
    procedure SetDisableVBA(const Value: boolean);
public
     //* ~exclude
     constructor Create;
     //* ~exclude
     destructor Destroy; override;
     //* Clears all VBA data:
     procedure Clear;
     //* ~exclude
     function  Read(Storage: IStorage): TVBAReadResult;
     //* ~exclude
     procedure Write(Storage: IStorage);
     //* Adds a new VBA module.
     //* ~param ModuleName The name of the module-
     //* ~param ModuleType The module type. Set this to vmtMacro.
     //* ~result The new module.
     function  AddModule(ModuleName: AxUCString; ModuleType: TVbaModuleType): TVBAModuleSource;
     //* Finds a module by it's name.
     //* ~param ModuleName The module to search for.
     //* ~result The module, if it was found, otherwise Nil is returned.
     function  ModuleByName(ModuleName: AxUCString): TVBAModuleSource;
     //* Finds a module by it's name.
     //* ~param ModuleName The module to search for.
     //* ~result The index of the module, if it was found, otherwise is -1 returned.
     function  ModuleIndexByName(ModuleName: AxUCString): integer;
     //* Deletes a module by it's name.
     //* ~param ModuleName The module to delete.
     procedure DeleteModule(ModuleName: AxUCString);
     //* Number of modules in the project.
     //* ~result Module count.
     function  Count: integer;
     //* Checks if there are any VBA modules.
     //* ~result True if there are any VBA modules.
     function  HasData: boolean;

     //* The list of VBA modules in the project.
     property Modules[Index: integer]: TVBAModuleSource read GetModules; default;
     //* Set EditVBA to True before adding or editing VBA data.
     property EditVBA: boolean read FEditVBA write SetEditVBA;
     property DisableVBA: boolean read FDisableVBA write SetDisableVBA;
     //* ~exclude
     property HighCompression: boolean read FHighCompression write FHighCompression;
{$ifdef BIFFSCAN}
     property VBAProject: TXLSVBAProject read FVBAProject;
     property Project: TXLSWideStringList read FProject;
     property OnRecord: TVbaRecordEvent read FRecordEvent write FRecordEvent;
{$endif}
     end;

implementation

const VBA_WINDOW_SZ = $1000;

const VBADIRREC_START            = $0001;
const VBADIRREC_ENCODING1        = $0002;
const VBADIRREC_VERSION1         = $0003;
const VBADIRREC_PROJECTNAME      = $0004;
const VBADIRREC_DESCRIPTION      = $0005;
const VBADIRREC_HELPFILENAME     = $0006;
const VBADIRREC_HELPCTXTID       = $0007;
const VBADIRREC_UNKNOWN_08       = $0008;
const VBADIRREC_OPPAPERSIZE      = $0009;
const VBADIRREC_CONDCOMPARGS     = $000C;
const VBADIRREC_EXTERNALLINKAGE  = $000D;
const VBADIRREC_OBJECTCOUNT      = $000F;
const VBADIRREC_EOF              = $0010;
const VBADIRREC_UNKNOWN_13       = $0013;
const VBADIRREC_ENCODING2        = $0014;
const VBADIRREC_STREAMTYPE       = $0016;
const VBADIRREC_USERNAME         = $0019;
const VBADIRREC_STREAMNAME       = $001A;
const VBADIRREC_UNKNOWN_1C       = $001C;
const VBADIRREC_UNKNOWN_1E       = $001E;
const VBADIRREC_TYPEMACRO        = $0021;
const VBADIRREC_TYPEDOCUMENT     = $0022;
const VBADIRREC_OBJECTEND        = $002B;
const VBADIRREC_UNKNOWN_2C       = $002C;
const VBADIRREC_VBAMODULEOFFSET  = $0031;
const VBADIRREC_UC_STREAMNAME    = $0032;
const VBADIRREC_UC_CONDCOMPARGS  = $003C;
const VBADIRREC_UC_HELPFILENAME  = $003D;
const VBADIRREC_UC_STREAMTYPE    = $003E;
const VBADIRREC_UC_DESCRIPTION   = $0040;
const VBADIRREC_UC_USERNAME      = $0047;
const VBADIRREC_UNKNOWN_48       = $0048;



type PLongwordArray = ^TLongwordArray;
     TLongwordArray = array[0..16383] of longWord;

procedure DumpBuf(Filename: XLS8String; Buf: PByteArray; Sz: integer);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AxUCString(Filename),fmCreate);
  try
    Stream.Write(Buf^,Sz);
  finally
    Stream.Free;
  end;
end;

function OleResult(Res: HResult): boolean;
begin
  Result := Res and $80000000 = 0;
end;

procedure WriteByte(Stream: IStream; B: byte);
var
  p: integer;
begin
  OleCheck(Stream.Write(@B,1,@p));
end;

procedure WriteWord(Stream: IStream; W: word);
var
  p: integer;
begin
  OleCheck(Stream.Write(@W,2,@p));
end;

procedure WriteString(Stream: IStream; S: XLS8String);
var
  p: integer;
begin
  OleCheck(Stream.Write(Pointer(S),Length(S),@p));
end;

procedure WriteUCString(Stream: IStream; WS: AxUCString);
var
  p: integer;
begin
  OleCheck(Stream.Write(Pointer(WS),Length(WS) * 2,@p));
end;

procedure WriteStringCRLF(Stream: IStream; S: XLS8String);
begin
  WriteString(Stream,S);
  WriteByte(Stream,$0D);
  WriteByte(Stream,$0A);
end;

{ TXLSVBA }

function TXLSVBA.AddModule(ModuleName: AxUCString; ModuleType: TVbaModuleType): TVBAModuleSource;
begin
  if ModuleByName(ModuleName) <> Nil then
    raise XLSRWException.Create('There is allready a Module with name ' + ModuleName);
  Result := TVBAModuleSource.Create;
  Result.Name := ModuleName;
  Result.ModuleType := ModuleType;
  Result.FAttributes.Add('VB_Name = "' + ModuleName + '"');
  case ModuleType of
    vmtMacro: ;
    vmtDocument: begin
      Result.FAttributes.Add('VB_Base = "0{00020819-0000-0000-C000-000000000046}"');
      Result.FAttributes.Add('VB_GlobalNameSpace = False');
      Result.FAttributes.Add('VB_Creatable = False');
      Result.FAttributes.Add('VB_PredeclaredId = True');
      Result.FAttributes.Add('VB_Exposed = True');
      Result.FAttributes.Add('VB_TemplateDerived = False');
      Result.FAttributes.Add('VB_Customizable = True');
    end;
  end;
  FModules.Add(Result);
end;

procedure TXLSVBA.Clear;
begin
  FModules.Clear;
  FDir.Clear;
  FVBAProject.Clear;
  FProject.Clear;
  SetLength(FProjectwm,0);
  FSavedStorage := Nil;
  FSavedLockbytes := Nil;
  FIsEditing := False;
  FEditStateLocked := False;
end;

function TXLSVBA.GetCompressShift(Index: integer): integer;
begin
  if Index <= $80 then begin
    if Index <= $20 then begin
      if Index <= $10 then Result := 12 else Result := 11;
    end
    else begin
      if Index <= $40 then Result := 10 else Result := 9;
    end;
  end
  else begin
    if Index <= $200 then begin
      if Index <= $100 then Result := 8 else Result := 7;
    end
    else if Index <= $800 then begin
      if Index <= $400 then Result := 6 else Result := 5;
    end
    else
      Result := 4;
  end;
end;

procedure TXLSVBA.CompressVBA(Buf: PByteArray; BufSize: integer; Stream: IStream);
var
  j,Blocks,BlockStart,BlockSize: integer;
  OutPos: integer;
  BlockShift,Mask: byte;
  P,OutBuf: PByteArray;
{$ifdef BABOON}
  p2: largeint;
{$else}
{$ifdef DELPHI_XE8_OR_LATER}
  p2: largeuint;
{$else}
  p2: largeint;
{$endif}
{$endif}

procedure CompressBuf(Ptr: PByteArray; Sz: integer);
var
  i,j,Cnt: integer;
begin
  Cnt := Sz div 8;
  for i := 1 to Cnt do begin
    WriteByte(Stream,0);
    Stream.Write(Ptr,8,@j);
    Ptr := PByteArray(NativeInt(Ptr) + 8);
    Inc(BlockSize,1 + 8);
  end;
  if (Sz mod 8) <> 0 then begin
    WriteByte(Stream,0);
    Stream.Write(Ptr,Sz mod 8,@j);
    Inc(BlockSize,1 + (Sz mod 8));
//    P2 := PByteArray(NativeInt(P2) + (Sz mod 8));
  end;
end;

procedure WriteBlockSize(BlockSz: integer; ReserveSpace: boolean);
begin
  OleCheck(Stream.Seek(BlockStart,STREAM_SEEK_SET,p2));
  if BlockSz >= VBA_WINDOW_SZ then
    raise XLSRWException.Create('VBA compression failed')
//    WriteWord(Stream,$B0B0)
  else
    WriteWord(Stream,$B000 or BlockSz);
  OleCheck(Stream.Seek(0,STREAM_SEEK_END,p2));
  BlockStart := p2;
  BlockSize := 0;
  if ReserveSpace then begin
    WriteByte(Stream,$00);
    WriteByte(Stream,$00);
  end;
end;

procedure WriteUncompressed;
var
  i: integer;
begin
  BlockStart := StreamSize(Stream) + 1;
  WriteByte(Stream,$01);
  WriteByte(Stream,$00);
  WriteByte(Stream,$00);
  BlockSize := 2;
  Blocks := BufSize div VBA_WINDOW_SZ;
  P := Buf;

  WriteByte(Stream,$00);
  Inc(BlockSize);
  j := 0;
  for i := 0 to BufSize - 1 do begin
    WriteByte(Stream,Buf[i]);
    Inc(BlockSize);
    Inc(j);
    if (BlockSize < VBA_WINDOW_SZ) and (j >= 8) then begin
      WriteByte(Stream,0);
      Inc(BlockSize);
      j := 0;
    end;
    if BlockSize >= VBA_WINDOW_SZ then begin
//      OleCheck(Stream.Seek(-1,STREAM_SEEK_CUR,p2));
//      WriteByte(Stream,$73);
      WriteBlockSize(BlockSize,True{ (BufSize - i) > VBA_WINDOW_SZ_W});
      WriteByte(Stream,$00);
      BlockSize := 3;
      j := 0;
    end;
  end;
  if BlockSize > 0 then begin
//    OleCheck(Stream.Seek(-1,STREAM_SEEK_CUR,p2));
//    WriteByte(Stream,$73);
    WriteBlockSize(BlockSize,False);
  end;

{
  for i := 1 to Blocks do begin
    CompressBuf(P,VBA_WINDOW_SZ_W);
//    WriteByte(Stream,$00);
//    WriteBlockSize(BlockSize,True);
    P := PByteArray(NativeInt(P) + VBA_WINDOW_SZ_W);
  end;
  if (BufSize mod VBA_WINDOW_SZ_W) <> 0 then
    CompressBuf(P,BufSize mod VBA_WINDOW_SZ_W);
}
{
  if Blocks > 0 then begin
    WriteByte(Stream,$00);
  end;
  WriteBlockSize(BlockSize,False);
}
{
  OleCheck(Stream.Seek(0,STREAM_SEEK_CUR,p2));
  Dec(p2,Sz + 4);
  if p2 > $0C0C then
    FileSz := $0C0C
  else
    FileSz := p2;
  OleCheck(Stream.Seek(Sz,STREAM_SEEK_SET,p2));
  WriteByte(Stream,$01);
  WriteByte(Stream,FileSz and $00FF);
  WriteByte(Stream,(FileSz shr 8) or $B0);
}
end;

procedure WriteToFile;
var
  p: integer;
begin
  if OutPos > 0 then begin
    OleCheck(Stream.Write(@Mask,1,@p));
    OleCheck(Stream.Write(OutBuf,OutPos,@p));
    Inc(BlockSize,1 + OutPos);
    OutPos := 0;
    Mask := 0;
    BlockShift := 0;
  end;
end;

function FindMatch(Buf: PByteArray; BufSize,Index: integer; var MatchPos,MatchSize: integer): boolean;
var
  i,j,MaxMatch,MaxSize: integer;
begin
  MaxSize := 2;
  Result := False;
  MaxMatch := (1 shl GetCompressShift(Index)) - 1;
  for i := Index - 1 downto 0 do begin
    j := 0;
    while (j < (BufSize - Index)) and (j < Index) do begin
      if Buf[Index + j] <> Buf[i + j] then
        Break;
      Inc(j);
    end;
    if (j > MaxSize) then begin
      if j > MaxMatch then
        j := MaxMatch;
      MaxSize := j;
      MatchSize := j;
      MatchPos := i;
      Result := True;
      // Stop on first match
      if not FHighCompression then
        Break;
    end;
  end;
end;

procedure WriteMatch(Buf: PByteArray; Index: integer; MatchPos,MatchSize: integer);
var
  Shift,Dist,Token: integer;
begin
  Shift := GetCompressShift(Index);
  Dist := Index - MatchPos - 1;
  Token := (Dist shl Shift) + ((MatchSize - 3 and ((1 shl Shift) - 1)));

  Mask := Mask or (1 shl BlockShift);
  PWordArray(NativeInt(OutBuf) + OutPos)[0] := Token;
  Inc(OutPos,2);
  Inc(BlockShift);
  if BlockShift >= 8 then
    WriteToFile;
end;

procedure WriteData(Buf: PByteArray; Index: integer);
begin
  OutBuf[OutPos] := Buf[Index];
  Inc(OutPos);
  Inc(BlockShift);
  if BlockShift >= 8 then
    WriteToFile;
end;

procedure CompressBlock(Buf: PByteArray; Size: integer);
var
  i: integer;
  MatchPos,MatchSize: integer;
begin
  i := 0;
  while i < Size do begin
    if FindMatch(Buf,Size,i,MatchPos,MatchSize) then begin
      WriteMatch(Buf,i,MatchPos,MatchSize);
      Inc(i,MatchSize);
    end
    else begin
      WriteData(Buf,i);
      Inc(i);
    end;
  end;
end;

procedure WriteCompressed;
var
  P: PByteArray;
  Sz: integer;
begin
  OutPos := 0;
  Mask := 0;
  BlockShift := 0;
  P := Buf;
  BlockStart := StreamSize(Stream) + 1;
  WriteByte(Stream,$01);
  WriteByte(Stream,$00);
  WriteByte(Stream,$00);
  BlockSize := 0;
  GetMem(OutBuf,VBA_WINDOW_SZ);
  try
    while BufSize > 0 do begin
      if BufSize > VBA_WINDOW_SZ then
        Sz := VBA_WINDOW_SZ
      else
        Sz := BufSize;
      CompressBlock(P,Sz);
      Dec(BufSize,Sz);
      if BufSize > 0 then begin
        WriteToFile;
        WriteBlockSize(BlockSize - 1,True);
      end;
      P := PByteArray(NativeInt(P) + Sz);
    end;
    WriteToFile;
    WriteBlockSize(BlockSize - 1,False);
  finally
    FreeMem(OutBuf);
  end;
end;

procedure WriteTest;
var
  i,Sz: integer;
  Strm: TFileStream;
begin
  Strm := TFileStream.Create('d:\temp\mod.bin',fmOpenRead);
  GetMem(P,BufSize);
  Sz := Strm.Read(P^,BufSize);
  Stream.Write(P,Sz,@i);
  FreeMem(P);
  Strm.Free;
end;

begin
//  DumpBuf('d:\temp\in.bin',Buf,BufSize);
  WriteCompressed;
end;

constructor TXLSVBA.Create;
begin
  inherited Create;
  FHighCompression := True;
  FModules := TObjectList.Create;
  FEditVBA := False;
  FVBAProject := TXLSVBAProject.Create;
  FDir := TRecordStorageVBA.Create;
  FProject := TXLSWideStringList.Create;
end;

function TXLSVBA.DecompressVBA(BufIn: PByteArray; var BufOut: PByteArray; BufSize, Offset: integer): integer;
var
  i,Index: integer;
  c,Flag: byte;
  Mask,Shift: integer;
  Pos,BlockPos,WinPos,Distance,SrcPos: integer;
  Token,Len: word;
  Clean: boolean;
  WinBuf: array[0..VBA_WINDOW_SZ - 1] of byte;

function GetByte: byte;
begin
  Result := BufIn[Index];
  Inc(Index);
  Inc(BlockPos);
end;

function GetWord: word;
var
  P: PWordArray;
begin
  P := PWordArray(NativeInt(BufIn) + Index);
  Result := P[0];
  Inc(Index,2);
  Inc(BlockPos,2);
end;

begin
  BufOut := Nil;
  Result := 0;
  if BufSize <= 0 then
    Exit;
  GetMem(BufOut,BufSize);
  Clean := True;
  Index := Offset;
  GetByte;
  GetWord;
  Pos := 0;
  BlockPos := 2;
  while Index < BufSize do begin
    Flag := GetByte;
    Mask := 1;
    while Mask < $0100 do begin
      if (BlockPos < VBA_WINDOW_SZ) and ((Flag and Mask) <> 0) then begin
        WinPos := Pos mod VBA_WINDOW_SZ;
        Shift := GetCompressShift(WinPos);
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
        // The size of a data block can never exceed VBA_WINDOW_SZ, netiher
        // decompressed nor in the stream. A special situation arises when
        // the compression "fails" (because of data that is too unequal), and
        // the size of the written stream is lager than the uncompressed data.
        // BlockPos is used keep track of the stream block size.
        if (BlockPos >= VBA_WINDOW_SZ) or ((Pos <> 0) and (Pos >= VBA_WINDOW_SZ) and Clean) then begin
          GetWord;
          Clean := False;
          ReAllocMem(BufOut,Result + Pos);
          System.Move(WinBuf,BufOut[Result],Pos);
          BlockPos := 2;
          Inc(Result,Pos);
          Pos := 0;
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
    ReAllocMem(BufOut,Result + (Pos mod VBA_WINDOW_SZ));
    System.Move(WinBuf,BufOut[Result],Pos mod VBA_WINDOW_SZ);
    Inc(Result,Pos mod VBA_WINDOW_SZ);
  end;
end;

procedure TXLSVBA.DeleteModule(ModuleName: AxUCString);
var
  i,p: integer;
  S: AxUCString;
begin
  i := ModuleIndexByName(ModuleName);
  if i < 0 then
    raise XLSRWException.Create('Can not find module ' + ModuleName);
  FModules.Delete(i);

  for i := 0 to FProject.Count - 1 do begin
    S := FProject[i];
    p := CPos('=',S);
    if p > 0 then begin
      S := Copy(S,p + 1,MAXINT);
      p := CPos('/',S);
      if p > 0 then
        S := Copy(S,1,p - 1);
{$ifdef DELPHI_5}
      if CompareText(S,ModuleName) = 0 then begin
{$else}
      if WideCompareText(S,ModuleName) = 0 then begin
{$endif}
        FProject.Delete(i);
        Break;
      end;
    end;
  end;
end;

destructor TXLSVBA.Destroy;
begin
  FVBAProject.Free;
  FDir.Free;
  FProject.Free;
  FModules.Free;
  inherited;
end;

function TXLSVBA.GetModules(Index: integer): TVBAModuleSource;
begin
  Result := TVBAModuleSource(FModules.Items[Index]);
end;

function TXLSVBA.HasData: boolean;
begin
  Result := (FSavedStorage <> Nil) or (Count > 0);
end;

procedure TXLSVBA.Int_AddModule(var Buf: PByteArray; Storage: IStorage);
var
{$ifdef BIFFSCAN}
  DBuf: TDynByteArray;
{$endif}
  WS: AxUCString;
  i,Sz: integer;
  Src: XLS8String;
  ModuleName: XLS8String;
  UCModuleName: AxUCString;
  SrcBuf,SrcBufComp: PByteArray;
  Module: TVBAModuleSource;
  Header: TVBARecHeader;
  Stream: IStream;
begin
  Module := TVBAModuleSource.Create;
  FModules.Add(Module);

  repeat
    Header.RecId := PWordArray(Buf)[0];
    Buf := PByteArray(NativeInt(Buf) + 2);
    Header.Length := PLongwordArray(Buf)[0];
    Buf := PByteArray(NativeInt(Buf) + 4);
{$ifdef BIFFSCAN}
    if Assigned(FRecordEvent) then begin
      SetLength(DBuf,Header.Length);
      System.Move(Buf^,DBuf[0],Header.Length);
      FRecordEvent(Header.RecId,'',DBuf,'');
    end;
{$endif}
    case Header.RecId of
      VBADIRREC_USERNAME: begin
        SetLength(ModuleName,Header.Length);
        System.Move(Buf^,Pointer(ModuleName)^,Header.Length);
        Module.FName := AxUCString(ModuleName);
      end;
      VBADIRREC_UC_USERNAME: begin
        SetLength(UCModuleName,Header.Length div 2);
        System.Move(Buf^,Pointer(UCModuleName)^,Header.Length);
        Module.FName := UCModuleName;
      end;
      VBADIRREC_VBAMODULEOFFSET: begin
        if UCModuleName <> '' then
          WS := UCModuleName
        else
          WS := AxUCString(ModuleName);
        OleCheck(Storage.OpenStream(PWideChar(WS),Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,Stream));
        try
          Sz := ReadStrmToBuf(Stream,SrcBufComp);
          try
            Sz := DecompressVBA(SrcBufComp,SrcBuf,Sz,PLongwordArray(Buf)[0]);
            SetLength(Src,Sz);
            System.Move(SrcBuf^,Pointer(Src)^,Sz);
            for i := Length(Src) downto 1 do begin
              if Byte(Src[i]) = $0A then
                Break
            end;
            Module.AddSrcExtractAttr(Copy(Src,1,i));
          finally
            FreeMem(SrcBuf);
          end;
        finally
          FreeMem(SrcBufComp);
          Stream := Nil;
        end;
        // Unknown data in modules are stripped when reading files.
        PLongwordArray(Buf)[0] := $0010;
        Module.FFileOffset := PLongwordArray(Buf)[0];
      end;
      VBADIRREC_TYPEMACRO:
        Module.ModuleType := vmtMacro;
      VBADIRREC_TYPEDOCUMENT:
        Module.ModuleType := vmtDocument;
    end;
    Module.FDir.AddRec(Header,Buf);
    Buf := PByteArray(NativeInt(Buf) + NativeInt(Header.Length));
  until (Header.RecId = VBADIRREC_OBJECTEND);
end;

procedure TXLSVBA.ReadDir(Storage: IStorage; Buf: PByteArray; BufSz: integer);
var
{$ifdef BIFFSCAN}
  DBuf: TDynByteArray;
{$endif}
  i: integer;
  P: PByteArray;
  ProjName: XLS8String;
  Header: TVBARecHeader;
  ModuleCount: integer;
begin
  ModuleCount := -1;
  P := Buf;
  repeat
    Header.RecId := PWordArray(P)[0];
    P := PByteArray(NativeInt(P) + 2);
    Header.Length := PLongwordArray(P)[0];
    P := PByteArray(NativeInt(P) + 4);

    case Header.RecId of
      VBADIRREC_PROJECTNAME: begin
         SetLength(ProjName,Header.Length);
         System.Move(P^,Pointer(ProjName)^,Header.Length);
       end;
      VBADIRREC_OPPAPERSIZE: Header.Length := Header.Length + 2;
      VBADIRREC_OBJECTCOUNT:
        ModuleCount := PWordArray(P)[0];
      VBADIRREC_USERNAME: begin
        P := PByteArray(NativeInt(P) - SizeOf(TVBARecHeader));
        for i := 1 to ModuleCount do
          Int_AddModule(P,Storage);
        if PWordArray(P)[0] <> VBADIRREC_EOF then
          raise XLSRWException.Create('EOF missing in VBA dir stream');
        Exit;
      end;
    end;
    FDir.AddRec(Header,P);
{$ifdef BIFFSCAN}
    if Assigned(FRecordEvent) then begin
      SetLength(DBuf,Header.Length);
      System.Move(P^,DBuf[0],Header.Length);
      FRecordEvent(Header.RecId,'',DBuf,'');
    end;
{$endif}
    P := PByteArray(NativeInt(P) + NativeInt(Header.Length));
  until (Header.RecId = VBADIRREC_EOF);
end;

function TXLSVBA.Read(Storage: IStorage): TVBAReadResult;
var
  Sz: integer;
  stgRoot,stgVBA: IStorage;
  Stream: IStream;
  Buf,B: PByteArray;
begin
  Result := vbarrFailure;
  if not OleResult(Storage.OpenStorage(VBA_ROOTDIR_NAME,Nil,STGM_READWRITE + STGM_SHARE_EXCLUSIVE,Nil,0,stgRoot)) then
    Exit;
  if not OleResult(stgRoot.OpenStorage(VBA_SUBDIR_NAME,Nil,STGM_READWRITE + STGM_SHARE_EXCLUSIVE,Nil,0,stgVBA)) then begin
    stgRoot := Nil;
    Exit;
  end;

  if FEditVBA then
    Strip(stgVBA);

  if not OleResult(stgVBA.OpenStream('dir',Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,Stream)) then begin
    stgRoot := Nil;
    stgVBA := Nil;
    Exit;
  end;
  Sz := ReadStrmToBuf(Stream,Buf);
  Stream := Nil;
  try
    if not FDisableVBA then begin
      try
        Sz := DecompressVBA(Buf,B,Sz,0);
  //      DumpBuf('d:\temp\out.bin',B,Sz);
        ReadDir(stgVBA,B,Sz);
      finally
        FreeMem(B);
      end;
    end;
    Result := vbarrReadOnly;
    FEditStateLocked := True;

    OleCheck(CreateILockBytesOnHGlobal(0,True,FSavedLockBytes));
    OleCheck(StgCreateDocfileOnILockBytes(FSavedLockBytes,STGM_CREATE + STGM_READWRITE + STGM_SHARE_EXCLUSIVE,0,FSavedStorage));
    OleCheck(stgRoot.CopyTo(0,Nil,Nil,FSavedStorage));
    if FEditVBA then begin
      OleCheck(FSavedStorage.DestroyElement(VBA_SUBDIR_NAME));
      OleCheck(FSavedStorage.DestroyElement('PROJECT'));
      OleCheck(FSavedStorage.DestroyElement('PROJECTwm'));
    end
    else
      Exit;

    if not OleResult(stgVBA.OpenStream(VBA_PROJDIR_NAME,Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,Stream)) then begin
      stgRoot := Nil;
      stgVBA := Nil;
      Exit;
    end;
{$ifdef BIFFSCAN}
    FVBAProject.Read(Stream);
{$endif}
    Stream := Nil;
  finally
    FreeMem(Buf);
  end;
  ReadPROJECT(stgRoot);
  ReadPROJECTwm(stgRoot);
  if FEditVBA and (Result = vbarrReadOnly) then begin
    FIsEditing := True;
    Result := vbarrOk;
  end;
end;

procedure TXLSVBA.ReadPROJECT(Storage: IStorage);
var
  Stream: IStream;
  S: XLS8String;
  p,Sz: integer;
begin
  if OleResult(Storage.OpenStream('PROJECT',Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,Stream)) then begin
    Sz := StreamSize(Stream);
    SetLength(S,Sz);
    OleCheck(Stream.Read(Pointer(S),Sz,@p));
    FProject.Text := AxUCString(S);
    Stream := Nil;
  end;
end;

procedure TXLSVBA.ReadPROJECTwm(Storage: IStorage);
var
  Stream: IStream;
  p,Sz: integer;
begin
  if OleResult(Storage.OpenStream('PROJECTwm',Nil,STGM_DIRECT or STGM_READ or STGM_SHARE_EXCLUSIVE,0,Stream)) then begin
    Sz := StreamSize(Stream);
    SetLength(FProjectwm,Sz);
    OleCheck(Stream.Read(FProjectwm,Sz,@p));
    Stream := Nil;
  end;
end;

function TXLSVBA.ModuleByName(ModuleName: AxUCString): TVBAModuleSource;
var
  i: integer;
begin
  for i := 0 to FModules.Count - 1 do begin
{$ifdef DELPHI_5}
    if CompareText(ModuleName,Modules[i].Name) = 0 then begin
{$else}
    if WideCompareText(ModuleName,Modules[i].Name) = 0 then begin
{$endif}
      Result := Modules[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXLSVBA.ModuleIndexByName(ModuleName: AxUCString): integer;
begin
  for Result := 0 to FModules.Count - 1 do begin
{$ifdef DELPHI_5}
    if CompareText(ModuleName,Modules[Result].Name) = 0 then
{$else}
    if WideCompareText(ModuleName,Modules[Result].Name) = 0 then
{$endif}
      Exit;
  end;
  Result := -1;
end;

function TXLSVBA.Count: integer;
begin
  Result := FModules.Count;
end;

procedure TXLSVBA.SetDisableVBA(const Value: boolean);
begin
  FDisableVBA := Value;
  if FDisableVBA then
    FEditVBA := False;
end;

procedure TXLSVBA.SetEditVBA(const Value: boolean);
begin
  if FEditStateLocked then
    raise XLSRWException.Create('VBA Edit state can not be changed after the file is opened.');
  FEditVBA := Value;
end;

function TXLSVBA.StreamSize(Stream: IStream): integer;
var
  Stat: TSTATSTG;
begin
  OleCheck(Stream.Stat(Stat,STATFLAG_NONAME));
  Result := Stat.cbSize;
end;

procedure TXLSVBA.Strip(stgVBA: IStorage);
var
  Enum: IEnumStatStg;
  Stat: TSTATSTG;
  Stream: IStream;
begin
  OleCheck(stgVBA.EnumElements(0,Nil,0,Enum));
  OleCheck(Enum.Reset);
  repeat
    OleCheck(Enum.Next(1,Stat,Nil));
    if Copy(Stat.pwcsName,1,6) = '__SRP_' then
      stgVBA.DestroyElement(Stat.pwcsName);
  until (Stat.pwcsName = Nil);

  stgVBA.Commit(STGC_DEFAULT);

  OleCheck(stgVBA.OpenStream(VBA_PROJDIR_NAME,Nil,STGM_DIRECT or STGM_READWRITE or STGM_SHARE_EXCLUSIVE,0,Stream));
  Stream.SetSize(0);
  WriteEmptyVBA_PROJECT(Stream);
  Stream := Nil;
end;

function TXLSVBA.ReadStrmToBuf(Stream: IStream; var Buf: PByteArray): integer;
var
  Stat: TSTATSTG;
  p: integer;
begin
  OleCheck(Stream.Stat(Stat,STATFLAG_NONAME));
  Result := Stat.cbSize;
  GetMem(Buf,Stat.cbSize);
  OleCheck(Stream.Read(Buf,Stat.cbSize,@p));
end;

procedure TXLSVBA.Write(Storage: IStorage);
var
  stgVBA: IStorage;
  Stream: IStream;
begin
  if FEditVBA and (Count > 0) then begin
    WritePROJECT(Storage);
    WritePROJECTwm(Storage);
    OleCheck(Storage.CreateStorage(VBA_SUBDIR_NAME, STGM_DIRECT or STGM_READWRITE or STGM_CREATE or STGM_SHARE_EXCLUSIVE,0,0,stgVBA));
    WriteModules(stgVBA);
    WriteDir(stgVBA);
    OleCheck(stgVBA.CreateStream(VBA_PROJDIR_NAME,STGM_DIRECT or STGM_READWRITE or STGM_SHARE_EXCLUSIVE,0,0,Stream));
    WriteEmptyVBA_PROJECT(Stream);
    Stream := Nil;
    stgVBA := Nil;
  end;
  if FSavedStorage <> Nil then begin
    OleCheck(FSavedStorage.CopyTo(0,Nil,Nil,Storage));
  end;
end;

procedure TXLSVBA.WriteDir(Storage: IStorage);
var
  i: integer;
  Stream: TMemoryStream;
  StreamOut: IStream;

procedure WriteDirHeader(RecId: word; Size: longword);
begin
  Stream.Write(RecId,2);
  Stream.Write(Size,4);
end;

procedure WriteDirWord(W: word);
begin
  Stream.Write(W,2);
end;

procedure WriteDirHeaderWord(RecId: word; W: word);
begin
  WriteDirHeader(RecId,2);
  Stream.Write(W,2);
end;

procedure WriteDirLWord(L: longword);
begin
  Stream.Write(L,4);
end;

procedure WriteDirHeaderLWord(RecId: word; L: longword);
begin
  WriteDirHeader(RecId,4);
  Stream.Write(L,4);
end;

procedure WriteDirStr(S: XLS8String);
begin
  Stream.Write(Pointer(S)^,Length(S));
end;

procedure WriteDirUCStr(WS: AxUCString);
begin
  Stream.Write(Pointer(WS)^,Length(WS) * 2);
end;

procedure WriteDirHeaderStr(RecId: word; S: XLS8String);
begin
  WriteDirHeader(RecId,Length(S));
  WriteDirStr(S);
end;

procedure WriteDirHeaderUCStr(RecId: word; WS: AxUCString);
begin
  WriteDirHeader(RecId,Length(WS) * 2);
  WriteDirUCStr(WS);
end;

begin
  Stream := TMemoryStream.Create;
  try
    if FIsEditing then begin
      FDir.Write(Stream);
      for i := 0 to FModules.Count - 1 do
        Modules[i].FDir.Write(Stream);
      WriteDirHeader(VBADIRREC_EOF,0);
    end
    else begin
      WriteDirHeaderLWord(VBADIRREC_START,$00000001);

      WriteDirHeaderLWord(VBADIRREC_ENCODING1,$00000409);

      WriteDirHeaderLWord(VBADIRREC_ENCODING2,$00000409);

      WriteDirHeaderWord(VBADIRREC_VERSION1,$04E4);

      WriteDirHeaderStr(VBADIRREC_PROJECTNAME,'VBAProject');

      WriteDirHeaderStr(VBADIRREC_DESCRIPTION,'');
      WriteDirHeaderUCStr(VBADIRREC_UC_DESCRIPTION,'');

      WriteDirHeaderStr(VBADIRREC_HELPFILENAME,'');
      WriteDirHeaderUCStr(VBADIRREC_UC_HELPFILENAME,'');

      WriteDirHeaderLWord(VBADIRREC_HELPCTXTID,$00000000);

      WriteDirHeaderLWord(VBADIRREC_UNKNOWN_08,$00000000);

      // Excel bug
      WriteDirHeader(VBADIRREC_OPPAPERSIZE,4);
      WriteDirWord($04B0);
      WriteDirWord($419A);
      WriteDirWord($0034);

      WriteDirHeaderStr(VBADIRREC_CONDCOMPARGS,'');
      WriteDirHeaderUCStr(VBADIRREC_UC_CONDCOMPARGS,'');

      WriteDirHeaderStr(VBADIRREC_STREAMTYPE,'stdole');
      WriteDirHeaderUCStr(VBADIRREC_UC_STREAMTYPE,'stdole');

      WriteDirHeaderWord(VBADIRREC_OBJECTCOUNT,FModules.Count);
      WriteDirHeaderWord(VBADIRREC_UNKNOWN_13,$6C1C);

      for i := 0 to FModules.Count - 1 do begin
        WriteDirHeaderStr(VBADIRREC_USERNAME,XLS8String(Modules[i].Name));
        WriteDirHeaderUCStr(VBADIRREC_UC_USERNAME,Modules[i].Name);
        WriteDirHeaderStr(VBADIRREC_STREAMNAME,XLS8String(Modules[i].Name));
        WriteDirHeaderUCStr(VBADIRREC_UC_STREAMNAME,Modules[i].Name);
        WriteDirHeader(VBADIRREC_UNKNOWN_1C,0);
        WriteDirHeader(VBADIRREC_UNKNOWN_48,0);
        WriteDirHeaderLWord(VBADIRREC_VBAMODULEOFFSET,Modules[i].FFileOffset);
        WriteDirHeaderLWord(VBADIRREC_UNKNOWN_1E,$00000000);
        WriteDirHeaderWord(VBADIRREC_UNKNOWN_2C,$869D);
        case Modules[i].ModuleType of
          vmtMacro:    WriteDirHeader(VBADIRREC_TYPEMACRO,0);
          vmtDocument: WriteDirHeader(VBADIRREC_TYPEDOCUMENT,0);
        end;
        WriteDirHeader(VBADIRREC_OBJECTEND,0);
      end;

      WriteDirHeader(VBADIRREC_EOF,0);
    end;
    Stream.Seek(0,soFromBeginning);

    OleCheck(Storage.CreateStream('dir',STGM_DIRECT or STGM_READWRITE or STGM_SHARE_EXCLUSIVE,0,0,StreamOut));

    CompressVba(Stream.Memory,Stream.Size,StreamOut);
  finally
    Stream.Free;
  end;
end;

procedure TXLSVBA.WriteEmptyVBA_PROJECT(Stream: IStream);
const
  HdrData: array[0..7] of byte = ($CC,$61,$A0,$00,$00,$01,$00,$00);
var
  i,p: integer;
begin
  OleCheck(Stream.Write(@HdrData,8,@p));                                           
  for i := 1 to 408 do
    WriteWord(Stream,$0000);
end;

procedure TXLSVBA.WriteModules(Storage: IStorage);
var
  i,j: integer;
  S: XLS8String;
  Stream: IStream;
begin
  for i := 0 to FModules.Count - 1 do begin
    OleCheck(Storage.CreateStream(PWideChar(Modules[i].Name),STGM_DIRECT or STGM_READWRITE or STGM_SHARE_EXCLUSIVE,0,0,Stream));
    for j := 1 to 8 do
      WriteWord(Stream,$0000);
    Modules[i].FFileOffset := StreamSize(Stream);
    S := XLS8String(Modules[i].FAttributes.PrefixText('Attribute') + #13#10 + Modules[i].Source.Text);
    CompressVBA(Pointer(S),Length(S),Stream);
  end;
end;

procedure TXLSVBA.WritePROJECT(Storage: IStorage);
var
  i: integer;
  Stream: IStream;
begin
  OleCheck(Storage.CreateStream('PROJECT',STGM_DIRECT or STGM_READWRITE or STGM_SHARE_EXCLUSIVE,0,0,Stream));
  if FIsEditing then
    WriteString(Stream,XLS8String(FProject.Text))
  else begin
    WriteStringCRLF(Stream,'ID="{921117CD-5376-4972-9C98-4AA04F3488F7}"');
    for i := 0 to FModules.Count - 1 do begin
      case Modules[i].ModuleType of
        vmtMacro:    WriteStringCRLF(Stream,XLS8String('Module=' + Modules[i].Name));
        vmtDocument: WriteStringCRLF(Stream,XLS8String('Document=' + Modules[i].Name + '/&H00000000'));
      end;
    end;
    WriteStringCRLF(Stream,'Name="VBAProject"');
    WriteStringCRLF(Stream,'HelpContextID="0"');
    WriteStringCRLF(Stream,'VersionCompatible32="393222000"');
    Stream := Nil;
  end;
end;

procedure TXLSVBA.WritePROJECTwm(Storage: IStorage);
var
  i,p: integer;
  Stream: IStream;
begin
  OleCheck(Storage.CreateStream('PROJECTwm',STGM_DIRECT or STGM_READWRITE or STGM_SHARE_EXCLUSIVE,0,0,Stream));
  if FIsEditing then
    Stream.Write(FProjectwm,Length(FProjectwm),@p)
  else begin
    for i := 0 to FModules.Count - 1 do begin
      WriteString(Stream,XLS8String(Modules[i].Name));
      WriteByte(Stream,$00);
      WriteUCString(Stream,Modules[i].Name);
      WriteWord(Stream,$0000);
    end;
    WriteWord(Stream,$0000);
    Stream := Nil;
  end;
end;

{ TVBAModuleSource }

procedure TVBAModuleSource.AddSrcExtractAttr(Src: XLS8String);
begin
  FSource.Text := AxUCString(Src);
  while FSource.Count > 0 do begin
    if Lowercase(Copy(FSource[0],1,9)) = 'attribute' then begin
      FAttributes.AddSkipFirstWord(FSource[0]);
      FSource.Delete(0);
    end
    else
      Break;
  end;
end;

constructor TVBAModuleSource.Create;
begin
  inherited Create;
  FSource := TStringList.Create;
end;

destructor TVBAModuleSource.Destroy;
begin
  FSource.Free;
  inherited;
end;

{ TXLSVBAProject }

procedure TXLSVBAProject.Clear;
begin
  SetLength(FLinkStrings,0);
  SetLength(FModules,0);
end;

function TXLSVBAProject.Read(Stream: IStream): boolean;
var
  p: longword;
{$ifdef BABOON}
  p2: largeint;
{$else}
{$ifdef DELPHI_XE8_OR_LATER}
  p2: largeuint;
{$else}
  p2: largeint;
{$endif}
{$endif}
  W: word;
  S: AxUCString;
  i,j,Cnt: integer;

procedure DebugRecord(const Buffer; Count: Longint; Name,Desc: AxUCString);
{$ifdef BIFFSCAN}
var
  D: TDynByteArray;
{$endif}
begin
{$ifdef BIFFSCAN}
  if Assigned(FRecordEvent) then begin
    SetLength(D,Count);
    Move(Buffer,D[0],Count);
    FRecordEvent(0,Name,D,Desc);
  end;
{$endif}
end;

function ReadWord: word;
begin
  OleCheck(Stream.Read(@Result,2,@p));
end;

function ReadLWord: longword;
begin
  OleCheck(Stream.Read(@Result,4,@p));
end;

function ReadUCString: AxUCString;
var
  B: PByteArray;
  L: word;
begin
  Result := '';
  L := ReadWord;
  if L <= 6 then
    Exit;
  GetMem(B,L + 10);
  try
    SetLength(Result,L div 2);
    OleCheck(Stream.Read(B,L + 10,@p));
    Move(B^,Pointer(Result)^,L);
    if not CharInSet(Result[3],['G','H','C','D']) then begin
      OleCheck(Stream.Seek(-8,STREAM_SEEK_CUR,p2));
      Result := '';
      Exit;
    end;
    DebugRecord(L,2,'String len',Format('Length = %d',[L div 2]));
    DebugRecord(B^,L,'String',Format('Text = %s',[Result]));
    DebugRecord(B[L],10,'String suffix','');

    L := ReadWord;
    DebugRecord(L,2,'Extra size','');
    if L > 0 then begin
      OleCheck(Stream.Read(B,2,@p));
      if PWordArray(B)[0] < 6 then begin
        OleCheck(Stream.Read(@B[2],18,@p));
        DebugRecord(B^,20,'Extra data','');
      end
      else
        OleCheck(Stream.Seek(-2,STREAM_SEEK_CUR,p2));
    end;
  finally
    FreeMem(B);
  end;

end;

begin
  Result := False;
  W := ReadWord;
  if W <> $61CC then
    Exit;
  DebugRecord(W,2,'File id','');
  OleCheck(Stream.Read(@FProjVersion,SizeOf(TRecVBAProjVersion),@p));
  // Mac is big endian...
  if FProjVersion.Mac <> $0100 then
    Exit;

  DebugRecord(FProjVersion,SizeOf(TRecVBAProjVersion),'Version',Format('VBA = %.4X' + #13 + 'Mac = %.4X' + #13 + 'Unknown = %.4X',[FProjVersion.ExcelVBA,FProjVersion.Mac,FProjVersion.Unknown]));

  OleCheck(Stream.Read(@FLangId,SizeOf(TRecVBALanguageId),@p));

  DebugRecord(FLangId,SizeOf(TRecVBALanguageId),'Language id',Format(
     'Id1 = %.4X' + #13 +
     'Id2 = %.4X' + #13 +
     'CharSet = %.2X' + #13 +
     'Len1 = %.2X' + #13 +
     'Unknown1 = %.4X' + #13 +
     'Unknown2 = %.4X' + #13 +
     'Len2 = %.2X' + #13 +
     'Len3 = %.2X' + #13 +
     'Len4 = %.2X' + #13,[
     FLangId.Id1,
     FLangId.Id2,
     FLangId.CharSet,
     FLangId.Len1,
     FLangId.Unknown1,
     FLangId.Unknown2,
     FLangId.Len2,
     FLangId.Len3,
     FLangId.Len4]));
{
  if FLangId.Len4 <> $0002 then
    Exit;
}
  Cnt := 0;
  SetLength(FLinkStrings,10);
  while True do begin
    S := ReadUCString;
    if Length(S) >= 20 then begin
      Inc(Cnt);
      if Cnt > Length(FLinkStrings) then
        SetLength(FLinkStrings,Length(FLinkStrings) + 10);
      FLinkStrings[Cnt - 1] := S;
    end
    else begin
      OleCheck(Stream.Seek(-2,STREAM_SEEK_CUR,p2));
      Break;
    end;
  end;
  SetLength(FLinkStrings,Cnt);
  W := ReadWord;
  OleCheck(Stream.Seek(W * 2,STREAM_SEEK_CUR,p2));
  W := ReadWord;
  OleCheck(Stream.Seek(W * 4,STREAM_SEEK_CUR,p2));
  OleCheck(Stream.Seek(2,STREAM_SEEK_CUR,p2));

  for i := 0 to 2 do begin
    W := ReadWord;
    if W <> $FFFF then
      OleCheck(Stream.Seek(W,STREAM_SEEK_CUR,p2));
  end;
  OleCheck(Stream.Seek(100,STREAM_SEEK_CUR,p2));

  W := ReadWord;
  if (W = $FFFF) or (W = 0) then
    Exit;
  SetLength(FModules,W);
  for I := 0 to High(FModules) do begin
    W := ReadWord;
    SetLength(S,W div 2);
    OleCheck(Stream.Read(Pointer(S),W,@p));
    FModules[i].Name := S;

    W := ReadWord;
    OleCheck(Stream.Seek(W,STREAM_SEEK_CUR,p2));

    W := ReadWord;
		if W = $FFFF then begin
      OleCheck(Stream.Seek(2,STREAM_SEEK_CUR,p2));
      W := ReadWord;
      OleCheck(Stream.Seek(W,STREAM_SEEK_CUR,p2));
    end
		else
      OleCheck(Stream.Seek(W + 2,STREAM_SEEK_CUR,p2));

    W := ReadWord;
		if W <> $FFFF then
      Exit;

    OleCheck(Stream.Seek(6,STREAM_SEEK_CUR,p2));
    W := ReadWord;
    for j := 0 to W - 1 do
      OleCheck(Stream.Seek(8,STREAM_SEEK_CUR,p2));

    OleCheck(Stream.Seek(5,STREAM_SEEK_CUR,p2));

    FModules[i].Offset := ReadLWord;
    OleCheck(Stream.Seek(2,STREAM_SEEK_CUR,p2));
  end;

end;

{ TRecordStorageVBA }

function TRecordStorageVBA.AddRec(Header: TVBARecHeader; Data: PByteArray): PVBARecordData;
begin
  GetMem(Result,Header.Length + SizeOf(TVBARecordData));
  Result.RecId := Header.RecId;
  Result.Length := Header.Length;
//  System.Move(Header,Result^,SizeOf(TVBARecHeader));
  System.Move(Data^,Result.Data,Header.Length);
  inherited Add(Result);
end;

procedure TRecordStorageVBA.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i] <> Nil then
      FreeMem(inherited Items[i]);
  end;
  inherited Clear;
end;

procedure TRecordStorageVBA.Delete(Index: integer);
begin
  if Items[Index] <> Nil then
    FreeMem(inherited Items[Index]);
  inherited Delete(Index);
end;

procedure TRecordStorageVBA.RenameModule(Name: AxUCString);
var
  i: integer;
begin
  i := 0;
  while i < Count do begin
     if (Items[i].RecId = VBADIRREC_USERNAME) then begin
       SetAsString(i,XLS8String(Name));
       Inc(i);
       if (i < Count) and (Items[i].RecId = VBADIRREC_UC_USERNAME) then
         SetAsWideString(i,Name);
       Inc(i);
       if (i < Count) and (Items[i].RecId = VBADIRREC_STREAMNAME) then
         SetAsString(i,XLS8String(Name));
       Inc(i);
       if (i < Count) and (Items[i].RecId = VBADIRREC_UC_STREAMNAME) then
         SetAsString(i,XLS8String(Name));
       Exit;
     end;
     Inc(i);
  end;
end;

procedure TRecordStorageVBA.SetAsString(Index: integer; const Value: XLS8String);
var
  P: PByteArray;
begin
  P := inherited Items[Index];
  Items[Index].Length := Length(Value);
  ReAllocMem(P,Items[Index].Length + SizeOf(TVBARecordData));
  System.Move(Pointer(Value)^,Items[Index].Data,Items[Index].Length);
  inherited Items[Index] := P;
  if GetAsString(Index) <> Value then
    raise XLSRWException.Create('Rename Error');
end;

procedure TRecordStorageVBA.SetAsWideString(Index: integer; const Value: AxUCString);
var
  P: PByteArray;
begin
  P := inherited Items[Index];
  Items[Index].Length := Length(Value) * 2;
  ReAllocMem(P,Items[Index].Length + SizeOf(TVBARecordData));
  System.Move(Pointer(Value)^,Items[Index].Data,Items[Index].Length);
  inherited Items[Index] := P;
  if GetAsWideString(Index) <> Value then
    raise XLSRWException.Create('Rename Error');
end;

destructor TRecordStorageVBA.Destroy;
begin

  inherited;
end;

function TRecordStorageVBA.GetAsString(Index: integer): XLS8String;
begin
  SetLength(Result,Items[Index].Length);
  System.Move(Items[Index].Data,Pointer(Result)^,Items[Index].Length);
end;

function TRecordStorageVBA.GetAsWideString(Index: integer): AxUCString;
begin
  SetLength(Result,Items[Index].Length div 2);
  System.Move(Items[Index].Data,Pointer(Result)^,Items[Index].Length);
end;

function TRecordStorageVBA.GetItems(Index: integer): PVBARecordData;
begin
  Result := PVBARecordData(Inherited Items[Index]);
end;

procedure TRecordStorageVBA.Write(Stream: TStream);
var
  i: integer;
  LW: longword;
begin
  for i := 0 to Count - 1 do begin
    Stream.Write(Items[i].RecId,2);
    // Excel bug
    if Items[i].RecId = VBADIRREC_OPPAPERSIZE then begin
      LW := 4;
      Stream.Write(LW,4);
    end
    else
      Stream.Write(Items[i].Length,4);
    Stream.Write(Items[i].Data,Items[i].Length);
  end;
end;

{ TVBAModule }

constructor TVBAModule.Create;
begin
  FAttributes := TXLSValueWideStringList.Create;
  FDir := TRecordStorageVBA.Create;
end;

destructor TVBAModule.Destroy;
begin
  FAttributes.Free;
  FDir.Free;
  inherited;
end;

procedure TVBAModule.SetName(const Value: AxUCString);
begin
  FName := Value;
  FDir.RenameModule(Value);
end;

end.
