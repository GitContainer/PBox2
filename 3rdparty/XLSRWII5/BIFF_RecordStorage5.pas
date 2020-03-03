unit BIFF_RecordStorage5;

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
     BIFF_RecsII5, BIFF_Stream5,
     XLSUtils5;

// Globals
const INTERNAL_PLACEHOLDER = $8000;
const INTERNAL_FORMATS     = $8001;
const INTERNAL_BOUNDSHEETS = $8002;
const INTERNAL_NAMES       = $8003;
const INTERNAL_MSODRWGRP   = $8004;
const INTERNAL_SST         = $8005;

// Worksheet
const INTERNAL_PAGEBREAKES = $8101;
const INTERNAL_HEADER      = $8102;
const INTERNAL_MARGINS     = $8103;
const INTERNAL_COLINFO     = $8104;
const INTERNAL_CELLDATA    = $8105;
const INTERNAL_SUFFIXDATA  = $8106;

type PRecordData = ^TRecordData;
     TRecordData = record
     RecId: word;
     Length: word;
     Data: record end;
     end;

type TBaseRecordStorage = class(TList)
private
     function  GetItems(Index: integer): PRecordData;
     procedure SetItems(Index: integer; const Value: PRecordData);
protected
public
     destructor Destroy; override;
     procedure Clear; override;
     function  AddRec(Header: TBIFFHeader; Data: PByteArray): PRecordData;
     function  UpdateRec(Index: integer; Len: integer; Data: PByteArray): PRecordData;
     procedure ReadUntilEOF(Stream: TXLSStream; PBuf: PByteArray);
     procedure WriteAllRecs(Stream: TXLSStream); virtual;
     procedure WriteRec(Index: integer; Stream: TXLSStream);
     function  Size: integer;
     function  FindRecord(RecId: word): integer;
     procedure Assign(Records: TBaseRecordStorage);

     property Items[Index: integer]: PRecordData read GetItems write SetItems; default;
     end;

type TRecordStorage = class(TBaseRecordStorage)
private
protected
public
     procedure SetDefaultData; virtual; abstract;
     function  PostCheck: integer; virtual; abstract;
     procedure UpdateDefault(Header: TBIFFHeader; Data: PByteArray); virtual; abstract;
     procedure UpdateInternal(Id: word); virtual; abstract;
     end;

type TRecordStorageEx2007Unknown = class(TBaseRecordStorage)
private
protected
public
     procedure WriteAllRecs(Stream: TXLSStream); override;
     end;


type TRecordStorageDefault = class(TRecordStorage)
protected
     FDefault: TBaseRecordStorage;

     function  AddDefRecWord(Id: integer; Data: word): PRecordData;
     function  AddDefRec(Id,Length: integer): PRecordData;
public
     constructor Create;
     destructor Destroy; override;
     procedure Clear; override;
     procedure MoveDefault(Rec,NewRec: PRecordData); virtual; abstract;
     procedure UpdateRec(var Rec: PRecordData; Header: TBIFFHeader; Data: PByteArray);
     procedure MoveAllDefault;
     end;

type TRecordStorageGlobals = class(TRecordStorageDefault)
private
     P_INT_FORMATS:     PRecordData;
     P_INT_BOUNDSHEETS: PRecordData;
     P_INT_NAMES:       PRecordData;
     P_INT_MSODRWGRP:   PRecordData;
     P_INT_SST:         PRecordData;

     PBOF:              PRecordData;
     PWRITEACCESS:      PRecordData;
     PCODEPAGE:         PRecordData;
     PDSF:              PRecordData;
     PTABID:            PRecordData;
     PFNGROUPCOUNT:     PRecordData;
     PWINDOWPROTECT:    PRecordData;

     PPROTECT:          PRecordData;
     PPASSWORD:         PRecordData;
     PPROT4REV:         PRecordData;

     PWINDOW1:          PRecordData;
     PBACKUP:           PRecordData;
     PHIDEOBJ:          PRecordData;
     P1904:             PRecordData;
     PPRECISION:        PRecordData;
     PREFRESHALL:       PRecordData;
     PBOOKBOOL:         PRecordData;
     PCOUNTRY:          PRecordData;
     PRECALCID:         PRecordData;
     PEOF:              PRecordData;

     function  GetCODEPAGE: word;
     function  GetWINDOW1: PRecWINDOW1;
     function  GetWINDOWPROTECT: boolean;
     function  GetWRITEACCESS: AxUCString;
     procedure SetCODEPAGE(const Value: word);
     procedure SetWINDOWPROTECT(const Value: boolean);
     procedure SetWRITEACCESS(const Value: AxUCString);
     function  GetBACKUP: boolean;
     procedure SetBACKUP(const Value: boolean);
     function  GetHIDEOBJ: word;
     procedure SetHIDEOBJ(const Value: word);
     function  GetPRECISION: boolean;
     procedure SetRECISION(const Value: boolean);
     function  GetREFRESHALL: boolean;
     procedure SetREFRESHALL(const Value: boolean);
     function  GetBOOKBOOL: boolean;
     procedure SetBOOKBOOL(const Value: boolean);
     function  GetCOUNTRY: PRecCOUNTRY;
     function  GetRECALCID: PRecRECALCID;
     function  GetBOF: PRecBOF8;
     function  GetDATE1904: boolean;
     function  GetPASSWORD: word;
protected
public
     constructor Create;
     procedure Clear; override;
     procedure SetDefaultData; override;
     function  PostCheck: integer; override;
     procedure MoveDefault(Rec,NewRec: PRecordData); override;
     procedure UpdateDefault(Header: TBIFFHeader; Data: PByteArray); override;
     procedure UpdateInternal(Id: word); override;

     property BOF: PRecBOF8 read GetBOF;
     property CODEPAGE: word read GetCODEPAGE write SetCODEPAGE;
     property WRITEACCESS: AxUCString read GetWRITEACCESS write SetWRITEACCESS;
     property WINDOWPROTECT: boolean read GetWINDOWPROTECT write SetWINDOWPROTECT;
     property WINDOW1: PRecWINDOW1 read GetWINDOW1;
     property BACKUP: boolean read GetBACKUP write SetBACKUP;
     property HIDEOBJ: word read GetHIDEOBJ write SetHIDEOBJ;
     property PRECISION: boolean read GetPRECISION write SetRECISION;
     property REFRESHALL: boolean read GetREFRESHALL write SetREFRESHALL;
     property BOOKBOOL: boolean read GetBOOKBOOL write SetBOOKBOOL;
     property COUNTRY: PRecCOUNTRY read GetCOUNTRY;
     property RECALCID: PRecRECALCID read GetRECALCID;
     property DATE1904: boolean read GetDATE1904;
     property PASSWORD: word read GetPASSWORD;
     end;

type TRecordStorageSheet = class(TRecordStorageDefault)
private
     P_INT_PAGEBREAKES: PRecordData;
     P_INT_HEADER:      PRecordData;
     P_INT_MARGINS:     PRecordData;
     P_INT_COLINFO:     PRecordData;
     P_INT_CELLDATA:    PRecordData;
     P_INT_SUFFIXDATA:  PRecordData;

     PBOF:              PRecordData;
     PCALCMODE:         PRecordData;
     PCALCCOUNT:        PRecordData;
     PDELTA:            PRecordData;
     PREFMODE:          PRecordData;
     PITERATION:        PRecordData;
     PSAVERECALC:       PRecordData;
     PPRINTHEADERS:     PRecordData;
     PPRINTGRIDLINES:   PRecordData;
     PGRIDSET:          PRecordData;
     PGUTS:             PRecordData;
     PDEFAULTROWHEIGHT: PRecordData;
     PWSBOOL:           PRecordData;
     PHCENTER:          PRecordData;
     PVCENTER:          PRecordData;
     PSETUP:            PRecordData;
     PPROTECT:          PRecordData;
     PDEFCOLWIDTH:      PRecordData;
     PDIMENSIONS:       PRecordData;
     PWINDOW2:          PRecordData;
     PEOF:              PRecordData;
     // Records with index are of variable size, and may need re-allocation.
//     IPLS:              integer;

     function  GetCALCMODE: word;
     procedure SetCALCMODE(const Value: word);
     function  GetDELTA: double;
     procedure SetDELTA(const Value: double);
     function  GetCALCCOUNT: word;
     procedure SetCALCCOUNT(const Value: word);
     function  GetITERATION: boolean;
     function  GetREFMODE: word;
     function  GetSAVERECALC: boolean;
     procedure SetITERATION(const Value: boolean);
     procedure SetREFMODE(const Value: word);
     procedure SetSAVERECALC(const Value: boolean);
     function  GetPRINTGRIDLINES: boolean;
     function  GetPRINTHEADERS: boolean;
     procedure SetPRINTGRIDLINES(const Value: boolean);
     procedure SetPRINTHEADERS(const Value: boolean);
     function  GetGRIDSET: word;
     procedure SetGRIDSET(const Value: word);
     function  GetGUTS: PRecGUTS;
     function  GetDEFAULTROWHEIGHT: PRecDEFAULTROWHEIGHT;
     function  GetWSBOOL: word;
     procedure SetWSBOOL(const Value: word);
     function  GetHCENTER: boolean;
     procedure SetHCENTER(const Value: boolean);
     function  GetVCENTER: boolean;
     procedure SetVCENTER(const Value: boolean);
     function  GetSETUP: PRecSETUP;
     function  GetDEFCOLWIDTH: word;
     procedure SetDEFCOLWIDTH(const Value: word);
     function  GetDIMENSIONS: PRecDIMENSIONS8;
     function  GetWINDOW2: PRecWINDOW2_8;
     function  GetBOF: PRecBOF8;
     procedure SetPROTECT(const Value: word);
     function  GetPROTECT: word;
protected
public
     constructor Create;
     procedure Clear; override;
     procedure SetDefaultData; override;
     function  PostCheck: integer; override;
     procedure MoveDefault(Rec,NewRec: PRecordData); override;
     procedure UpdateDefault(Header: TBIFFHeader; Data: PByteArray); override;
     procedure UpdateInternal(Id: word); override;

     property BOF: PRecBOF8 read GetBOF;
     property CALCMODE: word read GetCALCMODE write SetCALCMODE;
     property CALCCOUNT: word read GetCALCCOUNT write SetCALCCOUNT;
     property DELTA: double read GetDELTA write SetDELTA;
     property REFMODE: word read GetREFMODE write SetREFMODE;
     property ITERATION: boolean read GetITERATION write SetITERATION;
     property SAVERECALC: boolean read GetSAVERECALC write SetSAVERECALC;
     property PRINTHEADERS: boolean read GetPRINTHEADERS write SetPRINTHEADERS;
     property PRINTGRIDLINES: boolean read GetPRINTGRIDLINES write SetPRINTGRIDLINES;
     property GRIDSET: word read GetGRIDSET write SetGRIDSET;
     property GUTS: PRecGUTS read GetGUTS;
     property DEFAULTROWHEIGHT: PRecDEFAULTROWHEIGHT read GetDEFAULTROWHEIGHT;
     property WSBOOL: word read GetWSBOOL write SetWSBOOL;
     property HCENTER: boolean read GetHCENTER write SetHCENTER;
     property VCENTER: boolean read GetVCENTER write SetVCENTER;
//     property PLS: integer read IPLS;
     property SETUP: PRecSETUP read GetSETUP;
     property DEFCOLWIDTH: word read GetDEFCOLWIDTH write SetDEFCOLWIDTH;
     property DIMENSIONS: PRecDIMENSIONS8 read GetDIMENSIONS;
     property WINDOW2: PRecWINDOW2_8 read GetWINDOW2;
     property PROTECT: word read GetPROTECT write SetPROTECT;
     end;

implementation

{$ifdef ver130}
type PDouble = ^Double;
{$endif}

{ TBaseRecordStorage }

function TBaseRecordStorage.AddRec(Header: TBIFFHeader; Data: PByteArray): PRecordData;
begin
  GetMem(Result,Header.Length + SizeOf(TRecordData));
  Result.RecId := Header.RecID;
  Result.Length := Header.Length;
  System.Move(Data^,Result.Data,Header.Length);
  inherited Add(Result);
end;

procedure TBaseRecordStorage.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i] <> Nil then
      FreeMem(inherited Items[i]);
  end;
  inherited Clear;
end;

destructor TBaseRecordStorage.Destroy;
begin
  inherited;
end;

function TBaseRecordStorage.FindRecord(RecId: word): integer;
begin
  for Result := 0 to Count - 1 do begin
    if GetItems(Result).RecId = RecId then
      Exit;
  end;
  Result := -1;
end;

function TBaseRecordStorage.GetItems(Index: integer): PRecordData;
begin
  Result := PRecordData(Inherited Items[Index]);
end;

procedure TBaseRecordStorage.ReadUntilEOF(Stream: TXLSStream; PBuf: PByteArray);
var
  Header: TBIFFHeader;
begin
  while Stream.ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
    Stream.Read(PBuf^,Header.Length);
    AddRec(Header,PBuf);
    if Header.RecId = BIFFRECID_EOF then
      Break;
  end;
end;

procedure TBaseRecordStorage.WriteRec(Index: integer; Stream: TXLSStream);
var
  Buf: array[0..1] of word;
begin
  Buf[0] := Self[Index].RecId;
  Buf[1] := Self[Index].Length;
  Stream.Write(Buf,4);
  Stream.Write(Self[Index].Data,Self[Index].Length);
end;

procedure TBaseRecordStorage.WriteAllRecs(Stream: TXLSStream);
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    WriteRec(i,Stream);
end;

procedure TBaseRecordStorage.SetItems(Index: integer; const Value: PRecordData);
begin
  inherited Items[Index] := Value;
end;

// NOTE! Size gives the disk storage size, not the memory used.
function TBaseRecordStorage.Size: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Inc(Result,SizeOf(TBIFFHeader) + Items[i].Length);
end;

function TBaseRecordStorage.UpdateRec(Index: integer; Len: integer; Data: PByteArray): PRecordData;
begin
  Result := inherited Items[Index];
  Result.Length := Len;
  ReallocMem(Result,Result.Length + SizeOf(TRecordData));
  System.Move(Data^,Result.Data,Result.Length);
  inherited Items[Index] := Result;
end;

procedure TBaseRecordStorage.Assign(Records: TBaseRecordStorage);
var
  i: integer;
  H: TBIFFHeader;
begin
  for i := 0 to Count - 1 do begin
    H.RecID := Items[i].RecId;
    H.Length := Items[i].Length;
    Records.AddRec(H,@Items[i].Data);
  end;
end;

{ TRecordStorageGlobals }

procedure TRecordStorageGlobals.UpdateDefault(Header: TBIFFHeader; Data: PByteArray);
begin
  case Header.RecID of
    BIFFRECID_CODEPAGE:         UpdateRec(PCODEPAGE,Header,Data);
    BIFFRECID_DSF:              UpdateRec(PDSF,Header,Data);
    BIFFRECID_TABID:            UpdateRec(PTABID,Header,Data);
    BIFFRECID_FNGROUPCOUNT:     UpdateRec(PFNGROUPCOUNT,Header,Data);
    BIFFRECID_WRITEACCESS:      UpdateRec(PWRITEACCESS,Header,Data);
    BIFFRECID_WINDOWPROTECT:    UpdateRec(PWINDOWPROTECT,Header,Data);
    BIFFRECID_PROTECT:          UpdateRec(PPROTECT,Header,Data);
    BIFFRECID_PASSWORD:         UpdateRec(PPASSWORD,Header,Data);
    BIFFRECID_PROT4REV:         UpdateRec(PPROT4REV,Header,Data);
    BIFFRECID_WINDOW1:          UpdateRec(PWINDOW1,Header,Data);
    BIFFRECID_HIDEOBJ:          UpdateRec(PHIDEOBJ,Header,Data);
    BIFFRECID_1904:             UpdateRec(P1904,Header,Data);
    BIFFRECID_BACKUP:           UpdateRec(PBACKUP,Header,Data);
    BIFFRECID_PRECISION:        UpdateRec(PPRECISION,Header,Data);
    BIFFRECID_REFRESHALL:       UpdateRec(PREFRESHALL,Header,Data);
    BIFFRECID_BOOKBOOL:         UpdateRec(PBOOKBOOL,Header,Data);
    BIFFRECID_COUNTRY:          UpdateRec(PCOUNTRY,Header,Data);
    BIFFRECID_RECALCID:         UpdateRec(PRECALCID,Header,Data);
    BIFFRECID_EOF:              UpdateRec(PEOF,Header,Data);
    else begin

      if (Header.RecID and $FF) = BIFFRECID_BOF then
        UpdateRec(PBOF,Header,Data);

    end;
  end;
end;

procedure TRecordStorageGlobals.Clear;
begin
  inherited Clear;
end;

function TRecordStorageGlobals.PostCheck: integer;
begin
  Result := 0;
end;

function TRecordStorageGlobals.GetBACKUP: boolean;
begin
  Result := PWordArray(@PBACKUP.Data)[0] = 1;
end;

function TRecordStorageGlobals.GetBOOKBOOL: boolean;
begin
  Result := PWordArray(@PBOOKBOOL.Data)[0] = 1;
end;

function TRecordStorageGlobals.GetCODEPAGE: word;
begin
  Result := PWordArray(@PCODEPAGE.Data)[0];
end;

function TRecordStorageGlobals.GetCOUNTRY: PRecCOUNTRY;
begin
  Result := PRecCOUNTRY(@PCOUNTRY.Data);
end;

function TRecordStorageGlobals.GetRECALCID: PRecRECALCID;
begin
  Result := PRecRECALCID(@PRECALCID.Data);
end;
function TRecordStorageGlobals.GetHIDEOBJ: word;
begin
  Result := PWordArray(@PHIDEOBJ.Data)[0];
end;

function TRecordStorageGlobals.GetPASSWORD: word;
begin
  if PPASSWORD <> Nil then
    Result := PWordArray(@PPASSWORD.Data)[0]
  else
    Result := 0;
end;

function TRecordStorageGlobals.GetPRECISION: boolean;
begin
  Result := PWordArray(@PPRECISION.Data)[0] = 1;
end;

function TRecordStorageGlobals.GetREFRESHALL: boolean;
begin
  Result := PWordArray(@PREFRESHALL.Data)[0] = 1;
end;

function TRecordStorageGlobals.GetWINDOW1: PRecWINDOW1;
begin
  Result := PRecWINDOW1(@PWINDOW1.Data);
end;

function TRecordStorageGlobals.GetWINDOWPROTECT: boolean;
begin
  Result := PWordArray(@PWINDOWPROTECT.Data)[0] = 1;
end;

function TRecordStorageGlobals.GetWRITEACCESS: AxUCString;
var
  L: integer;
  S: XLS8String;
  WS: AxUCString;
begin
  Result := '';
  L := PWordArray(@PWRITEACCESS.Data)[0];
  if L <= 0 then
    Exit;
  if L > 109 then
    L := 109;
  if PByteArray(@PWRITEACCESS.Data)[2] = 0 then begin
    SetLength(S,L);
    System.Move(PByteArray(@PWRITEACCESS.Data)[3],Pointer(S)^,L);
    Result := AxUCString(S);
  end
  else if PByteArray(@PWRITEACCESS.Data)[2] = 1 then begin
    SetLength(WS,L);
    System.Move(PByteArray(@PWRITEACCESS.Data)[3],Pointer(WS)^,L * 2);
    Result := WS;
  end;
  Result := Trim(Result);
end;

procedure TRecordStorageGlobals.SetBACKUP(const Value: boolean);
begin
  PWordArray(@PBACKUP.Data)[0] := Word(Value);
end;

procedure TRecordStorageGlobals.SetBOOKBOOL(const Value: boolean);
begin
  PWordArray(@PBOOKBOOL.Data)[0] := Word(Value);
end;

procedure TRecordStorageGlobals.SetCODEPAGE(const Value: word);
begin
  PWordArray(@PCODEPAGE.Data)[0] := Value;
end;

procedure TRecordStorageGlobals.SetDefaultData;
begin
  Clear;

  PBOF := AddDefRec($0809,SizeOf(TRecBOF8));
  BOF.VersionNumber := $0600;
  BOF.BuildIdentifier := $18AF;
  BOF.BuildYear := $07CD;
  BOF.FileHistoryFlags := $000080C9;
  BOF.LowBIFF := $00000106;
  BOF.SubstreamType := $0005;

  PWRITEACCESS := AddDefRec(BIFFRECID_WRITEACCESS,112);
  WRITEACCESS  := '';

  PCODEPAGE := AddDefRecWord(BIFFRECID_CODEPAGE,$04B0);
  PDSF := AddDefRecWord(BIFFRECID_DSF,$0000);
  PTABID := AddDefRec(BIFFRECID_TABID,0);
  PFNGROUPCOUNT := AddDefRecWord(BIFFRECID_FNGROUPCOUNT,$000E);

  PWINDOWPROTECT := AddDefRecWord(BIFFRECID_WINDOWPROTECT,$0000);
  PPROTECT := AddDefRecWord(BIFFRECID_PROTECT,$0000);
  PPASSWORD := AddDefRecWord(BIFFRECID_PASSWORD,$0000);
  PPROT4REV := AddDefRecWord(BIFFRECID_PROT4REV,$0000);

  PWINDOW1 := AddDefRec(BIFFRECID_WINDOW1,SizeOf(TRecWINDOW1));
  WINDOW1.Left := 100;
  WINDOW1.Top := 100;
  WINDOW1.Width := 10000;
  WINDOW1.Height := 7000;
  WINDOW1.Options := $0038;
  WINDOW1.SelectedTabIndex := 0;
  WINDOW1.FirstDispTabIndex := 0;
  WINDOW1.SelectedTabs := 1;
  WINDOW1.TabRatio := $0258;

  PBACKUP := AddDefRecWord(BIFFRECID_BACKUP,0);
  PHIDEOBJ := AddDefRecWord(BIFFRECID_HIDEOBJ,0);
  P1904 := AddDefRecWord(BIFFRECID_1904,0);
  PPRECISION := AddDefRecWord(BIFFRECID_PRECISION,1);
  PREFRESHALL := AddDefRecWord(BIFFRECID_REFRESHALL,0);
  PBOOKBOOL := AddDefRecWord(BIFFRECID_BOOKBOOL,0);

  P_INT_FORMATS := AddDefRec(INTERNAL_FORMATS,0);
  P_INT_BOUNDSHEETS := AddDefRec(INTERNAL_BOUNDSHEETS,0);

  PCOUNTRY := AddDefRec(BIFFRECID_COUNTRY,SizeOf(TRecCOUNTRY));
  COUNTRY.DefaultCountryIndex := $0001;
  COUNTRY.WinIniCountry := $002E;

// According to MS doc's, this record shall only be written by excel, but if
// it's missing, excel asks if "changes shall be saved" when the file containes
// formulas. $00FFFFFF seems to be the highest possible value.
  PRECALCID := AddDefRec(BIFFRECID_RECALCID,SizeOf(TRecRECALCID));
  RECALCID.RecordIdRepeated := BIFFRECID_RECALCID;
  RECALCID.Reserved := 0;
  RECALCID.RecalcEngineId := $00FFFFFF;

  P_INT_NAMES := AddDefRec(INTERNAL_NAMES,0);
  P_INT_MSODRWGRP := AddDefRec(INTERNAL_MSODRWGRP,0);
  P_INT_SST := AddDefRec(INTERNAL_SST,0);

  PEOF := AddDefRec(BIFFRECID_EOF,0);
end;

procedure TRecordStorageGlobals.SetHIDEOBJ(const Value: word);
begin
  PWordArray(@PHIDEOBJ.Data)[0] := Value
end;

procedure TRecordStorageGlobals.SetRECISION(const Value: boolean);
begin
  PWordArray(@PPRECISION.Data)[0] := Word(Value);
end;

procedure TRecordStorageGlobals.SetREFRESHALL(const Value: boolean);
begin
  PWordArray(@PREFRESHALL.Data)[0] := Word(Value);
end;

procedure TRecordStorageGlobals.SetWINDOWPROTECT(const Value: boolean);
begin
  PWordArray(@PWINDOWPROTECT.Data)[0] := Word(Value);
end;

procedure TRecordStorageGlobals.SetWRITEACCESS(const Value: AxUCString);
var
  i: integer;
  WS: AxUCString;
begin
  WS := Copy(Value,1,109 - 5);
  FillChar(PByteArray(@PWRITEACCESS.Data)[0],112,' ');
  System.Move(Pointer(WS)^,PByteArray(@PWRITEACCESS.Data)[3],Length(WS) * 2);
  PByteArray(@PWRITEACCESS.Data)[0] := Length(WS);
  PByteArray(@PWRITEACCESS.Data)[1] := 0;
  PByteArray(@PWRITEACCESS.Data)[2] := 1;

  for i := 1 to 5 do
    PByteArray(@PWRITEACCESS.Data)[106 + i] := Byte(CurrentBuildNumber[i]);
end;

function TRecordStorageGlobals.GetBOF: PRecBOF8;
begin
  Result := PRecBOF8(@PBOF.Data);
end;

constructor TRecordStorageGlobals.Create;
begin
  inherited Create;
end;

procedure TRecordStorageGlobals.UpdateInternal(Id: word);
var
  Header: TBIFFHeader;
begin
  Header.RecID := Id;
  Header.Length := 0;
  case Id of
    INTERNAL_FORMATS:     UpdateRec(P_INT_FORMATS,Header,Nil);
    INTERNAL_BOUNDSHEETS: UpdateRec(P_INT_BOUNDSHEETS,Header,Nil);
    INTERNAL_NAMES:       UpdateRec(P_INT_NAMES,Header,Nil);
    INTERNAL_MSODRWGRP:   UpdateRec(P_INT_MSODRWGRP,Header,Nil);
    INTERNAL_SST:         UpdateRec(P_INT_SST,Header,Nil);
//    else
//      raise XLSRWException.CreateFmt('[int] Unknown int. default record %.4X',[Id]);
  end;
end;

procedure TRecordStorageGlobals.MoveDefault(Rec, NewRec: PRecordData);
begin
  case Rec.RecId of
    INTERNAL_FORMATS:           P_INT_FORMATS := NewRec;
    INTERNAL_BOUNDSHEETS:       P_INT_BOUNDSHEETS := NewRec;
    INTERNAL_NAMES:             P_INT_NAMES := NewRec;
    INTERNAL_MSODRWGRP:         P_INT_MSODRWGRP := NewRec;
    INTERNAL_SST:               P_INT_SST := NewRec;

    BIFFRECID_BOF8:             PBOF := NewRec;
    BIFFRECID_CODEPAGE:         PCODEPAGE := NewRec;
    BIFFRECID_DSF:              PDSF := NewRec;
    BIFFRECID_TABID:            PTABID := NewRec;
    BIFFRECID_FNGROUPCOUNT:     PFNGROUPCOUNT := NewRec;
    BIFFRECID_WRITEACCESS:      PWRITEACCESS := NewRec;
    BIFFRECID_WINDOWPROTECT:    PWINDOWPROTECT := NewRec;
    BIFFRECID_PROTECT:          PPROTECT := NewRec;
    BIFFRECID_PASSWORD:         PPASSWORD := NewRec;
    BIFFRECID_PROT4REV:         PPROT4REV := NewRec;
    BIFFRECID_WINDOW1:          PWINDOW1 := NewRec;
    BIFFRECID_HIDEOBJ:          PHIDEOBJ := NewRec;
    BIFFRECID_1904:             P1904 := NewRec;
    BIFFRECID_BACKUP:           PBACKUP := NewRec;
    BIFFRECID_PRECISION:        PPRECISION := NewRec;
    BIFFRECID_REFRESHALL:       PREFRESHALL := NewRec;
    BIFFRECID_BOOKBOOL:         PBOOKBOOL := NewRec;
    BIFFRECID_COUNTRY:          PCOUNTRY := NewRec;
    BIFFRECID_RECALCID:         PRECALCID := NewRec;
    BIFFRECID_EOF:              PEOF := NewRec;
//    else
//      raise XLSRWException.CreateFmt('[int] Unknown move default record %.4X',[Rec.RecID]);
  end;
end;

function TRecordStorageGlobals.GetDATE1904: boolean;
begin
  Result := PWordArray(@P1904.Data)[0] = 1;
end;

{ TRecordStorageSheet }

procedure TRecordStorageSheet.UpdateDefault(Header: TBIFFHeader; Data: PByteArray);
begin
  case Header.RecID of
    BIFFRECID_CALCMODE:         UpdateRec(PCALCMODE,Header,Data);
    BIFFRECID_CALCCOUNT:        UpdateRec(PCALCCOUNT,Header,Data);
    BIFFRECID_REFMODE:          UpdateRec(PREFMODE,Header,Data);
    BIFFRECID_ITERATION:        UpdateRec(PITERATION,Header,Data);
    BIFFRECID_DELTA:            UpdateRec(PDELTA,Header,Data);
    BIFFRECID_SAVERECALC:       UpdateRec(PSAVERECALC,Header,Data);
    BIFFRECID_PRINTGRIDLINES:   UpdateRec(PPRINTGRIDLINES,Header,Data);
    BIFFRECID_PRINTHEADERS:     UpdateRec(PPRINTHEADERS,Header,Data);
    BIFFRECID_GRIDSET:          UpdateRec(PGRIDSET,Header,Data);
    BIFFRECID_GUTS:             UpdateRec(PGUTS,Header,Data);
    BIFFRECID_DEFAULTROWHEIGHT: UpdateRec(PDEFAULTROWHEIGHT,Header,Data);
    BIFFRECID_WSBOOL:           UpdateRec(PWSBOOL,Header,Data);
    BIFFRECID_HCENTER:          UpdateRec(PHCENTER,Header,Data);
    BIFFRECID_VCENTER:          UpdateRec(PVCENTER,Header,Data);
//    BIFFRECID_PLS:              IPLS              := Count - 1;
    BIFFRECID_SETUP:            UpdateRec(PSETUP,Header,Data);
    BIFFRECID_PROTECT:          UpdateRec(PPROTECT,Header,Data);
    BIFFRECID_DEFCOLWIDTH:      UpdateRec(PDEFCOLWIDTH,Header,Data);
    BIFFRECID_DIMENSIONS:       UpdateRec(PDIMENSIONS,Header,Data);
    BIFFRECID_WINDOW2:          UpdateRec(PWINDOW2,Header,Data);
    BIFFRECID_EOF:              UpdateRec(PEOF,Header,Data);
    else begin
      if (Header.RecID and $FF) = BIFFRECID_BOF then begin
        Header.RecID := BIFFRECID_BOF8;
        UpdateRec(PBOF,Header,Data);
      end;
//      else
//        raise XLSRWException.CreateFmt('[int] Unknown default record %.4X',[Header.RecID]);
    end;
  end;
end;

function TRecordStorageSheet.GetDELTA: double;
begin
  Result := PDouble(@PDELTA.Data)^;
end;

function TRecordStorageSheet.GetCALCMODE: word;
begin
  Result := PWordArray(@PCALCMODE.Data)[0];
end;

procedure TRecordStorageSheet.SetDELTA(const Value: double);
begin
  PDouble(@PDELTA.Data)^ := Value;
end;

procedure TRecordStorageSheet.SetCALCMODE(const Value: word);
begin
  PWordArray(@PCALCMODE.Data)[0] := Value;
end;

procedure TRecordStorageSheet.SetDefaultData;
begin
  Clear;

  PBOF := AddDefRec($0809,SizeOf(TRecBOF8));
  BOF.VersionNumber := $0600;
  BOF.BuildIdentifier := $18AF;
  BOF.BuildYear := $07CD;
  BOF.FileHistoryFlags := $000080C9;
  BOF.LowBIFF := $00000106;
  BOF.SubstreamType := $0010;

  PCALCMODE := AddDefRecWord(BIFFRECID_CALCMODE,$0001);
  PCALCCOUNT := AddDefRecWord(BIFFRECID_CALCCOUNT,$0064);
  PREFMODE := AddDefRecWord(BIFFRECID_REFMODE,$0001);
  PITERATION := AddDefRecWord(BIFFRECID_ITERATION,$0000);

  PDELTA := AddDefRec(BIFFRECID_DELTA,SizeOf(double));
  DELTA := 0.001;

  PSAVERECALC := AddDefRecWord(BIFFRECID_SAVERECALC,$0001);
  PPRINTGRIDLINES := AddDefRecWord(BIFFRECID_PRINTGRIDLINES,$0000);
  PPRINTHEADERS := AddDefRecWord(BIFFRECID_PRINTHEADERS,$0000);
  PGRIDSET := AddDefRecWord(BIFFRECID_GRIDSET,$0001);

  PGUTS := AddDefRec(BIFFRECID_GUTS,SizeOf(TRecGUTS));
  GUTS.SizeRow := 0;
  GUTS.SizeCol := 0;
  GUTS.LevelRow := 0;
  GUTS.LevelCol := 0;

  PDEFAULTROWHEIGHT := AddDefRec(BIFFRECID_DEFAULTROWHEIGHT,SizeOf(TRecDEFAULTROWHEIGHT));
  DEFAULTROWHEIGHT.Options := $0000;
  DEFAULTROWHEIGHT.Height := $00FF;

  PWSBOOL := AddDefRecWord(BIFFRECID_WSBOOL,$04C1);

  P_INT_PAGEBREAKES := AddDefRec(INTERNAL_PAGEBREAKES,0);
  P_INT_HEADER := AddDefRec(INTERNAL_HEADER,0);
  P_INT_MARGINS := AddDefRec(INTERNAL_MARGINS,0);

  PHCENTER := AddDefRecWord(BIFFRECID_HCENTER,$0000);
  PVCENTER := AddDefRecWord(BIFFRECID_VCENTER,$0000);

  PSETUP := AddDefRec(BIFFRECID_SETUP,SizeOf(TRecSETUP));
  SETUP.PaperSize := 0;
  SETUP.Scale     := 100; // $FF;
  SETUP.PageStart := 1;
  SETUP.FitWidth  := 1;
  SETUP.FitHeight := 1;
  SETUP.Options   := $06B6;
  SETUP.Resolution:= 600;
  SETUP.VertResolution := 600;
  SETUP.HeaderMargin := 0.5;
  SETUP.FooterMargin := 0.5;
  SETUP.Copies    := 0;

  PPROTECT := AddDefRecWord(BIFFRECID_PROTECT,$0000);

  PDEFCOLWIDTH := AddDefRecWord(BIFFRECID_DEFCOLWIDTH,$0008);

  PDIMENSIONS := AddDefRec(BIFFRECID_DIMENSIONS,SizeOf(TRecDIMENSIONS8));
  DIMENSIONS.FirstCol := 0;
  DIMENSIONS.LastCol := 0;
  DIMENSIONS.FirstRow := 0;
  DIMENSIONS.LastRow := 0;

  P_INT_COLINFO := AddDefRec(INTERNAL_COLINFO,0);
  P_INT_CELLDATA := AddDefRec(INTERNAL_CELLDATA,0);

  PWINDOW2 := AddDefRec(BIFFRECID_WINDOW2,SizeOf(TRecWINDOW2_8));
  WINDOW2.Options := $06B6 and not ($0200);
  WINDOW2.TopRow := 0;
  WINDOW2.LeftCol := 0;
  WINDOW2.HeaderColorIndex := $00000040;
  WINDOW2.ZoomPreview := 0;
  WINDOW2.Zoom := 0;
  WINDOW2.Reserved := 0;

  P_INT_SUFFIXDATA  := AddDefRec(INTERNAL_SUFFIXDATA,0);
  
  PEOF := AddDefRec(BIFFRECID_EOF,0);
end;

function TRecordStorageSheet.GetCALCCOUNT: word;
begin
  Result := PWordArray(@PCALCCOUNT.Data)[0];
end;

procedure TRecordStorageSheet.SetCALCCOUNT(const Value: word);
begin
  PWordArray(@PCALCCOUNT.Data)[0] := Value;
end;

function TRecordStorageSheet.GetITERATION: boolean;
begin
  Result := PWordArray(@PITERATION.Data)[0] = 1;
end;

function TRecordStorageSheet.GetREFMODE: word;
begin
  Result := PWordArray(@PREFMODE.Data)[0];
end;

function TRecordStorageSheet.GetSAVERECALC: boolean;
begin
  Result := PWordArray(@PSAVERECALC.Data)[0] = 1;
end;

procedure TRecordStorageSheet.SetITERATION(const Value: boolean);
begin
  PWordArray(@PITERATION.Data)[0] := Word(Value);
end;

procedure TRecordStorageSheet.SetREFMODE(const Value: word);
begin
  PWordArray(@PREFMODE.Data)[0] := Value;
end;

procedure TRecordStorageSheet.SetSAVERECALC(const Value: boolean);
begin
  PWordArray(@PSAVERECALC.Data)[0] := Word(Value);
end;

function TRecordStorageSheet.GetPRINTGRIDLINES: boolean;
begin
  Result := PWordArray(@PPRINTGRIDLINES.Data)[0] = 1;
end;

function TRecordStorageSheet.GetPRINTHEADERS: boolean;
begin
  Result := PWordArray(@PPRINTHEADERS.Data)[0] = 1;
end;

procedure TRecordStorageSheet.SetPRINTGRIDLINES(const Value: boolean);
begin
  PWordArray(@PPRINTGRIDLINES.Data)[0] := Word(Value);
end;

procedure TRecordStorageSheet.SetPRINTHEADERS(const Value: boolean);
begin
  PWordArray(@PPRINTHEADERS.Data)[0] := Word(Value);
end;

function TRecordStorageSheet.GetGRIDSET: word;
begin
  Result := PWordArray(@PGRIDSET.Data)[0];
end;

procedure TRecordStorageSheet.SetGRIDSET(const Value: word);
begin
  PWordArray(@PGRIDSET.Data)[0] := Value;
end;

function TRecordStorageSheet.GetGUTS: PRecGUTS;
begin
  Result := PRecGUTS(@PGUTS.Data);
end;

function TRecordStorageSheet.GetDEFAULTROWHEIGHT: PRecDEFAULTROWHEIGHT;
begin
  Result := PRecDEFAULTROWHEIGHT(@PDEFAULTROWHEIGHT.Data);
end;

function TRecordStorageSheet.GetWSBOOL: word;
begin
  Result := PWordArray(@PWSBOOL.Data)[0];
end;

procedure TRecordStorageSheet.SetWSBOOL(const Value: word);
begin
  PWordArray(@PWSBOOL.Data)[0] := Value;
end;

function TRecordStorageSheet.GetHCENTER: boolean;
begin
  Result := PWordArray(@PHCENTER.Data)[0] = 1;
end;

procedure TRecordStorageSheet.SetHCENTER(const Value: boolean);
begin
  PWordArray(@PHCENTER.Data)[0] := Word(Value);
end;

function TRecordStorageSheet.GetVCENTER: boolean;
begin
  Result := PWordArray(@PVCENTER.Data)[0] = 1;
end;

procedure TRecordStorageSheet.SetVCENTER(const Value: boolean);
begin
  PWordArray(@PVCENTER.Data)[0] := Word(Value);
end;

procedure TRecordStorageSheet.Clear;
begin
  inherited Clear;
end;

function TRecordStorageSheet.GetSETUP: PRecSETUP;
begin
  Result := PRecSETUP(@PSETUP.Data);
end;

function TRecordStorageSheet.PostCheck: integer;
begin
  Result := -1;
end;

function TRecordStorageSheet.GetDEFCOLWIDTH: word;
begin
  Result := PWordArray(@PDEFCOLWIDTH.Data)[0];
end;

procedure TRecordStorageSheet.SetDEFCOLWIDTH(const Value: word);
begin
  PWordArray(@PDEFCOLWIDTH.Data)[0] := Value;
end;

function TRecordStorageSheet.GetDIMENSIONS: PRecDIMENSIONS8;
begin
  Result := PRecDIMENSIONS8(@PDIMENSIONS.Data);
end;

function TRecordStorageSheet.GetWINDOW2: PRecWINDOW2_8;
begin
  Result := PRecWINDOW2_8(@PWINDOW2.Data);
end;

constructor TRecordStorageSheet.Create;
begin
  inherited Create;
end;

procedure TRecordStorageSheet.UpdateInternal(Id: word);
var
  Header: TBIFFHeader;
begin
  Header.RecID := Id;
  Header.Length := 0;
  case Id of
    INTERNAL_PAGEBREAKES: UpdateRec(P_INT_PAGEBREAKES,Header,Nil);
    INTERNAL_HEADER:      UpdateRec(P_INT_HEADER,Header,Nil);
    INTERNAL_MARGINS:     UpdateRec(P_INT_MARGINS,Header,Nil);
    INTERNAL_COLINFO:     UpdateRec(P_INT_COLINFO,Header,Nil);
    INTERNAL_CELLDATA:    UpdateRec(P_INT_CELLDATA,Header,Nil);
    INTERNAL_SUFFIXDATA:  UpdateRec(P_INT_SUFFIXDATA,Header,Nil);
//    else
//      raise XLSRWException.CreateFmt('[int] Unknown int. default record %.4X',[Id]);
  end;
end;

function TRecordStorageSheet.GetBOF: PRecBOF8;
begin
  Result := PRecBOF8(@PBOF.Data);
end;

procedure TRecordStorageSheet.MoveDefault(Rec, NewRec: PRecordData);
begin
  case Rec.RecId of
    INTERNAL_PAGEBREAKES:       P_INT_PAGEBREAKES := NewRec;
    INTERNAL_HEADER:            P_INT_HEADER := NewRec;
    INTERNAL_MARGINS:           P_INT_MARGINS := NewRec;
    INTERNAL_COLINFO:           P_INT_COLINFO := NewRec;
    INTERNAL_CELLDATA:          P_INT_CELLDATA := NewRec;
    INTERNAL_SUFFIXDATA:        P_INT_SUFFIXDATA := NewRec;

    BIFFRECID_BOF8:             PBOF := NewRec;
    BIFFRECID_CALCMODE:         PCALCMODE := NewRec;
    BIFFRECID_CALCCOUNT:        PCALCCOUNT := NewRec;
    BIFFRECID_REFMODE:          PREFMODE := NewRec;
    BIFFRECID_ITERATION:        PITERATION := NewRec;
    BIFFRECID_DELTA:            PDELTA := NewRec;
    BIFFRECID_SAVERECALC:       PSAVERECALC := NewRec;
    BIFFRECID_PRINTGRIDLINES:   PPRINTGRIDLINES := NewRec;
    BIFFRECID_PRINTHEADERS:     PPRINTHEADERS := NewRec;
    BIFFRECID_GRIDSET:          PGRIDSET := NewRec;
    BIFFRECID_GUTS:             PGUTS := NewRec;
    BIFFRECID_DEFAULTROWHEIGHT: PDEFAULTROWHEIGHT := NewRec;
    BIFFRECID_WSBOOL:           PWSBOOL := NewRec;
    BIFFRECID_HCENTER:          PHCENTER := NewRec;
    BIFFRECID_VCENTER:          PVCENTER := NewRec;
    BIFFRECID_SETUP:            PSETUP := NewRec;
    BIFFRECID_PROTECT:          PPROTECT := NewRec;
    BIFFRECID_DEFCOLWIDTH:      PDEFCOLWIDTH := NewRec;
    BIFFRECID_DIMENSIONS:       PDIMENSIONS := NewRec;
    BIFFRECID_WINDOW2:          PWINDOW2 := NewRec;
    BIFFRECID_EOF:              PEOF := NewRec;
//    else
//      raise XLSRWException.CreateFmt('[int] Unknown move default record %.4X',[Rec.RecID]);
  end;
end;

procedure TRecordStorageSheet.SetPROTECT(const Value: word);
begin
  PWordArray(@PPROTECT.Data)[0] := Value;
end;

function TRecordStorageSheet.GetPROTECT: word;
begin
  Result := PWordArray(@PPROTECT.Data)[0];
end;

{ TRecordStorageDefault }

function TRecordStorageDefault.AddDefRec(Id, Length: integer): PRecordData;
var
  P: PRecordData;
begin
  GetMem(P,Length + SizeOf(TRecordData));
  P.RecId := Id;
  P.Length := Length;
  FDefault.Add(P);
  Result := P;
end;

function TRecordStorageDefault.AddDefRecWord(Id: integer; Data: word): PRecordData;
var
  P: PRecordData;
begin
  GetMem(P,SizeOf(TRecordData) + 2);
  P.RecId := Id;
  P.Length := 2;
  System.Move(Data,PByteArray(@P.Data)[0],2);
  FDefault.Add(P);
  Result := P;
end;

procedure TRecordStorageDefault.Clear;
begin
  inherited Clear;
  FDefault.Clear;
end;

constructor TRecordStorageDefault.Create;
begin
  inherited Create;
  FDefault := TBaseRecordStorage.Create;
end;

destructor TRecordStorageDefault.Destroy;
begin
  inherited;
  FDefault.Free;
end;

procedure TRecordStorageDefault.MoveAllDefault;
var
  i: integer;
  Rec: PRecordData;
begin
  inherited Clear;
  for i := 0 to FDefault.Count - 1 do begin
    GetMem(Rec,SizeOf(TRecordData) + FDefault[i].Length);
    System.Move(FDefault[i]^,Rec^,SizeOf(TRecordData) + FDefault[i].Length);
    Add(Rec);
  end;
end;

procedure TRecordStorageDefault.UpdateRec(var Rec: PRecordData; Header: TBIFFHeader; Data: PByteArray);
var
  i,j: integer;
  NewRec: PRecordData;

function FindRec: integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].RecId = Header.RecID then
      Exit;
  end;
  Result := -1;
end;

function FindDefRec: integer;
begin
  for Result := 0 to FDefault.Count - 1 do begin
    if (FDefault[Result] <> Nil) and (FDefault[Result].RecId = Header.RecID) then
      Exit;
  end;
  Result := -1;
end;

begin
  j := FindRec;
  if j >= 0 then begin
    if Items[j].Length <> Header.Length then
      raise XLSRWException.CreateFmt('Size missmatch on default record %.4X',[Header.RecID]);
    if Data <> Nil then
      System.Move(Data^,Items[j].Data,Header.Length);
  end
  else begin
    NewRec := Nil;
    j := FindDefRec;
    if j < 0 then
      raise XLSRWException.CreateFmt('Excpected default record %.4X not found.',[Header.RecId]);
    for i := 0 to j do begin
      if FDefault[i] <> Nil then begin
        if i = j then begin
          GetMem(NewRec,SizeOf(TRecordData) + Header.Length);
          NewRec.RecId := Header.RecID;
          NewRec.Length := Header.Length;
          System.Move(Data^,NewRec.Data,Header.Length);
        end
        else begin
          GetMem(NewRec,SizeOf(TRecordData) + FDefault[i].Length);
          System.Move(FDefault[i]^,NewRec^,SizeOf(TRecordData) + FDefault[i].Length);
        end;
        Add(NewRec);
        MoveDefault(NewRec,NewRec);
        FreeMem(FDefault[i]);
        FDefault[i] := Nil;
      end;
    end;
    System.Move(Data^,NewRec.Data,Header.Length);
  end;
end;

{ TRecordStorageEx2007Unknown }

procedure TRecordStorageEx2007Unknown.WriteAllRecs(Stream: TXLSStream);
const
  RecOrder: array[0..9] of word = ($0863,$087C,$087D,$0892,$088E,$089A,$08A3,$0896,$089B,$088C);
var
  i: integer;

procedure WRecs(Id: word);
var
  i: integer;
begin
  for I := 0 to Count - 1 do begin
    if Items[i].RecId = Id then
      WriteRec(i,Stream);
  end;
end;

begin
  inherited WriteAllRecs(Stream);
  Exit;

  for i := 0 to High(RecOrder) do
    WRecs(RecOrder[i]);
end;

end.
