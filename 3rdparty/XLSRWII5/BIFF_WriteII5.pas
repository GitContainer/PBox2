unit BIFF_WriteII5;

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

uses Classes, SysUtils, Contnrs,
{$ifdef BABOON}
{$else}
     Windows,
{$endif}
     BIFF_RecsII5, BIFF_SheetData5, BIFF_Stream5,
     BIFF_Utils5, BIFF5, BIFF_RecordStorage5, BIFF_EncodeFormulaII5,
     BIFF_ExcelFuncII5,
     Xc12Utils5, Xc12Manager5, Xc12DataStylesheet5, Xc12DataWorkbook5,
     Xc12DataWorksheet5, Xc12DataSST5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSFormulaTypes5, XLSDecodeFmla5;

type TBoundsheetType = (btSheet,btChart,btVBModule,btExcel4Macro);

type TRKData = record
     Cache: array[0..255] of TMulRK;
     CachePtr: integer;
     FirstCol,LastCol: integer;
     end;

type TMulblankData = record
     Cache: array[0..255] of word;
     CachePtr: integer;
     FirstCol,LastCol: integer;
     end;

type TBoundsheetData = class(TObject)
private
     FBoundsheetType: TBoundsheetType;
     FIndex: integer;
     FFilePos: integer;
     FName: AxUCString;
     FHidden: THiddenState;
public
     procedure WritePos(Stream: TXLSStream);
     property BoundsheetType: TBoundsheetType read FBoundsheetType write FBoundsheetType;
     property Index: integer read FIndex write FIndex;
     property FilePos: integer read FFilePos write FFilePos;
     property Name: AxUCString read FName write FName;
     property Hidden: THiddenState read FHidden write FHidden;
     end;

type TBoundsheetList = class(TObjectList)
private
     function GetCharts(Index: integer): TBoundsheetData;
     function GetItems(Index: integer): TBoundsheetData;
     function GetSheets(Index: integer): TBoundsheetData;
public
     function  Find(AName: AxUCString): integer;

     procedure AddSheet(Index: integer; Name: AxUCString; Hidden: THiddenState; WorksheetType: TWorksheetType);
     procedure AddChart(Index: integer; Name: AxUCString);

     property Items[Index: integer]: TBoundsheetData read GetItems; default;
     property Sheets[Index: integer]: TBoundsheetData read GetSheets;
     property Charts[Index: integer]: TBoundsheetData read GetCharts;
     end;

type TXLSWriteII = class(TObject)
protected
     FManager: TXc12Manager;
     FXLS: TBIFF5;
     FVersion: TExcelVersion;
     PBuf: PByteArray;
     FCurrSheet: integer;
     FCurrXc12Sheet: TXc12DataWorksheet;
     FXLSStream: TXLSStream;
     FBoundsheetList: TBoundsheetList;

     FRKData: TRKData;
     FMulblankData: TMulblankData;
     FExternalNamesWritten: boolean;

     procedure WriteRecId(RecId: word);
     procedure WriteWord(RecId,Value: word);
     procedure WriteBoolean(RecId: word; Value: boolean);
     procedure WriteBuf(RecId,Size: word); overload;
     procedure WriteBuf(RecId,Size: word; P: Pointer); overload;
     procedure WritePointer(RecId: word; P: Pointer; Size: word);

     // File prefix
     procedure WREC_BOF(SubStreamType: TSubStreamType);
     procedure WREC_WRITEACCESS;
     procedure WREC_CODEPAGE;
     procedure WREC_DSF;
     procedure WREC_EXCEL9FILE;
     procedure WREC_OBPROJ;
     procedure WREC_WINDOWPROTECT;
     procedure WREC_PASSWORD;
     procedure WREC_WINDOW1;
     procedure WREC_BACKUP;
     procedure WREC_HIDEOBJ;
     procedure WREC_1904;
     procedure WREC_PRECISION;
     procedure WREC_REFRESHALL;
     procedure WREC_BOOKBOOL;

     procedure WREC_FONT;
     procedure WREC_FORMAT;
     procedure WREC_XF;
     procedure WREC_STYLE;
     procedure WREC_PALETTE;
     procedure WREC_USESELFS;
     procedure WREC_BOUNDSHEET(Index: integer);
     procedure WREC_COUNTRY;
     procedure WREC_SST;
     procedure WREC_SUPBOOK;
     procedure WREC_EXTERNSHEET;
     procedure WREC_NAME;
     procedure WREC_EOF;

     // Sheet prefix
     procedure WREC_CALCMODE;
     procedure WREC_CALCCOUNT;
     procedure WREC_REFMODE;
     procedure WREC_ITERATION;
     procedure WREC_DELTA;
     procedure WREC_SAVERECALC;
     procedure WREC_PRINTHEADERS;
     procedure WREC_PRINTGRIDLINES;
     procedure WREC_GUTS;
     procedure WREC_GRIDSET;
     procedure WREC_DEFAULTROWHEIGHT;
     procedure WREC_WSBOOL;
     procedure WREC_HORIZONTALPAGEBREAKS;
     procedure WREC_VERTICALPAGEBREAKS;
     procedure WREC_HEADER;
     procedure WREC_FOOTER;
     procedure WREC_HCENTER;
     procedure WREC_VCENTER;
     procedure WREC_MARGINS;
     procedure WREC_DEFCOLWIDTH;
     procedure WREC_SETUP;
     procedure WREC_COLINFO;
     procedure WREC_DIMENSIONS;
     procedure WREC_ROW;

     // Sheet suffix
     procedure WREC_MSODRAWING;
     procedure WREC_NOTE;
     procedure WREC_WINDOW2;
     procedure WREC_SCL;
     procedure WREC_PANE;
     procedure WREC_SELECTION;
     procedure WREC_FEATHEADR;

     procedure WREC_MERGECELLS;
     procedure WREC_CONDFMT;
     procedure WREC_HLINK;
     procedure WREC_DVAL;

     procedure Write50FormatBlock;

     function  CheckDeafultXF(const AXFIndex: integer): integer;
     function  RkEncode(const Value: double; var RK: longword): boolean;
     procedure RkFlushCache(const ACurrRow: integer);
     function  RkAdd(const RK: longword; XF,Col: integer): boolean;
     function  MulblankAdd(const XF,Col: integer): boolean;
     procedure MulblankFlushCache(const ACurrRow: integer);
     procedure WriteSTRING(const AStr: AxUCString);
     procedure WriteCells;
public
     constructor Create(AXLS: TBIFF5);
     destructor Destroy; override;

     procedure WriteToStream(Stream: TStream);
     end;

implementation

type TWordArray = array[0..65535] of word;
     PWordArray = ^TWordArray;

const DEFAULT_FONT8_COUNT = 7;
const DEFAULT_XF8_COUNT   = 22;

type THelperWriteXF = class(TObject)
protected
     FXF: PRecXF8;
     FManager: TXc12Manager;
     FStream: TXLSStream;

     procedure SetIndent(const Value: byte);
     procedure SetRotation(const Value: smallint);
     procedure SetFFormatOptions(const Value: TXc12AlignmentOptions);
     procedure SetHorizAlignment(const Value: TXc12HorizAlignment);
     procedure SetProtection(const Value: TXc12CellProtections);
     procedure SetVertAlignment(const Value: TXc12VertAlignment);
     procedure SetMerged(const Value: boolean);
     procedure SetBorderBottomColor(const Value: TXc12IndexColor);
     procedure SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderDiagColor(const Value: TXc12IndexColor);
     procedure SetBorderDiagLines(const Value: TDiagLines);
     procedure SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderLeftColor(const Value: TXc12IndexColor);
     procedure SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderRightColor(const Value: TXc12IndexColor);
     procedure SetBorderRightStyle(const Value: TXc12CellBorderStyle);
     procedure SetBorderTopColor(const Value: TXc12IndexColor);
     procedure SetBorderTopStyle(const Value: TXc12CellBorderStyle);
     procedure SetFillPatternBackColor(const Value: TXc12IndexColor);
     procedure SetFillPatternForeColor(const Value: TXc12IndexColor);
     procedure SetFillPatternPattern(const Value: TXc12FillPattern);
public
     constructor Create(AManager: TXc12Manager; AStream: TXLSStream; ABuf: PByteArray);

     procedure Write;
     end;

type PExtSSTRec = ^TExtSSTRec;
     TExtSSTRec = record
     StreamPos: longword;
     RecPos: word;
     end;

type THelperWriteExtSST = class(TList)
private
     FBucketSize: integer;

     procedure SetStringCount(const Value: integer);
public
     procedure Clear; override;
     procedure Add(StreamPos: longword; RecPos: word);
     procedure Write(Stream: TXLSStream);

     property BucketSize: integer read FBucketSize;
     property StringCount: integer write SetStringCount;
     end;


type THelperWriteSST = class(TObject)
protected
     FBuf       : PByteArray;
     FManager   : TXc12Manager;
     FStream    : TXLSStream;

     FCurrIndex : integer;
     FTotalCount: integer;
     FRecSize   : integer;
     FRecPos    : integer;

     FExtSST    : THelperWriteExtSST;

     procedure WriteCONTINUE;
     procedure WriteString(P: PByteArray; Len: word; IsUnicode: boolean);
     procedure WriteRichData(P: PByteArray; Count: integer);
     procedure WriteData(P: PByteArray; Len: word);
     procedure WriteSSTString(AStr: PXLSString);
public
     constructor Create(AManager: TXc12Manager; AStream: TXLSStream; ABuf: PByteArray);
     destructor Destroy; override;

     procedure Write;
     end;

function Xc12ColorToIndex(Color: TXc12Color): TXc12IndexColor;
var
  RGB: longword;
begin
  if (Color.ColorType = exctIndexed) and (Color.Tint = 0) then
    Result := Color.Indexed
  else begin
    RGB := Xc12ColorToRGB(Color);
    Result := RGBToClosestIndexColor(RevRGB(RGB));
  end;
end;

{ TXLSWriteII }

function TXLSWriteII.CheckDeafultXF(const AXFIndex: integer): integer;
begin
  if AXFIndex = 0 then
    Result := XLS_STYLE_DEFAULT_XF_97
  else
    Result := AXFIndex - XLS_STYLE_DEFAULT_XF_COUNT;
end;

constructor TXLSWriteII.Create(AXLS: TBIFF5);
begin
  FXLS := AXLS;
  FManager := FXLS.Manager;
  FVersion := xvExcel97;
  GetMem(PBuf,FXLS.MaxBuffsize);
  FXLSStream := TXLSStream.Create(FXLS.VBA);
  FBoundsheetList := TBoundsheetList.Create;
end;

destructor TXLSWriteII.Destroy;
begin
  FreeMem(PBuf);
  FXLSStream.Free;
  FBoundsheetList.Free;
end;

function TXLSWriteII.MulblankAdd(const XF, Col: integer): boolean;
begin
  if FMulblankData.CachePtr > High(FMulblankData.Cache) then
    raise XLSRWException.Create('Mulblank Cache overflow');
  if FMulblankData.FirstCol < 0 then
    FMulblankData.FirstCol := Col;
  if (FMulblankData.LastCol >= 0) and (Col <> (FMulblankData.LastCol + 1)) then
    Result := False
  else begin
    FMulblankData.LastCol := Col;
    FMulblankData.Cache[FMulblankData.CachePtr] := XF;
    Inc(FMulblankData.CachePtr);
    Result := True;
  end;
end;

procedure TXLSWriteII.MulblankFlushCache(const ACurrRow: integer);
var
  RecBlank: TRecBlank;
begin
  if FMulblankData.CachePtr > 0 then begin
    if FMulblankData.CachePtr > 1 then begin
      FXLSStream.WriteHeader(BIFFRECID_MULBLANK,6 + FMulblankData.CachePtr * SizeOf(Word));
      FXLSStream.WWord(ACurrRow);
      FXLSStream.WWord(FMulblankData.FirstCol);
      FXLSStream.Write(FMulblankData.Cache,FMulblankData.CachePtr * SizeOf(Word));
      FXLSStream.WWord(FMulblankData.LastCol);
    end
    else begin
      RecBlank.Row := ACurrRow;
      RecBlank.Col := FMulblankData.FirstCol;
      RecBlank.FormatIndex := FMulblankData.Cache[0];
      WriteBuf(BIFFRECID_BLANK,SizeOf(TRecBLANK),@RecBlank);
    end;
    FMulblankData.CachePtr := 0;
    FMulblankData.FirstCol := -1;
    FMulblankData.LastCol := -1;
  end;
end;

function TXLSWriteII.RkAdd(const RK: longword; XF, Col: integer): boolean;
begin
  if FRKData.CachePtr > High(FRKData.Cache) then
    raise XLSRWException.Create('RK Cache overflow');
  if FRKData.FirstCol < 0 then
    FRKData.FirstCol := Col;
  if (FRKData.LastCol >= 0) and (Col <> (FRKData.LastCol + 1)) then
    Result := False
  else begin
    FRKData.LastCol := Col;
    FRKData.Cache[FRKData.CachePtr].XF := XF;
    FRKData.Cache[FRKData.CachePtr].RK := RK;
    Inc(FRKData.CachePtr);
    Result := True;
  end;
end;

function TXLSWriteII.RkEncode(const Value: double; var RK: longword): boolean;
var
  D: double;
  pL1, pL2: ^longword;
  Mask: longword;
  i: integer;
begin
  Result := True;
  for i := 0 to 1 do begin
    D := Value * (1 + 99 * i);
    pL1 := @d;
    pL2 := pL1;
    Inc(pL2);
    if (pL1^ = 0) and ((pL2^ and 3) = 0) then begin
      RK := pL2^ + Longword(i);
      Exit;
    end;
    Mask := $1FFFFFFF;
    if (Int(D) = D) and (D <= Mask) and (D >= -Mask - 1) then begin
      RK := Round(D) shl 2 + i + 2;
      Exit;
    end;
  end;
  Result := False;
end;

procedure TXLSWriteII.RkFlushCache(const ACurrRow: integer);
var
  RecRK: TRecRK;
begin
  if FRKData.CachePtr > 0 then begin
    if FRKData.CachePtr = 1 then begin
      RecRK.Row := ACurrRow;
      RecRK.Col := FRKData.FirstCol;
      RecRK.FormatIndex := FRKData.Cache[0].XF;
      RecRK.Value := FRKData.Cache[0].RK;
      WriteBuf(BIFFRECID_RK7,SizeOf(TRecRK),@RecRK);
    end
    else begin
      FXLSStream.WriteHeader(BIFFRECID_MULRK,FRKData.CachePtr * SizeOf(TMulRk) + 6);
      FXLSStream.WWord(ACurrRow);
      FXLSStream.WWord(FRKData.FirstCol);
      FXLSStream.Write(FRKData.Cache[0],FRKData.CachePtr * SizeOf(TMulRk));
      FXLSStream.WWord(FRKData.LastCol);
    end;
    FRKData.CachePtr := 0;
    FRKData.FirstCol := -1;
    FRKData.LastCol := -1;
  end;
end;

procedure TXLSWriteII.WriteRecId(RecId: word);
var
  Header: TBIFFHeader;
begin
  Header.RecId := RecID;
  Header.Length := 0;
  FXLSStream.Write(Header,SizeOf(TBIFFHeader));
end;

procedure TXLSWriteII.WriteSTRING(const AStr: AxUCString);
var
  L: integer;
  S8: XLS8String;
begin
  L := Length(AStr);
  if FVersion >= xvExcel97 then begin
    if L > 0 then begin
      FXLSStream.WriteHeader(BIFFRECID_STRING,Length(AStr) * 2 + 2 + 1);
      FXLSStream.Write(L,2);
      FXLSStream.WByte(1);
      FXLSStream.Write(Pointer(AStr)^,Length(AStr) * 2);
    end;
  end
  else begin
    S8 := XLS8String(AStr);
    FXLSStream.WriteHeader(BIFFRECID_STRING,L + 2);
    FXLSStream.Write(L,2);
    FXLSStream.Write(Pointer(S8)^,L);
  end;
end;

procedure TXLSWriteII.WriteWord(RecId,Value: word);
var
  Header: TBIFFHeader;
begin
  Header.RecId := RecID;
  Header.Length := SizeOf(word);
  FXLSStream.Write(Header,SizeOf(TBIFFHeader));
  FXLSStream.Write(Value,SizeOf(word));
end;

procedure TXLSWriteII.Write50FormatBlock;
const
FONTData: array[0..3,0..19] of byte = (
($C8,$00,$00,$00,$FF,$7F,$90,$01,$00,$00,$00,$00,$00,$00,$05,$41,$72,$69,$61,$6C),
($C8,$00,$00,$00,$FF,$7F,$90,$01,$00,$00,$00,$00,$00,$00,$05,$41,$72,$69,$61,$6C),
($C8,$00,$00,$00,$FF,$7F,$90,$01,$00,$00,$00,$00,$00,$00,$05,$41,$72,$69,$61,$6C),
($C8,$00,$00,$00,$FF,$7F,$90,$01,$00,$00,$00,$00,$00,$00,$05,$41,$72,$69,$61,$6C));

FORMATData1: array[0..27] of byte = (
$05,$00,$19,$23,$2C,$23,$23,$30,$5C,$20,$22,$6B,$72,$22,$3B,$5C,$2D,$23,$2C,$23,$23,$30,$5C,$20,$22,$6B,$72,$22);
FORMATData2: array[0..32] of byte = (
$06,$00,$1E,$23,$2C,$23,$23,$30,$5C,$20,$22,$6B,$72,$22,$3B,$5B,$52,$65,$64,$5D,$5C,$2D,$23,$2C,$23,$23,$30,$5C,$20,$22,$6B,$72,$22);
FORMATData3: array[0..33] of byte = (
$07,$00,$1F,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$22,$6B,$72,$22,$3B,$5C,$2D,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$22,$6B,$72,$22);
FORMATData4: array[0..38] of byte = (
$08,$00,$24,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$22,$6B,$72,$22,$3B,$5B,$52,$65,$64,$5D,$5C,$2D,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$22,$6B,$72,$22);
FORMATData5: array[0..59] of byte = (
$2A,$00,$39,$5F,$2D,$2A,$20,$23,$2C,$23,$23,$30,$5C,$20,$22,$6B,$72,$22,$5F,$2D,$3B,$5C,$2D,$2A,$20,$23,$2C,$23,$23,$30,$5C,$20,$22,$6B,$72,$22,$5F,$2D,$3B,$5F,$2D,$2A,$20,$22,$2D,$22,$5C,$20,$22,$6B,$72,$22,$5F,$2D,$3B,$5F,$2D,$40,$5F,$2D);
FORMATData6: array[0..59] of byte = (
$29,$00,$39,$5F,$2D,$2A,$20,$23,$2C,$23,$23,$30,$5C,$20,$5F,$6B,$5F,$72,$5F,$2D,$3B,$5C,$2D,$2A,$20,$23,$2C,$23,$23,$30,$5C,$20,$5F,$6B,$5F,$72,$5F,$2D,$3B,$5F,$2D,$2A,$20,$22,$2D,$22,$5C,$20,$5F,$6B,$5F,$72,$5F,$2D,$3B,$5F,$2D,$40,$5F,$2D);
FORMATData7: array[0..67] of byte = (
$2C,$00,$41,$5F,$2D,$2A,$20,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$22,$6B,$72,$22,$5F,$2D,$3B,$5C,$2D,$2A,$20,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$22,$6B,$72,$22,$5F,$2D,$3B,$5F,$2D,$2A,$20,$22,$2D,$22,$3F,$3F,$5C,$20,$22,$6B,$72,$22,$5F,$2D,$3B,$5F,$2D,$40,$5F,$2D);
FORMATData8: array[0..67] of byte = (
$2B,$00,$41,$5F,$2D,$2A,$20,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$5F,$6B,$5F,$72,$5F,$2D,$3B,$5C,$2D,$2A,$20,$23,$2C,$23,$23,$30,$2E,$30,$30,$5C,$20,$5F,$6B,$5F,$72,$5F,$2D,$3B,$5F,$2D,$2A,$20,$22,$2D,$22,$3F,$3F,$5C,$20,$5F,$6B,$5F,$72,$5F,$2D,$3B,$5F,$2D,$40,$5F,$2D);

XFData: array[0..20,0..15] of byte = (
($00,$00,$00,$00,$F5,$FF,$20,$00,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($02,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($02,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$F5,$FF,$20,$F4,$C0,$20,$00,$00,$00,$00,$00,$00),
($00,$00,$00,$00,$01,$00,$20,$00,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$2B,$00,$F5,$FF,$20,$F8,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$29,$00,$F5,$FF,$20,$F8,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$2C,$00,$F5,$FF,$20,$F8,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$2A,$00,$F5,$FF,$20,$F8,$C0,$20,$00,$00,$00,$00,$00,$00),
($01,$00,$09,$00,$F5,$FF,$20,$F8,$C0,$20,$00,$00,$00,$00,$00,$00));

STYLEData: array[0..5,0..3] of byte = (
($10,$80,$03,$FF),
($11,$80,$06,$FF),
($12,$80,$04,$FF),
($13,$80,$07,$FF),
($00,$80,$00,$FF),
($14,$80,$05,$FF));

var
  i: integer;
begin
  for i := 0 to 3 do begin
    FXLSStream.WriteHeader(BIFFRECID_FONT,Length(FONTData[i]));
    FXLSStream.Write(FONTData[i],Length(FONTData[i]));
  end;

  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData1));
  FXLSStream.Write(FORMATData1,Length(FORMATData1));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData2));
  FXLSStream.Write(FORMATData2,Length(FORMATData2));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData3));
  FXLSStream.Write(FORMATData3,Length(FORMATData3));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData4));
  FXLSStream.Write(FORMATData4,Length(FORMATData4));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData5));
  FXLSStream.Write(FORMATData5,Length(FORMATData5));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData6));
  FXLSStream.Write(FORMATData6,Length(FORMATData6));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData7));
  FXLSStream.Write(FORMATData7,Length(FORMATData7));
  FXLSStream.WriteHeader(BIFFRECID_FORMAT,Length(FORMATData8));
  FXLSStream.Write(FORMATData8,Length(FORMATData8));

  for i := 0 to 20 do begin
    FXLSStream.WriteHeader(BIFFRECID_XF,Length(XFData[i]));
    FXLSStream.Write(XFData[i],Length(XFData[i]));
  end;

  for i := 0 to 5 do begin
    FXLSStream.WriteHeader(BIFFRECID_STYLE,Length(STYLEData[i]));
    FXLSStream.Write(STYLEData[i],Length(STYLEData[i]));
  end;
end;

procedure TXLSWriteII.WriteBoolean(RecId: word; Value: boolean);
begin
  if Value then
    WriteWord(RecId,1)
  else
    WriteWord(RecId,0);
end;

procedure TXLSWriteII.WriteBuf(RecId,Size: word);
begin
  FXLSStream.WriteHeader(RecId,Size);
  if Size > 0 then
    FXLSStream.Write(PBuf^,Size);
end;

procedure TXLSWriteII.WriteBuf(RecId,Size: word; P: Pointer);
begin
  FXLSStream.WriteHeader(RecID,Size);
  if Size > 0 then
    FXLSStream.Write(P^,Size);
end;

procedure TXLSWriteII.WriteCells;
var
  Col,Row: integer;
  CurrRkRow: integer;
  iXF: integer;
  L: integer;
  HdrLen: integer;
  S: AxUCString;
  Fmla: AxUCString;
  V: double;
  P: PByteArray;
  Sz: integer;
  RK: longword;
  Ptgs: PXLSPtgs;
  Xc12Sheet: TXc12DataWorksheet;
  RecNUM: TRecNUMBER;
  RecLabelSST: TRecLABELSST;
  RecBOOL: TRecBOOLERR;
  RecFORMULA: TRecFORMULA_;
  RecARRAY: TRecARRAY;
  RecBLANK: TRecBLANK;
  Encoder: TEncodeFormula;
begin
  Xc12Sheet := FManager.Worksheets[FCurrSheet];

  FRKData.CachePtr := 0;
  FRKData.FirstCol := -1;
  FRKData.LastCol := -1;

  FMulblankData.CachePtr := 0;
  FMulblankData.FirstCol := -1;
  FMulblankData.LastCol := -1;

  CurrRkRow := -1;

  Xc12Sheet.Cells.BeginIterate;

  while Xc12Sheet.Cells.IterateNext do begin
    Col := Xc12Sheet.Cells.IterCellCol;
    Row := Xc12Sheet.Cells.IterCellRow;

    iXF := CheckDeafultXF(Xc12Sheet.Cells.IterGetStyleIndex);

    if (Col > MAXCOL_97) or (Row > MAXROW_97) then
      Continue;
    if Row <> CurrRkRow then begin
      RkFlushCache(CurrRkRow);
      MulblankFlushCache(CurrRkRow);
      CurrRkRow := Row;
    end;
    if not (Xc12Sheet.Cells.IterCellType in [xctBlank,xctFloat]) then begin
      RkFlushCache(CurrRkRow);
      MulblankFlushCache(CurrRkRow);
    end;

    case Xc12Sheet.Cells.IterCellType of
      xctBlank         : begin
        if not MulblankAdd(iXF,Col) then begin
          MulblankFlushCache(CurrRkRow);
          RecBLANK.Row := Row;
          RecBLANK.Col := Col;
          RecBLANK.FormatIndex := iXF;
          WriteBuf(BIFFRECID_BLANK,SizeOf(TRecBLANK),@RecBLANK);
        end;
      end;
      xctBoolean       : begin
        RecBool.Row := Row;
        RecBool.Col := Col;
        RecBool.FormatIndex := iXF;
        RecBool.BoolErr := Byte(Xc12Sheet.Cells.IterGetBoolean);
        RecBool.Error := 0;
        WriteBuf(BIFFRECID_BOOLERR,SizeOf(TRecBOOLERR),@RecBool);
      end;
      xctError         : begin
        RecBool.Row := Row;
        RecBool.Col := Col;
        RecBool.FormatIndex := iXF;
        case Xc12Sheet.Cells.IterGetError of
          errUnknown: raise XLSRWException.CreateFmt('Invalid error value in cell %d:%d',[RecBool.Row,RecBool.Col]);
          errNull:  RecBool.BoolErr := $00;
          errDiv0:  RecBool.BoolErr := $07;
          errValue: RecBool.BoolErr := $0F;
          errRef:   RecBool.BoolErr := $17;
          errName:  RecBool.BoolErr := $1D;
          errNum:   RecBool.BoolErr := $24;
          errNA:    RecBool.BoolErr := $2A;
        end;
        RecBool.Error := 1;
        WriteBuf(BIFFRECID_BOOLERR,SizeOf(TRecBOOLERR),@RecBool);
      end;
      xctString        : begin
        RecLabelSST.Row := Row;
        RecLabelSST.Col := Col;
        RecLabelSST.FormatIndex := iXF;
        RecLabelSST.SSTIndex := Xc12Sheet.Cells.IterGetStringIndex;
        WriteBuf(BIFFRECID_LABELSST,SizeOf(TRecLABELSST),@RecLabelSST);
      end;
      xctFloat         : begin
        V := Xc12Sheet.Cells.IterGetFloat;
        if not (RkEncode(V,RK) and RkAdd(RK,iXF,Col)) then begin
          RkFlushCache(CurrRkRow);
          RecNum.Row := Row;
          RecNum.Col := Col;
          RecNum.FormatIndex := iXF;
          RecNum.Value := V;
          WriteBuf(BIFFRECID_NUMBER,SizeOf(TRecNUMBER),@RecNum);
        end;
      end;
      xctFloatFormula,
      xctStringFormula,
      xctBooleanFormula,
      xctErrorFormula  : begin
        Xc12Sheet.Cells.RetriveFormula(Xc12Sheet.Cells.IterCell);

        RecFORMULA.Row := Row;
        RecFORMULA.Col := Col;
        RecFORMULA.FormatIndex := iXF;
        RecFORMULA.Reserved := 0;
        if FManager.Workbook.CalcPr.FullCalcOnLoad then
          RecFORMULA.Options := $02
        else
          RecFORMULA.Options := $00;

        case Xc12Sheet.Cells.IterCellType of
          xctFloatFormula  : begin
            RecFORMULA.Value := Xc12Sheet.Cells.FormulaHelper.AsFloat;
          end;
          xctStringFormula : begin
            S := Xc12Sheet.Cells.FormulaHelper.AsString;
            if S = '' then
              TByte8Array(RecFORMULA.Value)[0] := 3
            else
              TByte8Array(RecFORMULA.Value)[0] := 0;
            TByte8Array(RecFORMULA.Value)[1] := 0;
            TByte8Array(RecFORMULA.Value)[2] := 0;
            TByte8Array(RecFORMULA.Value)[3] := 0;
            TByte8Array(RecFORMULA.Value)[4] := 0;
            TByte8Array(RecFORMULA.Value)[5] := 0;
            TByte8Array(RecFORMULA.Value)[6] := $FF;
            TByte8Array(RecFORMULA.Value)[7] := $FF;
          end;
          xctBooleanFormula: begin
            TByte8Array(RecFORMULA.Value)[0] := 1;
            TByte8Array(RecFORMULA.Value)[1] := 0;
            TByte8Array(RecFORMULA.Value)[2] := Byte(Xc12Sheet.Cells.FormulaHelper.AsBoolean);
            TByte8Array(RecFORMULA.Value)[3] := 0;
            TByte8Array(RecFORMULA.Value)[4] := 0;
            TByte8Array(RecFORMULA.Value)[5] := 0;
            TByte8Array(RecFORMULA.Value)[6] := $FF;
            TByte8Array(RecFORMULA.Value)[7] := $FF;
          end;
          xctErrorFormula  : begin
            TByte8Array(RecFORMULA.Value)[0] := 1;
            TByte8Array(RecFORMULA.Value)[1] := 0;
            TByte8Array(RecFORMULA.Value)[2] := CellErrorToErrorCode(Xc12Sheet.Cells.FormulaHelper.AsError);
            TByte8Array(RecFORMULA.Value)[3] := 0;
            TByte8Array(RecFORMULA.Value)[4] := 0;
            TByte8Array(RecFORMULA.Value)[5] := 0;
            TByte8Array(RecFORMULA.Value)[6] := $FF;
            TByte8Array(RecFORMULA.Value)[7] := $FF;
          end;
        end;
        case Xc12Sheet.Cells.FormulaHelper.FormulaType of
          xcftNormal        : begin
            Ptgs := Xc12Sheet.Cells.FormulaHelper.Ptgs;
            if Ptgs.Id = xptg_EXCEL_97 then begin
              RecFORMULA.ParseLen := Xc12Sheet.Cells.FormulaHelper.PtgsSize - SizeOf(TXLSPtgs);
              FXLSStream.WriteHeader(BIFFRECID_FORMULA,SizeOf(TRecFORMULA_) + RecFORMULA.ParseLen);
              FXLSStream.Write(RecFORMULA,SizeOf(TRecFORMULA_));
              Ptgs := PXLSPtgs(NativeInt(Ptgs) + SizeOf(TXLSPtgs));
              FXLSStream.Write(Ptgs^,RecFORMULA.ParseLen);

              if Xc12Sheet.Cells.IterCellType = xctStringFormula then
                WriteSTRING(S);
            end
            else if Ptgs.Id = xptg_ARRAYCONSTS_97 then begin
              HdrLen := SizeOf(TXLSPtgs) + SizeOf(Word);
              Ptgs := PXLSPtgs(NativeInt(Ptgs) + SizeOf(TXLSPtgs));
              L := PWord(Ptgs)^;
              Ptgs := PXLSPtgs(NativeInt(Ptgs) + SizeOf(Word));

              RecFORMULA.ParseLen := Xc12Sheet.Cells.FormulaHelper.PtgsSize - HdrLen - L;

              FXLSStream.WriteHeader(BIFFRECID_FORMULA,SizeOf(TRecFORMULA_) + Xc12Sheet.Cells.FormulaHelper.PtgsSize - HdrLen);
              FXLSStream.Write(RecFORMULA,SizeOf(TRecFORMULA_));
              FXLSStream.Write(Ptgs^,RecFORMULA.ParseLen + L);

              if Xc12Sheet.Cells.IterCellType = xctStringFormula then
                WriteSTRING(S);
            end
            else begin
              Fmla := XLSDecodeFormula(FManager,Xc12Sheet.Cells,Xc12Sheet.Index,Ptgs,Xc12Sheet.Cells.FormulaHelper.PtgsSize);
              GetMem(P,1024);
              try
                Encoder := TEncodeFormula.Create;
                try
                  Sz := Encoder.Encode(Fmla,0,P,1024,fvExcel97);
                finally
                  Encoder.Free;
                end;
                RecFORMULA.ParseLen := Sz;
                FXLSStream.WriteHeader(BIFFRECID_FORMULA,SizeOf(TRecFORMULA_) + RecFORMULA.ParseLen);
                FXLSStream.Write(RecFORMULA,SizeOf(TRecFORMULA_));
                FXLSStream.Write(P^,RecFORMULA.ParseLen);

                if Xc12Sheet.Cells.IterCellType = xctStringFormula then
                  WriteSTRING(S);
              finally
                FreeMem(P);
              end;
            end;
          end;
          xcftArray         : begin
            Ptgs := Xc12Sheet.Cells.FormulaHelper.Ptgs;
            if Ptgs.Id = xptg_EXCEL_97 then
              Ptgs := PXLSPtgs(NativeInt(Ptgs) + SizeOf(TXLSPtgs))
            else if Ptgs.Id = xptg_ARRAYCONSTS_97 then begin
              raise XLSRWException.Create('TODO Array consts 97');
            end
            else begin
              // TODO Convert 2007 ptgs -> 97.
            end;

            RecARRAY.Col1 := Xc12Sheet.Cells.FormulaHelper.TargetRef.Col1;
            RecARRAY.Row1 := Xc12Sheet.Cells.FormulaHelper.TargetRef.Row1;
            RecARRAY.Col2 := Xc12Sheet.Cells.FormulaHelper.TargetRef.Col2;
            RecARRAY.Row2 := Xc12Sheet.Cells.FormulaHelper.TargetRef.Row2;
            if (Xc12Sheet.Cells.FormulaHelper.Options and Xc12FormulaOpt_ACA) <> 0 then
              RecARRAY.Options := $0001
            else
              RecARRAY.Options := $0000;

            RecARRAY.Unused := 0;
            RecARRAY.DataSize := Xc12Sheet.Cells.FormulaHelper.PtgsSize - SizeOf(TXLSPtgs);

            PRecPtgsExp(PBuf).Id := xptgExp97;
            PRecPtgsExp(PBuf).Row := Row;
            PRecPtgsExp(PBuf).Col := Col;
            RecFORMULA.ParseLen := SizeOf(TRecPtgsExp);
            FXLSStream.WriteHeader(BIFFRECID_FORMULA,SizeOf(TRecFORMULA_) + SizeOf(TRecPtgsExp));
            FXLSStream.Write(RecFORMULA,SizeOf(TRecFORMULA_));
            FXLSStream.Write(PBuf^,SizeOf(TRecPtgsExp));

            FXLSStream.WriteHeader(BIFFRECID_ARRAY,SizeOf(TRecARRAY) + RecARRAY.DataSize);
            FXLSStream.Write(RecARRAY,SizeOf(TRecARRAY));
            FXLSStream.Write(Ptgs^,RecARRAY.DataSize);
            if Xc12Sheet.Cells.IterCellType = xctStringFormula then
              WriteSTRING(S);
          end;
          xcftDataTable     : ;
          xcftArrayChild,
          xcftArrayChild97  : begin
            PRecPtgsExp(PBuf).Id := xptgExp97;
            PRecPtgsExp(PBuf).Row := Xc12Sheet.Cells.FormulaHelper.ParentRow;
            PRecPtgsExp(PBuf).Col := Xc12Sheet.Cells.FormulaHelper.ParentCol;
            RecFORMULA.ParseLen := SizeOf(TRecPtgsExp);
            FXLSStream.WriteHeader(BIFFRECID_FORMULA,SizeOf(TRecFORMULA_) + SizeOf(TRecPtgsExp));
            FXLSStream.Write(RecFORMULA,SizeOf(TRecFORMULA_));
            FXLSStream.Write(PBuf^,SizeOf(TRecPtgsExp));
            if Xc12Sheet.Cells.IterCellType = xctStringFormula then
              WriteSTRING(S);
          end;
          xcftDataTableChild: ;
        end;
      end;
    end;
  end;
  RkFlushCache(CurrRkRow);
  MulblankFlushCache(CurrRkRow);
end;

procedure TXLSWriteII.WritePointer(RecId: word; P: Pointer; Size: word);
var
  Header: TBIFFHeader;
begin
  Header.RecId := RecID;
  Header.Length := Size;
  FXLSStream.Write(Header,SizeOf(TBIFFHeader));
  if Size > 0 then
    FXLSStream.Write(P^,Header.Length);
end;

procedure TXLSWriteII.WriteToStream(Stream: TStream);
var
  i,j,k: integer;
  HasRecordData: boolean;
begin
  FBoundsheetList.Clear;
  try
    k := 0;
    for i := 0 to FXLS.Sheets.Count - 1 do begin
      FCurrXc12Sheet := FManager.Worksheets[i];

      while (k < FXLS.SheetCharts.Count) and (FXLS.SheetCharts[k].SheetIndex <= FXLS.Sheets[i]._Int_SheetIndex) do begin
        FBoundsheetList.AddChart(k,FXLS.SheetCharts[k].Name);
        Inc(k);
      end;
// TODO
//      if i = FXLS.Workbook.SelectedTab then begin
//        // Selected tab.
//        FXLS.Sheets[i]._Int_Records.WINDOW2.Options := FXLS.Sheets[i]._Int_Records.WINDOW2.Options or $0200;
//        FXLS.Sheets[i]._Int_Records.WINDOW2.Options := FXLS.Sheets[i]._Int_Records.WINDOW2.Options or $0400;
//      end
//      else begin
//        FXLS.Sheets[i]._Int_Records.WINDOW2.Options := FXLS.Sheets[i]._Int_Records.WINDOW2.Options and not $0200;
//        FXLS.Sheets[i]._Int_Records.WINDOW2.Options := FXLS.Sheets[i]._Int_Records.WINDOW2.Options and not $0400;
//      end;
      FBoundsheetList.AddSheet(i,FManager.Worksheets[i].Name,THiddenState(FCurrXc12Sheet.State),FXLS.Sheets[i].WorksheetType);
    end;
    // Charts after the last sheet.
    while k < FXLS.SheetCharts.Count do begin
      FBoundsheetList.AddChart(k,FXLS.SheetCharts[k].Name);
      Inc(k);
    end;


    HasRecordData := FXLS.Records.Count > 0;

    FXLSStream.ExtraObjects := FXLS.ExtraObjects;
    FXLSStream.TargetStream := Stream;

    FXLSStream.SaveVBA := FXLS.PreserveMacros;
    FXLSStream.OpenStorageWrite(FXLS.Filename,FXLS.Version);

    if not HasRecordData then
      FXLS.Records.MoveAllDefault;
    for i := 0 to FXLS.Records.Count - 1 do begin
      try
        case FXLS.Records[i].RecId of
          $0809: begin
            if FXLS.Version = xvExcel50 then begin
              PRecBOF8(PBuf).SubStreamType := $0005;
              PRecBOF8(PBuf).VersionNumber := $0500;
              // Can not be zero.
              PRecBOF8(PBuf).BuildIdentifier := $0DBB;
              // Can not be zero.
              PRecBOF8(PBuf).BuildYear := $07CE;
              WriteBuf($0809,SizeOf(TRecBOF7));
            end
            else
              FXLS.Records.WriteRec(i,FXLSStream);
            if (PRecBOF8(@FXLS.Records[i].Data).SubstreamType = $0005) and (FXLS.Password <> '') then begin
              PRecFILEPASS(PBuf).Options := $0001;
              PRecFILEPASS(PBuf).SillyPassword[0] := $01;
              PRecFILEPASS(PBuf).SillyPassword[1] := $00;
              PRecFILEPASS(PBuf).SillyPassword[2] := $01;
              PRecFILEPASS(PBuf).SillyPassword[3] := $00;
              FXLSStream.CreatePassword(PRecFILEPASS(PBuf),FXLS.Password);
              FXLSStream.WriteHeader(BIFFRECID_FILEPASS,SizeOf(TRecFILEPASS));
              FXLSStream.Write(PBuf^,SizeOf(TRecFILEPASS));
              // These records are not really required, but when encrypting files,
              // the first record after FILEPASS is unencrypted, and without
              // these, that record will be WRITEACCESS.
              if FXLS.Records.FindRecord(BIFFRECID_INTERFACEHDR) < 0 then begin
                FXLSStream.WriteWord(BIFFRECID_INTERFACEHDR,$04B0);
                FXLSStream.WriteWord(BIFFRECID_MMSADDMENUDELMENU,$0000);
                FXLSStream.WriteHeader(BIFFRECID_INTERFACEEND,0);
              end;
            end;
          end;

          BIFFRECID_WRITEACCESS: WREC_WRITEACCESS;

          BIFFRECID_OBPROJ:
            WREC_OBPROJ;

          BIFFRECID_PASSWORD:
            WREC_PASSWORD;
          BIFFRECID_WINDOW1:
            WREC_WINDOW1;
          BIFFRECID_BACKUP:
            WREC_BACKUP;
          BIFFRECID_1904:
            WREC_1904;
          BIFFRECID_PRECISION:
            WREC_PRECISION;
          BIFFRECID_REFRESHALL:
            WREC_REFRESHALL;
          BIFFRECID_BOOKBOOL:
            WREC_BOOKBOOL;

          INTERNAL_FORMATS: begin
            if FXLS.Version = xvExcel50 then
              Write50FormatBlock
            else begin
              WREC_FONT;
              WREC_FORMAT;
              WREC_XF;
              WREC_STYLE;
            end;
            WREC_PALETTE;
          end;
          INTERNAL_BOUNDSHEETS: begin
            for j := 0 to FBoundsheetList.Count - 1 do
              WREC_BOUNDSHEET(j);
          end;
          // Don't write this.
          BIFFRECID_RECALCID: ;
          INTERNAL_NAMES: begin
            if not FExternalNamesWritten then begin
//              FExternalNamesWritten := True;
              FXLS.FormulaHandler.ExternalNames.WriteRecords(FXLSStream);
              WREC_NAME;
            end;
          end;
          INTERNAL_MSODRWGRP: begin
            if FXLS.MSOPictures.HasData then begin
              FXLSStream.BeginCONTINUEWrite(MAXRECSZ_97);
              FXLS.MSOPictures.WriteToStream(FXLSStream,PBuf);
              FXLSStream.EndCONTINUEWrite;
            end;
          end;
          INTERNAL_SST: begin
            WREC_SST;
          end;
          BIFFRECID_TABID: begin
            if FXLS.Records[i].Length = 0 then begin
              FXLSStream.WriteHeader(BIFFRECID_TABID,FXLS.Sheets.Count * 2);
              for j := 1 to FXLS.Sheets.Count do
                FXLSStream.WWord(j);
            end;
          end;

          else begin
            FXLS.Records.WriteRec(i,FXLSStream);
          end;
        end;
      except
        on E: XLSRWException do
          raise XLSRWException.CreateFmt('Error on writing record # %d' + #13 + E.Message,[i]);
      end;
    end;
    k := 0;
    for i := 0 to FXLS.Sheets.Count - 1 do begin
      if FXLS.Sheets[i].SheetProtection <> DefaultSheetProtections then
        FXLS.Sheets[i]._Int_Records.PROTECT := 1
      else
        FXLS.Sheets[i]._Int_Records.PROTECT := 0;

      if not FXLS.Sheets[i]._Int_HasDefaultRecords then begin
        FXLS.Sheets[i]._Int_Records.MoveAllDefault;
        FXLS.Sheets[i]._Int_HasDefaultRecords := True;
      end;

      FCurrSheet := i;
      FCurrXc12Sheet := FManager.Worksheets[FCurrSheet];

      while (k < FXLS.SheetCharts.Count) and (FXLS.SheetCharts[k].SheetIndex < FXLS.Sheets[i]._Int_SheetIndex) do begin
        FBoundsheetList.Charts[k].WritePos(FXLSStream);
        FXLS.SheetCharts.SaveToStream(k,FXLSStream);
        Inc(k);
      end;
      FBoundsheetList.Sheets[i].WritePos(FXLSStream);

      for j := 0 to FXLS.Sheets[i]._Int_Records.Count - 1 do begin
        case FXLS.Sheets[i]._Int_Records[j].RecId of
          // Prefix data
          BIFFRECID_CALCMODE:
            WREC_CALCMODE;
          BIFFRECID_CALCCOUNT:
            WREC_CALCCOUNT;
          BIFFRECID_REFMODE:
            WREC_REFMODE;
          BIFFRECID_ITERATION:
            WREC_ITERATION;
          BIFFRECID_DELTA:
            WREC_DELTA;
          BIFFRECID_SAVERECALC:
            WREC_SAVERECALC;
          BIFFRECID_PRINTHEADERS:
            WREC_PRINTHEADERS;
          BIFFRECID_PRINTGRIDLINES:
            WREC_PRINTGRIDLINES;
          BIFFRECID_GUTS:
            WREC_GUTS;
          BIFFRECID_GRIDSET:
            WREC_GRIDSET;
          BIFFRECID_PROTECT: begin
            if soProtected in FXLS.Sheets[FCurrSheet].Options then
              WriteWord(BIFFRECID_PROTECT,1);
          end;
          BIFFRECID_DEFAULTROWHEIGHT:
            WREC_DEFAULTROWHEIGHT;
          BIFFRECID_DEFCOLWIDTH: begin
            WREC_DEFCOLWIDTH;
            FXLS.Sheets[i].Autofilters.SaveToStream(FXLSStream,PBuf);
          end;
          BIFFRECID_WSBOOL:
            WREC_WSBOOL;
          BIFFRECID_VCENTER:
            WREC_VCENTER;
          BIFFRECID_HCENTER:
            WREC_HCENTER;
          BIFFRECID_SETUP:
            WREC_SETUP;
          BIFFRECID_DIMENSIONS:
            WREC_DIMENSIONS;
          BIFFRECID_WINDOW2:
            WREC_WINDOW2;
          BIFFRECID_INDEX: ;


          INTERNAL_PAGEBREAKES: begin
            WREC_HORIZONTALPAGEBREAKS;
            WREC_VERTICALPAGEBREAKS;
          end;
          INTERNAL_HEADER: begin
            WREC_HEADER;
            WREC_FOOTER;
          end;
          INTERNAL_MARGINS: begin
            WREC_MARGINS;
          end;
          INTERNAL_COLINFO: begin
            WREC_COLINFO;
          end;
          INTERNAL_CELLDATA: begin
            WREC_ROW;
            WriteCells;

            if FXLS[i]._Int_CachedMSORecs.Count > 0 then
              FXLS[i]._Int_CachedMSORecs.WriteAllRecs(FXLSStream);
          end;

//          BIFFRECID_WINDOW2: begin
//            // If zoom in WINDOW2 is changed, SCL has to be written.
//            if FXLS.Version = xvExcel50 then begin
//              WREC_WINDOW1;
////              FXLS.Sheets[i]._Int_Records.WriteRec(j,FXLSStream);
//            end
//            else begin
//              WREC_SCL;
//              FXLS.Sheets[i]._Int_Records.WriteRec(j,FXLSStream);
//            end;
//            // If there are active panes, selections are written at WREC_PANE.
//            if FXLS.Sheets[FCurrSheet].Pane.PaneType = ptNone then
//              WREC_SELECTION(3);
//          end;

          BIFFRECID_PASSWORD:
            WREC_PASSWORD;

          INTERNAL_SUFFIXDATA : begin
            WREC_MSODRAWING;
            WREC_PANE;
            WREC_SELECTION;
            WREC_MERGECELLS;
            WREC_CONDFMT;
            WREC_HLINK;
            WREC_DVAL;
            WREC_FEATHEADR;
          end;
          else
            FXLS.Sheets[i]._Int_Records.WriteRec(j,FXLSStream);
        end;
      end;
//      FXLSStream.WriteHeader(BIFFRECID_EOF,0);
    end;
    // Charts after the last sheet.
    while k < FXLS.SheetCharts.Count do begin
      FBoundsheetList.Charts[k].WritePos(FXLSStream);
      FXLS.SheetCharts.SaveToStream(k,FXLSStream);
      Inc(k);
    end;
    if FXLS.Password <> '' then
      FXLSStream.EncryptFile;
    if FXLSStream.SaveVBA then
      FXLSStream.WriteVBA;
  finally
    FXLSStream.CloseStorage;
  end;
  FBoundsheetList.Clear;
end;

procedure TXLSWriteII.WREC_1904;
begin
  WriteWord(BIFFRECID_1904,BoolAsWord(FManager.Workbook.WorkbookPr.Date1904));
end;

procedure TXLSWriteII.WREC_BACKUP;
begin
  WriteWord(BIFFRECID_BACKUP,BoolAsWord(FManager.Workbook.WorkbookPr.BackupFile));
end;

procedure TXLSWriteII.WREC_BOF(SubStreamType: TSubStreamType);
begin
  case FXLS.Version of
    xvExcel40: begin
      PRecBOF4(PBuf).A := $0000;
      PRecBOF4(PBuf).B := $0010;
      PRecBOF4(PBuf).C := $18AF;
    end;
    xvExcel50: begin
      PRecBOF8(PBuf).VersionNumber := $0500;
      // Can not be zero.
      PRecBOF8(PBuf).BuildIdentifier := $0DBB;
      // Can not be zero.
      PRecBOF8(PBuf).BuildYear := $07CE;
    end;
    xvExcel97: begin
      PRecBOF8(PBuf).VersionNumber := $0600;
      PRecBOF8(PBuf).BuildIdentifier := $18AF;
      PRecBOF8(PBuf).BuildYear := $07CD;
      PRecBOF8(PBuf).FileHistoryFlags := 0;
      if FXLS.IsMac then
        PRecBOF8(PBuf).FileHistoryFlags := $00000010;
      PRecBOF8(PBuf).LowBIFF := $00000106;
    end;
  end;
  if FXLS.Version > xvExcel40 then begin
    case SubStreamType of
      stGlobals:      PRecBOF8(PBuf).SubStreamType := $0005;
      stVBModule:     PRecBOF8(PBuf).SubStreamType := $0006;
      stWorksheet:    PRecBOF8(PBuf).SubStreamType := $0010;
      stChart:        PRecBOF8(PBuf).SubStreamType := $0020;
      stExcel4Macro:  PRecBOF8(PBuf).SubStreamType := $0040;
      stWorkspace:    PRecBOF8(PBuf).SubStreamType := $0100;
    end;
  end;
  case FXLS.Version of
    xvExcel40: WriteBuf($0409,SizeOf(TRecBOF4));
    xvExcel50: WriteBuf($0809,SizeOf(TRecBOF7));
    xvExcel97: WriteBuf($0809,SizeOf(TRecBOF8));
  end;
end;

procedure TXLSWriteII.WREC_BOOKBOOL;
begin
  PWord(PBuf)^ := $0000;
  if FManager.Workbook.WorkbookPr.SaveExternalLinkValues then
    PWord(PBuf)^ := PWord(PBuf)^ and $0001;
  PWord(PBuf)^ := PWord(PBuf)^ or (Word(FManager.Workbook.WorkbookPr.UpdateLinks) shl 5);
  WriteWord(BIFFRECID_BOOKBOOL,PWord(PBuf)^);
end;

procedure TXLSWriteII.WREC_BOUNDSHEET(Index: integer);
var
  S: XLS8String;
begin
  if FXLS.Version >= xvExcel97 then begin
    FBoundsheetList[Index].FilePos := FXLSStream.Pos + SizeOf(TBIFFHeader);
    PRecBOUNDSHEET8(PBuf).BOFPos := 0;
    case FBoundsheetList[Index].FBoundsheetType of
      btSheet:       PRecBOUNDSHEET8(PBuf).Options := $0000;
      btChart:       PRecBOUNDSHEET8(PBuf).Options := $0200;
      btVBModule:    PRecBOUNDSHEET8(PBuf).Options := $0600;
      btExcel4Macro: PRecBOUNDSHEET8(PBuf).Options := $0100;
    end;
    PRecBOUNDSHEET8(PBuf).Options := PRecBOUNDSHEET8(PBuf).Options + Word(FBoundsheetList[Index].FHidden);
    PRecBOUNDSHEET8(PBuf).NameLen := Length(FBoundsheetList[Index].Name);
    PRecBOUNDSHEET8(PBuf).Name[0] := $01;
    Move(Pointer(FBoundsheetList[Index].Name)^,PRecBOUNDSHEET8(PBuf).Name[1],Length(FBoundsheetList[Index].Name) * 2);
    WriteBuf(BIFFRECID_BOUNDSHEET,7 + Length(FBoundsheetList[Index].Name) * 2 + 1);
  end
  else if FXLS.Version >= xvExcel50 then begin
    FBoundsheetList[Index].FilePos := FXLSStream.Pos + SizeOf(TBIFFHeader);
    PRecBOUNDSHEET7(PBuf).BOFPos := 0;
    case FBoundsheetList[Index].FBoundsheetType of
      btSheet: PRecBOUNDSHEET7(PBuf).Options := $0000;
      btChart: PRecBOUNDSHEET7(PBuf).Options := $0200;
    end;
    PRecBOUNDSHEET7(PBuf).Options := PRecBOUNDSHEET7(PBuf).Options + Word(FBoundsheetList[Index].FHidden);
    S := XLS8String(FBoundsheetList[Index].Name);
    PRecBOUNDSHEET7(PBuf).NameLen := Length(S);
    Move(Pointer(S)^,PRecBOUNDSHEET7(PBuf).Name,Length(S));
    WriteBuf(BIFFRECID_BOUNDSHEET,7 + Length(S));
  end;
end;

procedure TXLSWriteII.WREC_CODEPAGE;
begin
  WriteWord(BIFFRECID_CODEPAGE,FXLS.Codepage);
end;

procedure TXLSWriteII.WREC_COUNTRY;
begin
  if FXLS.Version < xvExcel97 then Exit;
  PRecCOUNTRY(PBuf).DefaultCountryIndex := FXLS.DefaultCountryIndex;
  PRecCOUNTRY(PBuf).WinIniCountry := FXLS.WinIniCountry;
  WriteBuf(BIFFRECID_COUNTRY,SizeOf(TRecCOUNTRY));
end;

procedure TXLSWriteII.WREC_DSF;
begin
  if FXLS.Version < xvExcel97 then Exit;
  WriteWord(BIFFRECID_DSF,$00);
end;

procedure TXLSWriteII.WREC_EXCEL9FILE;
begin
  WriteRecId(BIFFRECID_EXCEL9FILE);
end;

procedure TXLSWriteII.WREC_EXTERNSHEET;
var
  i: integer;
begin
  if FXLS.FormulaHandler.ExternalNames.HasEXTERNSHEET then
    FXLS.FormulaHandler.ExternalNames.WriteRecords(FXLSStream)
  else begin
    PRecEXTERNSHEET8(PBuf).XTICount := FManager.Worksheets.Count;
    for i := 0 to FManager.Worksheets.Count - 1 do begin
      PRecEXTERNSHEET8(PBuf).XTI[i].SupBook := 0;
      PRecEXTERNSHEET8(PBuf).XTI[i].FirstTab := i;
      PRecEXTERNSHEET8(PBuf).XTI[i].LastTab := i;
    end;
    WriteBuf(BIFFRECID_EXTERNSHEET,SizeOf(word) + SizeOf(TXTI) * FManager.Worksheets.Count);
  end;
end;

procedure TXLSWriteII.WREC_EOF;
begin
  WriteRecId(BIFFRECID_EOF);
end;

procedure TXLSWriteII.WREC_FONT;
var
  i: integer;
  S: XLS8String;
  Sz: integer;
  Xc12Font: TXc12Font;
  Color: TXc12IndexColor;

procedure WriteFONT40;
var
  i,Sz: integer;
begin
  for i := 0 to FXLS.Fonts.Count - 1 do begin
    PRecFont4(PBuf).Height := Round(FXLS.Fonts[i].Size * 20);
    PRecFont4(PBuf).Attributes := 0;
//    if xfsItalic    in FXLS.Fonts[i].Style then PRecFont4(PBuf).Attributes := PRecFont4(PBuf).Attributes or $02;
//    if xfsStrikeOut in FXLS.Fonts[i].Style then PRecFont4(PBuf).Attributes := PRecFont4(PBuf).Attributes or $08;

    PRecFont4(PBuf).Attributes := $7FFF;

    Sz := SizeOf(TRecFONT4) - 256;
    PRecFont4(PBuf).NameLen := Length(FXLS.Fonts[i].Name);
    Move(Pointer(FXLS.Fonts[i].Name)^,PRecFont4(PBuf).Name,PRecFont4(PBuf).NameLen);
    Inc(Sz,PRecFont4(PBuf).NameLen);
    if i <> 4 then
      WriteBuf(BIFFRECID_FONT,Sz);
  end;
end;

begin
  for i := XLS_STYLE_DEFAULT_XF_COUNT to FManager.StyleSheet.Fonts.Count - 1 do begin
    if i = 4 then
      Continue;
    Xc12Font := FManager.StyleSheet.Fonts[i];

    PRecFont(PBuf).Height := Round(Xc12Font.Size * 20);
    PRecFont(PBuf).Attributes := 0;
    if xfsItalic    in Xc12Font.Style then PRecFont(PBuf).Attributes := PRecFont(PBuf).Attributes or $02;
    if xfsStrikeOut in Xc12Font.Style then PRecFont(PBuf).Attributes := PRecFont(PBuf).Attributes or $08;
    Color := Xc12ColorToIndex(Xc12Font.Color);
    if Color = xc0 then
      PRecFont(PBuf).ColorIndex := FONT_COLOR_SYS_WINTEXT
    else
      PRecFont(PBuf).ColorIndex := Word(Color);
    if xfsBold in Xc12Font.Style then
      PRecFont(PBuf).Bold := $02BC
    else
      PRecFont(PBuf).Bold := $0190;
    case Xc12Font.SubSuperScript of
      xssNone:        PRecFont(PBuf).SubSuperScript := $00;
      xssSuperscript: PRecFont(PBuf).SubSuperScript := $01;
      xssSubscript:   PRecFont(PBuf).SubSuperScript := $02;
    end;
    case Xc12Font.Underline of
      xulNone:          PRecFont(PBuf).Underline := $00;
      xulSingle:        PRecFont(PBuf).Underline := $01;
      xulDouble:        PRecFont(PBuf).Underline := $02;
      xulSingleAccount: PRecFont(PBuf).Underline := $21;
      xulDoubleAccount: PRecFont(PBuf).Underline := $22;
    end;
    PRecFont(PBuf).Reserved := 0;
    PRecFont(PBuf).CharSet := Byte(Xc12Font.CharSet);
    PRecFont(PBuf).Family := Xc12Font.Family;
    PRecFont(PBuf).NameLen := Length(Xc12Font.Name);
    Sz := SizeOf(TRecFONT) - 256;
    if FVersion >= xvExcel97 then begin
      WideStringToByteStr(Xc12Font.Name,@PRecFont(PBuf).Name);
      Inc(Sz,Length(Xc12Font.Name) * 2 + 1);
    end
    else begin
      S := XLS8String(Xc12Font.Name);
      Move(Pointer(S)^,PRecFont(PBuf).Name,Length(S));
      Inc(Sz,PRecFont(PBuf).NameLen);
    end;

    WriteBuf(BIFFRECID_FONT,Sz);
  end;
end;

procedure TXLSWriteII.WREC_FORMAT;
var
  i: integer;
  S: AxUCString;
begin
  for i := 0 to FManager.StyleSheet.NumFmts.Count - 1 do begin
    if (i > High(ExcelStandardNumFormats)) or FManager.StyleSheet.NumFmts[i].StdRedefined then begin
      S := FManager.StyleSheet.NumFmts[i].Value;
      if S <> '' then begin
        if FXLS.Version < xvExcel97 then begin
          PRecFORMAT7(PBuf).Index := i;
          PRecFORMAT7(PBuf).Len := Length(S);
          WideStringToByteStr(S,@PRecFORMAT7(PBuf).Data);
          WriteBuf(BIFFRECID_FORMAT,SizeOf(TRecFORMAT7) + Length(S) * 2 + 1);
        end
        else begin
          PRecFORMAT8(PBuf).Index := i;
          PRecFORMAT8(PBuf).Len := Length(S);
          WideStringToByteStr(S,@PRecFORMAT8(PBuf).Data);
          WriteBuf(BIFFRECID_FORMAT,SizeOf(TRecFORMAT8) + Length(S) * 2 + 1);
        end
      end;
    end;
  end;
end;

procedure TXLSWriteII.WREC_GRIDSET;
begin
  WriteBoolean(BIFFRECID_GRIDSET,FCurrXc12Sheet.PrintOptions.GridLinesSet);
end;

procedure TXLSWriteII.WREC_GUTS;
var
  i: integer;
  Row: PXLSMMURowHeader;
  MaxOutline: integer;
begin
  PRecGUTS(PBuf).SizeRow := $0000;
  PRecGUTS(PBuf).SizeCol := $0000;

  MaxOutline := 0;
  FManager.Worksheets[FCurrSheet].Cells.BeginIterateRow;
  while FManager.Worksheets[FCurrSheet].Cells.IterateNextRow do begin
    Row := FManager.Worksheets[FCurrSheet].Cells.IterRow;
    if Row.OutlineLevel > MaxOutline then
      MaxOutline := Row.OutlineLevel;
  end;
  if MaxOutline > 7 then
    MaxOutline := 7;

  if MaxOutline > 0 then
    PRecGUTS(PBuf).LevelRow := MaxOutline + 1
  else
    PRecGUTS(PBuf).LevelRow := 0;

  MaxOutline := 0;
  for i := 0 to FCurrXc12Sheet.Columns.Count - 1 do begin
    if FCurrXc12Sheet.Columns[i].OutlineLevel > MaxOutline then
      MaxOutline := FCurrXc12Sheet.Columns[i].OutlineLevel;
  end;
  if MaxOutline > 7 then
    MaxOutline := 7;

  if MaxOutline > 0 then
    PRecGUTS(PBuf).LevelCol := MaxOutline + 1
  else
    PRecGUTS(PBuf).LevelCol := 0;

  WriteBuf(BIFFRECID_GUTS,SizeOf(TRecGUTS));
end;

procedure TXLSWriteII.WREC_HIDEOBJ;
begin
//  WriteWord(BIFFRECID_HIDEOBJ,Word(FXLS.OptionsDialog.ShowObjects));
end;

procedure TXLSWriteII.WREC_PASSWORD;
begin
  if (FCurrXc12Sheet = Nil) and (FManager.Workbook.WorkbookProtection.WorkbookPassword <> 0) then begin
    // If ste to True, the file will be invalid.
    WriteWord(BIFFRECID_PROTECT,0);
    WriteWord(BIFFRECID_PASSWORD,FManager.Workbook.WorkbookProtection.WorkbookPassword);
  end
  else if (FCurrXc12Sheet <> Nil) and (FCurrXc12Sheet.SheetProtection.Password <> 0) then begin
    WriteWord(BIFFRECID_PROTECT,1);
    WriteWord(BIFFRECID_SCENPROTECT,BoolAsWord(FCurrXc12Sheet.SheetProtection.Scenarios));
    WriteWord(BIFFRECID_OBJPROTECT,BoolAsWord(FCurrXc12Sheet.SheetProtection.Objects));
    WriteWord(BIFFRECID_PASSWORD,FCurrXc12Sheet.SheetProtection.Password);
  end
  else begin
    WriteWord(BIFFRECID_PROTECT,0);
    WriteWord(BIFFRECID_PASSWORD,0);
  end;
end;

procedure TXLSWriteII.WREC_PRECISION;
begin
  WriteWord(BIFFRECID_PRECISION,BoolAsWord(FManager.Workbook.CalcPr.FullPrecision));
end;

procedure TXLSWriteII.WREC_REFRESHALL;
begin
  WriteWord(BIFFRECID_REFRESHALL,BoolAsWord(FManager.Workbook.WorkbookPr.RefreshAllConnections));
end;

procedure TXLSWriteII.WREC_SST;
var
  SSTWriter: THelperWriteSST;
begin
 if FXLS.Version < xvExcel97 then
    Exit;
  SSTWriter := THelperWriteSST.Create(FManager,FXLSStream,PBuf);
  try
    SSTWriter.Write;
  finally
    SSTWriter.Free;
  end;
end;

procedure TXLSWriteII.WREC_STYLE;
const DefaultSTYLE: array[0..5] of longword =
($FF038010,$FF068011,$FF048012,$FF078013,$FF008000,$FF058014);
var
  i: integer;
  P: PRecSTYLE;
  PU: PRecSTYLE_USER;
  Sz: integer;
begin
  if FXLS.WriteDefaultData then begin
    for i := 0 to High(DefaultSTYLE) do
      WritePointer(BIFFRECID_STYLE,@DefaultSTYLE[i],SizeOf(DefaultSTYLE[i]));
  end
  else begin
    for i := 0 to FXLS.Styles.Count - 1 do begin
      P := FXLS.Styles[i];
      if (P.FormatIndex and $F000) = 0 then begin
        PU := PRecSTYLE_USER(P);
        if PU.Len = 0 then
          Sz := 0
        else if PByteArray(@PU.Data)[0] = 0 then
          Sz := PU.Len + 1
        else if PByteArray(@PU.Data)[0] = 1 then
          Sz := (PU.Len * 2) + 1
        // Excel 95
        else
          Sz := PU.Len;
        Inc(Sz,SizeOf(TRecSTYLE_USER));
      end
      else
        Sz := SizeOf(TRecSTYLE);

      Move(P^,PBuf^,Sz);
      WriteBuf(BIFFRECID_STYLE,Sz);
    end;
  end;
end;

procedure TXLSWriteII.WREC_SUPBOOK;
begin
 if not FXLS.FormulaHandler.ExternalNames.HasSUPBOOK then begin
   PBuf[0] := FManager.Worksheets.Count;
   PBuf[1] := $00;
   PBuf[2] := $01;
   PBuf[3] := $04;
   WriteBuf(BIFFRECID_SUPBOOK,4);
 end;
end;

procedure TXLSWriteII.WREC_OBPROJ;
begin
  if FXLS.PreserveMacros then
    FXLSStream.WriteHeader(BIFFRECID_OBPROJ,0);
end;                                                           

procedure TXLSWriteII.WREC_USESELFS;
begin
  if FXLS.Version < xvExcel97 then Exit;
  WriteWord(BIFFRECID_USESELFS,$0000);
end;

procedure TXLSWriteII.WREC_WINDOW1;
begin
  if FManager.Workbook.BookViews.Count <= 0 then begin
    PRecWINDOW1(PBuf).Left := 100;
    PRecWINDOW1(PBuf).Top := 100;
    PRecWINDOW1(PBuf).Width := 10000;
    PRecWINDOW1(PBuf).Height := 7000;
    PRecWINDOW1(PBuf).Options := $0038;
    PRecWINDOW1(PBuf).SelectedTabIndex := 0;
    PRecWINDOW1(PBuf).FirstDispTabIndex := 0;
    PRecWINDOW1(PBuf).SelectedTabs := 1;
    PRecWINDOW1(PBuf).TabRatio := $0258;
  end
  else begin
    PRecWINDOW1(PBuf).Left := FManager.Workbook.BookViews[0].XWindow;
    PRecWINDOW1(PBuf).Top := FManager.Workbook.BookViews[0].YWindow;
    PRecWINDOW1(PBuf).Width := FManager.Workbook.BookViews[0].WindowWidth;
    PRecWINDOW1(PBuf).Height := FManager.Workbook.BookViews[0].WindowHeight;
    PRecWINDOW1(PBuf).TabRatio := FManager.Workbook.BookViews[0].TabRatio;
    PRecWINDOW1(PBuf).SelectedTabIndex := FManager.Workbook.BookViews[0].ActiveTab;
    PRecWINDOW1(PBuf).FirstDispTabIndex := FManager.Workbook.BookViews[0].FirstSheet;
    PRecWINDOW1(PBuf).SelectedTabs := 1;

    PRecWINDOW1(PBuf).Options := $0000;

    if FManager.Workbook.BookViews[0].Visibility >= x12vHidden then
      PRecWINDOW1(PBuf).Options := PRecWINDOW1(PBuf).Options or $0001;

    if FManager.Workbook.BookViews[0].Minimized then
      PRecWINDOW1(PBuf).Options := PRecWINDOW1(PBuf).Options or $0002;
    if FManager.Workbook.BookViews[0].ShowHorizontalScroll then
      PRecWINDOW1(PBuf).Options := PRecWINDOW1(PBuf).Options or $0008;
    if FManager.Workbook.BookViews[0].ShowVerticalScroll then
      PRecWINDOW1(PBuf).Options := PRecWINDOW1(PBuf).Options or $0010;
    if FManager.Workbook.BookViews[0].ShowSheetTabs then
      PRecWINDOW1(PBuf).Options := PRecWINDOW1(PBuf).Options or $0020;
    if FManager.Workbook.BookViews[0].AutoFilterDateGrouping then
      PRecWINDOW1(PBuf).Options := PRecWINDOW1(PBuf).Options or $0040;
  end;
  WriteBuf(BIFFRECID_WINDOW1,SizeOf(TRecWINDOW1));
end;

procedure TXLSWriteII.WREC_WINDOWPROTECT;
begin
  // Not sure if this is correct.
  WriteWord(BIFFRECID_WINDOWPROTECT,BoolAsWord(FManager.Workbook.WorkbookProtection.LockWindows));
end;

procedure TXLSWriteII.WREC_WRITEACCESS;
var
  sz: byte;
  S: AXUCString;
begin
  if FXLS.Version < xvExcel40 then
    sz := 31
  else
    sz :=112;
  S := Copy(FManager.Workbook.UserName,1,sz);
  if FXLS.Version <= xvExcel50 then
    S := Char(Length(S)) + S;
  FillChar(PBuf^,sz,$20);
  PWord(PBuf)^ := Length(S);
  PByteArray(PBuf)[2] := 1;
  Move(Pointer(S)^,PByteArray(PBuf)[3],Length(S) * 2);
  WriteBuf(BIFFRECID_WRITEACCESS,sz);
end;

procedure TXLSWriteII.WREC_XF;
var
  XFWriter: THelperWriteXF;
begin
  XFWriter := THelperWriteXF.Create(FManager,FXLSStream,PBuf);
  try
    XFWriter.Write;
  finally
    XFWriter.Free;
  end;
end;

procedure TXLSWriteII.WREC_PALETTE;
var
  V,I: integer;
  W: word;
begin
  if FXLS.Version = xvExcel50 then begin
    FXLSStream.WriteHeader(BIFFRECID_PALETTE,2 + SizeOf(longword) * 56);
    W := 56;
    FXLSStream.Write(W,SizeOf(word));
    for i := 8 to 63 do begin
      V := Xc12IndexColorPalette[i];
      FXLSStream.Write(V,SizeOf(longword));
    end;
  end
  else if FXLS.Version >= xvExcel97 then begin
    for i := 8 to High(Xc12IndexColorPalette) do begin
      if Xc12IndexColorPalette[i] <> TXc12DefaultIndexColorPalette[i] then begin
        Move(Xc12IndexColorPalette[8],PBuf[2],SizeOf(longword) * 56);
        PWordArray(PBuf)[0] := 56;
        WriteBuf(BIFFRECID_PALETTE,2 + SizeOf(longword) * 56);
        Break;
      end;
    end;
  end;
end;

// Sheet prefix

procedure TXLSWriteII.WREC_CALCCOUNT;
begin
  WriteWord(BIFFRECID_CALCCOUNT,FManager.Workbook.CalcPr.IterateCount);
end;

procedure TXLSWriteII.WREC_CALCMODE;
begin
  case FManager.Workbook.CalcPr.CalcMode of
    cmManual      : WriteWord(BIFFRECID_CALCMODE,$0000);
    cmAutomatic   : WriteWord(BIFFRECID_CALCMODE,$0001);
    cmAutoExTables: WriteWord(BIFFRECID_CALCMODE,$FFFF);
  end;
end;

procedure TXLSWriteII.WREC_COLINFO;
var
  i: integer;
  Col: TXc12Column;
begin
  if FVersion < xvExcel97 then
    Exit;

  for i := 0 to FManager.Worksheets[FCurrSheet].Columns.Count - 1 do begin

    Col := FManager.Worksheets[FCurrSheet].Columns[i];
    PRecCOLINFO(PBuf).Col1 := Fork(Col.Min,0,255);
    PRecCOLINFO(PBuf).Col2 := Fork(Col.Max,0,255);
    PRecCOLINFO(PBuf).Width := Round(Col.Width * 255);
    PRecCOLINFO(PBuf).FormatIndex := CheckDeafultXF(Col.Style.Index);
    PRecCOLINFO(PBuf).Options := Col.OutlineLevel shl 8;

    PRecCOLINFO(PBuf).Options := $0000;
    if xcoHidden in Col.Options then
      PRecCOLINFO(PBuf).Options := PRecCOLINFO(PBuf).Options or $0001;
    if xcoCustomWidth in Col.Options then
      PRecCOLINFO(PBuf).Options := PRecCOLINFO(PBuf).Options or $0002;
    if xcoCollapsed in Col.Options then
      PRecCOLINFO(PBuf).Options := PRecCOLINFO(PBuf).Options or $0010;

    PRecCOLINFO(PBuf).Reserved := 0;

    WriteBuf(BIFFRECID_COLINFO,SizeOf(TRecCOLINFO));
  end;
end;

procedure TXLSWriteII.WREC_DEFAULTROWHEIGHT;
begin
  if FCurrXc12Sheet.SheetFormatPr.DefaultRowHeight > 0 then begin
    PRecDEFAULTROWHEIGHT(PBuf).Options := 1;
    PRecDEFAULTROWHEIGHT(PBuf).Height := Round(FCurrXc12Sheet.SheetFormatPr.DefaultRowHeight);
  end
  else begin
    PRecDEFAULTROWHEIGHT(PBuf).Options := 0;
    PRecDEFAULTROWHEIGHT(PBuf).Height := 255;
  end;
  WriteBuf(BIFFRECID_DEFAULTROWHEIGHT,SizeOf(TRecDEFAULTROWHEIGHT));
end;

procedure TXLSWriteII.WREC_DEFCOLWIDTH;
begin
  if FCurrXc12Sheet.SheetFormatPr.DefaultColWidth > 0 then
    WriteWord(BIFFRECID_DEFCOLWIDTH,Round(FCurrXc12Sheet.SheetFormatPr.DefaultColWidth))
end;

procedure TXLSWriteII.WREC_DELTA;
begin
  PDouble(PBuf)^ := FManager.Workbook.CalcPr.IterateDelta;
  WriteBuf(BIFFRECID_DELTA,SizeOf(Double))
end;

procedure TXLSWriteII.WREC_DIMENSIONS;
begin
  if FXLS.Version >= xvExcel97 then begin
    PRecDIMENSIONS8(PBuf).Reserved := 0;

    PRecDIMENSIONS8(PBuf).FirstCol := FCurrXc12Sheet.Dimension.Col1;
    PRecDIMENSIONS8(PBuf).FirstRow := FCurrXc12Sheet.Dimension.Row1;
    if FCurrXc12Sheet.Dimension.Col2 > XLS_MAXCOL_97 then
      PRecDIMENSIONS8(PBuf).LastCol := XLS_MAXCOL_97 + 1
    else
      PRecDIMENSIONS8(PBuf).LastCol := FCurrXc12Sheet.Dimension.Col2 + 1;
    if FCurrXc12Sheet.Dimension.Col2 > XLS_MAXROW_97 then
      PRecDIMENSIONS8(PBuf).LastRow := XLS_MAXROW_97 + 1
    else
      PRecDIMENSIONS8(PBuf).LastRow := FCurrXc12Sheet.Dimension.Row2 + 1;
    WriteBuf(BIFFRECID_DIMENSIONS,SizeOf(TRecDIMENSIONS8));
  end
  else begin
    PRecDIMENSIONS7(PBuf).Reserved := 0;
    PRecDIMENSIONS7(PBuf).FirstCol := FCurrXc12Sheet.Dimension.Col1;
    PRecDIMENSIONS7(PBuf).FirstRow := FCurrXc12Sheet.Dimension.Row1;
    if FCurrXc12Sheet.Dimension.Col2 > XLS_MAXCOL_97 then
      PRecDIMENSIONS7(PBuf).LastCol := XLS_MAXCOL_97 + 1
    else
      PRecDIMENSIONS7(PBuf).LastCol := FCurrXc12Sheet.Dimension.Col2 + 1;
    if FCurrXc12Sheet.Dimension.Col2 > XLS_MAXROW_97 then
      PRecDIMENSIONS7(PBuf).LastRow := XLS_MAXROW_97
    else
      PRecDIMENSIONS7(PBuf).LastRow := FCurrXc12Sheet.Dimension.Row2 + 1;
    WriteBuf(BIFFRECID_DIMENSIONS,SizeOf(TRecDIMENSIONS7));
  end;
end;

procedure TXLSWriteII.WREC_ROW;
var
  Row: PXLSMMURowHeader;
begin
  FManager.Worksheets[FCurrSheet].Cells.BeginIterateRow;
  while FManager.Worksheets[FCurrSheet].Cells.IterateNextRow do begin
    Row := FManager.Worksheets[FCurrSheet].Cells.IterRow;

    PRecROW(PBuf).Options := Row.OutlineLevel;

    PRecROW(PBuf).FormatIndex := CheckDeafultXF(Row.Style);
    if PRecROW(PBuf).FormatIndex <> XLS_STYLE_DEFAULT_XF_97 then
      PRecROW(PBuf).Options := PRecROW(PBuf).Options or $0080;

    PRecROW(PBuf).Row := FManager.Worksheets[FCurrSheet].Cells.IterCellRow;
    PRecROW(PBuf).Height := Row.Height;

    PRecROW(PBuf).Reserved1 := 0;
    PRecROW(PBuf).Reserved2 := 0;

    PRecROW(PBuf).Col1 := 0;
    PRecROW(PBuf).Col2 := FManager.Worksheets[FCurrSheet].Dimension.Col2;

    if xroCollapsed in Row.Options then
      PRecROW(PBuf).Options := PRecROW(PBuf).Options or $0010;
    if xroHidden in Row.Options then
      PRecROW(PBuf).Options := PRecROW(PBuf).Options or $0020;
    if xroCustomHeight in Row.Options then
      PRecROW(PBuf).Options := PRecROW(PBuf).Options or $0040;

    // Undocumented. Must be set.
    PRecROW(PBuf).Options := PRecROW(PBuf).Options or $0100;

    WriteBuf(BIFFRECID_ROW,SizeOf(TRecROW));
  end;
end;

procedure TXLSWriteII.WREC_FOOTER;
var
  L: integer;
  S: AxUCString;
begin
  S := FCurrXc12Sheet.HeaderFooter.FirstFooter;
  if S = '' then
    WriteRecId(BIFFRECID_FOOTER)
  else if FXLS.Version >= xvExcel97 then begin
    if S <> '' then begin
      L := Length(S);
      FXLSStream.WriteHeader(BIFFRECID_FOOTER,2 + 1 + L * 2);
      FXLSStream.WWord(L);
      WideStringToByteStr(S,PBuf);
      FXLSStream.Write(PBuf^,1 + L * 2);
    end
    else
      WriteBuf(BIFFRECID_FOOTER,0);
  end
  else begin
    PRecSTRING1Byte(PBuf).Len := Length(S);
    Move(Pointer(S)^,PRecSTRING1Byte(PBuf).Data,Length(S));
    WriteBuf(BIFFRECID_FOOTER,2 + Length(S));
  end;
end;

procedure TXLSWriteII.WREC_HCENTER;
begin
  WriteBoolean(BIFFRECID_HCENTER,FCurrXc12Sheet.PrintOptions.HorizontalCentered);
end;

procedure TXLSWriteII.WREC_HEADER;
var
  L: integer;
  S: AxUCString;
begin
  S := FCurrXc12Sheet.HeaderFooter.FirstHeader;
  if S = '' then
    WriteRecId(BIFFRECID_HEADER)
  else if FXLS.Version >= xvExcel97 then begin
    if S <> '' then begin
      L := Length(S);
      FXLSStream.WriteHeader(BIFFRECID_HEADER,2 + 1 + L * 2);
      FXLSStream.WWord(L);
      WideStringToByteStr(S,PBuf);
      FXLSStream.Write(PBuf^,1 + L * 2);
    end
    else
      WriteBuf(BIFFRECID_HEADER,0);
  end
  else begin
    PRecSTRING1Byte(PBuf).Len := Length(S);
    Move(Pointer(S)^,PRecSTRING1Byte(PBuf).Data,Length(S));
    WriteBuf(BIFFRECID_HEADER,2 + Length(S));
  end;
end;

procedure TXLSWriteII.WREC_ITERATION;
begin
  WriteBoolean(BIFFRECID_ITERATION,FManager.Workbook.CalcPr.Iterate);
end;

procedure TXLSWriteII.WREC_PRINTGRIDLINES;
begin
  WriteBoolean(BIFFRECID_PRINTGRIDLINES,FCurrXc12Sheet.PrintOptions.GridLines);
end;

procedure TXLSWriteII.WREC_PRINTHEADERS;
begin
  WriteBoolean(BIFFRECID_PRINTHEADERS,FCurrXc12Sheet.PrintOptions.Headings);
end;

procedure TXLSWriteII.WREC_REFMODE;
begin
  WriteBoolean(BIFFRECID_REFMODE,FManager.Workbook.CalcPr.RefMode = x12rmA1);
end;

procedure TXLSWriteII.WREC_MARGINS;
begin
  if FCurrXc12Sheet.PageMargins.Left > 0 then begin
    PRecMARGIN(PBuf).Value := FCurrXc12Sheet.PageMargins.Left;
    WriteBuf(BIFFRECID_LEFTMARGIN,SizeOf(TRecMARGIN));
  end;
  if FCurrXc12Sheet.PageMargins.Top > 0 then begin
    PRecMARGIN(PBuf).Value := FCurrXc12Sheet.PageMargins.Top;
    WriteBuf(BIFFRECID_TOPMARGIN,SizeOf(TRecMARGIN));
  end;
  if FCurrXc12Sheet.PageMargins.Right > 0 then begin
    PRecMARGIN(PBuf).Value := FCurrXc12Sheet.PageMargins.Right;
    WriteBuf(BIFFRECID_RIGHTMARGIN,SizeOf(TRecMARGIN));
  end;
  if FCurrXc12Sheet.PageMargins.Bottom > 0 then begin
    PRecMARGIN(PBuf).Value := FCurrXc12Sheet.PageMargins.Bottom;
    WriteBuf(BIFFRECID_BOTTOMMARGIN,SizeOf(TRecMARGIN));
  end;
end;

procedure TXLSWriteII.WREC_SETUP;
begin
  PRecSETUP(PBuf).PaperSize := FCurrXc12Sheet.PageSetup.PaperSize;
  PRecSETUP(PBuf).Scale := FCurrXc12Sheet.PageSetup.Scale;
  PRecSETUP(PBuf).PageStart := FCurrXc12Sheet.PageSetup.FirstPageNumber;
  PRecSETUP(PBuf).FitWidth := FCurrXc12Sheet.PageSetup.FitToWidth;
  PRecSETUP(PBuf).FitHeight := FCurrXc12Sheet.PageSetup.FitToHeight;
  PRecSETUP(PBuf).Resolution := FCurrXc12Sheet.PageSetup.HorizontalDpi;
  PRecSETUP(PBuf).VertResolution := FCurrXc12Sheet.PageSetup.VerticalDpi;
  PRecSETUP(PBuf).HeaderMargin := FCurrXc12Sheet.PageMargins.Header;
  PRecSETUP(PBuf).FooterMargin := FCurrXc12Sheet.PageMargins.Footer;
  PRecSETUP(PBuf).Copies := FCurrXc12Sheet.PageSetup.Copies;

  PRecSETUP(PBuf).Options := $0000;

  if FCurrXc12Sheet.PageSetup.PageOrder = x12poDownThenOver then
    PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or $0001;
  if FCurrXc12Sheet.PageSetup.Orientation = x12oPortrait then
    PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or $0002;
  if FCurrXc12Sheet.PageSetup.BlackAndWhite then
    PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or $0008;
  if FCurrXc12Sheet.PageSetup.Draft then
    PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or $0010;

  case FCurrXc12Sheet.PageSetup.CellComments of
    x12ccNone       : ;
    x12ccAsDisplayed: PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or $0020;
    x12ccAtEnd      : PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or $0220;
  end;

  PRecSETUP(PBuf).Options := PRecSETUP(PBuf).Options or (Word(FCurrXc12Sheet.PageSetup.Errors) shl 10);

  WriteBuf(BIFFRECID_SETUP,SizeOf(TRecSETUP));
end;

procedure TXLSWriteII.WREC_VCENTER;
begin
  WriteBoolean(BIFFRECID_VCENTER,FCurrXc12Sheet.PrintOptions.VerticalCentered);
end;

procedure TXLSWriteII.WREC_WSBOOL;
var
  W: word;
begin
  W := $0000;

  if FCurrXc12Sheet.SheetPr.PageSetupPr.AutoPageBreaks then
    W := W or $0001;
  if FCurrXc12Sheet.SheetPr.PageSetupPr.FitToPage then
    W := W or $0100;
  if FCurrXc12Sheet.SheetPr.OutlinePr.ApplyStyles then
    W := W or $0020;
  if FCurrXc12Sheet.SheetPr.OutlinePr.SummaryBelow then
    W := W or $0040;
  if FCurrXc12Sheet.SheetPr.OutlinePr.SummaryRight then
    W := W or $0080;

  WriteWord(BIFFRECID_WSBOOL,W);
end;

procedure TXLSWriteII.WREC_HORIZONTALPAGEBREAKS;
var
  i: integer;
  Brk: TXc12Break;
begin
  if FCurrXc12Sheet.RowBreaks.Count > 0 then begin
    PRecHORIZONTALPAGEBREAKS(PBuf).Count := FCurrXc12Sheet.RowBreaks.Count;
    for i := 0 to FCurrXc12Sheet.RowBreaks.Count - 1 do begin
      Brk := FCurrXc12Sheet.RowBreaks[i];
      PRecHORIZONTALPAGEBREAKS(PBuf).Breaks[i].Val1 := Brk.Id;
      PRecHORIZONTALPAGEBREAKS(PBuf).Breaks[i].Val2 := Brk.Min;
      PRecHORIZONTALPAGEBREAKS(PBuf).Breaks[i].Val3 := Brk.Max;
    end;
    WriteBuf(BIFFRECID_HORIZONTALPAGEBREAKS,2 + SizeOf(TPageBreak) * FCurrXc12Sheet.RowBreaks.Count);
  end;
end;

procedure TXLSWriteII.WREC_VERTICALPAGEBREAKS;
var
  i: integer;
  Brk: TXc12Break;
begin
  if FCurrXc12Sheet.ColBreaks.Count > 0 then begin
    PRecVERTICALPAGEBREAKS(PBuf).Count := FCurrXc12Sheet.ColBreaks.Count;
    for i := 0 to FCurrXc12Sheet.ColBreaks.Count - 1 do begin
      Brk := FCurrXc12Sheet.ColBreaks[i];
      PRecVERTICALPAGEBREAKS(PBuf).Breaks[i].Val1 := Brk.Id;
      PRecVERTICALPAGEBREAKS(PBuf).Breaks[i].Val2 := Brk.Min;
      PRecVERTICALPAGEBREAKS(PBuf).Breaks[i].Val3 := Brk.Max;
    end;
    WriteBuf(BIFFRECID_VERTICALPAGEBREAKS,2 + SizeOf(TPageBreak) * FCurrXc12Sheet.ColBreaks.Count);
  end;
end;

// Sheet suffix

procedure TXLSWriteII.WREC_MSODRAWING;
begin
  if FXLS.Version < xvExcel97 then Exit;

  if FXLS.Sheets[FCurrSheet]._Int_EscherDrawing.ShapeCount > 0 then
    FXLS.Sheets[FCurrSheet]._Int_EscherDrawing.SaveToStream(FXLSStream,PBuf);
end;

procedure TXLSWriteII.WREC_NAME;
var
  i: integer;
  P: Pointer;
  Sz: integer;
  Opts: word;
  Name: TXc12DefinedName;
begin
  if (FManager.Workbook.DefinedNames.Count > 0) and not FXLS.FormulaHandler.ExternalNames.HasSUPBOOK then begin
    FExternalNamesWritten := True;
    WREC_SUPBOOK;
    WREC_EXTERNSHEET;
  end;

  for i := 0 to FManager.Workbook.DefinedNames.Count - 1 do begin
    Name := FManager.Workbook.DefinedNames[i];
    Sz := SizeOf(TRecNAME);
    if Name.Ptgs.Id = xptg_EXCEL_97 then
      Inc(Sz,Name.PtgsSz - SizeOf(TXLSPtgs))
    else if Name.Ptgs.Id = xptg_ARRAYCONSTS_97 then begin
      raise XLSRWException.Create('TODO Array consts 97');
    end
    else // Can only write areas.
      Inc(Sz,1 + SizeOf(TPTGArea3d8));

    Inc(Sz,UCStringLenFile(Name.CustomMenu));
    Inc(Sz,UCStringLenFile(Name.Description));
    Inc(Sz,UCStringLenFile(Name.Help));
    Inc(Sz,UCStringLenFile(Name.StatusBar));

    if (Name.BuiltIn <> bnNone) or (Name.Unused97 <> 0) then
      Inc(Sz,2)
    else
      Inc(Sz,UCStringLenFile(Name.Name));

    FXLSStream.WriteHeader(BIFFRECID_NAME,Sz);

    Opts := $0000;
    if Name.Hidden      then Opts := Opts or $0001;
    if Name.Function_   then Opts := Opts or $0002;
    if Name.VbProcedure then Opts := Opts or $0004;
    if (Name.BuiltIn <> bnNone) or (Name.Unused97 <> 0) then
      Opts := Opts or $0020;

    Opts := Opts or (Name.FunctionGroupId shl 5);
    FXLSStream.WWord(Opts);

    FXLSStream.WByte(0);  // KeyShortcut
    if (Name.BuiltIn <> bnNone) or (Name.Unused97 <> 0) then
      FXLSStream.WByte(1)
    else
      FXLSStream.WByte(Length(Name.Name));
    if Name.Ptgs.Id = xptg_EXCEL_97 then
      FXLSStream.WWord(Name.PtgsSz - SizeOf(TXLSPtgs))
    else if Name.Ptgs.Id = xptg_ARRAYCONSTS_97 then begin
      raise XLSRWException.Create('TODO Array consts 97');
    end
    else
      FXLSStream.WWord(1 + SizeOf(TPTGArea3d8));
    FXLSStream.WWord(0);

    if (Name.LocalSheetId >= 0) and (Name.LocalSheetId < $FFFF) then
      FXLSStream.WWord(Name.LocalSheetId + 1)
    else
      FXLSStream.WWord(0);

    FXLSStream.WByte(Length(Name.CustomMenu));
    FXLSStream.WByte(Length(Name.Description));
    FXLSStream.WByte(Length(Name.Help));
    FXLSStream.WByte(Length(Name.StatusBar));

    if Name.BuiltIn <> bnNone then begin
      FXLSStream.WByte(0);
      case Name.BuiltIn of
        bnConsolidateArea: FXLSStream.WByte($00);
        bnExtract        : FXLSStream.WByte($03);
        bnDatabase       : FXLSStream.WByte($04);
        bnCriteria       : FXLSStream.WByte($05);
        bnPrintArea      : FXLSStream.WByte($06);
        bnPrintTitles    : FXLSStream.WByte($07);
        bnSheetTitle     : FXLSStream.WByte($0C);
        // Not correct, but something must be written to keep name index in sync.
        bnFilterDatabase : FXLSStream.WByte($04);
      end;
    end
    else if Name.Unused97 <> 0 then begin
      FXLSStream.WByte(0);
      FXLSStream.WByte(Name.Unused97);
    end
    else
      FXLSStream.WriteWideString(Name.Name,False);

    if Name.PtgsSz > 0 then begin
      if Name.Ptgs.Id = xptg_EXCEL_97 then begin
        P := Pointer(NativeInt(Name.Ptgs) + SizeOf(TXLSPtgs));
        FXLSStream.Write(P^,Name.PtgsSz - SizeOf(TXLSPtgs));
      end
      else if Name.Ptgs.Id = xptg_ARRAYCONSTS_97 then begin
        raise XLSRWException.Create('TODO Array consts 97');
      end
      else begin
        FXLSStream.WByte(xptgArea3d97);
        if Name.SimpleName in [xsntRef,xsntArea] then begin
          PPTGArea3d8(PBuf).Index := FXLS.FormulaHandler.ExternalNames.FindIndex(Name.SimpleArea.SheetIndex);
          PPTGArea3d8(PBuf).Col1 := Name.SimpleArea.Col1 and not COL_ABSFLAG;
          PPTGArea3d8(PBuf).Row1 := Name.SimpleArea.Row1 and not ROW_ABSFLAG;
          PPTGArea3d8(PBuf).Col2 := Name.SimpleArea.Col2 and not COL_ABSFLAG;
          PPTGArea3d8(PBuf).Row2 := Name.SimpleArea.Row2 and not ROW_ABSFLAG;
        end
        else begin
          // Write a faked name for complicated ptgs
          PPTGArea3d8(PBuf).Index := Name.SimpleArea.SheetIndex;
          PPTGArea3d8(PBuf).Col1 := 0;
          PPTGArea3d8(PBuf).Row1 := 0;
          PPTGArea3d8(PBuf).Col2 := 0;
          PPTGArea3d8(PBuf).Row2 := 0;
        end;
        FXLSStream.Write(PBuf^,SizeOf(TPTGArea3d8));
      end;
    end;
    if Name.CustomMenu  <> '' then FXLSStream.WriteWideString(Name.CustomMenu,False);
    if Name.Description <> '' then FXLSStream.WriteWideString(Name.Description,False);
    if Name.Help        <> '' then FXLSStream.WriteWideString(Name.Help,False);
    if Name.StatusBar   <> '' then FXLSStream.WriteWideString(Name.StatusBar,False);
  end;
end;

procedure TXLSWriteII.WREC_NOTE;
begin
//  Written by escher
end;

procedure TXLSWriteII.WREC_SELECTION;
var
  i,j: integer;
  Sel: TXc12Selection;
begin
  if FCurrXc12Sheet.SheetViews[0].Selection.Count > 0 then begin
    for i := 0 to FCurrXc12Sheet.SheetViews[0].Selection.Count - 1 do begin
      Sel := FCurrXc12Sheet.SheetViews[0].Selection[i];
      PRecSELECTION_2(PBuf).Pane := Word(Sel.Pane);
      PRecSELECTION_2(PBuf).ActiveCol := Sel.ActiveCell.Col1;
      PRecSELECTION_2(PBuf).ActiveRow := Sel.ActiveCell.Row1;
      PRecSELECTION_2(PBuf).ActiveRef := Sel.ActiveCellId;
      PRecSELECTION_2(PBuf).RefCount := Sel.SQRef.Count;

      FXLSStream.WriteHeader(BIFFRECID_SELECTION,SizeOf(TRecSELECTION_2) + (Sel.SQRef.Count * 6));
      FXLSStream.Write(PBuf^,SizeOf(TRecSELECTION_2));
      for j := 0 to Sel.SQRef.Count - 1 do begin
        FXLSStream.WWord(Sel.SQRef[j].Row1);
        FXLSStream.WWord(Sel.SQRef[j].Row2);
        FXLSStream.WByte(Sel.SQRef[j].Col1);
        FXLSStream.WByte(Sel.SQRef[j].Col2);
      end;
    end;
  end
  else begin
    PRecSELECTION_2(PBuf).Pane := 3;
    PRecSELECTION_2(PBuf).ActiveCol := 0;
    PRecSELECTION_2(PBuf).ActiveRow := 0;
    PRecSELECTION_2(PBuf).ActiveRef := 0;
    PRecSELECTION_2(PBuf).RefCount := 1;

    FXLSStream.WriteHeader(BIFFRECID_SELECTION,SizeOf(TRecSELECTION_2) + 6);
    FXLSStream.Write(PBuf^,SizeOf(TRecSELECTION_2));

    FXLSStream.WWord(0);
    FXLSStream.WWord(0);
    FXLSStream.WByte(0);
    FXLSStream.WByte(0);
  end;
end;

procedure TXLSWriteII.WREC_FEATHEADR;
var
  W: longword;
begin
  if FXLS.Sheets[FCurrSheet].SheetProtection <> DefaultSheetProtections then begin
    PRecFEATHEADR(PBuf).Frt := BIFFRECID_FEATHEADR;
    PRecFEATHEADR(PBuf).Flags := 0;
    PRecFEATHEADR(PBuf).Unused1[0] := 0;
    PRecFEATHEADR(PBuf).Unused1[1] := 0;
    PRecFEATHEADR(PBuf).Unused1[2] := 0;
    PRecFEATHEADR(PBuf).Unused1[3] := 0;
    PRecFEATHEADR(PBuf).Unused1[4] := 0;
    PRecFEATHEADR(PBuf).Unused1[5] := 0;
    PRecFEATHEADR(PBuf).Unused1[6] := 0;
    PRecFEATHEADR(PBuf).Unused1[7] := 0;
    PRecFEATHEADR(PBuf).SharedFeatureType := 2;
    PRecFEATHEADR(PBuf).Hdr := $01;
    PRecFEATHEADR(PBuf).HdrDataSz := $FFFFFFFF;

    W := $00000000;

    if FCurrXc12Sheet.SheetProtection.Objects             then W := W or $00000001;
    if FCurrXc12Sheet.SheetProtection.Scenarios           then W := W or $00000002;
    if FCurrXc12Sheet.SheetProtection.FormatCells         then W := W or $00000004;
    if FCurrXc12Sheet.SheetProtection.FormatColumns       then W := W or $00000008;
    if FCurrXc12Sheet.SheetProtection.FormatRows          then W := W or $00000010;
    if FCurrXc12Sheet.SheetProtection.InsertColumns       then W := W or $00000020;
    if FCurrXc12Sheet.SheetProtection.InsertRows          then W := W or $00000040;
    if FCurrXc12Sheet.SheetProtection.InsertHyperlinks    then W := W or $00000080;
    if FCurrXc12Sheet.SheetProtection.DeleteColumns       then W := W or $00000100;
    if FCurrXc12Sheet.SheetProtection.DeleteRows          then W := W or $00000200;
    if FCurrXc12Sheet.SheetProtection.SelectLockedCells   then W := W or $00000400;
    if FCurrXc12Sheet.SheetProtection.Sort                then W := W or $00000800;
    if FCurrXc12Sheet.SheetProtection.AutoFilter          then W := W or $00001000;
    if FCurrXc12Sheet.SheetProtection.PivotTables         then W := W or $00002000;
    if FCurrXc12Sheet.SheetProtection.SelectUnlockedCells then W := W or $00004000;

    PRecFEATHEADR(PBuf).Protections := W;

    WriteBuf(BIFFRECID_FEATHEADR,SizeOf(TRecFEATHEADR));
  end;
end;

procedure TXLSWriteII.WREC_DVAL;
begin
  if FXLS.Sheets[FCurrSheet].Validations.Count > 0 then
    FXLS.Sheets[FCurrSheet].Validations.SaveToStream(FXLSStream,PBuf);
end;

procedure TXLSWriteII.WREC_WINDOW2;
begin
  PRecWINDOW2_7(PBuf).Options := $0020 + $0080 + $0400;

  if FXLS.Version <= xvExcel50 then begin
    PRecWINDOW2_7(PBuf).TopRow := 0;
    PRecWINDOW2_7(PBuf).LeftCol := 0;
    PRecWINDOW2_7(PBuf).HeaderColorIndex := 0;
    WriteBuf(BIFFRECID_WINDOW2,SizeOf(TRecWINDOW2_7));
  end;
  if FXLS.Version > xvExcel40 then begin
    PRecWINDOW2_8(PBuf).LeftCol := 0; //FCurrXc12Sheet.SheetViews[0].TopLeftCell.Col1;
    PRecWINDOW2_8(PBuf).TopRow := 0; //FCurrXc12Sheet.SheetViews[0].TopLeftCell.Row1;

    PRecWINDOW2_8(PBuf).HeaderColorIndex := FCurrXc12Sheet.SheetViews[0].ColorId;
//    PRecWINDOW2_8(PBuf).Zoom := FCurrXc12Sheet.SheetViews[0].ZoomScaleNormal;
    PRecWINDOW2_8(PBuf).Zoom := FCurrXc12Sheet.SheetViews[0].ZoomScale;
    PRecWINDOW2_8(PBuf).ZoomPreview := FCurrXc12Sheet.SheetViews[0].ZoomScalePageLayoutView;

    PRecWINDOW2_8(PBuf).Options := $0020 + $0080 + $0400;
    if FCurrXc12Sheet.SheetViews[0].ShowFormulas then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0001;
    if FCurrXc12Sheet.SheetViews[0].ShowGridLines then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0002;
    if FCurrXc12Sheet.SheetViews[0].ShowRowColHeaders then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0004;
    if FCurrXc12Sheet.SheetViews[0].ShowZeros then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0010;
    if FCurrXc12Sheet.SheetViews[0].DefaultGridColor then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0020;
    if FCurrXc12Sheet.SheetViews[0].RightToLeft then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0040;
    if FCurrXc12Sheet.SheetViews[0].ShowOutlineSymbols then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0080;
    if FCurrXc12Sheet.SheetViews[0].TabSelected then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0200;

    if FCurrXc12Sheet.SheetViews[0].Pane.Excel97 and (FCurrXc12Sheet.SheetViews[0].Pane.State = x12psFrozen) then begin
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0008;
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0100;
    end;

    if (FCurrXc12Sheet.CustomSheetViews.Count > 0) and (FCurrXc12Sheet.CustomSheetViews[0].View = x12svtPageBreakPreview) then
      PRecWINDOW2_8(PBuf).Options := PRecWINDOW2_8(PBuf).Options or $0800;

    WriteBuf(BIFFRECID_WINDOW2,SizeOf(TRecWINDOW2_8));
  end;

  WREC_SCL;
end;

procedure TXLSWriteII.WREC_SAVERECALC;
begin
  WriteBoolean(BIFFRECID_SAVERECALC,FManager.Workbook.CalcPr.CalcOnSave);
end;

procedure TXLSWriteII.WREC_SCL;
begin
  if (FCurrXc12Sheet.SheetViews[0].ZoomScale > 0) and (FCurrXc12Sheet.SheetViews[0].ZoomScale <> 100) then begin
    PRecSCL(PBuf).Numerator := FCurrXc12Sheet.SheetViews[0].ZoomScale;
    PRecSCL(PBuf).Denominator := 100;
    WriteBuf(BIFFRECID_SCL,SizeOf(TRecSCL));
  end;
end;

procedure TXLSWriteII.WREC_PANE;
begin
  if FCurrXc12Sheet.SheetViews[0].Pane.Excel97 then begin
    PRecPANE(PBuf).PaneNumber := Word(FCurrXc12Sheet.SheetViews[0].Pane.ActivePane);
    if FCurrXc12Sheet.SheetViews[0].Pane.State = x12psFrozenSplit then begin
      PRecPANE(PBuf).X := Round(FCurrXc12Sheet.SheetViews[0].Pane.XSplit * 20);
      PRecPANE(PBuf).Y := Round(FCurrXc12Sheet.SheetViews[0].Pane.YSplit * 20);
    end
    else begin
      PRecPANE(PBuf).X := FCurrXc12Sheet.SheetViews[0].Pane.TopLeftCell.Col1;
      PRecPANE(PBuf).Y := FCurrXc12Sheet.SheetViews[0].Pane.TopLeftCell.Row1;
    end;
    PRecPANE(PBuf).LeftCol := FCurrXc12Sheet.SheetViews[0].Pane.TopLeftCell.Col1;
    PRecPANE(PBuf).TopRow := FCurrXc12Sheet.SheetViews[0].Pane.TopLeftCell.Row1;

    WriteBuf(BIFFRECID_PANE,SizeOf(TRecPANE));
  end;
end;

procedure TXLSWriteII.WREC_HLINK;
const
  UnSequence: array[0..23] of byte = ($FF,$FF,$AD,$DE,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00,$00);
var
  i,Sz: integer;
  DirUpCnt: word;
//  S: XLS8String;
  Dos8_3: XLS8String;
  HLink: TXc12Hyperlink;
  Options: word;

function ZUCLen(WS: AxUCString): integer;
begin
  if WS <> '' then
    Result := Length(WS) * 2 + 2
  else
    Result := 0;
end;

procedure WriteUSZLen(WS: AxUCString; CharLen: boolean = True);
var
  L: longword;
  Z: word;
begin
  Z := 0;
  if CharLen then
    L := Length(WS) + 1
  else
    L := Length(WS) * 2 + 2;
  FXLSStream.Write(L,4);
  FXLSStream.Write(Pointer(WS)^,Length(WS) * 2);
  FXLSStream.WWord(Z);
end;

begin
  if FXLS.Version < xvExcel97 then
    Exit;

  for i := 0 to FCurrXc12Sheet.Hyperlinks.Count - 1 do begin
    HLink := FCurrXc12Sheet.Hyperlinks[i];
//    if HLink.HyperlinkType = xhltUnknown then begin
//      FXLSStream.WriteHeader(BIFFRECID_HLINK,HLink.FBufLen);
//      FXLSStream.Write(HLink.FBuf^,HLink.FBufLen);
//      Continue;
//    end;

//    if HLink.FChanged then begin
      Options := $0000;
      case HLink.HyperlinkEnc of
        xhltURL: begin
          Options := Options and not $0160;
          Options := Options or $0003;
        end;
        xhltFile: begin
          Options := Options and not $0160;
          Options := Options or $0001;
        end;
        xhltUNC: begin
          Options := Options and not $0060;
          Options := Options or $0103;
        end;
        xhltWorkbook: begin
          Options := Options and not $0163;
          Options := Options or $0008;
        end;
      end;
      if HLink.Display <> '' then
        Options := Options or $0014
      else
        Options := Options and not $0014;
//    end;

    Sz := SizeOf(TRecHLINK);

    if HLink.Display <> '' then
      Inc(Sz,4 + ZUCLen(HLink.Display));
    if HLink.TargetFrame <> '' then
      Inc(Sz,4 + ZUCLen(HLink.TargetFrame));
    case HLink.HyperlinkEnc of
      xhltURL:
        Inc(Sz,16 + 4 + ZUCLen(HLink.RawAddress));
      xhltFile: begin
//        S := XLS8String(HLink.RawAddress);
//        // The DOS file path will only be included if GetShortPathName can find
//        // the disk file.
//        SetLength(Dos8_3,255);
//        SetLength(Dos8_3,GetShortPathName(PChar(S),PChar(Dos8_3),255));
//        if Dos8_3 <> '' then
//          Inc(Sz,16 + 2 + 4 + (Length(Dos8_3) + 1) + 24 + 4)
//        else
          Inc(Sz,16 + 2 + 4 + 24 + 4);
        if HLink.RawAddress <> '' then
          Inc(Sz,4 + 2 + Length(HLink.RawAddress) * 2);
      end;
      xhltUNC:
        Inc(Sz,4 + ZUCLen(HLink.RawAddress));
      xhltWorkbook: ;
    end;
    if HLink.ScreenTip <> '' then
      Inc(Sz,4 + ZUCLen(HLink.ScreenTip));

    FXLSStream.WriteHeader(BIFFRECID_HLINK,Sz);
    PRecHLINK(PBuf).Row1 := HLink.Row1;
    PRecHLINK(PBuf).Row2 := HLink.Row2;
    PRecHLINK(PBuf).Col1 := HLink.Col1;
    PRecHLINK(PBuf).Col2 := HLink.Col2;
    Move(GUID_HLINK_STDLINK,PRecHLINK(PBuf).GUID,Length(GUID_HLINK_STDLINK));
    PRecHLINK(PBuf).Reserved := $00000002;
    PRecHLINK(PBuf).Options := Options;
    FXLSStream.Write(PBuf^,SizeOf(TRecHLINK));
    if HLink.Display <> '' then
      WriteUSZLen(HLink.Display);
    if HLink.TargetFrame <> '' then
      WriteUSZLen(HLink.TargetFrame);
    case HLink.HyperlinkEnc of
      xhltURL: begin
        FXLSStream.Write(GUID_HLINK_URL,Length(GUID_HLINK_URL));
        WriteUSZLen(HLink.RawAddress,False);
      end;
      xhltFile: begin
        FXLSStream.Write(GUID_HLINK_FILE,Length(GUID_HLINK_FILE));
        DirUpCnt := 0;
        while Copy(HLink.RawAddress,1,3) = '..\' do begin
          Inc(DirUpCnt);
          HLink.RawAddress := Copy(HLink.RawAddress,4,MAXINT);
        end;
        FXLSStream.WWord(DirUpCnt);
        if Dos8_3 <> '' then begin
          FXLSStream.WLWord(Length(Dos8_3) + 1);
          FXLSStream.Write(Pointer(Dos8_3)^,Length(Dos8_3));
          FXLSStream.WByte(0);
        end
        else
          FXLSStream.WLWord(0);
        FXLSStream.Write(UnSequence[0],Length(UnSequence));
        if HLink.RawAddress <> '' then begin
          FXLSStream.WLWord(4 + 2 + Length(HLink.RawAddress) * 2);
          FXLSStream.WLWord(Length(HLink.RawAddress) * 2);
          FXLSStream.WByte($03);
          FXLSStream.WByte($00);
          FXLSStream.Write(Pointer(HLink.RawAddress)^,Length(HLink.RawAddress) * 2);
        end
        else
          FXLSStream.WLWord(0);
      end;
      xhltUNC: begin
        WriteUSZLen(HLink.RawAddress);
      end;
      xhltWorkbook: begin
      end;
    end;
    if HLink.ScreenTip <> '' then
      WriteUSZLen(HLink.ScreenTip);

    if HLink.ToolTip <> '' then begin
      FXLSStream.WriteHeader(BIFFRECID_HLINKTOOLTIP,SizeOf(TRecHLINKTOOLTIP) + Length(HLink.ToolTip) * 2 + 2);
      PRecHLINKTOOLTIP(PBuf).RecId := BIFFRECID_HLINKTOOLTIP;
      PRecHLINKTOOLTIP(PBuf).Row1 := HLink.Row1;
      PRecHLINKTOOLTIP(PBuf).Row2 := HLink.Row2;
      PRecHLINKTOOLTIP(PBuf).Col1 := HLink.Col1;
      PRecHLINKTOOLTIP(PBuf).Col2 := HLink.Col2;
      FXLSStream.Write(PBuf^,SizeOf(TRecHLINKTOOLTIP));
      FXLSStream.Write(Pointer(HLink.ToolTip)^,Length(HLink.ToolTip) * 2);
      FXLSStream.WWord(0);
    end;

  end;
//  FXLS.Sheets[FCurrSheet].Hyperlinks.SaveToStream(FXLSStream,PBuf);
end;

procedure TXLSWriteII.WREC_MERGECELLS;
var
  i,j: integer;
  MaxCount: integer;
  Xc12Sheet: TXc12DataWorksheet;
begin
  Xc12Sheet := FManager.Worksheets[FCurrSheet];
  if Xc12Sheet.MergedCells.Count > 0 then begin
    MaxCount := (MAXRECSZ_97 - 2) div 8;
    if Xc12Sheet.MergedCells.Count < MaxCount then
      MaxCount := Xc12Sheet.MergedCells.Count;
    j := 0;
    for i := 0 to Xc12Sheet.MergedCells.Count - 1 do begin
      if (Xc12Sheet.MergedCells[i].Col2 <= MAXCOL_97) and (Xc12Sheet.MergedCells[i].Row2 <= MAXROW_97) then begin
        PRecMERGEDCELLS(PBuf).Cells[j].Row1 := Xc12Sheet.MergedCells[i].Row1;
        PRecMERGEDCELLS(PBuf).Cells[j].Row2 := Xc12Sheet.MergedCells[i].Row2;
        PRecMERGEDCELLS(PBuf).Cells[j].Col1 := Xc12Sheet.MergedCells[i].Col1;
        PRecMERGEDCELLS(PBuf).Cells[j].Col2 := Xc12Sheet.MergedCells[i].Col2;
        Inc(j);
        if j >= MaxCount then begin
          PRecMERGEDCELLS(PBuf).Count := j;
          FXLSStream.WriteHeader(BIFFRECID_MERGEDCELLS,2 + j * 8);
          FXLSStream.Write(PBuf^,2 + j * 8);
          j := 0;
        end;
      end;
    end;
    if j > 0 then begin
      PRecMERGEDCELLS(PBuf).Count := j;
      FXLSStream.WriteHeader(BIFFRECID_MERGEDCELLS,2 + j * 8);
      FXLSStream.Write(PBuf^,2 + j * 8);
    end;
  end;
end;

procedure TXLSWriteII.WREC_CONDFMT;
begin
  FXLS.Sheets[FCurrSheet].ConditionalFormats.SaveToStream(FXLSStream,PBuf);
end;

{ TBoundsheetList }

procedure TBoundsheetList.AddChart(Index: integer; Name: AxUCString);
var
  BD: TBoundsheetData;
begin
  BD := TBoundsheetData.Create;
  BD.FBoundsheetType := btChart;
  BD.FIndex := Index;
  BD.FName := Name;
  inherited Add(BD); 
end;

procedure TBoundsheetList.AddSheet(Index: integer; Name: AxUCString; Hidden: THiddenState; WorksheetType: TWorksheetType);
var
  BD: TBoundsheetData;
begin
  BD := TBoundsheetData.Create;
  case WorksheetType of
    wtSheet:       BD.FBoundsheetType := btSheet;
    wtVBModule:    BD.FBoundsheetType := btVBModule;
    wtExcel4Macro: BD.FBoundsheetType := btExcel4Macro;
  end;
  BD.FIndex := Index;
  BD.FName := Name;
  BD.FHidden := Hidden;
  inherited Add(BD); 
end;

function TBoundsheetList.Find(AName: AxUCString): integer;
begin
  for Result := 0 to Count - 1 do begin
    if Items[Result].Name = AName then
      Exit;
  end;
  Result := -1;
end;

function TBoundsheetList.GetCharts(Index: integer): TBoundsheetData;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].FBoundsheetType = btChart) and (Items[i].FIndex = Index) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TBoundsheetList.GetItems(Index: integer): TBoundsheetData;
begin
  Result := TBoundsheetData(inherited Items[Index]);
end;

function TBoundsheetList.GetSheets(Index: integer): TBoundsheetData;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if (Items[i].FBoundsheetType in [btSheet,btVBModule,btExcel4Macro]) and (Items[i].FIndex = Index) then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

{ TBoundsheetData }

procedure TBoundsheetData.WritePos(Stream: TXLSStream);
var
  Pos: longint;
begin
  Pos := Stream.Pos;
  Stream.Seek(FFilePos,0);
  Stream.Write(Pos,SizeOf(longint));
  Stream.Seek(0,2);
end;

{ THelperWriteXF }

constructor THelperWriteXF.Create(AManager: TXc12Manager; AStream: TXLSStream; ABuf: PByteArray);
begin
  FManager := AManager;
  FStream := AStream;

  FXF := PRecXF8(ABuf);
end;

procedure THelperWriteXF.SetBorderBottomColor(const Value: TXc12IndexColor);
begin
  FXF.Data6 := FXF.Data6 and (not $00003F80);
  FXF.Data6 := FXF.Data6 + (Longword(Value) shl 7);
end;

procedure THelperWriteXF.SetBorderBottomStyle(const Value: TXc12CellBorderStyle);
begin
  FXF.Data4 := FXF.Data4 and (not $F000);
  FXF.Data4 := FXF.Data4 + (Word(Value) shl 12);
  if (FXF.Data4 <> 0) or ((FXF.Data6 and $01E00000) <> 0) then
    FXF.Data3 := FXF.Data3 or $2000
  else
    FXF.Data3 := FXF.Data3 and (not $2000);
end;

procedure THelperWriteXF.SetBorderDiagColor(const Value: TXc12IndexColor);
begin
  FXF.Data6 := FXF.Data6 and (not $001FC000);
  FXF.Data6 := FXF.Data6 + (Longword(Value) shl 14);
end;

procedure THelperWriteXF.SetBorderDiagLines(const Value: TDiagLines);
begin
  FXF.Data5 := FXF.Data5 and (not $C000);
  FXF.Data5 := FXF.Data5 + (Word(Value) shl 14);
end;

procedure THelperWriteXF.SetBorderDiagStyle(const Value: TXc12CellBorderStyle);
begin
  FXF.Data6 := FXF.Data6 and (not $01E00000);
  FXF.Data6 := FXF.Data6 + (Longword(Value) shl 21);
  if (FXF.Data4 <> 0) or ((FXF.Data6 and $01E00000) <> 0) then
    FXF.Data3 := FXF.Data3 or $2000
  else
    FXF.Data3 := FXF.Data3 and (not $2000);
end;

procedure THelperWriteXF.SetBorderLeftColor(const Value: TXc12IndexColor);
begin
  FXF.Data5 := FXF.Data5 and (not $007F);
  FXF.Data5 := FXF.Data5 + (Word(Value) shl 0);
end;

procedure THelperWriteXF.SetBorderLeftStyle(const Value: TXc12CellBorderStyle);
begin
  FXF.Data4 := FXF.Data4 and (not $000F);
  FXF.Data4 := FXF.Data4 + (Word(Value) shl 0);
  if (FXF.Data4 <> 0) or ((FXF.Data6 and $01E00000) <> 0) then
    FXF.Data3 := FXF.Data3 or $2000
  else
    FXF.Data3 := FXF.Data3 and (not $2000);
end;

procedure THelperWriteXF.SetBorderRightColor(const Value: TXc12IndexColor);
begin
  FXF.Data5 := FXF.Data5 and (not $3F80);
  FXF.Data5 := FXF.Data5 + (Word(Value) shl 7);
end;

procedure THelperWriteXF.SetBorderRightStyle(const Value: TXc12CellBorderStyle);
begin
  FXF.Data4 := FXF.Data4 and (not $00F0);
  FXF.Data4 := FXF.Data4 + (Word(Value) shl 4);
  if (FXF.Data4 <> 0) or ((FXF.Data6 and $01E00000) <> 0) then
    FXF.Data3 := FXF.Data3 or $2000
  else
    FXF.Data3 := FXF.Data3 and (not $2000);
end;

procedure THelperWriteXF.SetBorderTopColor(const Value: TXc12IndexColor);
begin
  FXF.Data6 := FXF.Data6 and (not $0000007F);
  FXF.Data6 := FXF.Data6 + (Longword(Value) shl 0);
end;

procedure THelperWriteXF.SetBorderTopStyle(const Value: TXc12CellBorderStyle);
begin
  FXF.Data4 := FXF.Data4 and (not $0F00);
  FXF.Data4 := FXF.Data4 + (Word(Value) shl 8);
  if (FXF.Data4 <> 0) or ((FXF.Data6 and $01E00000) <> 0) then
    FXF.Data3 := FXF.Data3 or $2000
  else
    FXF.Data3 := FXF.Data3 and (not $2000);
end;

procedure THelperWriteXF.SetFFormatOptions(const Value: TXc12AlignmentOptions);
begin
  FXF.Data2 := FXF.Data2 and (not ($0008 + $0010 + $0070));
  if foWrapText in Value then
    FXF.Data2 := FXF.Data2 + $0008;
  if foShrinkToFit in Value then
    FXF.Data3 := FXF.Data3 + $0010;
  if ((FXF.Data2 and $0008) <> 0) or ((FXF.Data2 and $0007) <> 0) then
    FXF.Data3 := FXF.Data3 or $1000
  else
    FXF.Data3 := FXF.Data3 and (not $1000);
end;

procedure THelperWriteXF.SetFillPatternBackColor(const Value: TXc12IndexColor);
begin
  FXF.Data7 := FXF.Data7 and (not $3F80);
  FXF.Data7 := FXF.Data7 + (Word(Value) shl 7);
  if (FXF.Data7 and ($007F + $3F80)) <> $2040 then
    FXF.Data3 := FXF.Data3 or $4000
  else
    FXF.Data3 := FXF.Data3 and (not $4000);
end;

procedure THelperWriteXF.SetFillPatternForeColor(const Value: TXc12IndexColor);
begin
  FXF.Data7 := FXF.Data7 and (not $007F);
  FXF.Data7 := FXF.Data7 + (Word(Value) shl 0);
  if (FXF.Data7 and ($007F + $3F80)) <> $2040 then
    FXF.Data3 := FXF.Data3 or $4000
  else
    FXF.Data3 := FXF.Data3 and (not $4000);
//  if Value <> xcAutomatic then begin
//    SetFillPatternPattern(efpSolid);
//    SetFillPatternBackColor(xcWhite);
//  end;
end;

procedure THelperWriteXF.SetFillPatternPattern(const Value: TXc12FillPattern);
begin
  FXF.Data6 := FXF.Data6 and (not $FC000000);
  FXF.Data6 := FXF.Data6 + ((Longword(Value) shl 26) and $FC000000);
  if Value <> efpNone then
    FXF.Data3 := FXF.Data3 or $4000
  else
    FXF.Data3 := FXF.Data3 and (not $4000);
end;

procedure THelperWriteXF.SetHorizAlignment(const Value: TXc12HorizAlignment);
begin
  FXF.Data2 := FXF.Data2 and (not $0007);
  FXF.Data2 := FXF.Data2 + (Word(Value) shl 0);
  if ((FXF.Data2 and $0008) <> 0) or ((FXF.Data2 and $0007) <> 0) or ((FXF.Data2 and $0070) <> 0) then
    FXF.Data3 := FXF.Data3 or $1000
  else
    FXF.Data3 := FXF.Data3 and (not $1000);
end;

procedure THelperWriteXF.SetIndent(const Value: byte);
begin
  FXF.Data3 := FXF.Data3 and (not $000F);
  FXF.Data3 := FXF.Data3 + (Word(Value and $0F) and $000F);
end;

procedure THelperWriteXF.SetMerged(const Value: boolean);
begin
  FXF.Data3 := FXF.Data3 and (not $0020);
  if Value then
    FXF.Data3 := FXF.Data3 + $0020;
end;

procedure THelperWriteXF.SetProtection(const Value: TXc12CellProtections);
begin
  FXF.Data1 := FXF.Data1 and (not $0003);
  if cpLocked in Value then
    FXF.Data1 := FXF.Data1 + $0001;
  if cpHidden in Value then
    FXF.Data1 := FXF.Data1 + $0002;
end;

procedure THelperWriteXF.SetRotation(const Value: smallint);
var
  V: byte;
begin
//  if Value >= 255 then
//    V := 255
//  else if Value > 90 then
//    V := 90
//  else if Value < -90 then
//    V := 180
//  else if Value < 0 then
//    V := -Value + 90
//  else
    V := Value;
  FXF.Data2 := (FXF.Data2 and $00FF) + (V shl 8);
end;

procedure THelperWriteXF.SetVertAlignment(const Value: TXc12VertAlignment);
begin
  FXF.Data2 := FXF.Data2 and (not $0070);
  FXF.Data2 := FXF.Data2 + (Word(Value) shl 4);
  if ((FXF.Data2 and $0008) <> 0) or ((FXF.Data2 and $0007) <> 0) or ((FXF.Data2 and $0070) <> 0) then
    FXF.Data3 := FXF.Data3 or $1000
  else
    FXF.Data3 := FXF.Data3 and (not $1000);
end;

procedure THelperWriteXF.Write;
var
  i: integer;
  Xc12XF: TXc12XF;
  Diag: TDiagLines;
begin
  for i := XLS_STYLE_DEFAULT_XF_COUNT to FManager.StyleSheet.XFs.Count - 1 do begin
    Xc12XF := FManager.StyleSheet.XFs[i];

    // XF is deleted, write default.
    if Xc12XF = Nil then
      Xc12XF := FManager.StyleSheet.XFs[XLS_STYLE_DEFAULT_XF];

    if Xc12XF.Font.Index > 0 then
      FXF.FontIndex := Xc12XF.Font.Index - XLS_STYLE_DEFAULT_FONT_COUNT
    else
      FXF.FontIndex := Xc12XF.Font.Index;

    FXF.NumFmtIndex := Xc12XF.NumFmt.Index;

    FXF.Data1 := Xc12XF.E97XfData1;
//    FXF.Data2 := Xc12XF.E97XfData2;
//    FXF.Data3 := Xc12XF.E97XfData3;
//    FXF.Data4 := Xc12XF.E97XfData4;
//    FXF.Data5 := Xc12XF.E97XfData5;
//    FXF.Data6 := Xc12XF.E97XfData6;
//    FXF.Data7 := Xc12XF.E97XfData7;

    FXF.Data2 := $0020;
    FXF.Data3 := $0000;
    FXF.Data4 := $0000;
    FXF.Data5 := $2040;
    FXF.Data6 := $00102040;
    FXF.Data7 := $20C0;

    SetIndent(Xc12XF.Alignment.Indent);
    SetRotation(Xc12XF.Alignment.Rotation);
    SetFFormatOptions(Xc12XF.Alignment.Options);
    SetHorizAlignment(Xc12XF.Alignment.HorizAlignment);
    SetProtection(Xc12XF.Protection);
    SetVertAlignment(Xc12XF.Alignment.VertAlignment);
    SetMerged(False); // Not in Xc12
    SetBorderBottomColor(Xc12ColorToIndex(Xc12XF.Border.Bottom.Color));
    SetBorderBottomStyle(Xc12XF.Border.Bottom.Style);
    SetBorderDiagColor(Xc12ColorToIndex(Xc12XF.Border.Diagonal.Color));

    if (ecboDiagonalDown in Xc12XF.Border.Options) and (ecboDiagonalUp in Xc12XF.Border.Options) then
      Diag := dlBoth
    else if ecboDiagonalDown in Xc12XF.Border.Options then
      Diag := dlDown
    else if ecboDiagonalUp in Xc12XF.Border.Options then
      Diag := dlUp
    else
      Diag := dlNone;
    SetBorderDiagLines(Diag);

    SetBorderDiagStyle(Xc12XF.Border.Diagonal.Style);
    SetBorderLeftColor(Xc12ColorToIndex(Xc12XF.Border.Left.Color));
    SetBorderLeftStyle(Xc12XF.Border.Left.Style);
    SetBorderRightColor(Xc12ColorToIndex(Xc12XF.Border.Right.Color));
    SetBorderRightStyle(Xc12XF.Border.Right.Style);
    SetBorderTopColor(Xc12ColorToIndex(Xc12XF.Border.Top.Color));
    SetBorderTopStyle(Xc12XF.Border.Top.Style);

    SetFillPatternPattern(Xc12XF.Fill.PatternType);

    SetFillPatternBackColor(Xc12ColorToIndex(Xc12XF.Fill.BgColor));
    SetFillPatternForeColor(Xc12ColorToIndex(Xc12XF.Fill.FgColor));

    FStream.WriteHeader(BIFFRECID_XF,SizeOf(TRecXF8));
    FStream.Write(FXF^,SizeOf(TRecXF8));
  end;
end;

{ THelperWriteSST }

constructor THelperWriteSST.Create(AManager: TXc12Manager; AStream: TXLSStream; ABuf: PByteArray);
begin
  FManager := AManager;
  FStream := AStream;
  FBuf := ABuf;

  FExtSST := THelperWriteExtSST.Create;
end;

destructor THelperWriteSST.Destroy;
begin
  FExtSST.Free;
  inherited;
end;

procedure THelperWriteSST.Write;
var
  i: integer;
  W: word;
  SSTRecPos: integer;
begin
  FCurrIndex := 0;

  FTotalCount := FManager.SST.TotalCount;
  FExtSST.StringCount := FTotalCount;

  FRecSize := MAXRECSZ_97;
  FStream.WriteHeader(BIFFRECID_SST,0);
  FRecPos := FStream.Pos - 2;
  SSTRecPos := FStream.Pos;
  Dec(FRecSize,FStream.Write(FTotalCount,4));
  // Is replaced later.
  Dec(FRecSize,FStream.Write(FTotalCount,4));

  FCurrIndex := 0;

  for i := 0 to FManager.SST.TotalCount - 1 do
    WriteSSTString(FManager.SST[i]);
//  Traverse(Pointer(Stream),TreeWriteStr);

  FStream.Seek(FRecPos,soFromBeginning);
  W := MAXRECSZ_97 - FRecSize;
  FStream.Write(W,2);

  FStream.Seek(SSTRecPos + 4,soFromBeginning);
  FStream.Write(FCurrIndex,4);

  FStream.Seek(0,soFromEnd);
  FExtSST.Write(FStream);
  FExtSST.Clear;
end;

procedure THelperWriteSST.WriteCONTINUE;
var
  W: word;
begin
  FStream.Seek(FRecPos,soFromBeginning);
  W := MAXRECSZ_97 - FRecSize;
  FStream.Write(W,2);
  FStream.Seek(0,soFromEnd);
  FStream.WriteHeader(BIFFRECID_CONTINUE,0);
  FRecPos := FStream.Pos - 2;
  FRecSize := MAXRECSZ_97;
end;

procedure THelperWriteSST.WriteData(P: PByteArray; Len: word);
begin
  if Len > FRecSize then begin
    FStream.Write(P^,FRecSize);
    Dec(Len,FRecSize);
    P := PByteArray(NativeInt(P) + FRecSize);
    FRecSize := 0;
    WriteCONTINUE;
    WriteData(P,Len);
  end
  else
    Dec(FRecSize,FStream.Write(P^,Len));
end;

procedure THelperWriteSST.WriteRichData(P: PByteArray; Count: integer);
var
  i: integer;
  n: integer;
  pFnt: PXc12FontRun;
begin
  for i := 0 to Count - 1 do begin
    if FRecSize < SizeOf(TFormatRun) then begin
      WriteCONTINUE;
    end;
    pFnt := PXc12FontRun(NativeInt(P) + (i * SizeOf(TXc12FontRun)));
    Dec(FRecSize,FStream.Write(pFnt.Index,2));
    n := pFnt.Font.Index - XLS_STYLE_DEFAULT_FONT_COUNT;
    Dec(FRecSize,FStream.Write(n,2));
  end;
end;

procedure THelperWriteSST.WriteSSTString(AStr: PXLSString);
var
  B: byte;
  W: word;
begin
  if (FCurrIndex mod FExtSST.BucketSize) = 0 then
    FExtSST.Add(FStream.Pos,FStream.Pos - FRecPos + SizeOf(TBIFFHeader) - 2);

  // Don't split string headers over CONTINUE records
  if FRecSize <= 15 then
    WriteCONTINUE;

  if AStr = Nil then begin
    W := 0;
    Dec(FRecSize,FStream.Write(W,2));
    B := 0;
    Dec(FRecSize,FStream.Write(B,1));
  end
  else begin
    Dec(FRecSize,FStream.Write(AStr.Len,2));
    Dec(FRecSize,FStream.Write(AStr.Options,1));
    case AStr.Options of
      STRID_COMPRESSED: WriteString(PByteArray(@AStr.Data),AStr.Len,False);
      STRID_UNICODE:  WriteString(PByteArray(@AStr.Data),AStr.Len * 2,True);
      STRID_RICH: begin
        Dec(FRecSize,FStream.Write(PXLSStringRich(AStr).FormatCount,2));
        WriteString(PByteArray(@PXLSStringRich(AStr).Data),PXLSStringRich(AStr).Len,False);
        WriteRichData(Pointer(NativeInt(@PXLSStringRich(AStr).Data) + PXLSStringRich(AStr).Len),PXLSStringRich(AStr).FormatCount);
      end;
      STRID_RICH_UNICODE: begin
        Dec(FRecSize,FStream.Write(PXLSStringRichUC(AStr).FormatCount,2));
        WriteString(PByteArray(@PXLSStringRichUC(AStr).Data),PXLSStringRichUC(AStr).Len * 2,True);
        WriteRichData(Pointer(NativeInt(@PXLSStringRichUC(AStr).Data) + PXLSStringRichUC(AStr).Len * 2),PXLSStringRichUC(AStr).FormatCount);
      end;
      STRID_FAREAST: begin
        Dec(FRecSize,FStream.Write(PXLSStringFarEast(AStr).FarEastDataSize,4));
        WriteString(PByteArray(@PXLSStringFarEast(AStr).Data),PXLSStringFarEast(AStr).Len,False);
        WriteData(Pointer(NativeInt(@PXLSStringFarEast(AStr).Data) + PXLSStringFarEast(AStr).Len),PXLSStringFarEast(AStr).FarEastDataSize);
      end;
      STRID_FAREAST_RICH: begin
        Dec(FRecSize,FStream.Write(PXLSStringFarEastRich(AStr).FormatCount,2));
        Dec(FRecSize,FStream.Write(PXLSStringFarEastRich(AStr).FarEastDataSize,4));
        WriteString(PByteArray(@PXLSStringFarEastRich(AStr).Data),PXLSStringFarEastRich(AStr).Len,False);
        WriteRichData(Pointer(NativeInt(@PXLSStringFarEastRich(AStr).Data) + PXLSStringFarEastRich(AStr).Len),PXLSStringFarEastRich(AStr).FormatCount);
        WriteData(Pointer(NativeInt(@PXLSStringFarEastRich(AStr).Data) + PXLSStringFarEastRich(AStr).Len + PXLSStringFarEastRich(AStr).FormatCount * SizeOf(TXc12FontRun)),PXLSStringFarEastRich(AStr).FarEastDataSize);
      end;
      STRID_FAREAST_UC: begin
        Dec(FRecSize,FStream.Write(PXLSStringFarEastUC(AStr).FarEastDataSize,4));
        WriteString(PByteArray(@PXLSStringFarEastUC(AStr).Data),PXLSStringFarEastUC(AStr).Len * 2,True);
        WriteData(Pointer(NativeInt(@PXLSStringFarEastUC(AStr).Data) + PXLSStringFarEastUC(AStr).Len * 2),PXLSStringFarEastUC(AStr).FarEastDataSize);
      end;
      STRID_FAREAST_RICH_UC: begin
        Dec(FRecSize,FStream.Write(PXLSStringFarEastRichUC(AStr).FormatCount,2));
        Dec(FRecSize,FStream.Write(PXLSStringFarEastRichUC(AStr).FarEastDataSize,4));
        WriteString(PByteArray(@PXLSStringFarEastRichUC(AStr).Data),PXLSStringFarEastRichUC(AStr).Len * 2,True);
        WriteRichData(Pointer(NativeInt(@PXLSStringFarEastRichUC(AStr).Data) + PXLSStringFarEastRichUC(AStr).Len * 2),PXLSStringFarEastRichUC(AStr).FormatCount);
        WriteData(Pointer(NativeInt(@PXLSStringFarEastRichUC(AStr).Data) + PXLSStringFarEastRichUC(AStr).Len * 2 + PXLSStringFarEastRichUC(AStr).FormatCount * SizeOf(TXc12FontRun)),PXLSStringFarEastRichUC(AStr).FarEastDataSize);
      end;
      else
        raise XLSRWException.CreateFmt('STT: Unhandled string type in Write (%.2X) # %d',[AStr.Options,FCurrIndex]);
    end;
  end;
  Inc(FCurrIndex);
end;

procedure THelperWriteSST.WriteString(P: PByteArray; Len: word; IsUnicode: boolean);
var
  Options: byte;
  IsSplittedChar: boolean;
begin
  if Len > FRecSize then begin
    // Do not split unicode characters
    IsSplittedChar := IsUnicode and Odd(Len - FRecSize);
    if IsSplittedChar then
      Dec(FRecSize);
    FStream.Write(P^,FRecSize);
    Dec(Len,FRecSize);
    P := PByteArray(NativeInt(P) + FRecSize);
    if IsSplittedChar then
      FRecSize := 1
    else
      FRecSize := 0;
    WriteCONTINUE;
    if IsUnicode then
      Options := $01
    else
      Options := $00;
    Dec(FRecSize,FStream.Write(Options,1));
    WriteString(P,Len,IsUnicode);
  end
  else
    Dec(FRecSize,FStream.Write(P^,Len));
end;

{ THelperWriteExtSST }

procedure THelperWriteExtSST.Add(StreamPos: longword; RecPos: word);
var
  P: PExtSSTRec;
begin
  New(P);
  P.StreamPos := StreamPos;
  P.RecPos := RecPos;
  inherited Add(P);
end;

procedure THelperWriteExtSST.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FreeMem(Items[i]);
  inherited;
end;

procedure THelperWriteExtSST.SetStringCount(const Value: integer);
var
  MaxBuckets: integer;
begin
  MaxBuckets := (MAXRECSZ_97 - SizeOf(TRecEXTSST)) div SizeOf(TRecISSTINF);
  FBucketSize := (Value div MaxBuckets) + 1;
  if FBucketSize < 8 then
    FBucketSize := 8;
end;

procedure THelperWriteExtSST.Write(Stream: TXLSStream);
var
  i: integer;
  Rec: TRecISSTINF;
begin
  Stream.WriteHeader(BIFFRECID_EXTSST,SizeOf(TRecEXTSST) + Count * SizeOf(TRecISSTINF));
  Stream.WWord(FBucketSize);
  for i := 0 to Count - 1 do begin
    Rec.Pos := PExtSSTRec(Items[i]).StreamPos;
    Rec.Offset := PExtSSTRec(Items[i]).RecPos;
    Rec.Reserved := 0;
    Stream.Write(Rec,SizeOf(TRecISSTINF));
  end;
end;

end.
