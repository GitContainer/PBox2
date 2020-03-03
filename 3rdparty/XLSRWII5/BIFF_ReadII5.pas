unit BIFF_ReadII5;

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

uses Classes, SysUtils, Math,
{$ifdef BABOON}
{$else}
     Windows, vcl.Clipbrd,
{$endif}
     BIFF_Utils5, BIFF_RecsII5, BIFF_SheetData5, BIFF_Stream5,
     BIFF5, BIFF_DecodeFormula5, BIFF_RecordStorage5, BIFF_Escher5,
     BIFF_MergedCells5, BIFF_WideStrList5,
     Xc12Utils5, Xc12Manager5, Xc12DataWorkbook5, Xc12DataWorksheet5, Xc12DataStylesheet5,
     Xc12Common5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSFormulaTypes5, XLSCellAreas5, XLSHyperlinks5;

type
  TSharedFormula = record
    Row1, Row2: word;
    Col1, Col2: byte;
    Len: word;
    Formula: PByteArray;
  end;

type
  TXLSReadII = class(TObject)
  protected
    FManager: TXc12Manager;
    PBuf: PByteArray;
    FXLS: TBIFF5;
    FCurrSheet: integer;
    FXLSStream: TXLSStream;
    FBoundsheets: TXLSWideStringList;
    FBoundsheetIndex: integer;
    Header: TBIFFHeader;
    FSharedFormulas: array of TSharedFormula;
    FBIFFVer: longword;
    FLastARRAY: PByteArray;
    FLastARRAYSize: integer;

    FMaxFormat: integer;
    FFormatCnt: integer;
    FEx40NumFmts: array of string;
    FEx40NumFmtsSaved: boolean;

    CurrRecs: TRecordStorage;
    InsertRecord: boolean;
    FCnt: integer;  // for debug

    FCurrXc12Sheet: TXc12DataWorksheet;
    FSkipMSO: boolean;

    // File prefix
    procedure RREC_FILEPASS;
    procedure RREC_WRITEACCESS;
    procedure RREC_FILESHAREING;
    procedure RREC_CODEPAGE;
    procedure RREC_DSF;
    procedure RREC_FNGROUPCOUNT;
    procedure RREC_WINDOWPROTECT;
    procedure RREC_PROTECT;
    procedure RREC_SCENPROTECT;
    procedure RREC_OBJPROTECT;
    procedure RREC_PASSWORD;
    procedure RREC_PROT4REV;
    procedure RREC_PROT4REVPASS;
    procedure RREC_WINDOW1;
    procedure RREC_BACKUP;
    procedure RREC_HIDEOBJ;
    procedure RREC_1904;
    procedure RREC_PRECISION;
    procedure RREC_REFRESHALL;
    procedure RREC_BOOKBOOL;
    procedure RREC_PALETTE;
    procedure RREC_FONT;
    procedure RRec_FONT_40;
    procedure RREC_FORMAT;
    procedure RREC_FORMAT_40;
    procedure RREC_XF;
    procedure RREC_XF_40;
    procedure RREC_STYLE;
    procedure RREC_NAME;
    procedure RREC_SUPBOOK;
    procedure RREC_EXTERNNAME;
    procedure RREC_XCT;
    procedure RREC_EXTERNCOUNT;
    procedure RREC_EXTERNSHEET;
    procedure RREC_BOUNDSHEET;
    procedure RREC_COUNTRY;
    procedure RREC_RECALCID;
    procedure RREC_MSODRAWINGGROUP;
    procedure RREC_MSO_0866;
    procedure RREC_SST;
    procedure RREC_EXTSST;
    procedure RREC_EOF;

    // Sheet prefix
    procedure RREC_CALCMODE;
    procedure RREC_CALCCOUNT;
    procedure RREC_REFMODE;
    procedure RREC_ITERATION;
    procedure RREC_DELTA;
    procedure RREC_SAVERECALC;
    procedure RREC_PRINTHEADERS;
    procedure RREC_PRINTGRIDLINES;
    procedure RREC_GRIDSET;
    procedure RREC_GUTS;
    procedure RREC_DEFAULTROWHEIGHT;
    procedure RREC_WSBOOL;
    procedure RREC_HORIZONTALPAGEBREAKS;
    procedure RREC_VERTICALPAGEBREAKS;
    procedure RREC_HEADER;
    procedure RREC_FOOTER;
    procedure RREC_HCENTER;
    procedure RREC_VCENTER;
    procedure RREC_SETUP;
    procedure RREC_LEFTMARGIN;
    procedure RREC_RIGHTMARGIN;
    procedure RREC_TOPMARGIN;
    procedure RREC_BOTTOMMARGIN;
    procedure RREC_DEFCOLWIDTH;
    procedure RREC_FILTERMODE;
    procedure RREC_AUTOFILTERINFO;
    procedure RREC_COLWIDTH;
    procedure RREC_COLINFO;
    procedure RREC_DIMENSIONS;

    // Sheet data
    procedure RREC_INTEGER_20;
    procedure RREC_NUMBER_20;
    procedure RREC_LABEL_20;

    procedure RREC_ROW;
    procedure RREC_BLANK;
    procedure RREC_BOOLERR;
    procedure RREC_FORMULA;
    procedure RREC_NUMBER;
    procedure RREC_RK;
    procedure RREC_MULRK;
    procedure RREC_MULBLANK;
    procedure RREC_LABELSST;
    procedure RREC_LABEL;
    procedure RREC_LABEL_X; // Non-excel export
    procedure RREC_RSTRING;
    procedure RREC_NOTE;

    procedure READ_SHRFMLA;

    // Sheet suffix
    procedure RREC_MSODRAWING;
    procedure RREC_MSODRAWINGSELECTION;
    procedure RREC_WINDOW2;
    procedure RREC_SCL;
    procedure RREC_PANE;
    procedure RREC_SELECTION;
    procedure RREC_SHEETEXT;
    procedure RREC_FEATHEADR;
    procedure RREC_DVAL;
    procedure RREC_MERGEDCELLS;
    procedure RREC_CONDFMT;
    procedure RREC_HLINK;

    function  CheckFmlaStrVal(AVal: AxUCString): AxUCString;

    procedure Clear;
    procedure ClearSharedFmla;
    procedure ReadFormulaVal(Col, Row, FormatIndex: integer; Value: double; Formula: PByteArray; Len, DataSz: integer);
    procedure FixupSharedFormula(LeftCol, TopRow, ACol, ARow: integer; FirstFormula: boolean);
    function  SSTReadCONTINUE: word;
    function  SSTReadString(var RecSize: word): PByteArray;
    procedure ReadSST(ARecSize: word);
    function  GetNAMEString(var P: PByteArray; Len: integer): AxUCString;
    procedure ReadExcel40;
    procedure CacheMSODrawing;
    procedure SkipChart;

    procedure DoDirectRead(const ACol,ARow: integer; AValue: double); overload;
    procedure DoDirectRead(const ACol,ARow: integer; AValue: boolean); overload;
    procedure DoDirectRead(const ACol,ARow: integer; AValue: TXc12CellError); overload;
    procedure DoDirectRead(const ACol,ARow: integer; AValue: AxUCString); overload;
  public
    constructor Create(AXLS: TBIFF5);
    destructor Destroy; override;

    procedure LoadFromStream(Stream: TStream);
    procedure LoadSheetNamesFromStream(Stream: TStream; AList: TStrings);

    property SkipMSO: boolean read FSkipMSO write FSkipMSO;
  end;

type
  PRecCellXF = ^TRecCellXF;

  TRecCellXF = packed record
    Row, Col: word;
    FormatIndex: word;
  end;

type
  TXFData = record
    XF: integer;
    Font: integer;
    NumFmt: integer;
  end;

type
  PDataRSTRING = ^TDataRSTRING;

  TDataRSTRING = packed record
    Row, Col: word;
    FormatIndex: word;
    Len: word;
    Options: byte;
    Data: record end;
  end;

type TXLSReadClipboard = class(TXLSReadII)
private
     procedure IncXF(Index: integer);
     procedure DoRSTRING;
     procedure SaveRSTRING;
     procedure AdjustCell;
protected
     FVersion: TExcelVersion;
     FFONTUsage,
     FFORMATUsage: array of integer;
     FXFUsage: array of TXFData;
     FXFCount,
     FFONTCount,
     FFORMATMax: integer;
     FSrcArea: TRecCellAreaI;
     FDestSheet: integer;
     FDestCol: integer;
     FDestRow: integer;

     function ReadPass1: boolean;
     function ReadPass2: boolean;
public
     function Read(AStream: TStream; const ADestSheet, ADestCol, ADestRow: integer): boolean;
     end;

implementation

var
  L_Count: integer;

const STRID_COMPRESSED      = $00;
const STRID_UNICODE         = $01;
const STRID_RICH            = $08;
const STRID_RICH_UNICODE    = STRID_RICH + STRID_UNICODE;
const STRID_FAREAST         = $04;
const STRID_FAREAST_RICH    = STRID_FAREAST + STRID_RICH;
const STRID_FAREAST_UC      = STRID_FAREAST + STRID_UNICODE;
const STRID_FAREAST_RICH_UC = STRID_FAREAST + STRID_UNICODE + STRID_RICH;

// Compressed
type PXLSString = ^TXLSString;
     TXLSString = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     Data: record end;
     end;

// Unicode
type PXLSStringUC = ^TXLSStringUC;
     TXLSStringUC = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     Data: record end;
     end;

// Compressed with formatting
type PXLSStringRich = ^TXLSStringRich;
     TXLSStringRich = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     Data: record end;
     end;

// Unicode with formatting
type PXLSStringRichUC = ^TXLSStringRichUC;
     TXLSStringRichUC = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     Data: record end;
     end;

// Compressed with Far East data
type PXLSStringFarEast = ^TXLSStringFarEast;
     TXLSStringFarEast = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     FarEastDataSize: longword;
     Data: record end;
     end;

// Unicode with Far East data
type PXLSStringFarEastUC = ^TXLSStringFarEastUC;
     TXLSStringFarEastUC = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     FarEastDataSize: longword;
     Data: record end;
     end;

// Compressed with Far East data and formatting
type PXLSStringFarEastRich = ^TXLSStringFarEastRich;
     TXLSStringFarEastRich = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     FarEastDataSize: longword;
     Data: record end;
     end;

// Unicode with Far East data and formatting
type PXLSStringFarEastRichUC = ^TXLSStringFarEastRichUC;
     TXLSStringFarEastRichUC = record
     Index: integer;
     RefCount: integer;
     Len: word;
     Options: byte;
     FormatCount: word;
     FarEastDataSize: longword;
     Data: record end;
     end;

constructor TXLSReadII.Create(AXLS: TBIFF5);
begin
  FXLS := AXLS;

  if FXLS <> Nil then begin
    FManager := FXLS.Manager;
    FXLSStream := TXLSStream.Create(FXLS.VBA);
    FXLSStream.SaveVBA := FXLS.PreserveMacros;
    FBoundsheets := TXLSWideStringList.Create;
  end
  else
    FXLSStream := TXLSStream.Create(Nil);

  FSkipMSO := False;
end;

destructor TXLSReadII.Destroy;
begin
  FXLSStream.Free;
  FBoundsheets.Free;
  ClearSharedFmla;
  if FLastARRAY <> Nil then
    FreeMem(FLastARRAY);
  inherited Destroy;
end;

procedure TXLSReadII.DoDirectRead(const ACol, ARow: integer; AValue: boolean);
begin
  FManager.EventCell.Clear;
  FManager.EventCell.SheetIndex := FCurrSheet;
  FManager.EventCell.SheetName := FXLS[FCurrSheet].Name;
  FManager.EventCell.Col := ACol;
  FManager.EventCell.Row := ARow;
  FManager.EventCell.AsBoolean := AValue;
  FManager.FireReadCellEvent;
end;

procedure TXLSReadII.DoDirectRead(const ACol, ARow: integer; AValue: TXc12CellError);
begin
  FManager.EventCell.Clear;
  FManager.EventCell.SheetIndex := FCurrSheet;
  FManager.EventCell.SheetName := FXLS[FCurrSheet].Name;
  FManager.EventCell.Col := ACol;
  FManager.EventCell.Row := ARow;
  FManager.EventCell.AsError := AValue;
  FManager.FireReadCellEvent;
end;

procedure TXLSReadII.DoDirectRead(const ACol, ARow: integer; AValue: AxUCString);
begin
  FManager.EventCell.Clear;
  FManager.EventCell.SheetIndex := FCurrSheet;
  FManager.EventCell.SheetName := FXLS[FCurrSheet].Name;
  FManager.EventCell.Col := ACol;
  FManager.EventCell.Row := ARow;
  FManager.EventCell.AsString := AValue;
  FManager.FireReadCellEvent;
end;

procedure TXLSReadII.DoDirectRead(const ACol, ARow: integer; AValue: double);
begin
  FManager.EventCell.Clear;
  FManager.EventCell.SheetIndex := FCurrSheet;
  FManager.EventCell.SheetName := FXLS[FCurrSheet].Name;
  FManager.EventCell.Col := ACol;
  FManager.EventCell.Row := ARow;
  FManager.EventCell.AsFloat := AValue;
  FManager.FireReadCellEvent;
end;

procedure TXLSReadII.CacheMSODrawing;
var
  H: integer;
begin
  FXLSStream.Read(PBuf^, Header.Length);
  FXLS[FCurrSheet]._Int_CachedMSORecs.AddRec(Header,PBuf);

  while True do begin
    H := FXLSStream.PeekHeader;
    if (H = BIFFRECID_MSODRAWING) or
       (H = BIFFRECID_TXO) or
       (H = BIFFRECID_OBJ) or
       (H = BIFFRECID_CONTINUE) or
       (H = BIFFRECID_NOTE) then begin

      FXLSStream.ReadHeader(Header);
      FXLSStream.Read(PBuf^, Header.Length);
      FXLS[FCurrSheet]._Int_CachedMSORecs.AddRec(Header,PBuf);
    end
    else
      Break;
  end;
end;

function TXLSReadII.CheckFmlaStrVal(AVal: AxUCString): AxUCString;
var
  i,j: integer;
begin
  SetLength(Result,Length(AVal));
  j := 0;
  for i := 1 to Length(AVal) do begin
    if Word(AVal[i]) >= 32 then begin
      Inc(j);
      Result[j] := AVal[i];
    end;
  end;
  SetLength(Result,j);
end;

procedure TXLSReadII.Clear;
begin
  FCurrSheet := -1;
  FBoundsheets.Clear;
  FBoundsheetIndex := -1;
end;

procedure TXLSReadII.ClearSharedFmla;
var
  i: integer;
begin
  for i := 0 to High(FSharedFormulas) do
    FreeMem(FSharedFormulas[i].Formula);
  SetLength(FSharedFormulas, 0);
end;

procedure TXLSReadII.LoadFromStream(Stream: TStream);
var
  V, ProgressCount: integer;
  Count: integer;
  InUSERSVIEW: boolean;
  WS: AxUCString;
begin
  InUSERSVIEW := False;
  Count := 0;
  Clear;
  FXLS.WriteDefaultData := False;
  try
    FXLSStream.ExtraObjects := FXLS.ExtraObjects;
    FXLSStream.SourceStream := Stream;

    FXLS.Version := FXLSStream.OpenStorageRead(FXLS.Filename);
    GetMem(PBuf, FXLS.MaxBuffsize);

    CurrRecs := FXLS.Records;

    ProgressCount := 0;
    if Assigned(FXLS.OnProgress) then
      FXLS.OnProgress(Self, 0);

    if (FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader)) and (Header.RecID = BIFFRECID_BOF4) then begin
      ReadExcel40;
      Exit;
    end
    else
      FXLSStream.Seek(0,soFromBeginning);

    try
      while FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
        if FXLS.Aborted then
          Exit;
        if Header.Length > FXLS.MaxBuffsize then begin
          // Invalid record size
          ReAllocMem(PBuf, Header.Length);
          // FXLSStream.Seek(Header.Length,soFromCurrent);
          // Continue;
        end
        else if (Header.RecID and INTERNAL_PLACEHOLDER) <> 0 then
          raise XLSRWException.Create('Bad record in file')
          // BOFPos in BOUNDSHEET will not have correect values when reading
          // encrypted files, as BOFPos is written unencrypted.

        else if ((Header.RecID and $FF) <> BIFFRECID_BOF) and
          (Header.RecID <> BIFFRECID_SST) and
        // (Header.RecId <> BIFFRECID_MSO_0866) and
          (Header.RecID <> BIFFRECID_MSODRAWINGGROUP) and
          (Header.RecID <> BIFFRECID_MSODRAWING) then
          FXLSStream.Read(PBuf^, Header.Length);
        if Assigned(FXLS.OnProgress) and (FXLSStream.Size > 0) then begin
          V := Round((FXLSStream.Pos / FXLSStream.Size) * 100);
          if V > ProgressCount then begin
            ProgressCount := V;
            FXLS.OnProgress(Self, ProgressCount);
          end;
        end;

        // if ((Header.RecID and $FF) <> BIFFRECID_BOF) and (Header.RecID <> BIFFRECID_EXCEL9FILE) then
        // CurrRecs.Add(Header,PBuf);

        if InUSERSVIEW then
        begin
          if Header.RecID = BIFFRECID_USERSVIEWEND then
            InUSERSVIEW := False;
          Continue;
        end;

        InsertRecord := True;
        case Header.RecID of
          BIFFRECID_EOF:
            begin
              CurrRecs.UpdateDefault(Header, PBuf);
              InsertRecord := False;
              ClearSharedFmla;
              if (FBoundsheets.Count <= 0) or
                (FBoundsheetIndex >= (FBoundsheets.Count - 1)) then
              begin
                Break;
              end;
            end;
          // File prefix
          // BIFFRECID_OBPROJ:              InsertRecord := False;
          // BIFFRECID_EXCEL9FILE:          InsertRecord := False;
          BIFFRECID_FILEPASS                : RREC_FILEPASS;
          BIFFRECID_WRITEACCESS             : RREC_WRITEACCESS;
          BIFFRECID_FILESHAREING            : RREC_FILESHAREING;
          BIFFRECID_CODEPAGE                : RREC_CODEPAGE;
          BIFFRECID_DSF                     : RREC_DSF;
          BIFFRECID_FNGROUPCOUNT            : RREC_FNGROUPCOUNT;

          BIFFRECID_WINDOWPROTECT           : RREC_WINDOWPROTECT;
          BIFFRECID_PROTECT                 : RREC_PROTECT;
          BIFFRECID_SCENPROTECT             : RREC_SCENPROTECT;
          BIFFRECID_OBJPROTECT              : RREC_OBJPROTECT;
          BIFFRECID_PASSWORD                : RREC_PASSWORD;
          BIFFRECID_PROT4REV                : RREC_PROT4REV;
          BIFFRECID_PROT4REVPASS            : RREC_PROT4REVPASS;

          BIFFRECID_WINDOW1                 : RREC_WINDOW1;
          BIFFRECID_BACKUP                  : RREC_BACKUP;
          BIFFRECID_HIDEOBJ                 : RREC_HIDEOBJ;
          BIFFRECID_1904                    : RREC_1904;
          BIFFRECID_PRECISION               : RREC_PRECISION;
          BIFFRECID_REFRESHALL              : RREC_REFRESHALL;
          BIFFRECID_BOOKBOOL                : RREC_BOOKBOOL;
          BIFFRECID_PALETTE                 : RREC_PALETTE;

          BIFFRECID_FONT,
          BIFFRECID_FONT_40                 : RREC_FONT;
          BIFFRECID_FORMAT                  : RREC_FORMAT;
          BIFFRECID_XF_30,
          BIFFRECID_XF_40                   : RREC_XF_40;
          BIFFRECID_XF                      : RREC_XF;

          BIFFRECID_STYLE                   : RREC_STYLE;

          BIFFRECID_NAME, $0218             : RREC_NAME;
          BIFFRECID_SUPBOOK                 : RREC_SUPBOOK;
          BIFFRECID_EXTERNNAME, $0223       : RREC_EXTERNNAME;
          BIFFRECID_XCT                     : RREC_XCT;
          BIFFRECID_EXTERNCOUNT             : RREC_EXTERNCOUNT;
          BIFFRECID_EXTERNSHEET             : RREC_EXTERNSHEET;

          BIFFRECID_BOUNDSHEET              : RREC_BOUNDSHEET;

          BIFFRECID_COUNTRY                 : RREC_COUNTRY;
          BIFFRECID_RECALCID                : RREC_RECALCID;

          BIFFRECID_MSODRAWINGGROUP         : begin
            if FSkipMSO then begin
              InsertRecord := False;
              FXLSStream.Read(PBuf^,Header.Length);
              CurrRecs.AddRec(Header, PBuf);
              while FXLSStream.PeekHeader in [BIFFRECID_MSODRAWINGGROUP,BIFFRECID_CONTINUE] do begin
                FXLSStream.ReadHeader(Header);
                FXLSStream.Read(PBuf^,Header.Length);
                CurrRecs.AddRec(Header, PBuf);
              end;
            end
            else
              RREC_MSODRAWINGGROUP;
          end;
          BIFFRECID_MSO_0866                : RREC_MSO_0866;

          BIFFRECID_SST                     : RREC_SST;
          BIFFRECID_EXTSST                  : RREC_EXTSST;

          // Sheet prefix
          BIFFRECID_CALCMODE                : RREC_CALCMODE;
          BIFFRECID_CALCCOUNT               : RREC_CALCCOUNT;
          BIFFRECID_REFMODE                 : RREC_REFMODE;
          BIFFRECID_ITERATION               : RREC_ITERATION;
          BIFFRECID_DELTA                   : RREC_DELTA;
          BIFFRECID_SAVERECALC              : RREC_SAVERECALC;

          BIFFRECID_PRINTHEADERS            : RREC_PRINTHEADERS;
          BIFFRECID_PRINTGRIDLINES          : RREC_PRINTGRIDLINES;
          BIFFRECID_GRIDSET                 : RREC_GRIDSET;
          BIFFRECID_GUTS                    : RREC_GUTS;
          BIFFRECID_DEFAULTROWHEIGHT        : RREC_DEFAULTROWHEIGHT;
          BIFFRECID_WSBOOL                  : RREC_WSBOOL;
          BIFFRECID_HORIZONTALPAGEBREAKS    : RREC_HORIZONTALPAGEBREAKS;
          BIFFRECID_VERTICALPAGEBREAKS      : RREC_VERTICALPAGEBREAKS;

          BIFFRECID_HEADER                  : RREC_HEADER;
          BIFFRECID_FOOTER                  : RREC_FOOTER;
          BIFFRECID_HCENTER                 : RREC_HCENTER;
          BIFFRECID_VCENTER                 : RREC_VCENTER;
          BIFFRECID_SETUP                   : RREC_SETUP;
          BIFFRECID_LEFTMARGIN              : RREC_LEFTMARGIN;
          BIFFRECID_RIGHTMARGIN             : RREC_RIGHTMARGIN;
          BIFFRECID_TOPMARGIN               : RREC_TOPMARGIN;
          BIFFRECID_BOTTOMMARGIN            : RREC_BOTTOMMARGIN;
          BIFFRECID_DEFCOLWIDTH             : RREC_DEFCOLWIDTH;
          BIFFRECID_FILTERMODE              : RREC_FILTERMODE;
          BIFFRECID_AUTOFILTERINFO          : RREC_AUTOFILTERINFO;
          BIFFRECID_COLINFO                 : RREC_COLINFO;
          BIFFRECID_DIMENSIONS_20           : begin
              // DIMENSIONS_20 has the value $0000. This record was used in excel
              // 5 files created with XLSReadWrite 1.x, in order to pad the file
              // size.
              if FXLS.Version = xvExcel50 then
                InsertRecord := False
              else
                RREC_DIMENSIONS;
            end;
          BIFFRECID_DIMENSIONS              : RREC_DIMENSIONS;

          // Sheet cells/data
          BIFFRECID_INDEX:
            begin
              InsertRecord := False;
              if FXLS.Version <= xvExcel40 then
                CurrRecs := FXLS.Sheets[FCurrSheet]._Int_Records;
            end;
          BIFFRECID_DBCELL                  : InsertRecord := False;

          BIFFRECID_INTEGER_20              : RREC_INTEGER_20;
          BIFFRECID_NUMBER_20               : RREC_NUMBER_20;
          BIFFRECID_LABEL_20                : RREC_LABEL_20;

          BIFFRECID_ROW                     : RREC_ROW;
          BIFFRECID_BLANK                   : RREC_BLANK;
          BIFFRECID_BOOLERR                 : RREC_BOOLERR;
          BIFFRECID_FORMULA                 : RREC_FORMULA;
          BIFFRECID_MULBLANK                : RREC_MULBLANK;
          BIFFRECID_RK, BIFFRECID_RK7       : RREC_RK;
          BIFFRECID_MULRK                   : RREC_MULRK;
          BIFFRECID_NUMBER                  : RREC_NUMBER;
          BIFFRECID_LABELSST                : RREC_LABELSST;
          BIFFRECID_LABEL                   : RREC_LABEL;

          BIFFRECID_NOTE                    : begin
            if FSkipMSO then
              InsertRecord := True
            else
              RREC_NOTE;
          end;

          // Sheet suffix
          BIFFRECID_MSODRAWING              : begin
            if FSkipMSO then begin
//              InsertRecord := True;
              FXLSStream.Read(PBuf^,Header.Length);
              CurrRecs.AddRec(Header, PBuf);
              SkipChart;
            end
//              CacheMSODrawing
            else
              RREC_MSODRAWING;
          end;

          BIFFRECID_MSODRAWINGSELECTION     : begin
            if FSkipMSO then begin
              InsertRecord := True;
              FXLSStream.Read(PBuf^, Header.Length);
            end
            else
              RREC_MSODRAWINGSELECTION;
          end;

          BIFFRECID_WINDOW2                 : RREC_WINDOW2;

          BIFFRECID_PANE                    : RREC_PANE;
          BIFFRECID_SELECTION               : RREC_SELECTION;
          BIFFRECID_SHEETEXT                : RREC_SHEETEXT;
          BIFFRECID_FEATHEADR               : RREC_FEATHEADR;
          BIFFRECID_MERGEDCELLS:
            begin
              TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_SUFFIXDATA);
              RREC_MERGEDCELLS;
            end;

          BIFFRECID_CONDFMT                 : RREC_CONDFMT;
          BIFFRECID_HLINK                   : RREC_HLINK;
          BIFFRECID_DV,
          BIFFRECID_DVAL                    : begin
            if FSkipMSO then
              InsertRecord := False
            else
              RREC_DVAL;
          end;

          BIFFRECID_USERBVIEW               : InsertRecord := False;
          BIFFRECID_USERSVIEWBEGIN:
            begin
              InsertRecord := False;
              InUSERSVIEW := True;
            end;
          // 0000:                          InsertRecord := False; // Dummy recors used by XLSReadWrite 1.37

          // Excel 2007 unknown
          $0863, $087C, $087D, $0892, $088E, $089A, $08A3, $0896, $089B, $088C:
            begin
              InsertRecord := False;
//              FXLS.RecordsEx2007.AddRec(Header, PBuf);
            end
          else
          begin
            if (Header.RecID and $FF) = BIFFRECID_BOF then begin
              ClearSharedFmla;
              FXLSStream.ReadUnencryptedSync(PBuf^, Header.Length);
              InsertRecord := False;
              case PRecBOF8(PBuf).SubStreamType of
                $0005:
                  begin // Workbook globals
                    if Header.Length > 256 then
                      raise XLSRWException.Create('Invalid BOF size');
                    PRecBOF8(PBuf).BuildIdentifier := $18AF;
                    CurrRecs := FXLS.Records;
                    CurrRecs.UpdateDefault(Header, PBuf);
                  end;
                $0040, // BIFF4 Macro sheet
                $0006, // VB Module
                $0010:
                  begin // Worksheet
                    Inc(FBoundsheetIndex);
                    Inc(FCurrSheet);
                    case PRecBOF8(PBuf).SubStreamType of
                      $0010: begin
                        if FXLS.Version < xvExcel40 then begin
                          FXLS.Sheets.Add(wtSheet);
                          FCurrXc12Sheet := FManager.Worksheets.Add(FManager.Worksheets.Count + 1);
                        end
                        else
                          FCurrXc12Sheet := FManager.Worksheets[FCurrSheet];
                      end
                      else   raise XLSRWException.Create('Invalid sheet type in file');
                    end;
                    FXLS.Sheets[FCurrSheet]._Int_HasDefaultRecords := True;
                    if FXLS.Version <= xvExcel40 then begin
                      FXLS.Sheets[FCurrSheet].Name := 'Sheet1';
                      FCurrXc12Sheet.Name := 'Sheet1';
                    end
                    else
                    begin
                      WS := FBoundsheets[FBoundsheetIndex];
                      FXLS.Sheets[FCurrSheet]._Int_SheetIndex := FBoundsheetIndex;
                      FXLS.Sheets[FCurrSheet].Name := WS;
                      FXLS.Sheets[FCurrSheet].Hidden := THiddenState(FBoundsheets.Objects[FBoundsheetIndex]);
                      CurrRecs := FXLS.Sheets[FCurrSheet]._Int_Records;
                      CurrRecs.UpdateDefault(Header, PBuf);

                      FCurrXc12Sheet.Name := WS;
                      case FXLS.Sheets[FCurrSheet].Hidden of
                        hsHidden    : FCurrXc12Sheet.State := x12vHidden;
                        hsVeryHidden: FCurrXc12Sheet.State := x12vVeryHidden;
                        else          FCurrXc12Sheet.State := x12vVisible;
                      end;
                    end;

                    FBIFFVer := PRecBOF8(PBuf).LowBIFF;
                  end;
                $0020:
                  begin // Chart
                    Inc(FBoundsheetIndex);
                    WS := FBoundsheets[FBoundsheetIndex];
                    FXLS.SheetCharts.LoadFromStream(FXLSStream, WS, PBuf, FManager.StyleSheet.Fonts, FBoundsheetIndex);
                    // raise XLSRWException.Create('Chart unhandled');
                  end;
                $0100:
                  begin // BIFF4 Workbook globals
                    raise XLSRWException.Create('BIFF4 Workbook unhandled');
                  end;
              end;
            end
          end;
        end;
        if InsertRecord then
          CurrRecs.AddRec(Header, PBuf);
        Inc(Count);
      end;
    except
      on EAbort do
        raise;
      on E: XLSRWException do
        raise XLSRWException.CreateFmt('Error on reading record # %d, %.4X Offs: %.8X' + #13 + E.Message,[Count, Header.RecID, FXLSStream.Pos]);
    end;
    if Assigned(FXLS.OnProgress) then
      FXLS.OnProgress(Self, 100);
  finally
    FXLSStream.CloseStorage;
    FreeMem(PBuf);
  end;

end;

procedure TXLSReadII.LoadSheetNamesFromStream(Stream: TStream; AList: TStrings);
var
  i: integer;
  WS: AxUCString;
  Version: TExcelVersion;
  InBoundsheet: boolean;
begin
  try
    FXLSStream.SourceStream := Stream;

    Version := FXLSStream.OpenStorageRead('');

    GetMem(PBuf, MAXRECSZ_97);

    InBoundsheet := False;

    if (FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader)) and (Header.RecID = BIFFRECID_BOF4) then begin
      ReadExcel40;
      Exit;
    end
    else
      FXLSStream.Seek(0,soFromBeginning);

    try
      while FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
        if Header.Length > MAXRECSZ_97 then begin
          ReAllocMem(PBuf, Header.Length);
        end;

        FXLSStream.Read(PBuf^, Header.Length);

        if Header.RecID = BIFFRECID_BOUNDSHEET then begin
          if Version < xvExcel97 then begin
              SetLength(WS, PRecBOUNDSHEET7(PBuf).NameLen);
              for i := 1 to PRecBOUNDSHEET7(PBuf).NameLen do
                WS[i] := WideChar(PRecBOUNDSHEET7(PBuf).Name[i - 1]);
//              Hidden := (PRecBOUNDSHEET7(PBuf).Options and $0300) shr 8;
          end
          else begin
            WS := ByteToWideString(@PRecBOUNDSHEET8(PBuf).Name,PRecBOUNDSHEET8(PBuf).NameLen);
//            Hidden := PRecBOUNDSHEET8(PBuf).Options and $0003;
          end;
          AList.Add(WS);

          InBoundsheet := True;
        end
        else if InBoundsheet then
          Break;
      end;
    except
      on EAbort do
        raise;
      on E: XLSRWException do
        raise XLSRWException.CreateFmt('Error on reading record: %.4X Offs: %.8X' + #13 + E.Message,[Header.RecID, FXLSStream.Pos]);
    end;
  finally
    FXLSStream.CloseStorage;
    FreeMem(PBuf);
  end;
end;

// Sheet cells

procedure TXLSReadII.RREC_BLANK;
begin
  InsertRecord := False;
  with PRecBLANK(PBuf)^ do begin
    FCurrXc12Sheet.Cells.StoreBlank(Col,Row,FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT);
  end;
end;

procedure TXLSReadII.RREC_ROW;
var
  Row: PXLSMMURowHeader;
begin
  InsertRecord := False;

  if (PRecROW(PBuf).Options and $0080) = $0080 then
    Row := FCurrXc12Sheet.Cells.AddRow(PRecROW(PBuf).Row,(PRecROW(PBuf).FormatIndex  and $0FFF) + XLS_STYLE_DEFAULT_XF_COUNT)
  else
    Row := FCurrXc12Sheet.Cells.AddRow(PRecROW(PBuf).Row,XLS_STYLE_DEFAULT_XF);

  Row.Height := PRecROW(PBuf).Height;
  Row.OutlineLevel := PRecROW(PBuf).Options and $0007;

  Row.Options := [];
  if (PRecROW(PBuf).Options and $0010) <> 0 then
    Row.Options := Row.Options + [xroCollapsed];
  if (PRecROW(PBuf).Options and $0020) <> 0 then
    Row.Options := Row.Options + [xroHidden];
  if (PRecROW(PBuf).Options and $0040) <> 0 then
    Row.Options := Row.Options + [xroCustomHeight];
end;

procedure TXLSReadII.RREC_BOOLERR;
begin
  InsertRecord := False;
  if PRecBOOLERR(PBuf).Error = 0 then begin
    if FManager.DirectRead then
      DoDirectRead(PRecBOOLERR(PBuf).Col,PRecBOOLERR(PBuf).Row,Boolean(PRecBOOLERR(PBuf).BoolErr))
    else
      FCurrXc12Sheet.Cells.StoreBoolean(PRecBOOLERR(PBuf).Col,PRecBOOLERR(PBuf).Row,PRecBOOLERR(PBuf).FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,Boolean(PRecBOOLERR(PBuf).BoolErr));
  end
  else begin
    if FManager.DirectRead then
      DoDirectRead(PRecBOOLERR(PBuf).Col,PRecBOOLERR(PBuf).Row,ErrorCodeToCellError(PRecBOOLERR(PBuf).BoolErr))
    else
      FCurrXc12Sheet.Cells.StoreError(PRecBOOLERR(PBuf).Col,PRecBOOLERR(PBuf).Row,PRecBOOLERR(PBuf).FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,ErrorCodeToCellError(PRecBOOLERR(PBuf).BoolErr));
  end;
end;

procedure TXLSReadII.ReadExcel40;
var
  i: integer;
begin
  SetLength(FEx40NumFmts,Length(Excel40StandardNumFormats));
  for i := 0 to High(Excel40StandardNumFormats) do
    FEx40NumFmts[i] := Excel40StandardNumFormats[i];

  FMaxFormat := FManager.StyleSheet.NumFmts.Count;
  FFormatCnt := 0;

  FXLSStream.Seek(0,soFromBeginning);

  FCurrXc12Sheet := FManager.Worksheets.Add(1);
  FCurrXc12Sheet.Name := 'Sheet1';

  if FManager.DirectRead and (FXLS.Sheets.Count < 1) then begin
    FXLS.Clear;
    FXLS.Sheets.Clear;
    FCurrSheet := 0;
  end;

  FXLS.Version := xvExcel50;

  while FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
    if Header.Length > FXLS.MaxBuffsize then
      Break;

    FXLSStream.Read(PBuf^, Header.Length);

    case Header.RecID of
      BIFFRECID_EOF      : Break;

      BIFFRECID_PALETTE   : RREC_PALETTE;
      BIFFRECID_FONT_40   : RREC_FONT_40;
      BIFFRECID_FORMAT    : RREC_FORMAT_40;
      BIFFRECID_XF_40     : RREC_XF_40;

      BIFFRECID_INTEGER_20: RREC_INTEGER_20;
      BIFFRECID_NUMBER_20 : RREC_NUMBER_20;
      BIFFRECID_LABEL_20  : RREC_LABEL_20;

      BIFFRECID_LABEL     : RREC_LABEL_X;
      BIFFRECID_RK7       : RREC_RK;
      BIFFRECID_MULRK     : RREC_MULRK;
      BIFFRECID_NUMBER    : RREC_NUMBER;
      BIFFRECID_BOOLERR   : RREC_BOOLERR;
      BIFFRECID_BLANK     : RREC_BLANK;
      BIFFRECID_MULBLANK  : RREC_MULBLANK;

      BIFFRECID_COLWIDTH  : RREC_COLWIDTH;

      BIFFRECID_COLINFO   : RREC_COLINFO;
      BIFFRECID_ROW       : RREC_ROW;
    end;
  end;
end;

procedure TXLSReadII.ReadFormulaVal(Col, Row, FormatIndex: integer; Value: double; Formula: PByteArray; Len, DataSz: integer);
var
  Fmla: PByteArray;
  S: AxUCString;
begin
  GetMem(Fmla, DataSz);
  try
    FCurrXc12Sheet.Cells.FormulaHelper.Clear;
    FCurrXc12Sheet.Cells.FormulaHelper.Col := Col;
    FCurrXc12Sheet.Cells.FormulaHelper.Row := Row;
    FCurrXc12Sheet.Cells.FormulaHelper.Style := FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT;
    if (DataSz - Len) > 0 then
      FCurrXc12Sheet.Cells.FormulaHelper.AllocPtgsArrayConsts97(PXLSPtgs(Formula),Len,DataSz - Len)
    else
      FCurrXc12Sheet.Cells.FormulaHelper.AllocPtgs97(PXLSPtgs(Formula),Len);
    // Formula points to PBuf, which is overwritten by the next read of the file.
    Move(Formula^, Fmla^, DataSz);
    case TByte8Array(Value)[0] of
      0: begin
          Header.RecID := FXLSStream.PeekHeader;
          if (Header.RecID = BIFFRECID_SHRFMLA) or (Header.RecID = $04BC) then begin
            raise XLSRWException.Create('Unxepected SHRFMLA');
            {
              ExtSz := SizeOf(TBIFFHeader) + Header.Length;
              FXLSStream.Read(PBuf^,Header.Length);
              FXLSStream.ReadHeader(Header);
            }
          end;
          if Header.RecID = BIFFRECID_STRING then begin
            FXLSStream.ReadHeader(Header);
            FXLSStream.Read(PBuf^, Header.Length);
            if FXLS.Version < xvExcel97 then
              S := Bit8StrToWideString(@PRecSTRING(PBuf).Data,PRecSTRING(PBuf).Len)
            else
              S := ByteStrToWideString(@PRecSTRING(PBuf).Data,PRecSTRING(PBuf).Len);
            {
              if ExtSz > 0 then begin
              Inc(ExtSz,SizeOf(TBIFFHeader) + Header.Length);
              FXLSStream.Seek2(-ExtSz,soFromCurrent);
              end;
            }
          end
          else if Header.RecID = BIFFRECID_STRING_20 then begin
            FXLSStream.ReadHeader(Header);
            FXLSStream.Read(PBuf^, Header.Length);
            S := ByteStrToWideString(@PRec2STRING(PBuf).Data,PRec2STRING(PBuf).Len);
          end
          else
            Exit;

        if FManager.DirectRead then
          DoDirectRead(Col,Row,S)
        else
          FCurrXc12Sheet.Cells.FormulaHelper.AsString := CheckFmlaStrVal(S);
        end;
      1: begin
        if FManager.DirectRead then
          DoDirectRead(Col,Row,Boolean(integer(TByte8Array(Value)[2])))
        else
          FCurrXc12Sheet.Cells.FormulaHelper.AsBoolean := Boolean(integer(TByte8Array(Value)[2]));
      end;
      2: begin
        if FManager.DirectRead then
          DoDirectRead(Col,Row,ErrorCodeToCellError(TByte8Array(Value)[2]))
        else
          FCurrXc12Sheet.Cells.FormulaHelper.AsError := ErrorCodeToCellError(TByte8Array(Value)[2]);
      end;
      // Undocumented: the value 3 indicates that the result is an empty string,
      // and therefore there is no STRING record following the formula.
      3: begin
        if FManager.DirectRead then
          DoDirectRead(Col,Row,'')
        else
          FCurrXc12Sheet.Cells.FormulaHelper.AsString := '';
      end;
    end;
    if not FManager.DirectRead then
      FCurrXc12Sheet.Cells.StoreFormula;
  finally
    FreeMem(Fmla);
  end;
end;

function TXLSReadII.GetNAMEString(var P: PByteArray; Len: integer): AxUCString;
begin
  Result := ByteToWideString(P,Len);
  if Result <> '' then begin
    if P[0] = 0 then
      P := PByteArray(NativeInt(P) + Len + 1)
    else
      P := PByteArray(NativeInt(P) + (Len * 2) + 1);
  end;
end;

procedure TXLSReadII.ReadSST(ARecSize: word);
var
  i: integer;
  Count: longword;
  TotalCount: longword;
begin
  Dec(ARecSize,FXLSStream.Read(TotalCount,4));
  Dec(ARecSize,FXLSStream.Read(Count,4));
  for i := 1 to Count do begin
    if ARecSize = 0 then
      ARecSize := SSTReadCONTINUE;
    SSTReadString(ARecSize);
  end;
end;

procedure TXLSReadII.RREC_FORMULA;
var
  WS: AxUCString;
  V: double;
  DataSz: integer;
  ArrayPtgsSz: integer;
  FirstFormula: boolean;

function ReadArraySTRING: boolean;
var
  pWS: PByteArray;
  Hdr: TBIFFHeader;
begin
  Result := FXLSStream.PeekHeader = BIFFRECID_STRING;
  if Result then begin
    FXLSStream.ReadHeader(Hdr);
    GetMem(pWS, Hdr.Length);
    try
      FXLSStream.Read(pWS^, Hdr.Length);
      WS := ByteStrToWideString(@PRecSTRING(pWS).Data, PRecSTRING(pWS).Len);
    finally
      FreeMem(pWS);
    end;
  end;
end;

begin
  Inc(L_Count);

  InsertRecord := False;
  with PRecFORMULA(PBuf)^ do begin
    DataSz := Header.Length - 22;
    // ATTN: Under some circumstances (don't know why) the Shared Formula bit is set
    // in the Options field, even if the formula not is part of a shared formula group.
    // There is no SHRFMLA record following the FORMULA record.
    // The formula contaions a complete expression.
    // This seems to only occure in XLS-97 files.
    // Bug in Excel?
    // One way to check this is to se if the first PTG in the formula is ptgExp, = part of shared formula.
    if ((Options and $0008) = $0008) and (Data[0] = xptgExp97) then begin
      Header.RecID := FXLSStream.PeekHeader;
      if (Header.RecID = BIFFRECID_SHRFMLA) or (Header.RecID = BIFFRECID_SHRFMLA_20) then begin
        FXLSStream.ReadHeader(Header);
        READ_SHRFMLA;
        FirstFormula := True;
      end
      else
        FirstFormula := False;

      FixupSharedFormula(PWordArray(@Data[1])[1], PWordArray(@Data[1])[0],Col, Row, FirstFormula);
      DataSz := ParseLen;
    end
    else if Data[0] = xptgExp97 then begin
      if FXLSStream.PeekHeader = BIFFRECID_ARRAY then begin
        FXLSStream.ReadHeader(Header);

        if FLastARRAY <> Nil then
          FreeMem(FLastARRAY);
        GetMem(FLastARRAY, Header.Length);

        FXLSStream.Read(FLastARRAY^, Header.Length);

        // Array ptgs is stored after the entire normal ptgs.
        ArrayPtgsSz := Header.Length - SizeOf(TRecARRAY) - PRecARRAY(FLastARRAY).DataSize;

        FCurrXc12Sheet.Cells.FormulaHelper.Clear;
        FCurrXc12Sheet.Cells.FormulaHelper.Col := Col;
        FCurrXc12Sheet.Cells.FormulaHelper.Row := Row;
        FCurrXc12Sheet.Cells.FormulaHelper.Style := FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT;
        FCurrXc12Sheet.Cells.FormulaHelper.SetTargetRef(PRecARRAY(FLastARRAY).Col1,PRecARRAY(FLastARRAY).Row1,PRecARRAY(FLastARRAY).Col2,PRecARRAY(FLastARRAY).Row2);
        FCurrXc12Sheet.Cells.FormulaHelper.HasApply := True;
        if ArrayPtgsSz > 0 then
          FCurrXc12Sheet.Cells.FormulaHelper.AllocPtgsArrayConsts97(@PRecARRAY(FLastARRAY).Data,PRecARRAY(FLastARRAY).DataSize,ArrayPtgsSz)
        else
          FCurrXc12Sheet.Cells.FormulaHelper.AllocPtgs97(@PRecARRAY(FLastARRAY).Data,PRecARRAY(FLastARRAY).DataSize);
        FCurrXc12Sheet.Cells.FormulaHelper.FormulaType := xcftArray;
        if (PRecARRAY(FLastARRAY).Options and $0001)  <> 0 then
          FCurrXc12Sheet.Cells.FormulaHelper.Options := FCurrXc12Sheet.Cells.FormulaHelper.Options + Xc12FormulaOpt_ACA;
        if ReadArraySTRING then begin
          FCurrXc12Sheet.Cells.FormulaHelper.AsString := CheckFmlaStrVal(WS);
        end
        else begin
          FCurrXc12Sheet.Cells.FormulaHelper.AsFloat := Value;
        end;
        FLastARRAYSize := Header.Length;

        if not FManager.DirectRead then
          FCurrXc12Sheet.Cells.StoreFormula;
        Exit;
      end
      else if (FLastARRAY <> Nil) and (Col >= PRecARRAY(FLastARRAY).Col1) and (Col <= PRecARRAY(FLastARRAY).Col2) and (Row >= PRecARRAY(FLastARRAY).Row1) and (Row <= PRecARRAY(FLastARRAY).Row2) then begin
        FCurrXc12Sheet.Cells.FormulaHelper.Clear;
        FCurrXc12Sheet.Cells.FormulaHelper.Col := Col;
        FCurrXc12Sheet.Cells.FormulaHelper.Row := Row;
        FCurrXc12Sheet.Cells.FormulaHelper.ParentCol := PRecARRAY(FLastARRAY).Col1;
        FCurrXc12Sheet.Cells.FormulaHelper.ParentRow := PRecARRAY(FLastARRAY).Row1;
        FCurrXc12Sheet.Cells.FormulaHelper.Style := FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT;
        FCurrXc12Sheet.Cells.FormulaHelper.FormulaType := xcftArrayChild97;
        if ReadArraySTRING then begin
          FCurrXc12Sheet.Cells.FormulaHelper.AsString := CheckFmlaStrVal(WS);
        end
        else begin
          FCurrXc12Sheet.Cells.FormulaHelper.AsFloat := Value;
        end;
        if not FManager.DirectRead then
          FCurrXc12Sheet.Cells.StoreFormula;
        Exit;
      end
      else begin
        if FLastARRAY <> Nil then
          FreeMem(FLastARRAY);
        FLastARRAY := Nil;
        FLastARRAYSize := 0;
      end;
    end;
    if (TByte8Array(Value)[0] in [$00, $01, $02, $03]) and (TByte8Array(Value)[6] = $FF) and (TByte8Array(Value)[7] = $FF) then
      ReadFormulaVal(Col, Row, FormatIndex, Value, @Data, ParseLen, DataSz)
    else if ParseLen <> 0 then begin
      ArrayPtgsSz := DataSz - ParseLen;

      FCurrXc12Sheet.Cells.FormulaHelper.Clear;
      FCurrXc12Sheet.Cells.FormulaHelper.Col := Col;
      FCurrXc12Sheet.Cells.FormulaHelper.Row := Row;
      FCurrXc12Sheet.Cells.FormulaHelper.Style := FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT;
      if ArrayPtgsSz > 0 then
        FCurrXc12Sheet.Cells.FormulaHelper.AllocPtgsArrayConsts97(@Data,ParseLen,ArrayPtgsSz)
      else
        FCurrXc12Sheet.Cells.FormulaHelper.AllocPtgs97(@Data,ParseLen);
      // This detects NAN values. A NAN number may cause an "Invalid Floating Point Operation" XLSRWException when used.
      if (TByte8Array(Value)[0] = $02) and (TByte8Array(Value)[6] = $FF) and (TByte8Array(Value)[7] = $FF) then
        V := 0
      else
        V := Value;

      if FManager.DirectRead then
        DoDirectRead(Col,Row,V)
      else
        FCurrXc12Sheet.Cells.FormulaHelper.AsFloat := V;
      if not FManager.DirectRead then
        FCurrXc12Sheet.Cells.StoreFormula;
    end;
  end;
end;

procedure TXLSReadII.RREC_NUMBER;
begin
  InsertRecord := False;
  with PRecNUMBER(PBuf)^ do begin
    if FManager.DirectRead then
      DoDirectRead(Col,Row,Value)
    else
      FCurrXc12Sheet.Cells.StoreFloat(Col,Row,FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,Value);
  end;
end;

procedure TXLSReadII.RREC_INTEGER_20;
begin
  InsertRecord := False;
  with PRec2INTEGER(PBuf)^ do begin
    if FManager.DirectRead then
      DoDirectRead(Col,Row,Value)
    else
      FCurrXc12Sheet.Cells.StoreFloat(Col,Row,XLS_STYLE_DEFAULT_XF,Value);
  end;
end;

procedure TXLSReadII.RREC_LABEL_20;
var
  S: XLS8String;
begin
  InsertRecord := False;
  with PRec2LABEL(PBuf)^ do begin
    SetLength(S, Len);
    Move(Data, Pointer(S)^, Len);
    if FManager.DirectRead then
      DoDirectRead(Col,Row,AxUCString(S))
    else
      FCurrXc12Sheet.Cells.StoreString(Col,Row,XLS_STYLE_DEFAULT_XF,AxUCString(S));
  end;
end;

procedure TXLSReadII.RREC_LABEL_X;
var
  i: integer;
  WS: AxUCString;
begin
  SetLength(WS,PRecLABEL(PBuf).Len);

  for i := 0 to PRecLABEL(PBuf).Len - 1 do
    WS[i + 1] := AxUCChar(PRecLABEL(PBuf).Data[i]);

  if FManager.DirectRead then
    DoDirectRead(PRecLABEL(PBuf).Col,PRecLABEL(PBuf).Row,WS)
  else
    FCurrXc12Sheet.Cells.StoreString(PRecLABEL(PBuf).Col,PRecLABEL(PBuf).Row,PRecLABEL(PBuf).FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,WS);
end;

procedure TXLSReadII.RREC_NUMBER_20;
begin
  InsertRecord := False;
  with PRec2NUMBER(PBuf)^ do begin
    if FManager.DirectRead then
      DoDirectRead(Col,Row,Value)
    else
      FCurrXc12Sheet.Cells.StoreFloat(Col,Row,XLS_STYLE_DEFAULT_XF,Value);
  end;
end;

procedure TXLSReadII.RREC_OBJPROTECT;
begin
//  FCurrXc12Sheet.SheetProtection.Objects := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_LABELSST;
begin
  InsertRecord := False;
  with PRecLABELSST(PBuf)^ do begin
    if FManager.DirectRead then
      DoDirectRead(Col,Row,FManager.SST.ItemText[SSTIndex])
    else
      FCurrXc12Sheet.Cells.StoreString(Col,Row,FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,SSTIndex);
  end;
end;

procedure TXLSReadII.RREC_LABEL;
var
  i: integer;
  P: PWordArray;
  WS: AxUCString;
begin
  InsertRecord := False;
  with PRecLABEL(PBuf)^ do begin
    if FXLS.Version >= xvExcel97 then begin
    if FManager.DirectRead then
      DoDirectRead(Col,Row,ByteStrToWideString(@Data[0], Len))
    else
      FCurrXc12Sheet.Cells.StoreString(Col,Row,FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,ByteStrToWideString(@Data[0], Len));
    end
    else begin
      SetLength(WS, Len);
      if Data[0] = 0 then begin
        for i := 1 to Len do
          WS[i] := AxUCChar(Data[i]);
      end
      else if Data[0] = 1 then begin
        P := @Data[1];
        for i := 0 to Len - 1 do
          WS[i + 1] := AxUCChar(P[i]);
      end
      else begin
        for i := 0 to Len - 1 do
          WS[i + 1] := AxUCChar(Data[i]);
      end;
      if FManager.DirectRead then
        DoDirectRead(Col,Row,WS)
      else
        FCurrXc12Sheet.Cells.StoreString(Col,Row,FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,WS);
    end;
  end;
end;

procedure TXLSReadII.RREC_RSTRING;
begin
  // RSTRING is only used when copying to/from the clipboard.
end;

procedure TXLSReadII.RREC_MULBLANK;
var
  i: integer;
begin
  InsertRecord := False;
  with PRecMULBLANK(PBuf)^ do begin
    for i := 0 to (Header.Length - 6) div 2 - 1 do begin
      FCurrXc12Sheet.Cells.StoreBlank(Col1 + i,Row,FormatIndexes[i] + XLS_STYLE_DEFAULT_XF_COUNT);
    end;
  end;
end;

procedure TXLSReadII.RREC_RK;
begin
  InsertRecord := False;
  with PRecRK(PBuf)^ do begin
    if FManager.DirectRead then
      DoDirectRead(Col,Row,DecodeRK(Value))
    else
      FCurrXc12Sheet.Cells.StoreFloat(Col,Row,FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT,DecodeRK(Value));
  end;
end;

procedure TXLSReadII.RREC_MULRK;
var
  i: integer;
begin
  InsertRecord := False;
  for i := 0 to (Header.Length - 6) div 6 - 1 do begin
    if FManager.DirectRead then
      DoDirectRead(PRecMULRK(PBuf).Col1 + i,PRecMULRK(PBuf).Row,DecodeRK(PRecMULRK(PBuf).RKs[i].RK))
    else
      FCurrXc12Sheet.Cells.StoreFloat(PRecMULRK(PBuf).Col1 + i,PRecMULRK(PBuf).Row,PRecMULRK(PBuf).RKs[i].XF + XLS_STYLE_DEFAULT_XF_COUNT,DecodeRK(PRecMULRK(PBuf).RKs[i].RK));
  end;
end;

procedure TXLSReadII.FixupSharedFormula(LeftCol, TopRow, ACol, ARow: integer; FirstFormula: boolean);
var
  i: integer;
begin
  if False { not FirstFormula } then
  begin
    // A shard formula identifies the definition it belongs to by the upper left
    // corner of the definition area. 070419 Not true. The entire shared formula
    // area has to be checked for hit, as below.
    for i := 0 to High(FSharedFormulas) do
    begin
      with FSharedFormulas[i] do
      begin
        if (LeftCol = Col1) and (TopRow = Row1) then
        begin
          Move(Formula^, PRecFORMULA(PBuf).Data, Len);
          PRecFORMULA(PBuf).ParseLen := Len;
          ConvertShrFmla(FXLS.Version = xvExcel97, @PRecFORMULA(PBuf).Data,PRecFORMULA(PBuf).ParseLen, ACol, ARow);
          Exit;
        end;
      end;
    end;
  end
  else
  begin
    for i := 0 to High(FSharedFormulas) do
    begin
      with FSharedFormulas[i] do
      begin
        // if (ACol >= Col1) and (ACol <= Col2) and (ARow >= Row1) and (ARow <= Row2) then begin
        if (LeftCol >= Col1) and (LeftCol <= Col2) and (TopRow >= Row1) and
          (TopRow <= Row2) then
        begin
          Move(Formula^, PRecFORMULA(PBuf).Data, Len);
          PRecFORMULA(PBuf).ParseLen := Len;
          ConvertShrFmla(FXLS.Version = xvExcel97, @PRecFORMULA(PBuf).Data,
            PRecFORMULA(PBuf).ParseLen, ACol, ARow);
          Exit;
        end;
      end;
    end;
  end;
  raise XLSRWException.CreateFmt('Fixup error in shared formula at C=%d,R=%d (%s)',
    [ACol, ARow, ColRowToRefStr(ACol, ARow, False, False)]);
end;

procedure TXLSReadII.READ_SHRFMLA;
var
  i: integer;
begin
  // ATTN: If Excel saves a SHRFMLA where some cells have been deleted,
  // (Ex: If you have 10 formulas in a row, saves it, open the sheet again,
  // deletes 4 of them and saves it again) then may the deleted formulas
  // still be saved in the SHRFMLA record!
  // The only way to check this is to look for the corresponding FORMULA
  // records. There is only FORMULA records for formulas that exsists on
  // the sheet. I think :-)

  SetLength(FSharedFormulas, Length(FSharedFormulas) + 1);
  i := High(FSharedFormulas);
  FXLSStream.Read(FSharedFormulas[i], 6);
  FXLSStream.Read(FSharedFormulas[i].Len, 2); // Reserved data
  FXLSStream.Read(FSharedFormulas[i].Len, 2);
  GetMem(FSharedFormulas[i].Formula, FSharedFormulas[i].Len);
  FXLSStream.Read(FSharedFormulas[i].Formula^, FSharedFormulas[i].Len);
end;

procedure TXLSReadII.RREC_NOTE;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then begin
    TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_SUFFIXDATA);
    with PRecNOTE(PBuf)^ do
      FXLS.Sheets[FCurrSheet]._Int_EscherDrawing.SetNoteData(Col, Row, Options, ObjId, ByteStrToWideString(PByteArray(@AuthorName), AuthorNameLen));
  end;
end;

procedure TXLSReadII.RREC_1904;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.WorkbookPr.Date1904 := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_AUTOFILTERINFO;
begin
  InsertRecord := False;
  if FXLS.Version >= xvExcel97 then
    FXLS.Sheets[FCurrSheet].Autofilters.LoadFromStream(FXLSStream, PBuf);
end;

procedure TXLSReadII.RREC_BACKUP;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.WorkbookPr.BackupFile := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_BOOKBOOL;
begin
  CurrRecs.UpdateDefault(Header, PBuf);
  InsertRecord := False;
  TRecordStorageGlobals(CurrRecs).UpdateInternal(INTERNAL_FORMATS);

  FManager.Workbook.WorkbookPr.SaveExternalLinkValues := (PWord(PBuf)^ and $0001) <> 0;
  FManager.Workbook.WorkbookPr.UpdateLinks := TXc12UpdateLinks((PWord(PBuf)^ and $0060) shr 5);
end;

procedure TXLSReadII.RREC_PALETTE;
var
  i: integer;
begin
  InsertRecord := False;

  if PRecPALETTE(PBuf).Count > (Length(Xc12IndexColorPalette) - 8) then
    PRecPALETTE(PBuf).Count := Length(Xc12IndexColorPalette) - 8;

  for i := 0 to PRecPALETTE(PBuf).Count + 8 - 1 do
    FManager.StyleSheet.InxColors.Add;

  for i := 0 to PRecPALETTE(PBuf).Count - 1 do begin
    FManager.StyleSheet.InxColors[i + 8].RGB := PRecPALETTE(PBuf).Color[i];
    Xc12IndexColorPalette[i + 8] := PRecPALETTE(PBuf).Color[i];
  end;
end;

procedure TXLSReadII.RREC_BOUNDSHEET;
var
  i: integer;
  WS: AxUCString;
  Hidden: byte;
  Sheet: TXc12DataWorksheet;
begin
  TRecordStorageGlobals(CurrRecs).UpdateInternal(INTERNAL_BOUNDSHEETS);
  InsertRecord := False;
  if FXLS.Version < xvExcel97 then begin
      SetLength(WS, PRecBOUNDSHEET7(PBuf).NameLen);
      for i := 1 to PRecBOUNDSHEET7(PBuf).NameLen do
        WS[i] := WideChar(PRecBOUNDSHEET7(PBuf).Name[i - 1]);
      Hidden := (PRecBOUNDSHEET7(PBuf).Options and $0300) shr 8;
  end
  else begin
    WS := ByteToWideString(@PRecBOUNDSHEET8(PBuf).Name,PRecBOUNDSHEET8(PBuf).NameLen);
    Hidden := PRecBOUNDSHEET8(PBuf).Options and $0003;
  end;
  FBoundsheets.AddObject(WS, TObject(Hidden));

  if (PRecBOUNDSHEET8(PBuf).Options and $0200) = 0 then begin
    FXLS.Sheets.Add(wtSheet);
    Sheet := FManager.Worksheets.Add(FManager.Worksheets.Count + 1);
    case Hidden of
      0: Sheet.State := x12vVisible;
      1: Sheet.State := x12vHidden;
      2: Sheet.State := x12vVeryHidden;
    end;
    Sheet.Name := WS;
  end;
end;

procedure TXLSReadII.RREC_CALCCOUNT;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.CalcPr.IterateCount := PWord(PBuf)^;
end;

procedure TXLSReadII.RREC_CALCMODE;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);

  case PSmallint(PBuf)^ of
    -1: FManager.Workbook.CalcPr.CalcMode := cmAutoExTables;
     0: FManager.Workbook.CalcPr.CalcMode := cmManual;
     1: FManager.Workbook.CalcPr.CalcMode := cmAutomatic;
  end;
end;

procedure TXLSReadII.RREC_CODEPAGE;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_COLINFO;
var
  Col: TXc12Column;
  XF: TXc12XF;
begin
//  if FXLS.Version = xvExcel97 then begin
    XF := FManager.StyleSheet.XFs[PRecCOLINFO(PBuf).FormatIndex + XLS_STYLE_DEFAULT_XF_COUNT];
    FManager.StyleSheet.XFEditor.UseStyle(XF);
    Col := FCurrXc12Sheet.Columns.Add(XF);
    Col.Min := PRecCOLINFO(PBuf).Col1;
    Col.Max := PRecCOLINFO(PBuf).Col2;
    Col.Width := PRecCOLINFO(PBuf).Width / 255;
    Col.OutlineLevel := (PRecCOLINFO(PBuf).Options shr 8) and $0007;

    Col.Options := [];
    if (PRecCOLINFO(PBuf).Options and $0001) <> 0 then
      Col.Options := Col.Options + [xcoHidden];
    if (PRecCOLINFO(PBuf).Options and $0002) <> 0 then
      Col.Options := Col.Options + [xcoCustomWidth];
    if (PRecCOLINFO(PBuf).Options and $0010) <> 0 then
      Col.Options := Col.Options + [xcoCollapsed];
//  end;
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_COLWIDTH;
var
  Col: TXc12Column;
  XF: TXc12XF;
begin
  XF := FManager.StyleSheet.XFs[0];
  FManager.StyleSheet.XFEditor.UseStyle(XF);

  Col := FCurrXc12Sheet.Columns.Add(XF);
  Col.Min := PRecCOLWIDTH(PBuf).Col1;
  Col.Max := PRecCOLWIDTH(PBuf).Col2;
  Col.Width := PRecCOLWIDTH(PBuf).Width / 255;
end;

procedure TXLSReadII.RREC_COUNTRY;
begin
  InsertRecord := False;
  if FXLS.Version >= xvExcel97 then
    CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_RECALCID;
begin
  InsertRecord := False;
  CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_DEFAULTROWHEIGHT;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    FCurrXc12Sheet.SheetFormatPr.DefaultRowHeight := PRecDEFAULTROWHEIGHT(PBuf).Height;
end;

procedure TXLSReadII.RREC_DEFCOLWIDTH;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
  begin
    // Fixes problem withe some non-excel files.
    if Header.Length = 4 then
      Header.Length := 2;
    FCurrXc12Sheet.SheetFormatPr.DefaultColWidth := PWord(PBuf)^;
  end;
end;

procedure TXLSReadII.RREC_DELTA;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.CalcPr.IterateDelta := PDouble(PBuf)^;
end;

procedure TXLSReadII.RREC_DIMENSIONS;
var
  C1, R1, C2, R2: integer;
begin
  InsertRecord := False;
  if FXLS.Version >= xvExcel97 then begin
    FCurrXc12Sheet.Cells.Dimension := SetCellArea(PRecDIMENSIONS8(PBuf).FirstCol,PRecDIMENSIONS8(PBuf).FirstRow,PRecDIMENSIONS8(PBuf).LastCol - 1,PRecDIMENSIONS8(PBuf).LastRow - 1);

    CurrRecs.UpdateDefault(Header, PBuf);
  end
  else if FXLS.Version > xvExcel40 then begin
    C1 := PRecDIMENSIONS7(PBuf).FirstCol;
    R1 := PRecDIMENSIONS7(PBuf).FirstRow;
    C2 := PRecDIMENSIONS7(PBuf).LastCol;
    R2 := PRecDIMENSIONS7(PBuf).LastRow;
    PRecDIMENSIONS8(PBuf).FirstCol := C1;
    PRecDIMENSIONS8(PBuf).FirstRow := R1;
    PRecDIMENSIONS8(PBuf).LastCol := C2;
    PRecDIMENSIONS8(PBuf).LastRow := R2;

    FCurrXc12Sheet.Cells.Dimension := SetCellArea(PRecDIMENSIONS8(PBuf).FirstCol,PRecDIMENSIONS8(PBuf).FirstRow,PRecDIMENSIONS8(PBuf).LastCol - 1,PRecDIMENSIONS8(PBuf).LastRow - 1);

    Header.Length := SizeOf(TRecDIMENSIONS8);
    CurrRecs.UpdateDefault(Header, PBuf);
  end;
end;

procedure TXLSReadII.RREC_DSF;
begin
  InsertRecord := False;
  CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_FNGROUPCOUNT;
begin
  InsertRecord := False;
  CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_EOF;
begin
  // This shall never happens.
  raise XLSRWException.Create('Unexpected EOF');
end;

procedure TXLSReadII.RREC_NAME;
var
  P: PByteArray;
  PDefDest: Pointer;
  PDefSrc: Pointer;
  Name: TXc12DefinedName;
  Rec: PRecNAME;
begin
  InsertRecord := False;

  if FXLS.Version = xvExcel97 then begin
    Rec := PRecNAME(PBuf);

    P := @Rec.Data;

    if ((Rec.Options and $0020) <> 0) then begin
      Name := FManager.Workbook.DefinedNames.Add(bnPrintArea,Rec.LocalSheetIndex - 1);
      P := Pointer(NativeInt(P) + 1);
      case PByte(P)^ of
        $00: Name.BuiltIn := bnConsolidateArea;
        $03: Name.BuiltIn := bnExtract;
        $04: Name.BuiltIn := bnDatabase;
        $05: Name.BuiltIn := bnCriteria;
        $06: Name.BuiltIn := bnPrintArea;
        $07: Name.BuiltIn := bnPrintTitles;
        $0C: Name.BuiltIn := bnSheetTitle;
        else Name.Unused97 := PByte(P)^;
      end;
      PDefSrc := Pointer(NativeInt(P) + 1);
    end
    else begin
      Name := FManager.Workbook.DefinedNames.Add(GetNAMEString(P,Rec.LenName),Rec.LocalSheetIndex - 1);
      PDefSrc := P;
      P := Pointer(NativeInt(P) + Rec.LenNameDef);
      if Rec.LenCustMenu   > 0 then Name.CustomMenu  := GetNAMEString(P,Rec.LenCustMenu);
      if Rec.LenDescText   > 0 then Name.Description := GetNAMEString(P,Rec.LenDescText);
      if Rec.LenHelpText   > 0 then Name.Help        := GetNAMEString(P,Rec.LenHelpText);
      if Rec.LenStatusText > 0 then Name.StatusBar   := GetNAMEString(P,Rec.LenStatusText);
    end;
//    Name.ShortcutKey := Rec.KeyShortcut;

    Name.Hidden      := (Rec.Options and $0001) <> 0;
    Name.Function_   := (Rec.Options and $0002) <> 0;
    Name.VbProcedure := (Rec.Options and $0004) <> 0;

    Name.FunctionGroupId := (Rec.Options and $0FC0) shr 5;

    GetMem(PDefDest,Rec.LenNameDef + SizeOf(TXLSPtgs));
    PXLSPtgs(PDefDest).Id := xptg_EXCEL_97;
    P := Pointer(NativeInt(PDefDest) + SizeOf(TXLSPtgs));
    Move(PDefSrc^,P^,Rec.LenNameDef);
    Name.Ptgs := PDefDest;
    Name.PtgsSz := Rec.LenNameDef + SizeOf(TXLSPtgs);
  end;
end;

procedure TXLSReadII.RREC_SUPBOOK;
begin
  case FXLS.Version of
    xvExcel97:
      begin
        TRecordStorageGlobals(CurrRecs).UpdateInternal(INTERNAL_NAMES);
        InsertRecord := False;
        FXLS.FormulaHandler.ExternalNames.SetSUPBOOK(PRecSUPBOOK(PBuf));
      end;
  end;
end;

procedure TXLSReadII.RREC_EXTERNNAME;
begin
  case FXLS.Version of
    xvExcel97:
      begin
        TRecordStorageGlobals(CurrRecs).UpdateInternal(INTERNAL_NAMES);
        InsertRecord := False;
        FXLS.FormulaHandler.ExternalNames.SetEXTERNNAME(PRecEXTERNNAME8(PBuf));
      end;
  end;
end;

procedure TXLSReadII.RREC_XCT;
var
  i, Count, Index: integer;
begin
  InsertRecord := False;
  Count := PRecXCT(PBuf).CRNCount;
  // Don't know why some Count are negative
  if Count > $7FFF then
    Count := $FFFF - Count + 1;
  Index := PRecXCT(PBuf).SheetIndex;
  for i := 0 to Count - 1 do begin
    // Skip any missing CRN records.
    if not (FXLSStream.PeekHeader = BIFFRECID_CRN) then
      Break;

    if (FXLSStream.ReadHeader(Header) <> SizeOf(TBIFFHeader)) or (Header.RecID <> BIFFRECID_CRN) then
      raise XLSRWException.Create('CRN record missing');
    FXLSStream.Read(PBuf^, Header.Length);
    FXLS.FormulaHandler.ExternalNames.SetCRN(Index, PRecCRN(PBuf),Header.Length);
  end;
end;

procedure TXLSReadII.RREC_EXTERNCOUNT;
begin
  // if FXLS.Version = xvExcel97 then
  // FXLS.NameDefs.ExternCount := PWord(PBuf)^;
end;

procedure TXLSReadII.RREC_EXTERNSHEET;
begin
  case FXLS.Version of
    xvExcel97:
      begin
        TRecordStorageGlobals(CurrRecs).UpdateInternal(INTERNAL_NAMES);
        FXLS.FormulaHandler.ExternalNames.SetEXTERNSHEET(PBuf);
        InsertRecord := False;
      end;
    xvExcel50:
      begin
      end;
  end;
end;

procedure TXLSReadII.RREC_FONT;
var
  Xc12Font: TXc12Font;
begin
  InsertRecord := False;

  Xc12Font := FManager.StyleSheet.Fonts.Add;

  Xc12Font.Name := ByteStrToWideString(@PRecFONT(PBuf).Name,PRecFONT(PBuf).NameLen);
  Xc12Font.Charset := PRecFONT(PBuf).CharSet;
  Xc12Font.Family := PRecFONT(PBuf).Family;

  if PRecFONT(PBuf).Bold >= $02BC then
    Xc12Font.Style := Xc12Font.Style + [xfsBold];
  if (PRecFONT(PBuf).Attributes and $02) = $02 then
    Xc12Font.Style := Xc12Font.Style + [xfsItalic];
  if (PRecFONT(PBuf).Attributes and $08) = $08 then
    Xc12Font.Style := Xc12Font.Style + [xfsStrikeout];
  Xc12Font.Underline := TXc12Underline(PRecFONT(PBuf).Underline);
  Xc12Font.SubSuperscript := TXc12SubSuperscript(PRecFONT(PBuf).SubSuperScript);

  if PRecFONT(PBuf).ColorIndex > $00FF then
    Xc12Font.Color := IndexColorToXc12(Integer(xc0))
  else
    Xc12Font.Color := IndexColorToXc12(PRecFONT(PBuf).ColorIndex);

  Xc12Font.Size := PRecFONT(PBuf).Height / 20;
  Xc12Font.Scheme := efsNone;

//  Xc12Font.Name := Font.Name;
//  Xc12Font.Charset := Font.Charset;
//  Xc12Font.Family := Font.Family;
//  Xc12Font.Style := TXc12FontStyles(Font.Style);
//  Xc12Font.Underline := TXc12Underline(Font.Underline);
//  Xc12Font.SubSuperscript := TXc12SubSuperscript(Font.SubSuperScript);
//  Xc12Font.Color := IndexColorToXc12(Integer(Font.ColorIndex));
//  Xc12Font.Size := Font.Size20 / 20;
//  Xc12Font.Scheme := efsMajor;

  if FManager.StyleSheet.Fonts.Count = 4 then
    FManager.StyleSheet.Fonts.Add;
end;

procedure TXLSReadII.RRec_FONT_40;
var
  S: XLS8String;
  Xc12Font: TXc12Font;
begin
  if FManager.StyleSheet.Fonts.Count = 4 then
    FManager.StyleSheet.Fonts.Add;

  Xc12Font := FManager.StyleSheet.Fonts.Add;

  Xc12Font.Size := PRecFONT4(PBuf).Height / 20;

  if (PRecFONT4(PBuf).Attributes and $0001) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsBold];
  if (PRecFONT4(PBuf).Attributes and $0002) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsItalic];
  if (PRecFONT4(PBuf).Attributes and $0004) <> 0 then
    Xc12Font.Underline := xulSingle;
  if (PRecFONT4(PBuf).Attributes and $0008) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsStrikeOut];
  if (PRecFONT4(PBuf).Attributes and $0010) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsOutline];
  if (PRecFONT4(PBuf).Attributes and $0020) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsShadow];
  if (PRecFONT4(PBuf).Attributes and $0040) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsCondense];
  if (PRecFONT4(PBuf).Attributes and $0080) <> 0 then
    Xc12Font.Style := Xc12Font.Style + [xfsExtend];

  Xc12Font.Color := IndexColorToXc12(PRecFONT4(PBuf).Color);

  SetLength(S,PRecFONT4(PBuf).NameLen);
  Move(PRecFONT4(PBuf).Name[0],Pointer(S)^,PRecFONT4(PBuf).NameLen);
  Xc12Font.Name := AxUCString(S);
end;

procedure TXLSReadII.RREC_HEADER;
begin
  InsertRecord := False;
  if FXLS.Version <= xvExcel40 then
    Exit;
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_HEADER);
  if FXLS.Version < xvExcel97 then begin
    Move(PBuf[1], PBuf[3], Header.Length - 1);
    PBuf[1] := 0;
    PBuf[2] := 0;
  end
  else if Header.Length > 0 then
    FCurrXc12Sheet.HeaderFooter.FirstHeader := ByteStrToWideString(@PByteArray(PBuf)[2], PWordArray(PBuf)[0]);
end;

procedure TXLSReadII.RREC_FOOTER;
begin
  InsertRecord := False;
  if FXLS.Version <= xvExcel40 then
    Exit;
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_HEADER);
  if FXLS.Version < xvExcel97 then begin
    Move(PBuf[1], PBuf[3], Header.Length - 1);
    PBuf[1] := 0;
    PBuf[2] := 0;
  end
  else if Header.Length > 0 then
    FCurrXc12Sheet.HeaderFooter.FirstFooter := ByteStrToWideString(@PByteArray(PBuf)[2], PWordArray(PBuf)[0]);
end;

procedure TXLSReadII.RREC_FORMAT;
begin
  InsertRecord := False;

  if FXLS.Version < xvExcel97 then
    FManager.StyleSheet.NumFmts.Add(Bit8StrToWideString(@PRecFORMAT7(PBuf).Data,PRecFORMAT7(PBuf).Len),PRecFORMAT7(PBuf).Index)
  else
    FManager.StyleSheet.NumFmts.Add(ByteStrToWideString(@PRecFORMAT8(PBuf).Data,PRecFORMAT8(PBuf).Len),PRecFORMAT8(PBuf).Index);
end;

procedure TXLSReadII.RRec_FORMAT_40;
var
  S: XLS8String;
begin
  SetLength(S,PRecFORMAT4(PBuf).Len);
  Move(PRecFORMAT4(PBuf).Data[0],Pointer(S)^,PRecFORMAT4(PBuf).Len);
  if FFormatCnt > High(FEx40NumFmts) then
    SetLength(FEx40NumFmts,FFormatCnt + 1);
  FEx40NumFmts[FFormatCnt] := AxUCString(S);

//  FManager.StyleSheet.NumFmts.Add(AxUCString(S),FMaxFormat + FFormatCnt);
  Inc(FFormatCnt);
end;

procedure TXLSReadII.RREC_GRIDSET;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FCurrXc12Sheet.PrintOptions.GridLinesSet := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_GUTS;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_HCENTER;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FCurrXc12Sheet.PrintOptions.HorizontalCentered := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_HIDEOBJ;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_HLINK;
var
  i: integer;
  P,P2: PByteArray;
  S: XLS8String;
  WS: AxUCString;
  HLink: TXc12Hyperlink;
  Header: TBIFFHeader;
  DirUpCnt: word;
  Options: word;
begin
  InsertRecord := False;
  if FXLS.Version < xvExcel97 then
    Exit;

  HLink := TXLSHyperlink(FCurrXc12Sheet.Hyperlinks.Add);

  HLink.Col1 := PRecHLINK(PBuf).Col1;
  HLink.Col2 := PRecHLINK(PBuf).Col2;
  HLink.Row1 := PRecHLINK(PBuf).Row1;
  HLink.Row2 := PRecHLINK(PBuf).Row2;
  Options := PRecHLINK(PBuf).Options;

  if ((Options and $0160) = 0) and ((Options and $0003) = $0003) then
    HLink.HyperlinkType := xhltURL
  else if ((Options and $0160) = 0) and ((Options and $0001) = $0001) then
    HLink.HyperlinkType := xhltFile
  else if ((Options and $0060) = 0) and ((Options and $0103) = $0103) then
    HLink.HyperlinkType := xhltUNC
  else if (Options and $0008) = $0008 then
    HLink.HyperlinkType := xhltWorkbook
  else begin
    HLink.HyperlinkType := xhltUnknown;
//    HLink.StoreUnknown(Len,PBuf);
    Exit;
  end;

  HLink.HyperlinkEnc := HLink.HyperlinkType;

  P := PByteArray(NativeInt(PBuf) + SizeOf(TRecHLINK));

  if CompareMem(P,@GUID_HLINK_URL,SizeOf(GUID_HLINK_URL)) or CompareMem(P,@GUID_HLINK_FILE,SizeOf(GUID_HLINK_FILE)) then begin
    Options := Options and not $0014;
    Options := Options and not $0080;
  end;

  if (Options and $0014) = $0014 then begin
    HLink.Display := BufUnicodeZToWS(@P[4],PLongwordArray(P)[0] * 2);
    P := PByteArray(Longword(P) + 4 + PLongwordArray(P)[0] * 2);
  end;
  if (Options and $0080) = $0080 then begin
    HLink.TargetFrame := BufUnicodeZToWS(@P[4],PLongwordArray(P)[0] * 2);
    P := PByteArray(Longword(P) + 4 + PLongwordArray(P)[0] * 2);
  end;

  if HLink.HyperlinkType in [xhltURL,xhltFile] then begin
    if CompareMem(P,@GUID_HLINK_URL,SizeOf(GUID_HLINK_URL)) then
      HLink.HyperlinkEnc := xhltURL
    else if CompareMem(P,@GUID_HLINK_FILE,SizeOf(GUID_HLINK_FILE)) then
      HLink.HyperlinkEnc := xhltFile;
  end;

  case HLink.HyperlinkEnc of
    xhltUnknown: begin
    end;
    xhltURL: begin
      P := PByteArray(NativeInt(P) + 16);
      HLink.RawAddress := BufUnicodeZToWS(@P[4],PLongwordArray(P)[0]);
      P := PByteArray(Longword(P) + 4 + PLongwordArray(P)[0]);
    end;
    xhltFile: begin
      P := PByteArray(NativeInt(P) + 16);
      DirUpCnt := PWordArray(P)[0];
      P := PByteArray(NativeInt(P) + 2);
      SetLength(S,PLongwordArray(P)[0] - 1);
      P := PByteArray(NativeInt(P) + 4);
      Move(P^,Pointer(S)^,Length(S));
      if S <> '' then begin
        P := PByteArray(NativeInt(P) + Length(S) + 1 + 24);
        if PLongwordArray(P)[0] > 0 then begin
          P2 := PByteArray(NativeInt(P) + 4);
          SetLength(WS,PLongwordArray(P2)[0] div 2);
          Move(P2[6],Pointer(WS)^,PLongwordArray(P2)[0]);
          HLink.RawAddress := WS;
        end
        else
          HLink.RawAddress := AxUCString(S);
      end
      else
        HLink.RawAddress := '';
      while DirUpCnt > 0 do begin
        HLink.RawAddress := HLink.RawAddress + '..\';
        Dec(DirUpCnt);
      end;
      P := PByteArray(Longword(P) + PLongwordArray(P)[0]);
    end;
    xhltUNC: begin
      HLink.RawAddress := BufUnicodeZToWS(@P[4],PLongwordArray(P)[0] * 2);
      P := PByteArray(Longword(P) + 4 + PLongwordArray(P)[0] * 2);
    end;
    xhltWorkbook: begin

    end;
  end;

  if (Options and $0008) = $0008 then begin
    // The reading of ScreenTip is probably not 100% correct. See OO test file.
    i := CPos('#',HLink.Display);
    if i > 1 then
      HLink.ScreenTip := Copy(HLink.Display,i + 1,MAXINT)
    else
      HLink.ScreenTip := BufUnicodeZToWS(@P[4],PLongwordArray(P)[0] * 2);
    if HLink.HyperlinkEnc = xhltWorkbook then
      HLink.RawAddress := HLink.ScreenTip;
  end;

  // Assume that the HLINKTOOLTIP belongs to the hyperlink.
  if FXLSStream.PeekHeader = BIFFRECID_HLINKTOOLTIP then begin
    FXLSStream.ReadHeader(Header);
    FXLSStream.Read(PBuf^,Header.Length);
    HLink.ToolTip := BufUnicodeZToWS(@PRecHLINKTOOLTIP(PBuf).Text,Header.Length - SizeOf(TRecHLINKTOOLTIP));
  end;
//  FXLS.Sheets[FCurrSheet].Hyperlinks.LoadFromStream(FXLSStream, Header.Length, PBuf, FCurrSheet);
end;

procedure TXLSReadII.RREC_FILEPASS;
var
  S: AxUCString;
begin
  if (PRecFILEPASS(PBuf).Options <> 1) or (PRecFILEPASS(PBuf).Options = 0) then
    FManager.Errors.Error('',XLSERR_FILEREAD_UNKNOWENCRYPT)
  else if not Assigned(FManager.OnPassword) and (FManager.Password = '') then begin
    if FXLSStream.SetReadDecrypt(PRecFILEPASS(PBuf),'VelvetSweatshop') then
      FXLS.Password := 'VelvetSweatshop'
    else
      FManager.Errors.Error('',XLSERR_FILEREAD_PASSWORDMISSING);
  end
  else begin
    S := FManager.Password;
    if Assigned(FManager.OnPassword) then
      FManager.OnPassword(FManager, S);
    if not FManager.Aborted then begin
      if S = '' then
        FManager.Errors.Error('',XLSERR_FILEREAD_PASSWORDMISSING)
      else if not FXLSStream.SetReadDecrypt(PRecFILEPASS(PBuf), S) then
        FManager.Errors.Error('',XLSERR_FILEREAD_WRONGPASSWORD)
    end;
  end;
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_FILESHAREING;
begin
  InsertRecord := False;
  FXLS.SetFILESHARING(PBuf, Header.Length);
end;

procedure TXLSReadII.RREC_FILTERMODE;
begin
  InsertRecord := False;
  if FXLS.Version >= xvExcel97 then
    FXLS.Sheets[FCurrSheet].Autofilters.LoadFromStream(FXLSStream, PBuf);
end;

procedure TXLSReadII.RREC_ITERATION;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.CalcPr.Iterate := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_MERGEDCELLS;
var
  i, Count: integer;
  Area: TXLSCellArea;
begin
  InsertRecord := False;

  Count := PRecMERGEDCELLS(PBuf).Count;
  for i := 0 to Count - 1 do begin
    SetCellArea(Area,PRecMERGEDCELLS(PBuf).Cells[i].Col1,PRecMERGEDCELLS(PBuf).Cells[i].Row1,PRecMERGEDCELLS(PBuf).Cells[i].Col2,PRecMERGEDCELLS(PBuf).Cells[i].Row2);
    // There can be duplicate merged cells. Bug in excel? Excel 2007 don't like duplicate merged cells.
    if FCurrXc12Sheet.MergedCells.FindArea(Area.Col1,Area.Row1,Area.Col2,Area.Row2) < 0 then
      FCurrXc12Sheet.MergedCells.Add(Area);
  end;
end;

procedure TXLSReadII.RREC_CONDFMT;
begin
  InsertRecord := False;
  FXLS.Sheets[FCurrSheet].ConditionalFormats.LoadFromStream(FXLSStream, PBuf);
end;

procedure TXLSReadII.RREC_MSODRAWING;
begin
  InsertRecord := False;

  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_SUFFIXDATA);
  FXLSStream.BeginCONTINUERead;
  FXLS.Sheets[FCurrSheet]._Int_EscherDrawing.LoadFromStream(FXLSStream, PBuf);
  FXLSStream.EndCONTINUERead;
end;

procedure TXLSReadII.RREC_MSODRAWINGGROUP;
begin
  InsertRecord := False;

  TRecordStorageGlobals(CurrRecs).UpdateInternal(INTERNAL_MSODRWGRP);
  FXLSStream.BeginCONTINUERead;
  try
    FXLS.MSOPictures.LoadFromStream(FXLSStream, PBuf);
  finally
    FXLSStream.EndCONTINUERead;
  end;
end;

procedure TXLSReadII.RREC_MSODRAWINGSELECTION;
begin
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_MSO_0866;
{
  var
  StreamType: word;
}
begin
  {
    if FCurrSheet < 0 then begin
    FXLSStream.BeginCONTINUERead;
    try
    FXLSStream.Read(PBuf^,12);
    FXLSStream.Read(StreamType,2);
    case StreamType of
    1: ;
    2: FXLS.PrintPictures.LoadFromStream(FXLSStream,PBuf);
    end;
    finally
    FXLSStream.EndCONTINUERead;
    end;
    end;
  }
end;

procedure TXLSReadII.RREC_PASSWORD;
begin
  // Silly password for not encrypted sheets.
  // password protected
  if FBoundsheetIndex < 0 then begin
    InsertRecord := False;
    FManager.Workbook.WorkbookProtection.WorkbookPassword := PWord(PBuf)^;
    CurrRecs.UpdateDefault(Header, PBuf);
  end
  else
    FCurrXc12Sheet.SheetProtection.Password := PWord(PBuf)^;
end;

procedure TXLSReadII.RREC_PRECISION;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.CalcPr.FullPrecision := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_PRINTGRIDLINES;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FCurrXc12Sheet.PrintOptions.GridLines := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_PRINTHEADERS;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FCurrXc12Sheet.PrintOptions.Headings := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_PROT4REV;
begin
  InsertRecord := False;
  CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_PROT4REVPASS;
begin

end;

procedure TXLSReadII.RREC_PROTECT;
begin
  InsertRecord := False;
  CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_REFMODE;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  if PWord(PBuf)^ = 0 then
    FManager.Workbook.CalcPr.RefMode := x12rmR1C1
  else
    FManager.Workbook.CalcPr.RefMode := x12rmA1;
end;

procedure TXLSReadII.RREC_REFRESHALL;
begin
  InsertRecord := False;
  CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.WorkbookPr.RefreshAllConnections := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_SAVERECALC;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FManager.Workbook.CalcPr.CalcOnSave := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_SELECTION;
var
  i: integer;
  P: PByteArray;
  Sel: TXc12Selection;
  Area: TCellArea;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then begin
//    if FCurrXc12Sheet.SheetViews[0].Selection.Count <= 0 then
      Sel := FCurrXc12Sheet.SheetViews[0].Selection.Add;
//    else
//      Sel := FCurrXc12Sheet.SheetViews[0].Selection[0];

    Sel.Pane := TXc12PaneEnum(PRecSELECTION_2(PBuf).Pane);
    Sel.ActiveCell := SetCellArea(PRecSELECTION_2(PBuf).ActiveCol,PRecSELECTION_2(PBuf).ActiveRow);
    Sel.ActiveCellId := PRecSELECTION_2(PBuf).ActiveRef;

    P := @PRecSELECTION_2(PBuf).Refs;
    for i := 0 to PRecSELECTION_2(PBuf).RefCount - 1 do begin
      if i > 0 then
        Area := Sel.SQRef.Add
      else
        Area := Sel.SQRef[0];

      Area.Col1 := PRecCellAreaShort(P).Col1;
      Area.Row1 := PRecCellAreaShort(P).Row1;
      Area.Col2 := PRecCellAreaShort(P).Col2;
      Area.Row2 := PRecCellAreaShort(P).Row2;
    end;
  end;
end;

procedure TXLSReadII.RREC_SHEETEXT;
begin
  FXLS.Sheets[FCurrSheet].TabColor := TXc12IndexColor(PWordArray(PBuf)[8]);
  FCurrXc12Sheet.SheetPr.TabColor := IndexColorToXc12(PRecSHEETEXT(PBuf).TabColor and $0000007F);
end;

procedure TXLSReadII.RREC_FEATHEADR;
begin
  // Global sheet protection? Seems only occure in Excel 2007 files saved in 97 format.
  if FCurrSheet < 0 then
    Exit;

  if PRecFEATHEADR(PBuf).SharedFeatureType <> 2 then
    Exit;

  InsertRecord := False;

  FCurrXc12Sheet.SheetProtection.Objects             := (PRecFEATHEADR(PBuf).Protections and $00000001) <> 0;
  FCurrXc12Sheet.SheetProtection.Scenarios           := (PRecFEATHEADR(PBuf).Protections and $00000002) <> 0;
  FCurrXc12Sheet.SheetProtection.FormatCells         := (PRecFEATHEADR(PBuf).Protections and $00000004) <> 0;
  FCurrXc12Sheet.SheetProtection.FormatColumns       := (PRecFEATHEADR(PBuf).Protections and $00000008) <> 0;
  FCurrXc12Sheet.SheetProtection.FormatRows          := (PRecFEATHEADR(PBuf).Protections and $00000010) <> 0;
  FCurrXc12Sheet.SheetProtection.InsertColumns       := (PRecFEATHEADR(PBuf).Protections and $00000020) <> 0;
  FCurrXc12Sheet.SheetProtection.InsertRows          := (PRecFEATHEADR(PBuf).Protections and $00000040) <> 0;
  FCurrXc12Sheet.SheetProtection.InsertHyperlinks    := (PRecFEATHEADR(PBuf).Protections and $00000080) <> 0;
  FCurrXc12Sheet.SheetProtection.DeleteColumns       := (PRecFEATHEADR(PBuf).Protections and $00000100) <> 0;
  FCurrXc12Sheet.SheetProtection.DeleteRows          := (PRecFEATHEADR(PBuf).Protections and $00000200) <> 0;
  FCurrXc12Sheet.SheetProtection.SelectLockedCells   := (PRecFEATHEADR(PBuf).Protections and $00000400) <> 0;
  FCurrXc12Sheet.SheetProtection.Sort                := (PRecFEATHEADR(PBuf).Protections and $00000800) <> 0;
  FCurrXc12Sheet.SheetProtection.AutoFilter          := (PRecFEATHEADR(PBuf).Protections and $00001000) <> 0;
  FCurrXc12Sheet.SheetProtection.PivotTables         := (PRecFEATHEADR(PBuf).Protections and $00002000) <> 0;
  FCurrXc12Sheet.SheetProtection.SelectUnlockedCells := (PRecFEATHEADR(PBuf).Protections and $00004000) <> 0;
end;

procedure TXLSReadII.RREC_DVAL;
begin
  FXLS.Sheets[FCurrSheet].Validations.LoadFromStream(FXLSStream, PBuf);
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_BOTTOMMARGIN;
begin
  InsertRecord := False;
  if FXLS.Version <= xvExcel40 then
    Exit;
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_MARGINS);
  FCurrXc12Sheet.PageMargins.Bottom := PRecMARGIN(PBuf).Value;
end;

procedure TXLSReadII.RREC_LEFTMARGIN;
begin
  InsertRecord := False;
  if FXLS.Version <= xvExcel40 then
    Exit;
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_MARGINS);
  FCurrXc12Sheet.PageMargins.Left := PRecMARGIN(PBuf).Value;
end;

procedure TXLSReadII.RREC_RIGHTMARGIN;
begin
  InsertRecord := False;
  if FXLS.Version <= xvExcel40 then
    Exit;
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_MARGINS);
  FCurrXc12Sheet.PageMargins.Right := PRecMARGIN(PBuf).Value;
end;

procedure TXLSReadII.RREC_TOPMARGIN;
begin
  InsertRecord := False;
  if FXLS.Version <= xvExcel40 then
    Exit;
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_MARGINS);
  FCurrXc12Sheet.PageMargins.Top := PRecMARGIN(PBuf).Value;
end;

procedure TXLSReadII.RREC_SETUP;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FCurrXc12Sheet.PageSetup.PaperSize := PRecSETUP(PBuf).PaperSize;
  FCurrXc12Sheet.PageSetup.Scale := PRecSETUP(PBuf).Scale;
  FCurrXc12Sheet.PageSetup.FirstPageNumber := PRecSETUP(PBuf).PageStart;
  FCurrXc12Sheet.PageSetup.FitToWidth := PRecSETUP(PBuf).FitWidth;
  FCurrXc12Sheet.PageSetup.FitToHeight := PRecSETUP(PBuf).FitHeight;
  FCurrXc12Sheet.PageSetup.HorizontalDpi := PRecSETUP(PBuf).Resolution;
  FCurrXc12Sheet.PageSetup.VerticalDpi := PRecSETUP(PBuf).VertResolution;
  FCurrXc12Sheet.PageMargins.Header := PRecSETUP(PBuf).HeaderMargin;
  FCurrXc12Sheet.PageMargins.Footer := PRecSETUP(PBuf).FooterMargin;
  FCurrXc12Sheet.PageSetup.Copies := PRecSETUP(PBuf).Copies;

  if (PRecSETUP(PBuf).Options and $0001) <> 0 then
    FCurrXc12Sheet.PageSetup.PageOrder := x12poOverThenDown
  else
    FCurrXc12Sheet.PageSetup.PageOrder := x12poDownThenOver;

  if (PRecSETUP(PBuf).Options and $0002) = 0 then
    FCurrXc12Sheet.PageSetup.Orientation := x12oLandscape
  else
    FCurrXc12Sheet.PageSetup.Orientation := x12oPortrait;

  FCurrXc12Sheet.PageSetup.BlackAndWhite := (PRecSETUP(PBuf).Options and $0008) <> 0;
  FCurrXc12Sheet.PageSetup.Draft := (PRecSETUP(PBuf).Options and $0010) <> 0;
  if (PRecSETUP(PBuf).Options and $0020) <> 0 then begin
    if (PRecSETUP(PBuf).Options and $0200) <> 0 then
      FCurrXc12Sheet.PageSetup.CellComments := x12ccAtEnd
    else
      FCurrXc12Sheet.PageSetup.CellComments := x12ccAsDisplayed;
  end
  else
    FCurrXc12Sheet.PageSetup.CellComments := x12ccNone;
  FCurrXc12Sheet.PageSetup.Errors := TXc12PrintError((PRecSETUP(PBuf).Options and $0C00) shr 10);

  // Below values are not valid in file.
  if (PRecSETUP(PBuf).Options and $0004) <> 0 then begin
    FCurrXc12Sheet.PageSetup.PaperSize := 9; // A4
    FCurrXc12Sheet.PageSetup.Scale := 100;
    FCurrXc12Sheet.PageSetup.HorizontalDpi := 600;
    FCurrXc12Sheet.PageSetup.VerticalDpi := 600;
    FCurrXc12Sheet.PageSetup.Copies := 1;
    FCurrXc12Sheet.PageSetup.Orientation := x12oPortrait;
  end;
end;

procedure TXLSReadII.RREC_SST;
begin
  InsertRecord := False;

  ReadSST(Header.Length);
end;

procedure TXLSReadII.RREC_EXTSST;
begin
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_STYLE;
begin
  InsertRecord := False;
  FXLS.Styles.Add(PRecSTYLE(PBuf));
end;

procedure TXLSReadII.RREC_VCENTER;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  FCurrXc12Sheet.PrintOptions.VerticalCentered := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_WINDOW1;
begin
  InsertRecord := False;
  if (FCurrSheet < 0) and (FXLS.Version >= xvExcel97) then begin
    CurrRecs.UpdateDefault(Header, PBuf);

    if FManager.Workbook.BookViews.Count <= 0 then
      FManager.Workbook.BookViews.Add;

    FManager.Workbook.BookViews[0].XWindow := PRecWINDOW1(PBuf).Left;
    FManager.Workbook.BookViews[0].YWindow := PRecWINDOW1(PBuf).Top;
    FManager.Workbook.BookViews[0].WindowWidth := PRecWINDOW1(PBuf).Width;
    FManager.Workbook.BookViews[0].WindowHeight := PRecWINDOW1(PBuf).Height;
    FManager.Workbook.BookViews[0].TabRatio := PRecWINDOW1(PBuf).TabRatio;
    FManager.Workbook.BookViews[0].ActiveTab := PRecWINDOW1(PBuf).SelectedTabIndex;
    FManager.Workbook.BookViews[0].FirstSheet := PRecWINDOW1(PBuf).FirstDispTabIndex;

    if (PRecWINDOW1(PBuf).Options and $01) <> 0 then
      FManager.Workbook.BookViews[0].Visibility := x12vHidden;

    FManager.Workbook.BookViews[0].Minimized := (PRecWINDOW1(PBuf).Options and $02) <> 0;
    FManager.Workbook.BookViews[0].ShowHorizontalScroll := (PRecWINDOW1(PBuf).Options and $08) <> 0;
    FManager.Workbook.BookViews[0].ShowVerticalScroll := (PRecWINDOW1(PBuf).Options and $10) <> 0;
    FManager.Workbook.BookViews[0].ShowSheetTabs := (PRecWINDOW1(PBuf).Options and $20) <> 0;
    FManager.Workbook.BookViews[0].AutoFilterDateGrouping := not (PRecWINDOW1(PBuf).Options and $40) <> 0;
  end;
end;

procedure TXLSReadII.RREC_WINDOW2;
var
  CSV: TXc12CustomSheetView;
begin
  if FXLS.Version <= xvExcel50 then begin
    PRecWINDOW2_8(PBuf).Zoom := 0;
    PRecWINDOW2_8(PBuf).ZoomPreview := 0;
    PRecWINDOW2_8(PBuf).Reserved := 0;
  end;
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then begin
    CurrRecs.UpdateDefault(Header, PBuf);
    FCurrXc12Sheet.SheetViews[0].TopLeftCell             := SetCellArea(PRecWINDOW2_8(PBuf).LeftCol,PRecWINDOW2_8(PBuf).TopRow);
    FCurrXc12Sheet.SheetViews[0].ColorId                 := PRecWINDOW2_8(PBuf).HeaderColorIndex;
    FCurrXc12Sheet.SheetViews[0].ZoomScaleNormal         := PRecWINDOW2_8(PBuf).Zoom;
    FCurrXc12Sheet.SheetViews[0].ZoomScalePageLayoutView := PRecWINDOW2_8(PBuf).ZoomPreview;
    FCurrXc12Sheet.SheetViews[0].ShowFormulas            := (PRecWINDOW2_8(PBuf).Options and $0001) <> 0;
    FCurrXc12Sheet.SheetViews[0].ShowGridLines           := (PRecWINDOW2_8(PBuf).Options and $0002) <> 0;
    FCurrXc12Sheet.SheetViews[0].ShowRowColHeaders       := (PRecWINDOW2_8(PBuf).Options and $0004) <> 0;
    FCurrXc12Sheet.SheetViews[0].ShowZeros               := (PRecWINDOW2_8(PBuf).Options and $0010) <> 0;
    FCurrXc12Sheet.SheetViews[0].DefaultGridColor        := (PRecWINDOW2_8(PBuf).Options and $0020) <> 0;
    FCurrXc12Sheet.SheetViews[0].RightToLeft             := (PRecWINDOW2_8(PBuf).Options and $0040) <> 0;
    FCurrXc12Sheet.SheetViews[0].ShowOutlineSymbols      := (PRecWINDOW2_8(PBuf).Options and $0080) <> 0;
    FCurrXc12Sheet.SheetViews[0].TabSelected             := (PRecWINDOW2_8(PBuf).Options and $0200) <> 0;

    if (PRecWINDOW2_8(PBuf).Options and $0008) <> 0 then
      FCurrXc12Sheet.SheetViews[0].Pane.State := x12psFrozen;

    if (PRecWINDOW2_8(PBuf).Options and $0800) <> 0 then begin
      CSV := FCurrXc12Sheet.CustomSheetViews.Add;
      CSV.View := x12svtPageBreakPreview;
    end;
  end;
end;

procedure TXLSReadII.RREC_SCENPROTECT;
begin
//  FCurrXc12Sheet.SheetProtection.Scenarios := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_SCL;
begin
  FXLS.Sheets[FCurrSheet].Zoom := Round((PRecSCL(PBuf).Numerator / PRecSCL(PBuf).Denominator) * 100);
end;

procedure TXLSReadII.RREC_PANE;
begin
  InsertRecord := False;

  FCurrXc12Sheet.SheetViews[0].Pane.Excel97 := True;
  FCurrXc12Sheet.SheetViews[0].Pane.ActivePane := TXc12PaneEnum(PRecPANE(PBuf).PaneNumber);
  FCurrXc12Sheet.SheetViews[0].Pane.XSplit := PRecPANE(PBuf).X / 20;
  FCurrXc12Sheet.SheetViews[0].Pane.YSplit := PRecPANE(PBuf).Y / 20;
  FCurrXc12Sheet.SheetViews[0].Pane.TopLeftCell := SetCellArea(PRecPANE(PBuf).LeftCol,PRecPANE(PBuf).TopRow);
end;

procedure TXLSReadII.RREC_WINDOWPROTECT;
begin
  InsertRecord := False;
  if FXLS.Version > xvExcel40 then
    CurrRecs.UpdateDefault(Header, PBuf);
  // Not sure if this is correct.
  FManager.Workbook.WorkbookProtection.LockWindows := PWord(PBuf)^ = 1;
end;

procedure TXLSReadII.RREC_WRITEACCESS;
begin
  InsertRecord := False;
  if FXLS.Version < xvExcel97 then begin
    Move(PBuf[1], PBuf[3], Header.Length - 2);
    PBuf[1] := 0;
    PBuf[2] := 0;
  end;
  if Header.Length < 112 then begin
    FillChar(PBuf[Header.Length], 112 - Header.Length, ' ');
    Header.Length := 112;
  end;
  FManager.Workbook.UserName := ByteStrToWideString(@PBuf[2],PWord(PBuf)^);
  CurrRecs.UpdateDefault(Header, PBuf);
end;

procedure TXLSReadII.RREC_WSBOOL;
begin
  if (FXLS.Version > xvExcel40) and not (CurrRecs is TRecordStorageGlobals) then begin
    InsertRecord := False;
    CurrRecs.UpdateDefault(Header, PBuf);
    FCurrXc12Sheet.SheetPr.PageSetupPr.AutoPageBreaks := (PWord(PBuf)^ and $0001) <> 0;
    FCurrXc12Sheet.SheetPr.PageSetupPr.FitToPage := (PWord(PBuf)^ and $0100) <> 0;
    FCurrXc12Sheet.SheetPr.OutlinePr.ApplyStyles := (PWord(PBuf)^ and $0020) <> 0;
    FCurrXc12Sheet.SheetPr.OutlinePr.SummaryBelow := (PWord(PBuf)^ and $0040) <> 0;
    FCurrXc12Sheet.SheetPr.OutlinePr.SummaryRight := (PWord(PBuf)^ and $0080) <> 0;
  end;
end;

procedure TXLSReadII.RREC_HORIZONTALPAGEBREAKS;
var
  i: integer;
  Brk: TXc12Break;
begin
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_PAGEBREAKES);
  if FXLS.Version >= xvExcel97 then begin
    for i := 0 to PRecHORIZONTALPAGEBREAKS(PBuf).Count - 1 do begin
      Brk := FCurrXc12Sheet.RowBreaks.Add;
      Brk.Id := PRecHORIZONTALPAGEBREAKS(PBuf).Breaks[i].Val1 - 1;
      Brk.Min := PRecHORIZONTALPAGEBREAKS(PBuf).Breaks[i].Val2 - 1;
      Brk.Max := PRecHORIZONTALPAGEBREAKS(PBuf).Breaks[i].Val3 - 1;
    end;
  end
  else if FXLS.Version = xvExcel50 then begin
    for i := 1 to PWordArray(PBuf)[0] - 1 do begin
      Brk := FCurrXc12Sheet.RowBreaks.Add;
      Brk.Id := PWordArray(PBuf)[i] - 1;
      Brk.Min := 0;
      Brk.Max := MAXROW_97;
    end;
  end;
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_VERTICALPAGEBREAKS;
var
  i: integer;
  Brk: TXc12Break;
begin
  TRecordStorageSheet(CurrRecs).UpdateInternal(INTERNAL_PAGEBREAKES);
  if FXLS.Version >= xvExcel97 then begin
    for i := 0 to PRecVERTICALPAGEBREAKS(PBuf).Count - 1 do begin
      Brk := FCurrXc12Sheet.ColBreaks.Add;
      Brk.Id := PRecVERTICALPAGEBREAKS(PBuf).Breaks[i].Val1 - 1;
      Brk.Min := PRecVERTICALPAGEBREAKS(PBuf).Breaks[i].Val2 - 1;
      Brk.Max := PRecVERTICALPAGEBREAKS(PBuf).Breaks[i].Val3 - 1;
    end;
  end
  else if FXLS.Version = xvExcel50 then begin
    for i := 1 to PWordArray(PBuf)[0] - 1 do begin
      Brk := FCurrXc12Sheet.ColBreaks.Add;
      Brk.Id := PWordArray(PBuf)[i] - 1;
      Brk.Min := 0;
      Brk.Max := MAXROW_97;
    end;
  end;
  InsertRecord := False;
end;

procedure TXLSReadII.RREC_XF;
var
  XF: TXc12XF;
  X7: PRecXF7;
  X8: PRecXF8;
  W: word;
begin
  InsertRecord := False;

  if FXLS.Version = xvExcel50 then begin
    XF := TXc12XF.Create(Nil);

    X7 := PRecXF7(PBuf);

    if (X7.FontIndex + XLS_STYLE_DEFAULT_FONT_COUNT) < FManager.StyleSheet.Fonts.Count  then
      XF.Font := FManager.StyleSheet.Fonts[X7.FontIndex + XLS_STYLE_DEFAULT_FONT_COUNT]
    else
      XF.Font := FManager.StyleSheet.Fonts[0];

    if X7.FormatIndex < FManager.StyleSheet.NumFmts.Count then
      XF.NumFmt := FManager.StyleSheet.NumFmts[X7.FormatIndex]
    else
      XF.NumFmt := FManager.StyleSheet.NumFmts[0];

    XF.Fill := FManager.StyleSheet.Fills[0];
    XF.Border := FManager.StyleSheet.Borders[0];

    FManager.StyleSheet.XFs.Add(XF);
  end
  else if FXLS.Version = xvExcel97 then begin
    XF := TXc12XF.Create(Nil);

    X8 := PRecXF8(PBuf);

    XF.E97XfData1 := X8.Data1;
//    XF.E97XfData2 := X8.Data2;
//    XF.E97XfData3 := X8.Data3;
//    XF.E97XfData4 := X8.Data4;
//    XF.E97XfData5 := X8.Data5;
//    XF.E97XfData6 := X8.Data6;
//    XF.E97XfData7 := X8.Data7;

    XF.Font := FManager.StyleSheet.Fonts[X8.FontIndex + XLS_STYLE_DEFAULT_FONT_COUNT];

    XF.NumFmt := FManager.StyleSheet.NumFmts[X8.NumFmtIndex];

    XF.Protection := TXc12CellProtections(Byte(X8.Data1 and $0003));

    XF.Alignment.HorizAlignment := TXc12HorizAlignment(X8.Data2 and $0007);
    XF.Alignment.VertAlignment := TXc12VertAlignment((X8.Data2 and $0070) shr 4);
    XF.Alignment.Indent := X8.Data3 and $000F;
    XF.Alignment.Rotation := X8.Data2 shr 8;
//    if (XF.Alignment.Rotation > 90) and (XF.Alignment.Rotation <> 255) then
//      XF.Alignment.Rotation := -(XF.Alignment.Rotation - 90);
    XF.Alignment.Options := [];
    if ((X8.Data2 and $0008) shr 3) > 0 then
      XF.Alignment.Options := [foWrapText];
    if ((X8.Data3 and $0010) shr 4) > 0 then
      XF.Alignment.Options := XF.Alignment.Options + [foShrinkToFit];

    XF.Fill := FManager.StyleSheet.Fills.Add;
    XF.Fill.FgColor := IndexColorToXc12((X8.Data7 and $007F) shr 0);

    W := X8.Data7 and $3F80;
    W := W shr 7;
    if W > Word(Ord(xcAutomatic)) then
      W := Ord(xcAutomatic);
    XF.Fill.BgColor := IndexColorToXc12(W);

    XF.Fill.PatternType := TXc12FillPattern((X8.Data6 and $FC000000) shr 26);

    XF.Border := FManager.StyleSheet.Borders.Add;

    XF.Border.Top.Color := IndexColorToXc12((X8.Data6 and $0000007F) shr 0);
    XF.Border.Top.Style := TXc12CellBorderStyle((X8.Data4 and $0F00) shr 8);

    XF.Border.Left.Color := IndexColorToXc12((X8.Data5 and $007F) shr 0);
    XF.Border.Left.Style := TXc12CellBorderStyle((X8.Data4 and $000F) shr 0);

    XF.Border.Right.Color := IndexColorToXc12((X8.Data5 and $3F80) shr 7);
    XF.Border.Right.Style := TXc12CellBorderStyle((X8.Data4 and $00F0) shr 4);

    XF.Border.Bottom.Color := IndexColorToXc12((X8.Data6 and $00003F80) shr 7);
    XF.Border.Bottom.Style := TXc12CellBorderStyle((X8.Data4 and $F000) shr 12);

    XF.Border.Options := [];
    case (X8.Data5 and $C000) shr 14 of
      1: XF.Border.Options := [ecboDiagonalDown];
      2: XF.Border.Options := [ecboDiagonalUp];
      3: XF.Border.Options := [ecboDiagonalDown,ecboDiagonalUp];
    end;

//    XF.Border.Diagonal.Color := IndexColorToXc12((X8.Data6 and $001FC000) shr 14);
//    XF.Border.Diagonal.Style := TXc12CellBorderStyle((X8.Data6 and $01E00000) shr 21);

    FManager.StyleSheet.XFs.Add(XF);
    if FManager.StyleSheet.XFs.Count <= (XLS_STYLE_DEFAULT_XF_COUNT + XLS_STYLE_DEFAULT_XF_COUNT_97) then
      XF.Locked := True;

    FManager.StyleSheet.XFEditor.UseStyle(XF);
  end;
end;

procedure TXLSReadII.RREC_XF_40;
var
  i: integer;
  XF: TXc12XF;
  Pattern: word;
  FgColor: word;
  BgColor: word;
  Fill: TXc12Fill;
  Border: TXc12Border;

procedure DoBorder(const ABorder: byte; AXc12Border: TXc12BorderPr);
var
  Color: word;
  Style: byte;
begin
  if ABorder <> 0 then begin
    Style := ABorder and $07;
    Color := (ABorder and $F8) shr 3;
    AXc12Border.Color := IndexColorToXc12(Color);
    AXc12Border.Style := TXc12CellBorderStyle(Style);
  end;
end;

begin
  if not FEx40NumFmtsSaved then begin
    for i := 0 to High(FEx40NumFmts) do
      FManager.StyleSheet.NumFmts.Add(FEx40NumFmts[i],FMaxFormat + i);
    FEx40NumFmtsSaved := True;
  end;

  XF := FManager.StyleSheet.XFs.Add;

  if (PRecXF4(PBuf).FontIndex + XLS_STYLE_DEFAULT_FONT_COUNT) < FManager.StyleSheet.Fonts.Count then
    XF.Font := FManager.StyleSheet.Fonts[PRecXF4(PBuf).FontIndex + XLS_STYLE_DEFAULT_FONT_COUNT]
  else
    XF.Font := FManager.StyleSheet.Fonts.DefaultFont;

  if (PRecXF4(PBuf).FormatIndex + FMaxFormat) < FManager.StyleSheet.NumFmts.Count then
    XF.NumFmt := FManager.StyleSheet.NumFmts[PRecXF4(PBuf).FormatIndex + FMaxFormat]
  else
    XF.NumFmt := FManager.StyleSheet.NumFmts.DefaultNumFmt;

  Pattern := PRecXF4(PBuf).CellColor and $003F;
  FgColor := (PRecXF4(PBuf).CellColor and $07C0) shr 6;
  BgColor := (PRecXF4(PBuf).CellColor and $F800) shr 11;

  if (FgColor <= $17) or (BgColor <= $17) then begin
    Fill := FManager.StyleSheet.Fills.Add;
    if FgColor <= $17 then
      Fill.FgColor := IndexColorToXc12(FgColor);
    if BgColor <= $17 then
      Fill.BgColor := IndexColorToXc12(BgColor);
    if Pattern <= Integer(High(TXc12FillPattern)) then
      Fill.PatternType := TXc12FillPattern(Pattern);
    XF.Fill := Fill;
  end
  else
    XF.Fill := FManager.StyleSheet.Fills.DefaultFill;

  if (PRecXF4(PBuf).LeftBorder > 0) or (PRecXF4(PBuf).RightBorder > 0) or (PRecXF4(PBuf).TopBorder > 0) or (PRecXF4(PBuf).BottomBorder > 0) then begin
    Border := FManager.StyleSheet.Borders.Add;
    DoBorder(PRecXF4(PBuf).LeftBorder,Border.Left);
    DoBorder(PRecXF4(PBuf).RightBorder,Border.Right);
    DoBorder(PRecXF4(PBuf).TopBorder,Border.Top);
    DoBorder(PRecXF4(PBuf).BottomBorder,Border.Bottom);
    XF.Border := Border;
  end
  else
    XF.Border := FManager.StyleSheet.Borders.DefaultBorder;

  XF.Alignment.HorizAlignment := TXc12HorizAlignment(PRecXF4(PBuf).Alignment and $07);
  if (PRecXF4(PBuf).Alignment and $08) <> 0 then
    XF.Alignment.Options := [foWrapText];
  XF.Alignment.VertAlignment := TXc12VertAlignment((PRecXF4(PBuf).Alignment and $30) shr 4);

  case (PRecXF4(PBuf).Alignment and $C0) shr 6 of
    2: XF.Alignment.Rotation := 90;
    3: XF.Alignment.Rotation := 180;
  end;

end;

procedure TXLSReadII.SkipChart;
var
  p: integer;
begin
  if FSkipMSO and (FXLSStream.PeekHeader = BIFFRECID_OBJ) then begin
    p := FXLSStream.Pos;
    FXLSStream.ReadHeader(Header);
    FXLSStream.Read(PBuf^,Header.Length);
    if (PRecOBJ(PBuf).ObjType = OBJTYPE_CHART) and ((FXLSStream.PeekHeader and $FF) = BIFFRECID_BOF) then begin
      CurrRecs.AddRec(Header, PBuf);
      while FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader) do begin
        if Header.RecID = BIFFRECID_EOF then
          Break;
        FXLSStream.Read(PBuf^,Header.Length);
        CurrRecs.AddRec(Header, PBuf);
      end;
    end
    else
      FXLSStream.Seek(p,soFromBeginning);
  end;
end;

function TXLSReadII.SSTReadCONTINUE: word;
var
  Header: TBIFFHeader;
begin
  FXLSStream.ReadHeader(Header);
  if Header.RecID <> BIFFRECID_CONTINUE then
    raise XLSRWException.Create('CONTINUE record is missing in SST');
  Result := Header.Length;
end;

function TXLSReadII.SSTReadString(var RecSize: word): PByteArray;
var
  i: integer;
  S: AxUCString;
  Len,FmtCount: word;
  Options: byte;
  MemLen,StrMemLen: word;
  FarEastDataSize: longword;
  FontRuns: TXc12DynFontRunArray;

procedure ReadSplittedString(P: PByteArray; Len: word; Unicode: boolean);
type
  PStrPart = ^TStrPart;
  TStrPart = record
  Unicode: byte;
  Len: word;
  PStr: PByteArray;
  end;
var
  i,j: integer;
  Parts: TList;
  SPart: PStrPart;

procedure ReadPart(Opt: byte);
var
  BytesToRead: integer;
begin
  if Opt = $FF then
    Dec(RecSize,FXLSStream.Read(Opt,1));
  if Opt = $01 then begin
    BytesToRead := Len * 2;
    Options := Options or $01;
  end
  else
    BytesToRead := Len;
  if BytesToRead > RecSize then
    BytesToRead := RecSize;
  New(SPart);
  SPart.Unicode := Opt;
  GetMem(SPart.PStr,BytesToRead);
  Dec(RecSize,FXLSStream.Read(SPart.PStr^,BytesToRead));
  if Opt = $01 then
    SPart.Len := BytesToRead div 2
  else
    SPart.Len := BytesToRead;
  Dec(Len,SPart.Len);
  Parts.Add(SPart);
end;

begin
  Parts := TList.Create;
  try
    ReadPart(Byte(Unicode));
    while Len > 0 do begin
      RecSize := SSTReadCONTINUE;
      ReadPart($FF);
    end;
    for i := 0 to Parts.Count - 1 do begin
      if (Options and $01) = $01 then begin
        if (PStrPart(Parts[i]).Unicode and $01) = $01 then
          Move(PStrPart(Parts[i]).PStr^,P^,PStrPart(Parts[i]).Len * 2)
        else begin
          for j := 0 to PStrPart(Parts[i]).Len - 1 do
            PWordArray(P)[j] := PStrPart(Parts[i]).PStr[j];
        end;
        P := PByteArray(NativeInt(P) + PStrPart(Parts[i]).Len * 2);
      end
      else begin
        if (PStrPart(Parts[i]).Unicode and $01) = $01 then
          raise XLSRWException.Create('SST split error: unicode part in compressed string.');
        Move(PStrPart(Parts[i]).PStr^,P^,PStrPart(Parts[i]).Len);
        P := PByteArray(NativeInt(P) + PStrPart(Parts[i]).Len);
      end;
      FreeMem(PStrPart(Parts[i]).PStr);
      FreeMem(PStrPart(Parts[i]));
    end;
  finally
    Parts.Free;
  end;
end;

procedure ReadString(P: PByteArray; Len: word; Unicode: boolean);
begin
  if MemLen > RecSize then
    ReadSplittedString(P,Len,Unicode)
  else
    Dec(RecSize,FXLSStream.Read(P^,MemLen));
end;

procedure ReadData(P: PByteArray; Len: word);
begin
  if Len > RecSize then begin
    FXLSStream.Read(P^,RecSize);
    Dec(Len,RecSize);
    P := PByteArray(NativeInt(P) + RecSize);
    RecSize := SSTReadCONTINUE;
    ReadData(P,Len);
  end
  else
    Dec(RecSize,FXLSStream.Read(P^,Len));
end;

function ConvertFormatRunsDynFontRuns(P: PByteArray; Count: integer): TXc12DynFontRunArray;
var
  i: integer;
  pFmt: PFormatRun;
begin
  SetLength(Result,Count);
  for i := 0 to Count - 1 do begin
    pFmt := PFormatRun(NativeInt(P) + (i * SizeOf(TFormatRun)));
    Result[i].Index := pFmt.Index;
    Result[i].Font := FManager.StyleSheet.Fonts[pFmt.FontIndex + XLS_STYLE_DEFAULT_FONT_COUNT];
  end;
end;

procedure CheckIfSplittedString;
var
  L: integer;
  AdjOptions: byte;
begin
  L := 0;
  case Options of
    STRID_COMPRESSED,
    STRID_UNICODE:
      L := 0;
    STRID_RICH,
    STRID_RICH_UNICODE:
      L := 2;
    STRID_FAREAST,
    STRID_FAREAST_UC:
      L := 4;
    STRID_FAREAST_RICH,
    STRID_FAREAST_RICH_UC:
      L := 5;
  end;
  if (Options and STRID_UNICODE) = STRID_UNICODE then
    Inc(L,Len * 2)
  else
    Inc(L,Len);
  if L > RecSize then
    AdjOptions := Options or $01
  else
    AdjOptions := Options;
  if (AdjOptions and $01) = $01 then
    MemLen := Len * 2
  else
    MemLen := Len;
end;

begin
//  Result := Nil;
  Dec(RecSize,FXLSStream.Read(Len,2));
  Dec(RecSize,FXLSStream.Read(Options,1));
  CheckIfSplittedString;

  if (Options and $01) = $01 then
    StrMemLen := Len * 2
  else
    StrMemLen := Len;

//  StrMemLen := MemLen;

  case Options of
    STRID_COMPRESSED: begin
      GetMem(Result,SizeOf(TXLSString) + MemLen);
      ReadString(PByteArray(@PXLSString(Result).Data),Len,False);

      FManager.SST.AddRawString(Len,PByteArray(@PXLSString(Result).Data));
    end;
    STRID_UNICODE: begin
      GetMem(Result,SizeOf(TXLSStringUC) + MemLen);
      ReadString(@PXLSStringUC(Result).Data,Len,True);

      FManager.SST.AddRawString(Len,PWordArray(@PXLSString(Result).Data));
    end;
    STRID_RICH: begin
      Dec(RecSize,FXLSStream.Read(FmtCount,2));
      GetMem(Result,SizeOf(TXLSStringRich) + MemLen + FmtCount * SizeOf(TXc12FontRun));
      ReadString(PByteArray(@PXLSStringRich(Result).Data),Len,False);
      ReadData(Pointer(NativeInt(@PXLSStringRich(Result).Data) + StrMemLen),FmtCount * SizeOf(TFormatRun));
      FontRuns := ConvertFormatRunsDynFontRuns(Pointer(NativeInt(@PXLSStringRich(Result).Data) + StrMemLen),FmtCount);

      SetLength(S,Len);
      for i := 0 to Len - 1 do
        S[i + 1] := AxUCChar(PByteArray(@PXLSStringRich(Result).Data)[i]);

      FManager.SST.AddFormattedString(S,FontRuns);
    end;
    STRID_RICH_UNICODE: begin
      Dec(RecSize,FXLSStream.Read(FmtCount,2));
      GetMem(Result,SizeOf(TXLSStringRichUC) + MemLen + FmtCount * SizeOf(TXc12FontRun));
      ReadString(@PXLSStringRichUC(Result).Data,Len,True);
      ReadData(Pointer(NativeInt(@PXLSStringRichUC(Result).Data) + StrMemLen),FmtCount * SizeOf(TFormatRun));
      ConvertFormatRunsDynFontRuns(Pointer(NativeInt(@PXLSStringRichUC(Result).Data) + StrMemLen),FmtCount);

      SetLength(S,Len);
      for i := 0 to Len - 1 do
        S[i + 1] := AxUCChar(PWordArray(@PXLSStringRichUC(Result).Data)[i]);

      FManager.SST.AddFormattedString(S,FontRuns);
    end;
    STRID_FAREAST: begin
      Dec(RecSize,FXLSStream.Read(FarEastDataSize,4));
      GetMem(Result,SizeOf(TXLSStringFarEast) + MemLen + FarEastDataSize);
      ReadString(@PXLSStringFarEast(Result).Data,Len,False);
      ReadData(Pointer(NativeInt(@PXLSStringFarEast(Result).Data) + StrMemLen),FarEastDataSize);
      PXLSStringFarEast(Result).FarEastDataSize := FarEastDataSize;

      FManager.SST.AddRawString(Len,PByteArray(@PXLSStringFarEast(Result).Data));
    end;
    STRID_FAREAST_UC: begin
      Dec(RecSize,FXLSStream.Read(FarEastDataSize,4));
      GetMem(Result,SizeOf(TXLSStringFarEastUC) + MemLen + FarEastDataSize);
      ReadString(@PXLSStringFarEastUC(Result).Data,Len,True);
      ReadData(Pointer(NativeInt(@PXLSStringFarEastUC(Result).Data) + StrMemLen),FarEastDataSize);
      PXLSStringFarEastUC(Result).FarEastDataSize := FarEastDataSize;

      FManager.SST.AddRawString(Len,PWordArray(@PXLSStringFarEastUC(Result).Data));
    end;
    STRID_FAREAST_RICH: begin
      Dec(RecSize,FXLSStream.Read(FmtCount,2));
      Dec(RecSize,FXLSStream.Read(FarEastDataSize,4));
      GetMem(Result,SizeOf(TXLSStringFarEastRich) + MemLen + FmtCount * SizeOf(TXc12FontRun) + FarEastDataSize);
      ReadString(@PXLSStringFarEastRich(Result).Data,Len,True);
      ReadData(Pointer(NativeInt(@PXLSStringFarEastRich(Result).Data) + StrMemLen),FmtCount * SizeOf(TFormatRun));
      ConvertFormatRunsDynFontRuns(Pointer(NativeInt(@PXLSStringFarEastRich(Result).Data) + StrMemLen),FmtCount);
      ReadData(Pointer(NativeInt(@PXLSStringFarEastRich(Result).Data) + StrMemLen + FmtCount * SizeOf(TXc12FontRun)),FarEastDataSize);
      PXLSStringFarEastRich(Result).FormatCount := FmtCount;
      PXLSStringFarEastRich(Result).FarEastDataSize := FarEastDataSize;

      SetLength(S,Len);
      for i := 0 to Len - 1 do
        S[i + 1] := AxUCChar(PByteArray(@PXLSStringFarEastRich(Result).Data)[i]);

      FManager.SST.AddFormattedString(S,FontRuns);
    end;
    STRID_FAREAST_RICH_UC: begin
      Dec(RecSize,FXLSStream.Read(FmtCount,2));
      Dec(RecSize,FXLSStream.Read(FarEastDataSize,4));
      GetMem(Result,SizeOf(TXLSStringFarEastRichUC) + MemLen + FmtCount * SizeOf(TXc12FontRun) + FarEastDataSize);
      ReadString(@PXLSStringFarEastRichUC(Result).Data,Len,True);
      ReadData(Pointer(NativeInt(@PXLSStringFarEastRichUC(Result).Data) + StrMemLen),FmtCount * SizeOf(TFormatRun));
      ConvertFormatRunsDynFontRuns(Pointer(NativeInt(@PXLSStringFarEastRichUC(Result).Data) + StrMemLen),FmtCount);
      ReadData(Pointer(NativeInt(@PXLSStringFarEastRichUC(Result).Data) + StrMemLen + FmtCount * SizeOf(TXc12FontRun)),FarEastDataSize);
      PXLSStringFarEastRichUC(Result).FormatCount := FmtCount;
      PXLSStringFarEastRichUC(Result).FarEastDataSize := FarEastDataSize;

      SetLength(S,Len);
      for i := 0 to Len - 1 do
        S[i + 1] := AxUCChar(PWordArray(@PXLSStringFarEastRichUC(Result).Data)[i]);

      FManager.SST.AddFormattedString(S,FontRuns);
    end;
    else
      raise XLSRWException.CreateFmt('STT: Unhandled string type in Read (%.2X)',[Options]);
  end;

  FreeMem(Result);
end;

{ TXLSReadClipboard }

procedure TXLSReadClipboard.AdjustCell;
begin
  PRecCellXF(PBuf).FormatIndex := FXFUsage[PRecCellXF(PBuf).FormatIndex].XF;
  PRecCellXF(PBuf).Col := (PRecCellXF(PBuf).Col - FSrcArea.Col1) + FDestCol;
  PRecCellXF(PBuf).Row := (PRecCellXF(PBuf).Row - FSrcArea.Row1) + FDestRow;
end;

procedure TXLSReadClipboard.DoRSTRING;
var
  i, Sz, N: integer;
  P: PFormatRun;
begin
  if (PDataRSTRING(PBuf).Options and $01) = $01 then
    Sz := SizeOf(TDataRSTRING) + PDataRSTRING(PBuf).Len * 2
  else
    Sz := SizeOf(TDataRSTRING) + PDataRSTRING(PBuf).Len;

  N := PWordArray(@PBuf[Sz])[0];
  P := @PWordArray(@PBuf[Sz])[1];
  for i := 0 to N - 1 do
  begin
    FFONTUsage[P.FontIndex] := FFONTUsage[P.FontIndex] + 1;
    P := PFormatRun(NativeInt(P) + SizeOf(TFormatRun));
  end;
end;

procedure TXLSReadClipboard.IncXF(Index: integer);
begin
  if Index <= High(FXFUsage) then
    FXFUsage[Index].XF := FXFUsage[Index].XF + 1;
end;

function TXLSReadClipboard.Read(AStream: TStream; const ADestSheet, ADestCol, ADestRow: integer): boolean;
begin
  FDestSheet := ADestSheet;
  FDestCol := ADestCol;
  FDestRow := ADestRow;
  try
    FXLSStream.SourceStream := AStream;
    FVersion := FXLSStream.OpenStorageRead('');
    GetMem(PBuf, FXLS.MaxBuffsize);

    Result := ReadPass1;
    if Result then
      Result := ReadPass2;
  finally
    FXLSStream.CloseStorage;
    FreeMem(PBuf);
  end;
end;

function TXLSReadClipboard.ReadPass1: boolean;
var
  i, Count: integer;
  C1, C2, R1, R2: integer;
begin
  Clear;
  FXFCount := 0;
  FFONTCount := 0;
  FFORMATMax := 0;
  FCurrSheet := 0;
  while FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader) do
  begin
    FXLSStream.Read(PBuf^, Header.Length);
    case Header.RecID of
      BIFFRECID_FONT:
        begin
          Inc(FFONTCount);
          if FFONTCount > Length(FFONTUsage) then SetLength(FFONTUsage, Length(FFONTUsage) + 64);
        end;
      BIFFRECID_FORMAT:
        FFORMATMax := Max(FFORMATMax, PRecFORMAT8(PBuf).Index);
      BIFFRECID_XF:
        begin
          Inc(FXFCount);
          if FXFCount > Length(FXFUsage) then
            SetLength(FXFUsage, Length(FXFUsage) + 256);
          FXFUsage[FXFCount - 1].XF := 0;
          FXFUsage[FXFCount - 1].Font := PRecXF8(PBuf).FontIndex;
          FXFUsage[FXFCount - 1].NumFmt := PRecXF8(PBuf).NumFmtIndex;
        end;

      BIFFRECID_BLANK,
      BIFFRECID_BOOLERR,
      BIFFRECID_FORMULA,
      BIFFRECID_RK,
      BIFFRECID_RK7,
      BIFFRECID_NUMBER,
      BIFFRECID_LABEL:
        IncXF(PRecCellXF(PBuf).FormatIndex);
      BIFFRECID_RSTRING:
        begin
          IncXF(PRecCellXF(PBuf).FormatIndex);
          DoRSTRING;
        end;

      BIFFRECID_MULRK:
        begin
          with PRecMULRK(PBuf)^ do
          begin
            for i := 0 to (Header.Length - 6) div 6 - 1 do
              IncXF(RKs[i].XF);
          end;
        end;
      BIFFRECID_MERGEDCELLS:
        begin
          Count := PRecMERGEDCELLS(PBuf).Count;
          for i := 0 to Count - 1 do
          begin
            R1 := (PRecMERGEDCELLS(PBuf).Cells[i].Row1 - FSrcArea.Row1) + FDestRow;
            R2 := (PRecMERGEDCELLS(PBuf).Cells[i].Row2 - FSrcArea.Row1) + FDestRow;
            C1 := (PRecMERGEDCELLS(PBuf).Cells[i].Col1 - FSrcArea.Col1) + FDestCol;
            C2 := (PRecMERGEDCELLS(PBuf).Cells[i].Col2 - FSrcArea.Col1) + FDestCol;
            if FXLS.Sheets[FCurrSheet].MergedCells.AreaInAreas(C1, R1, C2, R2) then
            begin
              FXLS.InteractError := xieCanNotChangeMerged;
              Result := False;
              Exit;
            end;
          end;
        end;
      BIFFRECID_MULBLANK:
        begin
          with PRecMULBLANK(PBuf)^ do
          begin
            for i := 0 to (Header.Length - 6) div 2 - 1 do
              IncXF(FormatIndexes[i]);
          end;
        end;
      // BIFFRECID_NOTE:                RREC_NOTE;
      BIFFRECID_SELECTION:
        begin
          if PRecSELECTION(PBuf).Refs > 0 then
          begin
            FSrcArea.Col1 := PRecSELECTION(PBuf).Col1;
            FSrcArea.Col2 := PRecSELECTION(PBuf).Col2;
            FSrcArea.Row1 := PRecSELECTION(PBuf).Row1;
            FSrcArea.Row2 := PRecSELECTION(PBuf).Row2;
          end;
        end;
    end;
  end;
  if FFONTCount > 4 then
    Inc(FFONTCount);

  SetLength(FFONTUsage, FFONTCount);
  SetLength(FFORMATUsage, FFORMATMax);

  for i := 0 to FXFCount - 1 do
  begin
    if FXFUsage[i].XF > 0 then
    begin
      if FXFUsage[i].Font > 4 then
        FFONTUsage[FXFUsage[i].Font] := FFONTUsage[FXFUsage[i].Font] + 1
      else
        FFONTUsage[FXFUsage[i].Font] := FFONTUsage[FXFUsage[i].Font] + 1;
      FFORMATUsage[FXFUsage[i].NumFmt] := FFORMATUsage[FXFUsage[i].NumFmt] + 1;
    end;
  end;
  Result := True;
end;

function TXLSReadClipboard.ReadPass2: boolean;
var
  i, Count: integer;
begin
  FXFCount := 0;
  FFONTCount := 0;
  FXLSStream.Seek(0, soFromBeginning);
  while FXLSStream.ReadHeader(Header) = SizeOf(TBIFFHeader) do
  begin
    FXLSStream.Read(PBuf^, Header.Length);
    case Header.RecID of
      BIFFRECID_FONT:
        begin
          if FFONTUsage[FFONTCount] > 0 then
          begin
            RREC_FONT;
            FFONTUsage[FFONTCount] := FXLS.Fonts.Count - 1;
          end;
          Inc(FFONTCount);
          if FFONTCount = 4 then
            Inc(FFONTCount);
        end;
      BIFFRECID_FORMAT:
        begin
          with PRecFORMAT8(PBuf)^ do
          begin
//            if FFORMATUsage[Index] > 0 then
//              FFORMATUsage[Index] := FXLS.Formats.NumberFormats.Add (ByteStrToWideString(@Data, Len)).IndexId;
          end;
        end;
      BIFFRECID_XF:
        begin
          if FXFUsage[FXFCount].XF > 0 then
          begin
            with PRecXF8(PBuf)^ do
            begin
              FontIndex := FFONTUsage[FontIndex];
              NumFmtIndex := FFORMATUsage[NumFmtIndex];
              RREC_XF;
//              FXFUsage[FXFCount].XF := FXLS.Formats.Count - 1;
            end;
          end;
          Inc(FXFCount);
        end;
      BIFFRECID_RK:
        begin
          AdjustCell;
          RREC_RK;
        end;
      BIFFRECID_RK7:
        begin
          AdjustCell;
          RREC_RK;
        end;
      BIFFRECID_BLANK:
        begin
          AdjustCell;
          RREC_BLANK;
        end;
      BIFFRECID_MULBLANK:
        begin
          with PRecMULBLANK(PBuf)^ do
          begin
            Col1 := (Col1 - FSrcArea.Col1) + FDestCol;
            Row := (Row - FSrcArea.Row1) + FDestRow;
            for i := 0 to (Header.Length - 6) div 2 - 1 do
            begin
//              FormatIndexes[i] := FXFUsage[FormatIndexes[i]].XF;
//              FXLS.Sheets[FCurrSheet].IntWriteBlank(Col1 + i, Row,FormatIndexes[i]);
            end;
          end;
        end;
      BIFFRECID_NUMBER:
        begin
          AdjustCell;
          RREC_NUMBER;
        end;
      BIFFRECID_LABEL:
        begin
          AdjustCell;
          RREC_LABEL;
        end;
      BIFFRECID_RSTRING:
        begin
          AdjustCell;
          SaveRSTRING; ;
        end;
      BIFFRECID_ROW:
        begin
          if (PRecROW(PBuf).Row >= FSrcArea.Row1) and (PRecROW(PBuf).Row <= FSrcArea.Row2) then
          begin
//            R := (PRecROW(PBuf).Row - FSrcArea.Row1) + FDestRow;
//            FXLS.Sheets[FCurrSheet].Rows.AddIfNone(R);
//            if PRecROW(PBuf).Height > FXLS.Sheets[FCurrSheet].Rows[R].Height then
//              FXLS.Sheets[FCurrSheet].Rows[R].Height := PRecROW(PBuf).Height;
          end;
        end;
      BIFFRECID_MERGEDCELLS:
        begin
          Count := PRecMERGEDCELLS(PBuf).Count;
          for i := 0 to Count - 1 do begin
            with TMergedCell(FXLS.Sheets[FCurrSheet].MergedCells.Add) do begin
              Row1 := (PRecMERGEDCELLS(PBuf).Cells[i].Row1 - FSrcArea.Row1) + FDestRow;
              Row2 := (PRecMERGEDCELLS(PBuf).Cells[i].Row2 - FSrcArea.Row1) + FDestRow;
              Col1 := (PRecMERGEDCELLS(PBuf).Cells[i].Col1 - FSrcArea.Col1) + FDestCol;
              Col2 := (PRecMERGEDCELLS(PBuf).Cells[i].Col2 - FSrcArea.Col1) + FDestCol;
            end;
          end;
        end;
    end;
  end;
  Result := True;
end;

procedure TXLSReadClipboard.SaveRSTRING;
var
  S: AxUCString;
  i, Sz, N: integer;
  P: PFormatRun;
  FormatRuns: array of TFormatRun;
begin
  SetLength(S, PDataRSTRING(PBuf).Len);
  if (PDataRSTRING(PBuf).Options and $01) = $01 then
  begin
    Sz := SizeOf(TDataRSTRING) + PDataRSTRING(PBuf).Len * 2;
    Move(PDataRSTRING(PBuf).Data, Pointer(S)^, PDataRSTRING(PBuf).Len * 2);
  end
  else
  begin
    Sz := SizeOf(TDataRSTRING) + PDataRSTRING(PBuf).Len;
    for i := 0 to PDataRSTRING(PBuf).Len - 1 do
      S[i + 1] := WideChar(PByteArray(@PDataRSTRING(PBuf).Data)[i]);
  end;

  N := PWordArray(@PBuf[Sz])[0];
  P := @PWordArray(@PBuf[Sz])[1];
  SetLength(FormatRuns, N);
  for i := 0 to N - 1 do
  begin
    FormatRuns[i].Index := P.Index;
    FormatRuns[i].FontIndex := FFONTUsage[P.FontIndex];
    P := PFormatRun(NativeInt(P) + SizeOf(TFormatRun));
  end;

//  FXLS.Sheets[FCurrSheet].IntWriteSSTRichString(PRecCellXF(PBuf).Col, PRecCellXF(PBuf).Row, PRecCellXF(PBuf).FormatIndex, S, FormatRuns);
end;

initialization
  L_Count := 0;

end.
