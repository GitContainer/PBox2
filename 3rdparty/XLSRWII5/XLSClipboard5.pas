unit XLSClipboard5;

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
{$ifdef BABOON}
{$else}
     Windows, vcl.Clipbrd,
{$endif}
     Xc12Utils5, Xc12DataStyleSheet5, Xc12Manager5,
     BIFF_Utils5, BIFF_RecsII5, BIFF_Stream5,
     BIFF12_Recs5,
     XLSUtils5, XLSFormulaTypes5, XLSMMU5, XLSCellMMU5, XLSReadWriteOPC5, XLSFormula5;

type TXLSClipboardFormat = (xcfNone,xcfText,xcfSYLK,xcfBIFF8,xcfBIFF12{,xcfXML});

type PDataRSTRING = ^TDataRSTRING;
     TDataRSTRING = packed record
     Row, Col: word;
     FormatIndex: word;
     Len: word;
     Options: byte;
     Data: record end;
     end;

var TextSepChars: array[0..3] of char = (',',';',':',#9);

type TXLSTextClipboard = class(TObject)
protected
     FManager: TXc12Manager;

     FList   : TStrings;

     FSheet  : integer;
     FCol    : integer;
     FRow    : integer;

     FHasQuoteChar: boolean;
     FSepChar: AXUCChar;

     FFormulas: TXLSFormulaHandler;

     function  CPosSkipQuote(C: AXUCChar; S: AXUCString): integer;
     function  StrTokenSkipQuote(var S: AXUCString; Token: AXUCChar): AXUCString;
public
     constructor Create(AManager: TXc12Manager; const ASheet,ACol,ARow: integer);
     destructor Destroy; override;

     procedure Read(AList: TStrings);
     end;

type TXLSBIFF8Clipboard = class(TObject)
protected
     FManager: TXc12Manager;

     FStream : TXLSStream;

     FSheet  : integer;
     FCol    : integer;
     FRow    : integer;

     FHeader : TBIFFHeader;

     FBuf    : PByteArray;
     FBufSz  : integer;

     function  BufAsString(const AData: PByteArray; const ALen: integer): AxUCString;
public
     constructor Create(AManager: TXc12Manager; const ASheet,ACol,ARow: integer);
     destructor Destroy; override;

     procedure Read(AStream: TStream);
     end;

type TXLSBIFF12Clipboard = class(TObject)
protected
     FManager: TXc12Manager;

     FStream : TStream;

     FSheet  : integer;
     FCol    : integer;
     FRow    : integer;

     FRecId  : integer;
     FRecSz  : integer;

     FBuf    : PByteArray;
     FBufSz  : integer;

     function  ReadHeader: boolean;
     function  BIFFString(const ALength: integer; AString: PWordArray): AxUCString;
public
     constructor Create(AManager: TXc12Manager; const ASheet,ACol,ARow: integer);
     destructor Destroy; override;

     procedure Read(AStream: TStream);
     end;

type TXLSXMLClipboard = class(TObject)
protected
     FManager: TXc12Manager;

     FStream : TStream;

     FSheet  : integer;
     FCol    : integer;
     FRow    : integer;
public
     constructor Create(AManager: TXc12Manager; const ASheet,ACol,ARow: integer);
     destructor Destroy; override;

     procedure Read(AStream: TStream);
     end;

//* Handles the clipboard.
type TXLSClipboard = class(TObject)
protected
     FManager: TXc12Manager;

     FWinCF  : integer;

     function  WantThis(const AFormat: TXLSClipboardFormat; const AFormats: array of TXLSClipboardFormat): boolean;
     procedure Open;
     function  Available(const AFormats: array of TXLSClipboardFormat): TXLSClipboardFormat;
     function  ReadAsStream(AStream: TStream): boolean;
     function  ReadAsStrings(AList: TStrings): boolean;
     procedure Close;

     procedure DoPasteText(const ASheet,ACol,ARow: integer);
     procedure DoPasteBIFF8(const ASheet,ACol,ARow: integer);
     procedure DoPasteBIFF12(const ASheet,ACol,ARow: integer);
     procedure DoPasteXML(const ASheet,ACol,ARow: integer);
public
     constructor Create(AManager: TXc12Manager);
     destructor Destroy; override;

     //* Pasts content on the clipboard top the location at ASheet,ACol,ARow.
     //* Format on the clipboard can be Excel 2007, Excel 97 or Text.
     //* If the format is Text, the following special parsing is done:
     //* True, False: boolean values.
     //* Date or time: Date or time values if the text is as system date/time.
     //* Formulas: If the text is preceeded with an equal sign, the text is
     //* parsed as a formula. If that fails, it is inserted as text.
     procedure PasteFromClipboard(const ASheet,ACol,ARow: integer);
     end;

implementation

{ TXLSClipboard }

function TXLSClipboard.Available(const AFormats: array of TXLSClipboardFormat): TXLSClipboardFormat;
{$ifndef BABOON}
var
  i: integer;
  S: AxUCString;
  XCF: TXLSClipboardFormat;
  Buf: array[0..1024] of char;
{$endif}
begin
  Result := xcfNone;
{$ifdef BABOON}
{$else}
  XCF := xcfNone;
  for i := 0 to CountClipboardFormats - 1 do begin
    FWinCF := EnumClipboardFormats(FWinCF);
    if GetClipboardFormatName(FWinCF,Buf,1024) > 0 then begin
      S := Uppercase(StrPas(Buf));
      if S = 'BIFF8' then
        XCF := xcfBIFF8
      else if S = 'BIFF12' then
        XCF := xcfBIFF12;
//      else if S = 'XML SPREADSHEET' then
//        XCF := xcfXML;
    end
    else begin
      case FWinCF of
        CF_TEXT,
        CF_UNICODETEXT: XCF := xcfText;
        CF_SYLK       : XCF := xcfSYLK;
      end;
    end;
    if (XCF <> xcfNone) and WantThis(XCF,AFormats) then begin
      Result := XCF;
      Exit;
    end;
  end;
{$endif}
end;

procedure TXLSClipboard.Close;
begin
{$ifdef BABOON}
{$else}
  CloseClipboard;
{$endif}
end;

procedure TXLSClipboard.PasteFromClipboard(const ASheet, ACol, ARow: integer);
begin
  Open;
  try
    case Available([xcfText,xcfBIFF8,xcfBIFF12{,xcfXML}]) of
      xcfText  : DoPasteText(ASheet,ACol,ARow);
      xcfBIFF8 : DoPasteBIFF8(ASheet,ACol,ARow);
      xcfBIFF12: DoPasteBIFF12(ASheet,ACol,ARow);
//      xcfXML   : DoPasteXML(ASheet,ACol,ARow);
    end;
  finally
    Close;
  end;
end;

constructor TXLSClipboard.Create(AManager: TXc12Manager);
begin
  FManager := AManager;
end;

destructor TXLSClipboard.Destroy;
begin

  inherited;
end;

procedure TXLSClipboard.DoPasteBIFF12(const ASheet,ACol,ARow: integer);
var
  Stream: TMemoryStream;
  StreamSheet: TStream;
  OPC: TOPC_XLSX;
  Item: TOPCItem;
  BIFF12: TXLSBIFF12Clipboard;
begin
  Stream := TMemoryStream.Create;
  try
    if ReadAsStream(Stream) then begin
      OPC := TOPC_XLSX.Create;
      try
        OPC.OpenRead(Stream,False);

        Stream.Seek(0,soFromBeginning);

        Item := OPC.FindSheetBin(1);
        if Item <> Nil then begin
          StreamSheet := OPC.OpenAndReadSheet(Item.Id);
          try
            BIFF12 := TXLSBIFF12Clipboard.Create(FManager,ASheet,ACol,ARow);
            try
              BIFF12.Read(StreamSheet);
            finally
              BIFF12.Free;
            end;
          finally
            StreamSheet.Free;
          end;
        end;

        OPC.Close;
      finally
        OPC.Free;
      end;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSClipboard.DoPasteBIFF8(const ASheet,ACol,ARow: integer);
var
  BIFF: TXLSBIFF8Clipboard;
  Stream: TMemoryStream;
begin
  Stream := TMemoryStream.Create;
  try
    if ReadAsStream(Stream) then begin
      BIFF := TXLSBIFF8Clipboard.Create(FManager,ASheet,ACol,ARow);
      try
        BIFF.Read(Stream);
      finally
        BIFF.Free;
      end;
    end;
  finally
    Stream.Free;
  end;
end;

procedure TXLSClipboard.DoPasteText(const ASheet,ACol,ARow: integer);
var
  Lines: TStringList;
  Text: TXLSTextClipboard;
begin
  Lines := TStringList.Create;
  try
    if ReadAsStrings(Lines) then begin
      Text := TXLSTextClipboard.Create(FManager,ASheet,ACol,ARow);
      try
        Text.Read(Lines);
      finally
        Text.Free;
      end;
    end;
  finally
    Lines.Free;
  end;
end;

procedure TXLSClipboard.DoPasteXML(const ASheet, ACol, ARow: integer);
begin

end;

procedure TXLSClipboard.Open;
begin
  FWinCF := -1;
{$ifdef BABOON}
{$else}
  OpenClipboard(0);
{$endif}
end;

function TXLSClipboard.ReadAsStream(AStream: TStream): boolean;
{$ifndef BABOON}
var
  Data: THandle;
  DataPtr: Pointer;
{$endif}
begin
{$ifdef BABOON}
  Result := False;
{$else}
  Data := GetClipboardData(FWinCF);
  Result := Data > 0;
  if Result then begin
    DataPtr := GlobalLock(Data);

    AStream.Write(DataPtr^,GlobalSize(Data));
    AStream.Seek(0,soFromBeginning);

    GlobalUnLock(Data);
  end;
{$endif}
end;

function TXLSClipboard.ReadAsStrings(AList: TStrings): boolean;
{$ifndef BABOON}
var
  S: AxUCString;
  S8: XLS8String;
  Data: THandle;
  DataPtr: Pointer;
{$endif}
begin
{$ifdef BABOON}
  Result := False;
{$else}
  Data := GetClipboardData(FWinCF);
  Result := Data > 0;
  if Result then begin
    DataPtr := GlobalLock(Data);

    if FWinCF = CF_UNICODETEXT then begin
      SetLength(S,GlobalSize(Data) div 2);
      System.Move(DataPtr^,Pointer(S)^,GlobalSize(Data));
    end
    else begin
      SetLength(S8,GlobalSize(Data));
      System.Move(DataPtr^,Pointer(S8)^,GlobalSize(Data));
      S := AxUCString(S8);
    end;
    AList.Text := S;

    GlobalUnLock(Data);
  end;
{$endif}
end;

function TXLSClipboard.WantThis(const AFormat: TXLSClipboardFormat; const AFormats: array of TXLSClipboardFormat): boolean;
var
  i: integer;
begin
  for i := 0 to High(AFormats) do begin
    if AFormat = AFormats[i] then begin
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

{ TXLSBIFF12Clipboard }

function TXLSBIFF12Clipboard.BIFFString(const ALength: integer; AString: PWordArray): AxUCString;
begin
  SetLength(Result,ALength);
  Move(AString^,Pointer(Result)^,ALength * 2);
end;

constructor TXLSBIFF12Clipboard.Create(AManager: TXc12Manager; const ASheet,ACol,ARow: integer);
begin
  FManager := AManager;
  FSheet := ASheet;
  FCol := ACol;
  FRow := ARow;
end;

destructor TXLSBIFF12Clipboard.Destroy;
begin
  if FBuf <> Nil then
    FreeMem(FBuf);

  inherited;
end;

procedure TXLSBIFF12Clipboard.Read(AStream: TStream);
var
  Found  : boolean;
  Cells  : TXLSCellMMU;
  CurrRow: integer;
  SrcCol : integer;
  SrcRow : integer;
  vFloat : double;
  vStr   : AxUCString;
  vBool  : boolean;
  vErr   : TXc12CellError;
begin
  FStream := AStream;

  ReadHeader;
  if FRecId <> BIFF12_REC_BEGINSHEET then
    Exit;

  CurrRow := 0;
  SrcCol := 0;
  SrcRow := 0;

  Cells := FManager.Worksheets[FSheet].Cells;

  Found := False;
  while ReadHeader do begin
    case FRecId of
      BIFF12_REC_WSDIM: begin
        SrcCol := PBIFF12_WsDim(FBuf).Col1;
        SrcRow := PBIFF12_WsDim(FBuf).Row1;
      end;
      BIFF12_REC_BEGINSHEETDATA: begin
        Found := True;
        Break;
      end;
    end;
  end;
  if not Found then
    Exit;

  while ReadHeader do begin
    case FRecId of
      BIFF12_REC_ROWHDR: CurrRow := PBIFF12_RowHdr(FBuf).Row;
      BIFF12_REC_CELLRK: begin
        vFloat := DecodeRK(PBIFF12_CellRK(FBuf).RK);
        Cells.UpdateFloat(FCol + (PBIFF12_CellRK(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vFloat);
      end;
      BIFF12_REC_CELLERROR: begin
        vErr := ErrorCodeToCellError(PBIFF12_CellError(FBuf).Error);
        Cells.UpdateError(FCol + (PBIFF12_CellError(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vErr);
      end;
      BIFF12_REC_CELLBOOL: begin
        vBool := PBIFF12_CellBool(FBuf).Bool = 1;
        Cells.UpdateBoolean(FCol + (PBIFF12_CellBool(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vBool);
      end;
      BIFF12_REC_CELLREAL: begin
        vFloat := PBIFF12_CellReal(FBuf).Num;
        Cells.UpdateFloat(FCol + (PBIFF12_CellReal(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vFloat);
      end;
      BIFF12_REC_FMLANUM: begin
        vFloat := PBIFF12_FmlaNum(FBuf).Num;
        Cells.UpdateFloat(FCol + (PBIFF12_FmlaNum(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vFloat);
      end;
      BIFF12_REC_FMLASTRING: begin
        vStr := BIFFString(PBIFF12_FmlaString(FBuf).Length,@PBIFF12_FmlaString(FBuf).Str[0]);
        Cells.UpdateString(FCol + (PBIFF12_FmlaString(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vStr);
      end;
      BIFF12_REC_FMLABOOL: begin
        vBool := PBIFF12_FmlaBool(FBuf).Bool = 1;
        Cells.UpdateBoolean(FCol + (PBIFF12_FmlaBool(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vBool);
      end;
      BIFF12_REC_FMLAERROR: begin
        vErr := ErrorCodeToCellError(PBIFF12_FmlaError(FBuf).Error);
        Cells.UpdateError(FCol + (PBIFF12_FmlaError(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vErr);
      end;
      BIFF12_REC_CELLST,
      BIFF12_REC_CELLRSTRING: begin
        vStr := BIFFString(PBIFF12_CellRString(FBuf).Length,@PBIFF12_CellRString(FBuf).Str[0]);
        Cells.UpdateString(FCol + (PBIFF12_CellRString(FBuf).Col - SrcCol),FRow + (CurrRow - SrcRow),vStr);
      end;
    end;
  end;
end;

function TXLSBIFF12Clipboard.ReadHeader: boolean;
var
  B1,B2,B3,B4: byte;
begin
  Result := False;
  if FStream.Read(B1,1) = 1 then begin
    if (B1 and $80) <> 0 then begin
      FStream.Read(B2,1);
      FRecId := B2 shl 7 + (B1 and $7F);
    end
    else
      FRecId := B1;

    Result := True;

    FStream.Read(B1,1);
    FRecSz := B1 and $7F;
    if (B1 and $80) <> 0 then begin
      FStream.Read(B2,1);
      Inc(FRecSz,(B2 and $7F) shl 7);
      if (B2 and $80) <> 0 then begin
        FStream.Read(B3,1);
        Inc(FRecSz,(B3 and $7F) shl 16);
        if (B3 and $80) <> 0 then begin
          FStream.Read(B4,1);
          Inc(FRecSz,B4 shl 24);
        end;
      end;
    end;
    if FRecSz > 0 then begin
      if FRecSz > FBufSz then begin
        FBufSz := FRecSz;
        ReallocMem(FBuf,FBufSz);
      end;
      FStream.Read(FBuf^,FRecSz);
    end;
  end;
end;

{ TXLSBIFF8Clipboard }

function TXLSBIFF8Clipboard.BufAsString(const AData: PByteArray; const ALen: integer): AxUCString;
var
  Sz: integer;
  vStr8: XLS8String;
begin
  if (AData[0] and $01) = $01 then begin
    Sz := ALen * 2;
    SetLength(Result,PDataRSTRING(FBuf).Len);
    Move(AData[1],Pointer(Result)^,Sz);
  end
  else begin
    Sz := ALen;
    SetLength(vStr8,Sz);
    Move(AData[1],Pointer(vStr8)^,Sz);
    Result := AxUCString(vStr8);
  end;
end;

constructor TXLSBIFF8Clipboard.Create(AManager: TXc12Manager; const ASheet, ACol, ARow: integer);
begin
  FManager := AManager;
  FSheet := ASheet;
  FCol := ACol;
  FRow := ARow;

  GetMem(FBuf,MAXRECSZ_97);

  FStream := TXLSStream.Create(Nil);
end;

destructor TXLSBIFF8Clipboard.Destroy;
begin
  FreeMem(FBuf);

  FStream.Free;
  inherited;
end;

procedure TXLSBIFF8Clipboard.Read(AStream: TStream);
var
  Found  : boolean;
  i      : integer;
  Sz     : integer;
  Cells  : TXLSCellMMU;
  SrcCol : integer;
  SrcRow : integer;
  C,R    : integer;
  vFloat : double;
  vStr   : AxUCString;
  vStr8  : XLS8String;
begin
  FStream.SourceStream := AStream;
  FStream.OpenStorageRead('');

  FStream.Seek(20,soFromBeginning);

  Cells := FManager.Worksheets[FSheet].Cells;

  try
    Found := False;
    while FStream.ReadHeader(FHeader) = SizeOf(TBIFFHeader) do begin
      FStream.Read(FBuf^,FHeader.Length);
      case FHeader.RecID of
        BIFFRECID_BOUNDSHEET: begin
          FStream.Seek(PRecBOUNDSHEET8(FBuf).BOFPos,soFromBeginning);
          Found := True;
          Break;
        end;
        BIFFRECID_EOF: Exit;
      end;
    end;
    if not Found then
      Exit;

    SrcCol := 0;
    SrcRow := 0;

    while FStream.ReadHeader(FHeader) = SizeOf(TBIFFHeader) do begin
      FStream.Read(FBuf^,FHeader.Length);
      case FHeader.RecID of
        BIFFRECID_DIMENSIONS: begin
          SrcCol := PRecDIMENSIONS8(FBuf).FirstCol;
          SrcRow := PRecDIMENSIONS8(FBuf).FirstRow;
        end;
        BIFFRECID_BOOLERR: begin
          if PRecBOOLERR(FBuf).Error = 0 then
            Cells.UpdateBoolean(FCol + (PRecBOOLERR(FBuf).Col - SrcCol),FRow + (PRecBOOLERR(FBuf).Row - SrcRow),Boolean(PRecBOOLERR(FBuf).BoolErr))
          else
            Cells.UpdateError(FCol + (PRecBOOLERR(FBuf).Col - SrcCol),FRow + (PRecBOOLERR(FBuf).Row - SrcRow),ErrorCodeToCellError(PRecBOOLERR(FBuf).BoolErr));
        end;
        BIFFRECID_RK,
        BIFFRECID_RK7: begin
          vFloat := DecodeRK(PRecRK(FBuf).Value);
          Cells.UpdateFloat(FCol + (PRecRK(FBuf).Col - SrcCol),FRow + (PRecRK(FBuf).Row - SrcRow),vFloat);
        end;
        BIFFRECID_MULRK: begin
          for i := 0 to (FHeader.Length - 6) div 6 - 1 do begin
            vFloat := DecodeRK(PRecMULRK(FBuf).RKs[i].RK);
            Cells.UpdateFloat(FCol + (PRecMULRK(FBuf).Col1 + i - SrcCol),FRow + (PRecMULRK(FBuf).Row - SrcRow),vFloat);
          end;
        end;
        BIFFRECID_NUMBER: begin
          vFloat := PRecNUMBER(FBuf).Value;
          Cells.UpdateFloat(FCol + (PRecRK(FBuf).Col - SrcCol),FRow + (PRecRK(FBuf).Row - SrcRow),vFloat);
        end;
        BIFFRECID_LABEL: begin
          if PRecLABEL(FBuf).Data[0] = 1 then begin
            Sz := SizeOf(TDataRSTRING) + PDataRSTRING(FBuf).Len * 2;
            SetLength(vStr,PRecLABEL(FBuf).Len);
            Move(PRecLABEL(FBuf).Data[1],Pointer(vStr)^,Sz);
          end
          else begin
            Sz := PRecLABEL(FBuf).Len;
            SetLength(vStr8,PRecLABEL(FBuf).Len);
            Move(PRecLABEL(FBuf).Data[1],Pointer(vStr8)^,Sz);
            vStr := AxUCString(vStr8);
          end;

          Cells.UpdateString(FCol + (PRecLABEL(FBuf).Col - SrcCol),FRow + (PRecLABEL(FBuf).Row - SrcRow),vStr);
        end;
        BIFFRECID_RSTRING: begin
          if (PDataRSTRING(FBuf).Options and $01) = $01 then begin
            Sz := PDataRSTRING(FBuf).Len * 2;
            SetLength(vStr,PDataRSTRING(FBuf).Len);
            Move(PDataRSTRING(FBuf).Data,Pointer(vStr)^,Sz);
          end
          else begin
            Sz := PDataRSTRING(FBuf).Len;
            SetLength(vStr8,Sz);
            Move(PDataRSTRING(FBuf).Data,Pointer(vStr8)^,Sz);
            vStr := AxUCString(vStr8);
          end;

          Cells.UpdateString(FCol + (PDataRSTRING(FBuf).Col - SrcCol),FRow + (PDataRSTRING(FBuf).Row - SrcRow),vStr);
        end;
        BIFFRECID_FORMULA: begin
          C := PRecFORMULA(FBuf).Col;
          R := PRecFORMULA(FBuf).Row;
          vFloat := PRecFORMULA(FBuf).Value;
          if (TByte8Array(vFloat)[0] in [$00, $01, $02, $03]) and (TByte8Array(vFloat)[6] = $FF) and (TByte8Array(vFloat)[7] = $FF) then begin
            case TByte8Array(vFloat)[0] of
              0: begin
                if FStream.PeekHeader = BIFFRECID_STRING then begin
                  FStream.ReadHeader(FHeader);
                  FStream.Read(FBuf^,FHeader.Length);
                  vStr := BufAsString(@PRecSTRING(FBuf).Data,PRecSTRING(FBuf).Len);
                  Cells.UpdateString(FCol + (C - SrcCol),FRow + (R - SrcRow),vStr);
                end;
              end;
              1: Cells.UpdateBoolean(FCol + (C - SrcCol),FRow + (R - SrcRow),Boolean(integer(TByte8Array(vFloat)[2])));
              2: Cells.UpdateError(FCol + (C - SrcCol),FRow + (R - SrcRow),ErrorCodeToCellError(TByte8Array(vFloat)[2]));
              3: ; // Empty string and no STRING record.
            end;
          end
          else if PRecFORMULA(FBuf).ParseLen <> 0 then begin
          // This detects NAN values. A NAN number may cause an "Invalid Floating Point Operation" exception when used.
            if (TByte8Array(vFloat)[0] = $02) and (TByte8Array(vFloat)[6] = $FF) and (TByte8Array(vFloat)[7] = $FF) then
              vFloat := 0;
            Cells.UpdateFloat(FCol + (PRecRK(FBuf).Col - SrcCol),FRow + (PRecRK(FBuf).Row - SrcRow),vFloat);
          end;
        end;
        BIFFRECID_EOF: Exit;
      end;
    end;
  finally
    FStream.CloseStorage;
  end;
end;

{ TXLSTextClipboard }

function TXLSTextClipboard.CPosSkipQuote(C: AXUCChar; S: AXUCString): integer;
var
  InQuote: boolean;
begin
  InQuote := False;
  for Result := 1 to Length(S) do begin
    if CharInSet(S[Result],['"','''']) then
      InQuote := not InQuote;
    if not InQuote and (S[Result] = C) then
      Exit;
  end;
  Result := -1;
end;

constructor TXLSTextClipboard.Create(AManager: TXc12Manager; const ASheet, ACol, ARow: integer);
begin
  FManager := AManager;
  FSheet := ASheet;
  FCol := ACol;
  FRow := ARow;

  FSepChar := #9;

  FFormulas := TXLSFormulaHandler.Create(FManager);
end;

destructor TXLSTextClipboard.Destroy;
begin
  FFormulas.Free;

  inherited;
end;

procedure TXLSTextClipboard.Read(AList: TStrings);
var
  i,C,L  : integer;
  Sz     : integer;
  Col    : integer;
  Row    : integer;
  V      : double;
  D      : TDateTime;
  DateFmt: integer;
  iXF    : integer;
  XF     : TXc12XF;
  S,Token: AXUCString;
  vErr   : TXc12CellError;
  TempDS : AXUCChar;
  Ptgs   : PXLSPtgs;
  Cell   : TXLSCellItem;
  Cells  : TXLSCellMMU;
begin
  FList := AList;

  Cells := FManager.Worksheets[FSheet].Cells;

  Col := FCol;
  Row := FRow;
  TempDS := FormatSettings.DecimalSeparator;
  try
    for i := 0 to FList.Count - 1 do begin
      S := FList[i];
      C := 0;
      while S <> '' do begin
        if FHasQuoteChar then
          Token := Trim(StrTokenSkipQuote(S,FSepChar))
        else
          Token := Trim(SplitAtChar(FSepChar,S));
        if Token <> '' then begin
          L := Length(Token);
          if FHasQuoteChar and (((Token[1] = '"') and (Token[L] = '"')) or ((Token[1] = '''') and (Token[L] = ''''))) then begin
            Token := Copy(Token,2,L - 2);

            Cells.UpdateString(Col + C,Row + i,Token);
          end
          else begin
            if Uppercase(Token) = G_StrTRUE then
              Cells.UpdateBoolean(Col + C,Row + i,True)
            else if Uppercase(Token) = G_StrFALSE then
              Cells.UpdateBoolean(Col + C,Row + i,False)
            else begin
              if TryStrToFloat(Token,V) then
                Cells.UpdateFloat(Col + C,Row + i,V)
              else begin
                if FormatSettings.DecimalSeparator = '.' then
                  FormatSettings.DecimalSeparator := ','
                else
                  FormatSettings.DecimalSeparator := '.';
                if TryStrToFloat(Token,V) then
                  Cells.UpdateFloat(Col + C,Row + i,V)
                else begin
                  if TryStrToDateTime(Token,D) then begin
                    if (Int(D) <> 0) and (Frac(D) <> 0) then
                      DateFmt := XLS_NUMFMT_STD_DATETIME
                    else if Int(D) <> 0 then
                      DateFmt := XLS_NUMFMT_STD_DATE
                    else if Frac(D) <> 0 then
                      DateFmt := XLS_NUMFMT_STD_TIME
                    else
                      DateFmt := XLS_NUMFMT_STD_DATE;

                    Cell := Cells.FindCell(Col + C,Row + i);
                    if Cell.Data <> Nil then begin
                      iXF := Cells.GetStyle(@Cell);
                      FManager.StyleSheet.XFEditor.BeginEdit(FManager.StyleSheet.XFs[iXF]);
                    end
                    else
                      FManager.StyleSheet.XFEditor.BeginEdit(FManager.StyleSheet.XFs.DefaultXF);

                    FManager.StyleSheet.XFEditor.NumberFormat := ExcelStandardNumFormats[DateFmt];

                    XF := FManager.StyleSheet.XFEditor.EndEdit;

                    Cells.UpdateFloat(Col + C,Row + i,D,XF.Index);
                  end
                  else begin
                    vErr := ErrorTextToCellError(Token);
                    if vErr <> errUnknown then
                      Cells.UpdateError(Col + C,Row + i,vErr)
                    else begin
                      if Copy(Token,1,1) = '=' then begin
                        Sz := FFormulas.EncodeFormula(Copy(Token,2,MAXINT),Ptgs,FSheet,FCol,FRow);
                        if Sz > 0 then begin
                          Cells.AddFormula(Col + C,Row + i,XLS_STYLE_DEFAULT_XF,Ptgs,Sz,0);
                          FreeMem(Ptgs);
                        end
                        else
                          Cells.UpdateString(Col + C,Row + i,Token);
                      end
                      else
                        Cells.UpdateString(Col + C,Row + i,Token);
                    end;
                  end;
                end;
              end;
            end;
          end;
        end;
        Inc(C);
      end;
    end;
  finally
    FormatSettings.DecimalSeparator := TempDS;
  end;
end;

function TXLSTextClipboard.StrTokenSkipQuote(var S: AXUCString; Token: AXUCChar): AXUCString;
var
  p: integer;
begin
  p := CPosSkipQuote(Token,S);
  if p >= 1 then begin
    Result := Copy(S,1,p - 1);
    S := Copy(S,p + 1,MAXINT);
  end
  else begin
    Result := S;
    S := '';
  end;
end;

{ TXLSXMLClipboard }

constructor TXLSXMLClipboard.Create(AManager: TXc12Manager; const ASheet, ACol, ARow: integer);
begin
  FManager := AManager;
  FSheet := ASheet;
  FCol := ACol;
  FRow := ARow;
end;

destructor TXLSXMLClipboard.Destroy;
begin

  inherited;
end;

procedure TXLSXMLClipboard.Read(AStream: TStream);
begin

end;

end.
