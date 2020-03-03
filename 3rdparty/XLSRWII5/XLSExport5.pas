unit XLSExport5;

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
     XLSUtils5, XLSReadWriteII5;

type TXLSExportOption = (xeoIncludeEmptyLeftCols,xeoIncludeEmptyTopRows,xeoSeparateFiles);
     TXLSExportOptions = set of TXLSExportOption;

//* Base class for objects that exports TXLSReadWriteII data to other file formats.
type TXLSExport5 = class(TComponent)
private
    procedure SetFilename(const Value: AxUCString);
    function  GetDirectory: AxUCString;
    procedure SetDirectory(const Value: AxUCString);
protected
    FOptions: TXLSExportOptions;

    // If different file for each sheet, this is the directory.
    FFilename: AxUCString;
    FFileExtension: AxUCString;
    // Don't has to be the same as FFilename if there is different files for
    // each sheet.
    FCurrFilename: AxUCString;
    FCurrMultiSheet: integer;
    FDirectory: AxUCString;
    FXLS: TXLSReadWriteII5;
    FCurrSheetIndex: integer;
    FFirstRow: integer;
    FFirstCol: integer;
    FCol1,FCol2: integer;
    FRow1,FRow2: integer;
    FSheets: array of integer;

    function  ExportImages: boolean; virtual;

    procedure OpenFile;        virtual;
    procedure WriteFilePrefix; virtual;
    procedure WritePagePrefix; virtual;
    procedure WriteRowPrefix(const ARow: integer);  virtual;
    procedure WriteCell(const Col,Row: integer; const IsFirstCol,IsFirstRow: boolean); virtual;
    procedure WriteRowSuffix;  virtual;
    procedure WritePageSuffix; virtual;
    procedure WriteFileSuffix; virtual;
    procedure CloseFile;       virtual;
    procedure WriteData;       virtual;

    procedure DoWriteSingle(const AFilename: AxUCString);
    procedure DoWriteMulti;

    function  CheckWriteSheet(const AIndex: integer): boolean;
public
    constructor Create(AOwner: TComponent); override;
    //* Writes the data to the file set with ~[link Filename]
    procedure Write;
    //* Writes the data to a stream.
    procedure SaveToStream(Stream: TStream); virtual;
    procedure SaveToFile(const AFilename: AxUCString); virtual;
    //* Array of sheets to write
    procedure Sheets(const ASheetIndexes: array of integer);
published
    property Options: TXLSExportOptions read FOptions write FOptions;

    //* Left column of the source area on the worksheet.
    property Col1: integer read FCol1 write FCol1;
    //* Right column of the source area on the worksheet.
    property Col2: integer read FCol2 write FCol2;
    //* Name of the destination file that the data shall be written to,
    //* or if different file for each sheet, this is the directory.
    property Filename: AxUCString read FFilename write SetFilename;
    property FileExtension: AxUCString read FFileExtension write FFileExtension;
    //* Directory where files used by the export (like images for html) will go.
    //* Default is the same directory as Filename.
    property Directory: AxUCString read GetDirectory write SetDirectory;
    //* Top row of the source area on the worksheet.
    property Row1: integer read FRow1 write FRow1;
    //* Bottom row of the source area on the worksheet.
    property Row2: integer read FRow2 write FRow2;
    //* The source TXLSReadWriteII2 object.
    property XLS: TXLSReadWriteII5 read FXLS write FXLS;
    end;

implementation

{ TXLSExport }

function TXLSExport5.CheckWriteSheet(const AIndex: integer): boolean;
var
  i: integer;
begin
  Result := True;
  if Length(FSheets) > 0 then begin
    Result := False;
    for i := 0 to High(FSheets) do begin
      Result := AIndex = FSheets[i];
      if Result then
        Break;
    end;
  end;
  if Result and (xeoSeparateFiles in FOptions) then
    Result := FCurrMultiSheet = AIndex;
end;

procedure TXLSExport5.CloseFile;
begin

end;

constructor TXLSExport5.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  FCol1 := -1;
  FCol2 := -1;
  FRow1 := -1;
  FRow2 := -1;
end;

procedure TXLSExport5.DoWriteMulti;
var
  i: integer;
begin
  for i := 0 to FXLS.Count - 1 do begin
    FCurrMultiSheet := i;
    DoWriteSingle(Directory + FXLS[i].Name + '.' + FileExtension);
  end;
end;

procedure TXLSExport5.DoWriteSingle(const AFilename: AxUCString);
begin
  FCurrFilename := AFilename;

  OpenFile;
  try
    WriteData;
  finally
    CloseFile;
  end;
end;

function TXLSExport5.ExportImages: boolean;
begin
  Result := False;
end;

function TXLSExport5.GetDirectory: AxUCString;
begin
  if FDirectory <> '' then
    Result := FDirectory
  else
    Result := ExtractFilePath(FFilename);

  if (Result <> '') and (Result[Length(Result)] <> G_DirSepChar) then
    Result := Result + G_DirSepChar;
end;

procedure TXLSExport5.OpenFile;
begin

end;

procedure TXLSExport5.SaveToFile(const AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  FFilename := AFilename;
  Stream := TFileStream.Create(FFilename,fmCreate);
  try
    SaveToStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TXLSExport5.SaveToStream(Stream: TStream);
begin
  if FXLS = Nil then
    raise XLSRWException.Create('No TXLSReadWriteII defined');
  WriteData;
end;

procedure TXLSExport5.SetDirectory(const Value: AxUCString);
begin
  FDirectory := Value;
  if (FDirectory <> '') and (Copy(FDirectory,1,1) <> '\') then
    FDirectory := FDirectory + '\';
end;

procedure TXLSExport5.SetFilename(const Value: AxUCString);
begin
  FFilename := Value;
end;

procedure TXLSExport5.Sheets(const ASheetIndexes: array of integer);
var
  i: integer;
begin
  SetLength(FSheets,Length(ASheetIndexes));

  for i := 0 to High(ASheetIndexes) do
    FSheets[i] := ASheetIndexes[i];
end;

procedure TXLSExport5.Write;
begin
  if FXLS = Nil then
    raise XLSRWException.Create('No TXLSReadWriteII defined');
  if FFilename = '' then
    FXLS.Manager.Errors.Error('export HTML',XLSERR_FILENAME_IS_MISSING)
  else begin
    if xeoSeparateFiles in FOptions then
      DoWriteMulti
    else
      DoWriteSIngle(FFilename);
  end;
end;

procedure TXLSExport5.WriteCell(const Col,Row: integer; const IsFirstCol,IsFirstRow: boolean);
begin

end;

procedure TXLSExport5.WriteData;
var
  i: integer;
  Col,Row: integer;
  C1,C2,R1,R2: integer;
  IC2,IR2: integer;
  RowCount: integer;
  RowCounter: integer;
  RowDiv: integer;
begin
  WriteFilePrefix;

  RowCount := 0;
  RowCounter := 0;
  for i := 0 to FXLS.Count - 1 do begin
    FXLS[i].CalcDimensions;
    if not CheckWriteSheet(i) then
      Continue;
    Inc(RowCount,FXLS[i].LastRow);
  end;

  RowDiv := RowCount div 1000;
  if RowDiv = 0 then
    RowDiv := 25;

  FXLS.Manager.BeginProgress(xptExport,RowCount);

  for i := 0 to FXLS.Count - 1 do begin
    if not CheckWriteSheet(i) then
      Continue;

    FCurrSheetIndex := i;

    if xeoIncludeEmptyLeftCols in FOptions then
      FFirstCol := 0
    else
      FFirstCol := FXLS[i].FirstCol;

    if xeoIncludeEmptyTopRows in FOptions then
      FFirstRow := 0
    else
      FFirstRow := FXLS[i].FirstRow;

    if ExportImages then
      FXLS[i].Drawing.Images.MaxDimension(IC2,IR2)
    else begin
      IC2 := 0;
      IR2 := 0;
    end;

    if FCol1 >= 0 then C1 := FCol1 else C1 := FFirstCol;
    if FCol2 >= 0 then C2 := FCol2 else C2 := FXLS[i].LastCol;
    if FRow1 >= 0 then R1 := FRow1 else R1 := FFirstRow;
    if FRow2 >= 0 then R2 := FRow2 else R2 := FXLS[i].LastRow;

    C2 := Max(C2,IC2);
    R2 := Max(R2,IR2);

    WritePagePrefix;
    for Row := R1 to R2 do begin
     if ((RowCounter + Row) mod RowDiv) = 0 then
       FXLS.Manager.WorkProgress(RowCounter + Row);

      WriteRowPrefix(Row);
      for Col := C1 to C2 do
        WriteCell(Col,Row,Col = FFirstCol,Row = FFirstRow);
      WriteRowSuffix;
    end;
    WritePageSuffix;

    Inc(RowCounter,R2);
  end;
  WriteFileSuffix;

  FXLS.Manager.EndProgress;
end;

procedure TXLSExport5.WriteFilePrefix;
begin

end;

procedure TXLSExport5.WriteFileSuffix;
begin

end;

procedure TXLSExport5.WritePagePrefix;
begin

end;

procedure TXLSExport5.WritePageSuffix;
begin

end;

procedure TXLSExport5.WriteRowPrefix;
begin

end;

procedure TXLSExport5.WriteRowSuffix;
begin

end;

end.
