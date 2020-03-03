unit BIFF5;

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
{$ifdef MSWINDOWS}
     Windows, Winspool,
{$endif}
{$ifdef BABOON}

{$else}
     vcl.Graphics, vcl.Printers, vcl.Forms,
{$endif}
     BIFF_Utils5, BIFF_RecsII5, BIFF_SheetData5,
     BIFF_Stream5, BIFF_RecordStorage5, BIFF_VBA5,
     BIFF_Names5, BIFF_ExcelFuncII5, BIFF_Escher5, BIFF_FormulaHandler5,
     BIFF_DrawingObjChart5,
     Xc12Utils5, Xc12DataStylesheet5, Xc12Manager5,
     XLSUtils5;

//* ~[userDoc ..\help\metadata.txt CANT_FIND METADATA]

type TXLSInteractError = (xieNone,xieCanNotChangeMerged);

type TWorkbookOption = (woHidden,   //* The window is hidden.
                        woIconized, //* The window is displayed as an icon.
                        woHScroll,  //* The horizontal scroll bar is displayed.
                        woVScroll,  //* The vertical scroll bar is displayed.
                        woTabs      //* The workbook tabs are displayed.
                        );
     TWorkbookOptions = set of TWorkbookOption;

//* Calculation mode.
type TCalcMode = (cm97Manual,      //* Manual calculation.
                  cm97Automatic,   //* Automatic calculation.
                  cm97AutoExTables //* Automatic calculation, except for tables.
                  );

//* ~exclude                  
type TStyles = class(TList)
private
     function GetItems(Index: integer): PRecSTYLE;
public
     destructor Destroy; override;
     procedure Clear; override;
     procedure Add(Style: PRecSTYLE);

     property Items[Index: integer]: PRecSTYLE read GetItems; default;
     end;

//* This is the TXLSReadWriteII component.
//* ~example
//* ~[sample
//*  // Here are some basic examples.
//*  // Write a value to cell C3.
//*  XLS.Sheet[0~[].AsString[2,2~[] := 'Hello';
//*  // Read a value from cell C3.
//*  S := XLS.Sheet[0~[].AsString[2,2~[];
//*  // Read the same cell, but the position is given as a string.
//*  S := XLS.Sheet[0~[].AsStringRef['C:3'~[];
//*
//*  // Add a named cell.
//*  with XLS.InternalNames.Add do begin
//*    Name := 'MyCell';
//*    Definition := 'Sheet1!C3';
//*  end;
//*  // Read the value from cell C3 by using the previous defined name.
//*  S := XLS.NameAsString['MyCell'~[];
//*
//*  // Set the color of a cell.
//*  XLS.Sheet[0~[].Cell[2,2~[].FillPatternForeColor := xcYellow;
//*  // Set the color of another cell. As there is in cell at this position, add a blank cell.
//*  XLS.Sheet[0~[].AsBlank[2,3~[] := True;
//*  XLS.Sheet[0~[].Cell[2,3~[].FillPatternForeColor := xcYellow;
//* ]
type TBIFF5 = class(TPersistent{, ICellOperations})
private
      FManager : TXc12Manager;
      FFilename: AxUCString;
      FTempStream: TStream;
      FVersion: TExcelVersion;
      FOrigVersion: TExcelVersion;
      FSheets: TSheets;
      FFonts: TXc12Fonts;
      FMaxBuffsize: integer;
      FIsMac: boolean;
      FWriteDefaultData: boolean;
      FDefaultCountryIndex: integer;
      FWinIniCountry: integer;
      FShowFormulas: boolean;
      FAfterLoad: TNotifyEvent;
      FFormulaHandler: TFormulaHandler;
      FProgressEvent: TIntegerEvent;
      FFunctionEvent: TFunctionEvent;
      FPasswordEvent: TXLSPasswordEvent;
      FStyles: TStyles;
      FRecords: TRecordStorageGlobals;
//      FDevMode: PDeviceModeW;
      FExtraObjects: TExtraObjects;
      FVBA: TXLSVBA;
      FMSOPictures: TMSOPictures;
//      FPrintPictures: TMSOPictures;
      FSheetCharts: TSheetCharts;
      FPreserveMacros: boolean;
      FPassword: AxUCString;
      FDefaultPaperSize: TXc12PaperSize;
      FCellCount: integer;
      FProgressCount: integer;
      FFILESHARING: array of byte;
      FAborted: boolean;
      FSkipDrawing: boolean;
      FSaveFormulasAs2007: boolean;

      FInteractError: TXLSInteractError;

      function  GetRecomendReadOnly: boolean;
      procedure SetRecomendReadOnly(const Value: boolean);
      procedure SetVersion(Value: TExcelVersion);
      procedure SetCodepage(Value: word);
      procedure FormulaHandlerSheetName(Name: AxUCString; var Index,Count: integer);
      function  FormulaHandlerSheetData(DataType: TSheetDataType; SheetIndex,Col,Row: integer): AxUCString;
      function  GetCodepage: word;
      function  GetUserName: AxUCString;
      procedure SetUserName(const Value: AxUCString);
      function  GetBookProtected: boolean;
      procedure SetBookProtected(const Value: boolean);
      function  GetBackup: boolean;
      procedure SetBackup(const Value: boolean);
      function  GetRefreshAll: boolean;
      procedure SetRefreshAll(const Value: boolean);
      function  GetVersionNumber: string;
      procedure SetVerionNumber(const Value: string);
      function  GetInternalNames: TInternalNames;
      procedure SetInternalNames(const Value: TInternalNames);
      procedure SetStrFALSE(const Value: AxUCString);
      procedure SetStrTRUE(const Value: AxUCString);
      function  GetPalette(Index: integer): TColor;
      procedure SetPalette(Index: integer; const Value: TColor);
      function  GetSheet(Index: integer): TSheet;
      procedure SetFilename(const Value: AxUCString);
      function  GetStrFALSE: AxUCString;
      function  GetStrTRUE: AxUCString;
      procedure WritePrepare;
      function  GetReadMacros: boolean;
      procedure SetReadMacros(const Value: boolean);
protected
      procedure InternalNameDeleted(Sender: TObject; const Delta: integer);
      procedure InternalName(Name: AxUCString; var ID: integer);
public
      //* ~exclude
      constructor Create(AManager : TXc12Manager);
      //* ~exclude
      destructor Destroy; override;
      //* Clears all cell values in all sheets. See also ~[link Clear]
      procedure ClearCells;
      //* Clears all data from the workbook.
      procedure Clear;
      //* Reads the file given by ~[link Filename]
      procedure Read;
      //* Writes the file given by ~[link Filename]. Please note that if
      //* the file allready is opened by another application, such as Excel, it
      //* is not possible to write to it, as Excel locks the file for exclusive
      //* use.
      procedure Write;
      //* Writes the workbook to a stream instead of writing to a file.
      //* ~param Stream The stream to write the workbook to.
      procedure WriteToStream(Stream: TStream);
      //* ~exclude
      procedure LoadFromStream(Stream: TStream);
      //* Returns a password that can be used to unprotect worksheets.
      //* As the password is calculated, this may not be the same password used to protect it.
      //* ~result A valid password that can unprotect the file. If there is no protection, an empty string is returned.
      function GetWeakPassword: string;
      //* The max number of rows there can be in a worksheet. Normally 65536.
      //* ~result Max number of rows.
      function  MaxRowCount: integer;
      //* Use BeginUpdate when you do a lot of changes to the workbook.
      //* BeginUpdate is most usefull when adding many string cells.
      //* EndUpdate must be called after the changes are done.
      //* See also ~[link EndUpdate]
      procedure BeginUpdate;
      //* Writes any pending changes since BeginUpdate was called.
      //* See also ~[link BeginUpdate]
      procedure EndUpdate;

      //* Abort reading of a file. After the reading is aborted ~[link Aborted] is set tor True.
      procedure Abort;
      //* Checks ~[link Abort] was called to stop the file reading.
      //* ~result True if the read operation was aborted.
      function  Aborted: boolean;

      property DefaultPaperSize: TXc12PaperSize read FDefaultPaperSize write FDefaultPaperSize;

      { ********************************************* }
      { ********** For internal use only. *********** }
      { ********************************************* }

      property Manager : TXc12Manager read FManager;

      // Set to False when loading an Excel file, as there is default data in the file.
      property WriteDefaultData: boolean read FWriteDefaultData write FWriteDefaultData;
      //* ~exclude
      property MaxBuffSize: integer read FMaxBuffsize;
      //* ~exclude
      property Styles: TStyles read FStyles;
      //* ~exclude
      property OnAfterLoad: TNotifyEvent read FAfterLoad write FAfterLoad;

      //* ~exclude
      property Records: TRecordStorageGlobals read FRecords;
      //* ~exclude
      property ExtraObjects: TExtraObjects read FExtraObjects;
      //* ~exclude
      property Fonts: TXc12Fonts read FFonts write FFonts;
      //* ~exclude
      function  GetExternNameValue(NameIndex,SheetIndex: integer): TFormulaValue;
      //* ~exclude
      procedure SetFILESHARING(PBuf: PByteArray; Size: integer);
      //* ~exclude
      function  GetFILESHARING(PBuf: PByteArray): integer;
      // @exclude
      // *********************************************
      // ********** End internal use only. ***********
      // *********************************************
      //* User defined names for cells and areas.
      property InternalNames: TInternalNames read GetInternalNames write SetInternalNames;
      //* Charts that are on it's own sheet.
      //* See also ~[link TSheetCharts]
      property SheetCharts: TSheetCharts read FSheetCharts;

      //* ~exclude
      property DefaultCountryIndex: integer read FDefaultCountryIndex write FDefaultCountryIndex;
      //* ~exclude
      property WinIniCountry: integer read FWinIniCountry write FWinIniCountry;
      //* ~exclude
      property FormulaHandler: TFormulaHandler read FFormulaHandler;
      //* ~exclude
      property Codepage: word read GetCodepage write SetCodepage;
      //* The palette for the colors used by Excel. The palette has 64 entries,
      //* but the first 8 are fixed and can not be changed.
      property Palette[Index: integer]: TColor read GetPalette write SetPalette;
      //* THe worksheets in the workbook.
      //* See ~[link TSheet]
      property Sheet[Index: integer]: TSheet read GetSheet; default;
      //* Object for handling Visual Basic macros.
      property InteractError: TXLSInteractError read FInteractError write FInteractError;
      property VBA: TXLSVBA read FVBA;
      //* The Excel file version for the workbook.
      //* See also ~[link TExcelVersion]
      property Version: TExcelVersion read FVersion write SetVersion;
      //* The original file version when reading a file. This can not be changed.
      //* When createing a new file, OrigVersion is set to xvNone;
      property OrigVersion: TExcelVersion read FOrigVersion;
      //* The worksheets in the workbook.
      //* See ~[link TSheet]
      property Sheets: TSheets read FSheets write FSheets;
      //* The name of the creator of the workbook.
      property UserName: AxUCString read GetUserName write SetUserName;
      //* The BookProtect property stores the protection state for a sheet or workbook.
      property BookProtected: boolean read GetBookProtected write SetBookProtected;
      //* The Backup property specifies whether Microsoft Excel should save
      //* backup versions of a file, when the file is opened with Excel.
      property Backup: boolean read GetBackup write SetBackup;
      //* Set the RefreshAll property to True if all external data should be
      //* refreshed when the workbook is loaded by Excel.
      property RefreshAll: boolean read GetRefreshAll write SetRefreshAll;
      //* Change StrTRUE property to change the string representation of the
      //* boolean value True. The default is 'True'.
      //* See also ~[link StrFALSE]
      property StrTRUE: AxUCString read GetStrTRUE write SetStrTRUE;
      //* Change StrFALSE property to change the string representation of the
      //* boolean value False. The default is 'False'.
      //* See also ~[link StrTRUE]
      property StrFALSE: AxUCString read GetStrFALSE write SetStrFALSE;
      //* Set the ShowFormulas property to True if functions which reads cells
      //*and return string values shall return the formula itself or de result
      //* (value) of the formula.
      property ShowFormulas: boolean read FShowFormulas write FShowFormulas;
      //* Set the Filename property to the name of the file you want to read or
      //* write.
      //* See also ~[link Read], ~[link Write], ~[link LoadFromStream], ~[link WriteToStream]
      property Filename: AxUCString read FFilename write SetFilename;
      //* True if the file is created by Excel for Macintosh.
      property IsMac: boolean read FIsMac write FIsMac;
      //* Set PreserveMacros to True if macros (VBA script) shall be preserved
      //* when a file is read. When set to False, all macros are deleted.
      property PreserveMacros: boolean read FPreserveMacros write FPreserveMacros;

      property ReadMacros: boolean read GetReadMacros write SetReadMacros;
      //* The version of the TXLSReadWriteII2 component.
      property ComponentVersion: string read GetVersionNumber write SetVerionNumber;
      //* Global storage for pictures used in the worksheets.
      property MSOPictures: TMSOPictures read FMSOPictures write FMSOPictures;
//      property PrintPictures: TMSOPictures read FPrintPictures write FPrintPictures;
      //* Password when reading/writing encrypted files.
      property Password: AxUCString read FPassword write FPassword;
      //* Mark the file as recomended read only.
      property RecomendReadOnly: boolean read GetRecomendReadOnly write SetRecomendReadOnly;

      property SkipDrawing: boolean read FSkipDrawing write FSkipDrawing;
      property SaveFormulasAs2007: boolean read FSaveFormulasAs2007 write FSaveFormulasAs2007;

      //* Event fired while a file is read or written. The Value parameter
      //* increases from 0 to 100 while the file is read or written.
      property OnProgress: TIntegerEvent read FProgressEvent write FProgressEvent;
      //* Use the OnFunction event to do the calculation of formulas which not
      //* are calculated by TXLSReadWriteII2.
      //* See also ~[link Calculate]
      property OnFunction: TFunctionEvent read FFunctionEvent write FFunctionEvent;
      //* Event fired when a password protected (encrypted) file is read, and
      //* the password is required.
      property OnPassword: TXLSPasswordEvent read FPasswordEvent write FPasswordEvent;
      end;

implementation

uses BIFF_ReadII5, BIFF_WriteII5, BIFF_DecodeFormula5;

{ TStyle }

procedure TStyles.Add(Style: PRecSTYLE);
var
  P: PRecSTYLE;
  PU: PRecSTYLE_USER;
  Sz: integer;
begin
  if (Style.FormatIndex and $F000) = 0 then begin
    PU := PRecSTYLE_USER(Style);
    if PU.Len = 0 then
      Sz := 0
    else if PByteArray(@PU.Data)[0] = 0 then
      Sz := PU.Len + 1
    else if PByteArray(@PU.Data)[0] = 1 then
      Sz := (PU.Len * 2) + 1
    // Excel 95
    else
      Sz := PU.Len;
    GetMem(P,SizeOf(TRecSTYLE_USER) + Sz);
    System.Move(Style^,P^,SizeOf(TRecSTYLE_USER) + Sz);
  end
  else begin
    New(P);
    System.Move(Style^,P^,SizeOf(TRecSTYLE));
  end;
  inherited Add(P);
end;

procedure TStyles.Clear;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    FreeMem(inherited Items[i]);
  inherited Clear;
end;

destructor TStyles.Destroy;
begin
  Clear;
  inherited;
end;

function TStyles.GetItems(Index: integer): PRecSTYLE;
begin
  Result := inherited Items[Index];
end;

constructor TBIFF5.Create(AManager: TXc12Manager);
begin
  FManager := AManager;

  Move(TXc12DefaultIndexColorPalette[0],Xc12IndexColorPalette[0],SizeOf(Xc12IndexColorPalette));

{
  FCodepage := $04B0;
}

  // GetDEVMODE can cause all kind of strange errors. It's only used for
  // reading the default paper size.
//  FDevMode := Nil; // GetDEVMODE;
//  FDefaultPaperSize := psA4;
//  PB := Nil;
//
//  Sz := GetLocaleInfo(GetUserDefaultLCID,$100A {LOCALE_IPAPERSIZE},PChar(PB),0);
//  GetMem(PB,Sz * 2);
//  try
//    if GetLocaleInfo(LOCALE_SYSTEM_DEFAULT,$100A {LOCALE_IPAPERSIZE},PChar(PB),Sz) <> 0 then begin
//      case Char(PB[0]) of
//        '1': FDefaultPaperSize := psLegal;
//        '5': FDefaultPaperSize := psLetter;
//        '8': FDefaultPaperSize := psA3;
//        '9': FDefaultPaperSize := psA4;
//      end;
//    end;
//  finally
//     FreeMem(PB);
//  end;

  FRecords := TRecordStorageGlobals.Create;
  FRecords.SetDefaultData;
  FPreserveMacros := True;
  FExtraObjects := TExtraObjects.Create;
  FIsMac := False;
  FWriteDefaultData := True;
  FFonts := FManager.StyleSheet.Fonts;
  FFormulaHandler := TFormulaHandler.Create(Self);
  FFormulaHandler._OnName := InternalName;
  FFormulaHandler.OnSheetName := FormulaHandlerSheetName;
  FFormulaHandler.OnSheetData := FormulaHandlerSheetData;
  FFormulaHandler.InternalNames.OnNameDeleted := InternalNameDeleted;
  SetStrTRUE('TRUE');
  SetStrFALSE('FALSE');
  FMSOPictures := TMSOPictures.Create;
//  FPrintPictures := TMSOPictures.Create(Self);
  FStyles := TStyles.Create;
  FSheets := TSheets.Create(Self,FFormulaHandler,FManager);
//  Clear;
  FSheets.Add;
  FSheetCharts := TSheetCharts.Create(Self,MSOPictures);
  SetVersion(xvExcel97);
end;

destructor TBIFF5.Destroy;
begin
  try
//    Clear;
  except
  end;
  FSheets.Free;
  FFormulaHandler.Free;
  FStyles.Free;
//  FreeMem(FDevMode);
  FRecords.Free;
  FMSOPictures.Free;
//  FPrintPictures.Free;
  FExtraObjects.Free;
  FSheetCharts.Free;
  if FVBA <> Nil then
    FVBA.Free;
  if FTempStream <> Nil then
    FTempStream.Free;

  inherited;
end;

procedure TBIFF5.ClearCells;
begin
  FSheets.Clear;
end;

procedure TBIFF5.Clear;
begin
  FFormulaHandler.Clear;
  FRecords.SetDefaultData;
  FWriteDefaultData := True;
  ClearCells;
  FStyles.Clear;
  FMSOPictures.Clear;
//  FPrintPictures.Free;
  FExtraObjects.Clear;
  if FVBA <> Nil then
    FVBA.Clear;
  FSheetCharts.Clear;
  if FTempStream <> Nil then
    FreeAndNil(FTempStream);
//  if FDevMode <> Nil then
//    FreeMem(FDevMode);
//  FDevMode := Nil;
  SetLength(FFILESHARING,0);
  FAborted := False;
  FOrigVersion := xvNone;
end;

procedure TBIFF5.SetVersion(Value: TExcelVersion);
begin
  FVersion := Value;
  if FVersion >= xvExcel97 then
    FMaxBuffsize := MAXRECSZ_97
  else
    FMaxBuffsize := MAXRECSZ_40;
end;

procedure TBIFF5.SetFilename(const Value: AxUCString);
begin
  FFilename := Value;
end;

procedure TBIFF5.SetFILESHARING(PBuf: PByteArray; Size: integer);
begin
  SetLength(FFILESHARING,Size);
  Move(PBuf^,FFILESHARING[0],Size);
end;

procedure TBIFF5.SetCodepage(Value: word);
begin
  if Value = 0 then
    FRecords.CODEPAGE := $04E4
  else
    FRecords.CODEPAGE := Value;
end;

function TBIFF5.GetCodepage: word;
begin
  Result := FRecords.CODEPAGE;
end;

procedure TBIFF5.Read;
var
  Ext: AxUCString;
begin
  Ext := Lowercase(ExtractFileExt(FFilename));
  LoadFromStream(Nil);
  FOrigVersion := FVersion;
end;

procedure TBIFF5.LoadFromStream(Stream: TStream);
var
  i: integer;
  List: TList;
  XLSRead: TXLSReadII;
begin
  Clear;
  FSheets.ClearAll;

  BeginUpdate;
  XLSRead := TXLSReadII.Create(Self);
  try
    XLSRead.SkipMSO := FSkipDrawing;
    XLSRead.LoadFromStream(Stream);
//    FixupPictureData;
    FFormulaHandler.ExternalNames.FilePath := ExtractFilePath(FFilename);
    if Assigned(FAfterLoad) then
      FAfterLoad(Self);
  finally
    XLSRead.Free;
    EndUpdate;
    if FSheets.Count < 1 then begin
      Clear;
      FSheets.Clear;
    end;
  end;
  FRecords.PostCheck;
  List := TList.Create;
  try
    FMSOPictures.GetBlipIds(List);
    for i := 0 to FSheets.Count - 1 do begin
      FSheets[i]._Int_EscherDrawing.AssignBlipIds(List);
      FSheets[i].AfterFileRead;
    end;
    for i := 0 to FSheetCharts.Count - 1 do begin
      if FSheetCharts[i].Drawing <> Nil then
        FSheetCharts[i].Drawing.AssignBlipIds(List);
    end;
  finally
    List.Free;
  end;
end;

procedure TBIFF5.Write;
begin
  WriteToStream(Nil);
end;

procedure TBIFF5.WritePrepare;
var
  i: integer;
begin
  if FVersion in [xvExcel21,xvExcel30] then
    raise XLSRWException.Create('Can not write Excel 2.1/3.0 files');

  FAborted := False;
  if Assigned(FProgressEvent) then begin
    FProgressEvent(Self,0);
    FCellCount := 0;
    FProgressCount := 0;
    for i := 0 to FSheets.Count - 1 do
      Inc(FCellCount,FSheets[i]._Int_Records.Count);
  end;

  FFormulaHandler.ExternalNames.UpdateIntSupbooks(FSheets.Count);

//  FMSOPictures.ResetBlipRefCount;
  for i := 0 to FSheets.Count - 1 do
    FSheets[i]._Int_EscherDrawing.SetBlipRefCount;
end;

procedure TBIFF5.WriteToStream(Stream: TStream);
var
  XLSWrite: TXLSWriteII;
begin
  WritePrepare;

  XLSWrite := TXLSWriteII.Create(Self);
  try
    XLSWrite.WriteToStream(Stream)
  finally
    XLSWrite.Free;
  end;
  if Assigned(FProgressEvent) then
    FProgressEvent(Self,100);
end;

procedure TBIFF5.Abort;
begin
  Clear;
  FAborted := True;
end;

function TBIFF5.Aborted: boolean;
begin
  Result := FAborted;
end;

function TBIFF5.MaxRowCount: integer;
begin
  Result := MAXROW;
end;

function TBIFF5.GetExternNameValue(NameIndex, SheetIndex: integer): TFormulaValue;
begin
  Result := FFormulaHandler.ExternalNames.GetNameValue(NameIndex,SheetIndex);
end;

function TBIFF5.GetUserName: AxUCString;
begin
  Result := FRecords.WRITEACCESS;
end;

procedure TBIFF5.SetUserName(const Value: AxUCString);
begin
  FRecords.WRITEACCESS := Value;
end;

function TBIFF5.GetBookProtected: boolean;
begin
  Result := FRecords.WINDOWPROTECT;
end;

procedure TBIFF5.SetBookProtected(const Value: boolean);
begin
  FRecords.WINDOWPROTECT := Value;
end;

function TBIFF5.GetBackup: boolean;
begin
  Result := FRecords.BACKUP;
end;

procedure TBIFF5.SetBackup(const Value: boolean);
begin
  FRecords.BACKUP := Value;
end;

function TBIFF5.GetReadMacros: boolean;
begin
  Result := FVBA <> Nil;
end;

function TBIFF5.GetRecomendReadOnly: boolean;
begin
  if Length(FFILESHARING) < 2 then
    Result := False
  else
    Result := PWordArray(FFILESHARING)[0] = 1;
end;

function TBIFF5.GetRefreshAll: boolean;
begin
  Result := FRecords.REFRESHALL;
end;

procedure TBIFF5.SetReadMacros(const Value: boolean);
begin
  if Value and (FVBA = Nil) then
    FVBA := TXLSVBA.Create
  else if not Value and (FVBA <> Nil) then
    FVBA.Free;
end;

procedure TBIFF5.SetRecomendReadOnly(const Value: boolean);
var
  Z: boolean;
  i: integer;
begin
  if Value then begin
    if Length(FFILESHARING) < 2 then begin
      SetLength(FFILESHARING,6);
      FillChar(FFILESHARING[0],Length(FFILESHARING),#0);
    end;
    PWordArray(FFILESHARING)[0] := 1;
  end
  else begin
    if Length(FFILESHARING) < 2 then
      SetLength(FFILESHARING,0)
    else begin
      PWordArray(FFILESHARING)[0] := 0;
      Z := True;
      for i := 0 to High(FFILESHARING) do begin
        if FFILESHARING[i] <> 0 then begin
          Z := False; 
          Break;
        end;
      end;
      if Z then
        SetLength(FFILESHARING,0);
    end;
  end;
end;

procedure TBIFF5.SetRefreshAll(const Value: boolean);
begin
  FRecords.REFRESHALL := Value;
end;

function TBIFF5.GetVersionNumber: string;
begin
  Result := CurrentVersionNumber;
end;

function TBIFF5.GetWeakPassword: string;
var
  P: word;
  Chars,S: string;
  i,j,k: integer;

function MakeWeakPassword(S: string): word;
var
  C,T: word;
  i: integer;
begin
  Result := 0;
  S := Copy(S,1,15);
  for i := 1 to Length(S) do begin
    T := Byte(S[i]);
    C := ((T shl i) + (T shr (15 - i))) and $7FFF;
    Result := Result xor C;
  end;
  Result := Result xor Length(S) xor $CE4B;
end;

begin
  Result := '';
  P := FRecords.PASSWORD;
  if P = 0 then
    Exit;
  Chars := 'abcdefghijklmnopqrstuwvxyzABCDEFGHIJKLMNOPQRSTUWVXYZ0123456789_';
  Randomize;
  for k := 1 to 1000000 do begin
    i := Random(15) + 1;
    SetLength(S,i);
    for j := 1 to i do
      S[j] := Chars[Random(Length(Chars)) + 1];
    if P = MakeWeakPassword(S) then begin
      Result := S;
      Exit;
    end;
  end;
end;

procedure TBIFF5.SetVerionNumber(const Value: string);
begin
//
end;

procedure TBIFF5.FormulaHandlerSheetName(Name: AxUCString; var Index,Count: integer);
var
  Sheet: TSheet;
begin
//  if not FManager.LoadedFromFile then begin
//    Index := FManager.Worksheets.Find(Name);
//    if Index >= 0 then
//      Count := FManager.Worksheets.Count
//    else
//      Count := 0;
//  end
//  else begin
    Sheet := FSheets.SheetByName(Name);
    if Sheet = Nil then begin
      Index := -1;
      Count := 0;
    end
    else begin
      Index := Sheet.Index;
      Count := FSheets.Count;
    end;
//  end;
end;

function TBIFF5.GetInternalNames: TInternalNames;
begin
  Result := FFormulaHandler.InternalNames;
end;

procedure TBIFF5.SetInternalNames(const Value: TInternalNames);
begin

end;

procedure TBIFF5.BeginUpdate;
begin

end;

procedure TBIFF5.EndUpdate;
begin
  
end;

procedure TBIFF5.SetStrFALSE(const Value: AxUCString);
begin
  G_StrFALSE := Value;
end;

procedure TBIFF5.SetStrTRUE(const Value: AxUCString);
begin
  G_StrTRUE := Value;
end;

function TBIFF5.GetFILESHARING(PBuf: PByteArray): integer;
begin
  Move(FFILESHARING[0],PBuf^,Length(FFILESHARING));
  Result := Length(FFILESHARING);
end;

function TBIFF5.GetPalette(Index: integer): TColor;
begin
  if (Index < 0) or (Index > High(Xc12IndexColorPalette)) then
    raise XLSRWException.Create('Palette index out of range');
  Result := Xc12IndexColorPalette[Index];
end;

procedure TBIFF5.SetPalette(Index: integer; const Value: TColor);
begin
  if (Index < 8) or (Index > High(Xc12IndexColorPalette)) then
    raise XLSRWException.Create('Palette index out of range');
  Xc12IndexColorPalette[Index] := Value;
end;

function TBIFF5.GetSheet(Index: integer): TSheet;
begin
  Result := FSheets[Index];
end;

function TBIFF5.GetStrFALSE: AxUCString;
begin
  Result := G_StrFALSE;
end;

function TBIFF5.GetStrTRUE: AxUCString;
begin
  Result := G_StrTRUE;
end;

procedure TBIFF5.InternalName(Name: AxUCString; var ID: integer);
begin
  ID := FManager.Workbook.DefinedNames.FindId(Name,-1);
end;

procedure TBIFF5.InternalNameDeleted(Sender: TObject; const Delta: integer);
var
  i: integer;
begin
  for i := 0 to FSheets.Count - 1 do
    FSheets[i].NameIndexChanged(Delta);
end;

function TBIFF5.FormulaHandlerSheetData(DataType: TSheetDataType; SheetIndex, Col, Row: integer): AxUCString;
begin
  case DataType of
    sdtName: Result := FSheets[SheetIndex].Name;
    sdtCell: Result := '';
  end;
end;

end.

