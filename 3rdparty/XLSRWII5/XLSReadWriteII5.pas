unit XLSReadWriteII5;

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

//* @anchor(a_ExcelFiles)About Excel files.
//* Excel files comes in two major models. Files up to Excel 2006 are binary
//* files and called Excel 97 as the file format was introduced in Excel 97.
//* Files after 2006 are called Excel 2007. These files are mostly XML files
//* stored in a ZIP files. You you change the extension on an Excel 2007+ file
//* to ZIP, you can open the file with WinZIP or similar. The Excel 2007 file
//* format is an official ISO/IEC standard called Office Open.

{.$define SHAREWARE}

uses Classes, SysUtils, Contnrs,
{$ifdef DEMO_TIMELIMIT}
     Forms,
{$endif}
{$ifndef BABOON}
     vcl.Dialogs, Windows, vcl.ExtCtrls,
{$endif}
{$ifdef DELPHI_XE3_OR_LATER}
     System.UITypes,
{$endif}
     xpgParseDrawing,
     Xc12Utils5, Xc12Manager5, Xc12DataWorkbook5, Xc12DefaultData5,
{$ifdef XLS_BIFF}
     BIFF5, BIFF_Utils5, BIFF_ReadII5, BIFF_Escher5, BIFF_EscherTypes5, BIFF_DrawingObj5,
     BIFF_SheetData5,
{$endif}
     XLSUtils5, XLSReadXLSX5, XLSWriteXLSX5, XLSSheetData5, XLSFormula5, XLSNames5,
     XLSMergedCells5, XLSHyperlinks5, XLSValidate5, XLSAutofilter5, XLSCondFormat5,
     XLSClassFactory5, XLSEvaluateFmla5, XLSDrawing5, XLSRelCells5,
{$ifdef XLS_CRYPTO_SUPPORT}
     XLSEncryption5,
{$endif}
     XLSZip5;

const XLSExcelFilesFilter = 'Excel files|*.xls;*.xlt;*.xlsx;*.xlsm;*.xlst|All files|*.*';
const XLSPictureFilesFilter = 'Picture files|*.bmp;*.jpg;*.jpeg;*.png|All files|*.*';

type TXLSReadWriteII5 = class;

     TXLSClassFactoryImpl = class(TXLSClassFactory)
protected
     FOwner   : TXLSReadWriteII5;
public
     constructor Create(AOwner: TXLSReadWriteII5);

     function CreateAClass(AClassType: TXLSClassFactoryType; AClassOwner: TObject = Nil): TObject; override;
//     function GetAClass(AClassType: TXLSClassFactoryType): TObject; override;
     end;

{$ifdef DELPHI_XE2_OR_LATER}
     [ComponentPlatformsAttribute (pidWin32 or pidWin64 or pidOSX32)]
{$endif}
     TXLSReadWriteII5 = class(TXLSWorkbook)
private
     function  GetFilename: AxUCString;
     procedure SetFilename(const Value: AxUCString);
     function  GetProgressEvent: TXLSProgressEvent;
     procedure SetProgressEvent(const Value: TXLSProgressEvent);
     function  GetVersionNumber: AxUCString;
     procedure SetVerionNumber(const Value: AxUCString);
     function  GetVersion: TExcelVersion;
     procedure SetVersion(const Value: TExcelVersion);
     function  GetStrTRUE: AxUCString;
     procedure SetStrTRUE(const Value: AxUCString);
     function  GetStrFALSE: AxUCString;
     procedure SetStrFALSE(const Value: AxUCString);
     function  GetPassword: AxUCString;
     procedure SetPassword(const Value: AxUCString);
     function  GetDirectRead: boolean;
     function  GetDirectWrite: boolean;
     procedure SetDirectRead(const Value: boolean);
     procedure SetDirectWrite(const Value: boolean);
     function  GetCellReadEvent: TCellReadWriteEvent;
     function  GetCellWriteEvent: TCellReadWriteEvent;
     procedure SetCellReadEvent(const Value: TCellReadWriteEvent);
     procedure SetCellWriteEvent(const Value: TCellReadWriteEvent);
     function  GetUseAlternateZip: boolean;
     procedure SetUseAlternateZip(const Value: boolean);
     function  GetUserFunctionEvent: TUserFunctionEvent;
     procedure SetUserFunctionEvent(const Value: TUserFunctionEvent);
     function  GetManager: TXc12Manager;
     function  GetCompressStrings: boolean;
     procedure SetCompressStrings(const Value: boolean);
     function  GetMonitorFile: boolean;
     procedure SetMonitorFile(const Value: boolean);
{$ifdef XLS_BIFF}
     function  GetBIFF: TBIFF5;
     function  GetFunctionEvent: TFunctionEvent;
     procedure SetFunctionEvent(const Value: TFunctionEvent);
{$endif}
protected
     FErrors          : TXLSErrorManager;
     FClassFactory    : TXLSClassFactoryImpl;
{$ifdef XLS_BIFF}
     FReadVBA         : boolean;
     FSkipExcel97Drw  : boolean;
{$endif}
     FPasswordEvent   : TXLSPasswordEvent;

     FHMonitorObj     : THandle;
     FHMonitorFile    : THandle;
{$ifdef BABOON}
{$else}
     FMonitorFileTimer: TTimer;
     FMonitorFileEvent: TNotifyEvent;
{$endif}

{$ifdef SHAREWARE}
     FNagMsgShown     : boolean;
{$endif}
{$ifdef XLS_BIFF}
     function  AddImage97(AAnchor: TCT_TwoCellAnchor; const ASheetIndex: integer): boolean;  override;
{$endif}

     procedure BeforeRead;
     procedure AfterRead; virtual;
     procedure BeforeWrite;
     procedure AddBIFFImages;
     procedure AddBIFFComments;
     procedure SaveBIFFComments;
     procedure FixupRelativelCells;

     function  DoBeginFileMonitor: boolean;
     procedure DoEndFileMonitor;
     procedure DoMonitorFileEvent(ASender: TObject);
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     //* Deletes all data from the workbook and adds ADefaulSheetsCount
     //* worksheets.
     procedure Clear(ADefaulSheetsCount: integer = 1);

     //* Reads the Excel file give by @link(Filename).
     procedure Read;
     //* Does the same thing as @link(Filename) and @link(Read) but in one call.
     //* The @link(Filename) property is set to AFilname.
     procedure LoadFromFile(AFilename: AxUCString);
     //* Reads an Excel file from a stream. If AAutoDetect is True the method
     //* will try to detect if the stream is an Excel 2007 or an Excel 97 stream.
     //* If you explicite want to read an Excel 97 file, use @link(LoadFromStream97).
     procedure LoadFromStream(AStream: TStream; const AAutoDetect: boolean = True);
     //* Reads an Excel 97 file from a stream. If you want to read an
     //* Excel 2007+ file, use @link(LoadFromStream).
     procedure LoadFromStream97(AStream: TStream);

     //* Writes the data the file give by @link(Filename).
     procedure Write;
     //* Does the same thing as @link(Filename) and @link(Write) but in one call.
     //* The @link(Filename) property is set to AFilname.
     procedure SaveToFile(AFilename: AxUCString); overload;
     procedure SaveToFile(AFilename: AxUCString; AVersion: TExcelVersion); overload;
     //* Saves the data to a stream. If AVersion is xvNone, the value of the Version
     //* property vill be used. Only xvExcel97 and xvExcel2007 are supported.
     procedure SaveToStream(AStream: TStream; AVersion: TExcelVersion = xvNone);

     //* Reads only the seet names fron an Excel file or stream.
     //* AList is filled with the sheet names.
     //* Returns the number of sheets in the workbook.
     function  GetSheetNames(const AFilename: AxUCString; AList: TStrings): integer;

     //* Sets the are you want to fill with cell values when using @link(DirectWrite).
     procedure SetDirectWriteArea(const ASheetIndex,ACol1,ARow1,ACol2,ARow2: integer);
{$ifdef _AXOLOT_DEBUG}
     procedure CheckIntegrity(const AList: TStrings);
{$endif}
{$ifdef XLS_BIFF}
     // The BIFF object contains data for Excel 97 files.
     property BIFF: TBIFF5 read GetBIFF;

     // Only Excel 97 macros are read.
     property ReadVBA: boolean read FReadVBA write FReadVBA;

     property SkipExcel97Drawing: boolean read FSkipExcel97Drw write FSkipExcel97Drw;
{$endif}
     //* Change StrTRUE property to change the string representation of the
     //* boolean value True. The default is 'TRUE'.
     //* See also @link(StrFALSE)
     property StrTRUE: AxUCString read GetStrTRUE write SetStrTRUE;
     //* Change StrFALSE property to change the string representation of the
     //* boolean value False. The default is 'FALSE'.
     //* See also @link(StrTRUE)
     property StrFALSE: AxUCString read GetStrFALSE write SetStrFALSE;
     //* The password for encrypted files.
     property Password: AxUCString read GetPassword write SetPassword;

     // Pictures will not work when using Delphi's ZIP.
     property UseAlternateZip : boolean read GetUseAlternateZip write SetUseAlternateZip;

     property CompressStrings: boolean read GetCompressStrings write SetCompressStrings;
     //* When True, the OnMonitorFile event is fired when there are any changes
     //* to the open file.
     property MonitorFile: boolean read GetMonitorFile write SetMonitorFile;

     // TODO remove on final.
     property Manager: TXc12Manager read GetManager;
published
     //* The version of the TXLSReadWriteII5 component.
     property ComponentVersion: AxUCString read GetVersionNumber write SetVerionNumber;

     //* The Excel file version for the workbook.
     //* When reading files, Version is set according to the file extension.@br
     //* xls and xlm = xvExcel97
     //* xlsx and xlsm = xvExcel2007
     //* See also @link(Xc12Utils5.TExcelVersion), @link(a_ExcelFiles)
     property Version: TExcelVersion read GetVersion write SetVersion;

     //* The name of the Excel file.
     property Filename: AxUCString read GetFilename write SetFilename;
     //* When DirectRead mode is active, the cell values are not stored in
     //* memory. Instead is the OnReadCell event fired when a cell is read.
     //* You can then access the cell value in the @link(Xc12Manager5.TXLSEventCell)
     //* object, ACell parameter, in the event. @br
     //* The advantage with the DirectRead mode is that no memory is used for
     //* storing cell values. As Excel 2007+ files can have 100 of millions of
     //* cells this can be a greate memory save. If you know that you just
     //* want a few cells, or going to store them elsewhere, this is a good
     //* choice. See also @link(DirectWrite).
     property DirectRead: boolean read GetDirectRead write SetDirectRead;
     //* DirectWrite mode writes cell values directly to the disk file instead
     //* of first storing them in memory. This can save hughe ammount of memory
     //* if you are writing large files. See also @link(DirectRead).
     property DirectWrite: boolean read GetDirectWrite write SetDirectWrite;
     //* Error manager. See @link(XLSUtils5.TXLSErrorManager).
     property Errors: TXLSErrorManager read FErrors;

     property OnReadCell    : TCellReadWriteEvent read GetCellReadEvent write SetCellReadEvent;
     property OnWriteCell   : TCellReadWriteEvent read GetCellWriteEvent write SetCellWriteEvent;
{$ifdef XLS_BIFF}
     property OnFunction    : TFunctionEvent read GetFunctionEvent write SetFunctionEvent;
{$endif}
     property OnPassword    : TXLSPasswordEvent read FPasswordEvent write FPasswordEvent;
     property OnProgress    : TXLSProgressEvent read GetProgressEvent write SetProgressEvent;
     property OnUserFunction: TUserFunctionEvent read GetUserFunctionEvent write SetUserFunctionEvent;
{$ifdef BABOON}
{$else}
     property OnMonitorFile : TNotifyEvent read FMonitorFileEvent write FMonitorFileEvent;
{$endif}
     end;

implementation

procedure _FileMonitorCallback(Context: Pointer; Success: Boolean); stdcall;
begin
{$ifdef BABOON}
{$else}
  TXLSReadWriteII5(Context).FMonitorFileTimer.Enabled := True;
{$endif}

  TXLSReadWriteII5(Context).DoBeginFileMonitor;
end;


{ TXLSReadWriteII5 }

{$ifdef XLS_BIFF}
function TXLSReadWriteII5.AddImage97(AAnchor: TCT_TwoCellAnchor; const ASheetIndex: integer): boolean;
var
  MSO      : TMSOPicture;
  DrwPict  : TDrwPicture;
  PPIX,PPIY: integer;

function EMUToPixels(APPI,AEMU: int64): int64;
begin
  Result := (AEMU * APPI) div 914400;
end;

begin
  Result := FBIFF <> Nil;

  if Result then begin
    FManager.StyleSheet.PixelsPerInchXY(PPIX,PPIY);

    MSO := FBIFF.MSOPictures.Add;
    MSO.Filename := AAnchor.Objects.Pic.NvPicPr.CNvPr.Descr;
    DrwPict := FBIFF.Sheets[ASheetIndex].DrawingObjects.Pictures.Add;
    DrwPict.PictureName := MSO.Filename;

    DrwPict.Col1       := AAnchor.From.Col;
    DrwPict.Col1Offset := EMUToPixels(PPIX,AAnchor.From.ColOff) / Sheets[ASheetIndex].Columns[AAnchor.From.Col].PixelWidth;
    DrwPict.Row1       := AAnchor.From.Row;
    DrwPict.Row1Offset := EMUToPixels(PPIX,AAnchor.From.RowOff) / Sheets[ASheetIndex].Rows[AAnchor.From.Row].PixelHeight;

    DrwPict.Col2       := AAnchor.To_.Col;
    DrwPict.Col2Offset := EMUToPixels(PPIX,AAnchor.To_.ColOff) / Sheets[ASheetIndex].Columns[AAnchor.To_.Col].PixelWidth;
    DrwPict.Row2       := AAnchor.To_.Row;
    DrwPict.Row2Offset := EMUToPixels(PPIX,AAnchor.To_.RowOff) / Sheets[ASheetIndex].Rows[AAnchor.To_.Row].PixelHeight;
  end;
end;
{$endif}

procedure TXLSReadWriteII5.AddBIFFComments;
{$ifdef XLS_BIFF}
var
  i,j: integer;
  Sht: TSheet;
  N: TDrwNote;
{$endif}
begin
{$ifdef XLS_BIFF}
  for i := 0 to FBIFF.Sheets.Count - 1 do begin
    if i < Count then begin
      Sht := FBIFF[i];
      for j := 0 to Sht.DrawingObjects.Notes.Count - 1 do begin
        N := Sht.DrawingObjects.Notes[j];
        Items[i].Comments.AsPlainText[N.CellCol,N.CellRow] := N.Text;
      end;
    end;
  end;
{$endif}
end;

procedure TXLSReadWriteII5.AddBIFFImages;
{$ifdef XLS_BIFF}
var
  i,j: integer;
  Pict: TDrwPicture;
  MSOPict: TMSOPicture;
  Stream: TMemoryStream;
{$endif}
begin
{$ifdef XLS_BIFF}
  for i := 0 to Count - 1 do begin
    for j := 0 to Items[i].Drawing.BIFFDrawing.Pictures.Count - 1 do begin
      Pict := Items[i].Drawing.BIFFDrawing.Pictures[j];
      if Pict.PictureId > 0 then begin
        MSOPict := FBIFF.MSOPictures[Pict.PictureId - 1];
        Stream := TMemoryStream.Create;
        try
          MSOPict.SaveToStream(Stream);
          Stream.Seek(0,soFromBeginning);
          case MSOPict.PictType of
  //          msoblipERROR: ;
  //          msoblipUNKNOWN: ;
            msoblipEMF : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.emf',[i + 1,j + 1]),Stream,Pict);
            msoblipWMF : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.wmf',[i + 1,j + 1]),Stream,Pict);
  //          msoblipPICT: ;
            msoblipJPEG: Items[i].Drawing.InsertImage97(Format('Pict%d_%d.jpg',[i + 1,j + 1]),Stream,Pict);
            msoblipPNG : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.png',[i + 1,j + 1]),Stream,Pict);
            msoblipDIB : Items[i].Drawing.InsertImage97(Format('Pict%d_%d.bmp',[i + 1,j + 1]),Stream,Pict);
          end;
        finally
          Stream.Free;
        end;
      end;
    end;
  end;
{$endif}
end;

procedure TXLSReadWriteII5.AfterRead;
var
  i: integer;
begin
  inherited AfterRead;
  FManager.AfterRead;

  FixupRelativelCells;

{$ifdef SHAREWARE}
  for i := 0 to Count - 1 do
    Items[i].AsString[0,0] := 'XLSReadWriteII Copyright(c) 2017 Axolot Data';
{$endif}

{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    for i := 0 to Count - 1 do
      Items[i].Drawing.BIFFDrawing := FBIFF.Sheet[i].DrawingObjects;

    AddBIFFImages;
    AddBIFFComments;
  end;
{$endif}
end;

procedure TXLSReadWriteII5.BeforeRead;
{$ifdef SHAREWARE}
var
  i: integer;
{$endif}
begin
  FManager.BeforeRead;
{$ifdef SHAREWARE}
  for i := 0 to Count - 1 do
    Items[i].AsString[0,0] := 'XLSReadWriteII Copyright(c) 2017 Axolot Data';
{$endif}
end;

procedure TXLSReadWriteII5.BeforeWrite;
{$ifdef SHAREWARE}
var
  i: integer;
{$endif}
begin
  inherited BeforeWrite;

{$ifdef SHAREWARE}
  for i := 0 to Count - 1 do
    Items[i].AsString[0,0] := 'XLSReadWriteII Copyright(c) 2017 Axolot Data';
{$endif}

  SaveBIFFComments;
end;

{$ifdef _AXOLOT_DEBUG}
procedure TXLSReadWriteII5.CheckIntegrity(const AList: TStrings);
begin
  inherited CheckIntegrity(AList);
end;
{$endif}

procedure TXLSReadWriteII5.Clear(ADefaulSheetsCount: integer = 1);
var
  i: integer;
begin
  FManager.Clear;
  inherited Clear;
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    FBIFF.Free;
    FBIFF := Nil;
  end;
{$endif}
  ADefaulSheetsCount := Fork(ADefaulSheetsCount,0,XLS_MAXSHEETS);

  for i := 0 to ADefaulSheetsCount - 1 do
    Add;
end;

constructor TXLSReadWriteII5.Create(AOwner: TComponent);
{$ifndef BABOON}
{$ifdef SHAREWARE}
var
  A,B,C,D: boolean;
{$endif}
{$endif}
begin
  inherited Create(AOwner);

{$ifdef DEMO_TIMELIMIT}
  if Now > EncodeDate(2018,09,01) then begin
    MessageDlg('Application demo time expired.',mtInformation,[mbOk],0);

    Application.Terminate;
  end;
{$endif}

{$ifdef BABOON}
{$else}
  FMonitorFileTimer := TTimer.Create(Self);
  FMonitorFileTimer.Enabled := False;
  FMonitorFileTimer.OnTimer := DoMonitorFileEvent;
{$endif}

{$ifndef BABOON}
{$ifdef SHAREWARE}
{$ifdef SHAREWARE_DLL}
  FNagMsgShown := True;
{$endif}
//  FNagMsgShown := True;

  if not FNagMsgShown then begin
    A := (FindWindow('TApplication', nil) = 0);
//    B := (FindWindow('TAlignPalette', nil) = 0);
//    C := (FindWindow('TPropertyInspector', nil) = 0);
    B := False;
    C := False;
    D := (FindWindow('TAppBuilder', nil) = 0);
    if A or B or C or D then begin
      MessageDlg('This application was built with a demo version of' + #13 +
                  'the XLSReadWriteII components.' + #13 + #13 +
                  'Distributing an application based upon this version' + #13 +
                  'of the components are against the licensing agreement.' + #13 + #13 +
                  'Please see http://www.axolot.com for more information' + #13 +
                  'on purchasing.',mtInformation,[mbOk],0);
      FNagMsgShown := True;
    end;
  end;
{$endif}
{$endif}

  FErrors := TXLSErrorManager.Create;

  FClassFactory := TXLSClassFactoryImpl.Create(Self);

  FManager := TXc12Manager.Create(FClassFactory,FErrors);
  FFormulas := TXLSFormulaHandler.Create(FManager);

  FManager.CreateMembers;

  CreateMembers;

  Clear;
end;

destructor TXLSReadWriteII5.Destroy;
begin
{$ifdef BABOON}
{$else}
  FMonitorFileTimer.Enabled := False;
  FMonitorFileTimer.Free;
{$endif}

  FManager.Free;
  FFormulas.Free;
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    FBIFF.Free;
{$endif}
  FErrors.Free;

  FClassFactory.Free;
  inherited;
end;

function TXLSReadWriteII5.DoBeginFileMonitor: boolean;
{$ifdef DELPHI_2007_OR_LATER}
var
  Dir  : AxUCString;
{$endif}
begin
{$ifdef DELPHI_2007_OR_LATER}
  Result := FManager.Filename <> '';
  if not Result then
    Exit;

  Dir := ExtractFileDir(FManager.Filename);

{$ifdef BABOON}
{$else}
  FHMonitorFile := Windows.FindFirstChangeNotificationW(PWideChar(Dir),False,FILE_NOTIFY_CHANGE_LAST_WRITE);

  Result := FHMonitorFile <> INVALID_HANDLE_VALUE;
  if not Result then
    Exit;

  Result := Windows.RegisterWaitForSingleObject(FHMonitorObj,FHMonitorFile,_FileMonitorCallback,Self,INFINITE,WT_EXECUTEONLYONCE);
{$endif}
{$else}
  Result := False;
{$endif}
end;

procedure TXLSReadWriteII5.DoEndFileMonitor;
begin
{$ifdef DELPHI_2007_OR_LATER}
{$ifdef BABOON}
{$else}
  if FHMonitorObj <> 0 then
    Windows.UnregisterWait(FHMonitorObj);
{$endif}
{$endif}

  FHMonitorObj := 0;
end;

procedure TXLSReadWriteII5.DoMonitorFileEvent(ASender: TObject);
begin
{$ifdef BABOON}
{$else}
  FMonitorFileTimer.Enabled := False;

  if Assigned(FMonitorFileEvent) then
    FMonitorFileEvent(Self);
{$endif}
end;

procedure TXLSReadWriteII5.FixupRelativelCells;
var
  i,j   : integer;
  S     : AxUCString;
  R     : AxUCString;
  VCells: TXLSRelCellsImpl;
  C1,R1,
  C2,R2 : integer;
begin
  for i := 0 to FManager.VirtualCells.Count - 1 do begin
    VCells := TXLSRelCellsImpl(FManager.VirtualCells[i]);
    R := VCells.__Ref;

    S := SplitAtChar('!',R);
    if R <> '' then begin
      // [x]SheetName!... = External sheet. Not handled.
      if (S[1] = '[') and (CPos(']',S) > 2) then begin
        j := -1;
      end
      else begin
        StripQuotes(S);
        j := FManager.Worksheets.Find(S);
      end;
      if j >= 0 then begin
        VCells.Sheet := Sheets[j];
        AreaStrToColRow(R,C1,R1,C2,R2);
        VCells.Col1 := C1;
        VCells.Row1 := R1;
        VCells.Col2 := C2;
        VCells.Row2 := R2;
        VCells.Command := CmdFormat.Commands;
      end;
    end;
  end;

  FManager.VirtualCells.Clear;
end;

{$ifdef XLS_BIFF}
function TXLSReadWriteII5.GetBIFF: TBIFF5;
begin
  Result := FBIFF;
end;
{$endif}

function TXLSReadWriteII5.GetCellReadEvent: TCellReadWriteEvent;
begin
  Result := FManager.OnReadCell;
end;

function TXLSReadWriteII5.GetCellWriteEvent: TCellReadWriteEvent;
begin
  Result := FManager.OnWriteCell;
end;

function TXLSReadWriteII5.GetCompressStrings: boolean;
begin
  Result := FManager.SST.CompressStrings;
end;

function TXLSReadWriteII5.GetDirectRead: boolean;
begin
  Result := FManager.DirectRead;
end;

function TXLSReadWriteII5.GetDirectWrite: boolean;
begin
  Result := FManager.DirectWrite;
end;

function TXLSReadWriteII5.GetFilename: AxUCString;
begin
  Result := FManager.Filename;
end;

{$ifdef XLS_BIFF}
function TXLSReadWriteII5.GetFunctionEvent: TFunctionEvent;
begin
  Result := Nil;
end;
{$endif}

function TXLSReadWriteII5.GetManager: TXc12Manager;
begin
  Result := FManager;
end;

function TXLSReadWriteII5.GetMonitorFile: boolean;
begin
  Result := FHMonitorFile <> 0;
end;

function TXLSReadWriteII5.GetPassword: AxUCString;
begin
  Result := FManager.Password;
end;

function TXLSReadWriteII5.GetProgressEvent: TXLSProgressEvent;
begin
  Result := FManager.OnProgress;
end;

function TXLSReadWriteII5.GetSheetNames(const AFilename: AxUCString; AList: TStrings): integer;
var
  Ext: AxUCString;

procedure DoSheetNames97;
{$ifdef XLS_BIFF}
var
  Stream : TFileStream;
  XLSRead: TXLSReadII;
{$endif}
begin
{$ifdef XLS_BIFF}
  Stream := TFileStream.Create(AFilename,fmOpenRead or fmShareDenyNone);
  try
    XLSRead := TXLSReadII.Create(Nil);
    try
      XLSRead.LoadSheetNamesFromStream(Stream,AList);
    finally
      XLSRead.Free;
    end;
  finally
    Stream.Free;
  end;
{$endif}
end;

procedure DoSheetNames2007;
var
  Stream : TFileStream;
  XLSRead: TXLSReadXLSX;
begin
  XLSRead := TXLSReadXLSX.Create(FManager,Nil);
  Stream := TFileStream.Create(AFilename,fmOpenRead or fmShareDenyNone);
  try
    XLSRead.LoadSheetNamesFromStream(Stream,AList);
  finally
    XLSRead.Free;
    Stream.Free;
  end;
end;

begin
  Ext := Lowercase(ExtractFileExt(AFilename));

  if (Ext = '.xls') or (Ext = '.xlm') or (Ext = '.xlt') then
    DoSheetNames97
  else if (Ext = '.xlsx') or (Ext = '.xlsm') or (Ext = '.xlst') then
    DoSheetNames2007;

  Result := AList.Count;
end;

function TXLSReadWriteII5.GetStrFALSE: AxUCString;
begin
  Result := G_StrFALSE;
end;

function TXLSReadWriteII5.GetStrTRUE: AxUCString;
begin
  Result := G_StrTRUE;
end;

function TXLSReadWriteII5.GetUseAlternateZip: boolean;
begin
  Result := FManager.UseAlternateZip;
end;

function TXLSReadWriteII5.GetUserFunctionEvent: TUserFunctionEvent;
begin
  Result := FFormulas.OnUserFunction;
end;

function TXLSReadWriteII5.GetVersion: TExcelVersion;
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then
    Result := xvExcel97
  else
{$endif}
    Result := xvExcel2007;
end;

function TXLSReadWriteII5.GetVersionNumber: AxUCString;
begin
  Result := CurrentVersionNumber;
{$ifdef SHAREWARE}
  Result := Result + 'a';
{$endif}
end;

procedure TXLSReadWriteII5.LoadFromFile(AFilename: AxUCString);
var
  Ext   : AxUCString;
  Stream: TFileStream;
{$ifdef XLS_CRYPTO_SUPPORT}
  Enc   : TXLSEncryption;
{$endif}
  XLSX  : TXLSReadXLSX;
begin
  SetFilename(AFilename);

  Ext := AnsiLowercase(ExtractFileExt(FManager.Filename));
  if (Ext = '.xls') or (Ext = '.xlm') then begin
    FManager.Version := xvExcel97;

    Clear(0);

    BeforeRead;
{$ifdef XLS_BIFF}
    FBIFF := TBIFF5.Create(FManager);
    FBIFF.Filename := AFilename;
    FBIFF.OnPassword := FPasswordEvent;
    FBIFF.ReadMacros := FReadVBA;
    FBIFF.SkipDrawing := FSkipExcel97Drw;
    FManager._ExtNames97 := FBIFF.FormulaHandler.ExternalNames;
    FManager.Names97 := FBIFF.FormulaHandler.InternalNames;
    FBIFF.Read;
{$endif}
    AfterRead;
  end
  else begin
    FManager.Version := xvExcel2007;
    Stream := TFileStream.Create(FManager.Filename,fmOpenRead or fmShareDenyNone);
    try
      case FileTypeFromMagic(Stream) of
        xkftCompound: begin
{$ifdef XLS_CRYPTO_SUPPORT}
          Enc := TXLSEncryption.Create;
          try
            Enc.Password := FManager.Password;
            Enc.OnPassword := FPasswordEvent;
            Enc.LoadFromStream(Stream);

            if Enc.CryptoResult <> xcrOk then begin
              case Enc.CryptoResult of
                xcrUnknown              : raise XLSRWException.Create('Unknown crypto error');
                xcrMissingPassword      : raise XLSRWException.Create('Missing password');
                xcrWrongPassword        : raise XLSRWException.Create('Wrong password');
                xcrUnsupportedEncryption: raise XLSRWException.Create('Unsupported encryption');
              end;

              Exit;
            end;

            XLSX := TXLSReadXLSX.Create(FManager,FFormulas);
            try
              Clear(0);

              BeforeRead;

              XLSX.LoadFromStream(Enc.OutStream);

              AfterRead;
            finally
              XLSX.Free;
            end;
          finally
            Enc.Free;
          end;
{$else}
         raise XLSRWException.Create('Can not read encrypted file');
{$endif}
        end;
        xkftZIP     : LoadFromStream(Stream);
        else raise XLSRWException.Create('XLSX File read error.');
      end;

    finally
      Stream.Free;
    end;

//    if FReadVBA then begin
//      Strm := FManager.FileData.StreamByName('vbaProject.bin');
//      if Strm <> Nil then begin
//
//      end;
//    end;
  end;
end;

procedure TXLSReadWriteII5.LoadFromStream(AStream: TStream; const AAutoDetect: boolean = True);
var
  XLSX: TXLSReadXLSX;
begin
  case FileTypeFromMagic(AStream) of
    xkftUnknown : ;
    xkftCompound: begin
      LoadFromStream97(AStream);
    end;
    xkftZIP: begin
      XLSX := TXLSReadXLSX.Create(FManager,FFormulas);
      try
        Clear(0);

        BeforeRead;

        XLSX.LoadFromStream(AStream);

        AfterRead;
      finally
        XLSX.Free;
      end;
    end;
    xkftRTF: begin

    end;
  end;
end;

procedure TXLSReadWriteII5.LoadFromStream97(AStream: TStream);
var
  i: integer;
begin
  FManager.Version := xvExcel97;

  Clear(0);

  BeforeRead;
{$ifdef XLS_BIFF}
  FBIFF := TBIFF5.Create(FManager);
  FBIFF.SkipDrawing := FSkipExcel97Drw;
  FBIFF.OnPassword := FPasswordEvent;
  FManager._ExtNames97 := FBIFF.FormulaHandler.ExternalNames;
  FManager.Names97 := FBIFF.FormulaHandler.InternalNames;
  for i := 0 to 1 do
    FBIFF.Sheets.Add;
  FBIFF.LoadFromStream(AStream);
{$endif}
  AfterRead;
end;

procedure TXLSReadWriteII5.Read;
begin
  LoadFromFile(FManager.Filename);
end;

procedure TXLSReadWriteII5.SaveBIFFComments;
{$ifdef XLS_BIFF}
var
  i,j: integer;
  N : TDrwNote;
{$endif}
begin
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    for i := 0 to Count - 1  do begin
      FBIFF[i].DrawingObjects.Notes.Clear;
      for j := 0 to Items[i].Comments.Count - 1 do begin
        N := FBIFF[i].DrawingObjects.Notes.Add;
        N.CellCol := Items[i].Comments[j].Col;
        N.CellRow := Items[i].Comments[j].Row;
        N.Text := Items[i].Comments[j].PlainText;
        N.Author := Items[i].Comments[j].Author;
      end;
    end;
  end;
{$endif}
end;

procedure TXLSReadWriteII5.SaveToFile(AFilename: AxUCString);
var
  Stream: TFileStream;
begin
  SetFilename(AFilename);
{$ifdef XLS_BIFF}
  if FBIFF <> Nil then begin
    BeforeWrite;
{$ifdef _AXOLOT_DEBUG}
    Filename := ChangeFileExt(Filename,'.xls');
{$endif}
    FBIFF.Filename := Filename;
    FBIFF.Write;
  end
  else begin
{$endif}
    SetFilename(Filename);

    Stream := TFileStream.Create(FManager.Filename,fmCreate);
    try
      SaveToStream(Stream);
    finally
      Stream.Free;
    end;
{$ifdef XLS_BIFF}
  end;
{$endif}
end;

procedure TXLSReadWriteII5.SaveToFile(AFilename: AxUCString; AVersion: TExcelVersion);
var
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmCreate);
  try
    SaveToStream(Stream,AVersion);
  finally
    Stream.Free;
  end;
end;

procedure TXLSReadWriteII5.SaveToStream(AStream: TStream; AVersion: TExcelVersion = xvNone);
var
  XLSX: TXLSWriteXLSX;
  Strm: TMemoryStream;
{$ifdef XLS_CRYPTO_SUPPORT}
  Enc : TXLSEncryption;
{$endif}
begin
  if AVersion = xvNone then
    AVersion := GetVersion;

  case AVersion of
    xvExcel97: begin
      BeforeWrite;

{$ifdef XLS_BIFF}
      if FBIFF <> Nil then
        FBIFF.WriteToStream(AStream)
      else
        raise XLSRWException.Create('Can not convert XLSX to XLS');
{$endif}
    end;
    xvExcel2007: begin
      BeforeWrite;

      CalcDimensions;

      FManager.BeforeWrite;

      if FManager.Password <> '' then begin
{$ifdef XLS_CRYPTO_SUPPORT}
        Strm := TMemoryStream.Create;
        XLSX := TXLSWriteXLSX.Create(FManager,FFormulas);
        try
          XLSX.SaveToStream(Strm);
          Strm.Seek(0,soBeginning);

          Enc := TXLSEncryption.Create;
          try
            Enc.InStream := Strm;
            Enc.Password := FManager.Password;
            Enc.SaveToStream(AStream);
          finally
            Enc.Free;
          end;
        finally
          XLSX.Free;
          Strm.Free;
        end;
{$else}
        raise XLSRWException.Create('Can not write encrypted files');
{$endif}
      end
      else begin
        XLSX := TXLSWriteXLSX.Create(FManager,FFormulas);
        try
          XLSX.SaveToStream(AStream);
        finally
          XLSX.Free;
        end;
      end;
    end;
  end;
end;

procedure TXLSReadWriteII5.SetCellReadEvent(const Value: TCellReadWriteEvent);
begin
  FManager.OnReadCell := Value;
end;

procedure TXLSReadWriteII5.SetCellWriteEvent(const Value: TCellReadWriteEvent);
begin
  FManager.OnWriteCell := Value;
end;

procedure TXLSReadWriteII5.SetCompressStrings(const Value: boolean);
begin
  FManager.SST.CompressStrings := Value;
end;

procedure TXLSReadWriteII5.SetDirectRead(const Value: boolean);
begin
  FManager.DirectRead := Value;
end;

procedure TXLSReadWriteII5.SetDirectWrite(const Value: boolean);
begin
  FManager.DirectWrite := Value;
end;

procedure TXLSReadWriteII5.SetDirectWriteArea(const ASheetIndex, ACol1, ARow1, ACol2, ARow2: integer);
begin
  FManager.EventCell.SetTargetArea(ASheetIndex, ACol1, ARow1, ACol2, ARow2);
end;

procedure TXLSReadWriteII5.SetFilename(const Value: AxUCString);
begin
  FManager.Filename := Value;
end;

procedure TXLSReadWriteII5.SetMonitorFile(const Value: boolean);
begin
  if Value then
    DoBeginFileMonitor
  else
    DoEndFileMonitor;
end;

{$ifdef XLS_BIFF}
procedure TXLSReadWriteII5.SetFunctionEvent(const Value: TFunctionEvent);
begin

end;
{$endif}

procedure TXLSReadWriteII5.SetPassword(const Value: AxUCString);
begin
  FManager.Password := Value;
end;

procedure TXLSReadWriteII5.SetProgressEvent(const Value: TXLSProgressEvent);
begin
  FManager.OnProgress := Value;
end;

procedure TXLSReadWriteII5.SetStrFALSE(const Value: AxUCString);
begin
  G_StrFALSE := Value;
end;

procedure TXLSReadWriteII5.SetStrTRUE(const Value: AxUCString);
begin
  G_StrTRUE := Value;
end;

procedure TXLSReadWriteII5.SetUseAlternateZip(const Value: boolean);
begin
  FManager.UseAlternateZip := Value;
end;

procedure TXLSReadWriteII5.SetUserFunctionEvent(const Value: TUserFunctionEvent);
begin
  FFormulas.OnUserFunction := Value;
end;

procedure TXLSReadWriteII5.SetVerionNumber(const Value: AxUCString);
begin

end;

procedure TXLSReadWriteII5.SetVersion(const Value: TExcelVersion);
var
  ZIP: TXLSZipArchive;
  Stream: TPointerMemoryStream;
  Stream97: TMemoryStream;
begin
  if Value <> GetVersion then begin

    if Value = xvExcel97 then begin
      if (Count > 1) or ((Count = 1) and not Items[0].IsEmpty) then
        FErrors.Warning('',XLSWARN_WORKBOOK_NOT_EMPTY);

      Clear;

      FManager.Version := xvExcel97;

      Stream := TPointerMemoryStream.Create;
      try
        Stream.SetStreamData(@XLS_DEFAULT_FILE_97,Length(XLS_DEFAULT_FILE_97));
        ZIP := TXLSZipArchive.Create;
        try
          ZIP.OpenRead(Stream);
          Stream97 := TMemoryStream.Create;
          try
            ZIP[0]._SaveToStream(Stream97);
            Stream97.Seek(0,soFromBeginning);
            LoadFromStream97(Stream97);

            FManager.StyleSheet.XFEditor.LockAll;
          finally
            Stream97.Free;
          end;
        finally
          ZIP.Free;
        end;
      finally
        Stream.Free;
      end;
    end
    else begin
{$ifdef XLS_BIFF}
      if FBIFF <> Nil then begin
        FBIFF.Free;
        FBIFF := Nil;
      end;
{$endif}
      FManager.Version := xvExcel2007;
    end;
  end;
end;

procedure TXLSReadWriteII5.Write;
begin
  SaveToFile(FManager.Filename);
end;

{ TXLSClassFactoryImpl }

constructor TXLSClassFactoryImpl.Create(AOwner: TXLSReadWriteII5);
begin
  FOwner := AOwner;
end;

function TXLSClassFactoryImpl.CreateAClass(AClassType: TXLSClassFactoryType; AClassOwner: TObject): TObject;
begin
  case AClassType of
    xcftNames                : Result := TXLSNames.Create(Self,FOwner.FManager,FOwner.FFormulas);
    xcftNamesMember          : Result := TXLSName.Create(FOwner.Names);
    xcftMergedCells          : Result := TXLSMergedCells.Create(Self);
    xcftMergedCellsMember    : Result := TXLSMergedCell.Create;
    xcftHyperlinks           : Result := TXLSHyperlinks.Create(Self,FOwner.Manager);
    xcftHyperlinksMember     : begin
      Result := TXLSHyperlink.Create;
      TXLSHyperlink(Result).Owner := TXLSHyperlinks(AClassOwner);
    end;
    xcftDataValidations      : Result := TXLSDataValidations.Create(Self);
    xcftDataValidationsMember: Result := TXLSDataValidation.Create;
    xcftConditionalFormat    : Result := TXLSConditionalFormat.Create(FOwner.FManager,FOwner.FFormulas);
    xcftConditionalFormats   : Result := TXLSConditionalFormats.Create(Self);
    xcftAutofilter           : Result := TXLSAutofilter.Create(Self,AClassOwner,FOwner.Names);
    xcftAutofilterColumn     : Result := TXLSAutofilterColumn.Create;
    xcftDrawing              : Result := TXPGDocXLSXDrawing.Create(FOwner.FManager.GrManager);
    xcftVirtualCells         : Result := TXLSRelCellsImpl.Create(Nil);
    else                       Result := Nil;
  end;
end;

//function TXLSClassFactoryImpl.GetAClass(AClassType: TXLSClassFactoryType): TObject;
//begin
//  case AClassType of
//    xcftDrawingManager       : Result := FOwner.FManager.Drawings;
//    else                       Result := Nil;
//  end;
//end;

end.
