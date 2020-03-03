unit XBookPrint2;

{-
********************************************************************************
******* XLSSpreadSheet V3.00                                             *******
*******                                                                  *******
******* Copyright(C) 2006,2017 Lars Arvidsson, Axolot Data               *******
*******                                                                  *******
******* email: components@axolot.com                                     *******
******* URL:   http://www.axolot.com                                     *******
********************************************************************************
** Users of the XLSSpreadSheet component must accept the following            **
** disclaimer of warranty:                                                    **
**                                                                            **
** XLSSpreadSheet is supplied as is. The author disclaims all warranties,     **
** expressed or implied, including, without limitation, the warranties of     **
** merchantability and of fitness for any purpose. The author assumes no      **
** liability for damages, direct or consequential, which may result from the  **
** use of XLSSpreadSheet.                                                     **
********************************************************************************
}

{$B-}
{$H+}
{$R-}
{$I AxCompilers.inc}

interface

uses Classes, SysUtils, Windows, vcl.Controls, Math, vcl.Printers, vcl.Graphics,
     XBookTypes2, XBook_System_2, XLSBook2, XBookPaintGDI2, XBookSkin2, XBookSheet2,
     XBookWindows2, XBookPaintLayers2,
     Xc12Utils5, Xc12DataStylesheet5,
     XLSUtils5, XLSMMU5, XLSCellMMU5, XLSNames5, XLSCellAreas5, XLSSheetData5,
     XLSReadWriteII5;


type TXLSBookPrinter = class(TObject)
private
     procedure SetXSS(const Value: TXLSSpreadSheet);
     function  GetGDI: TAXWGDI;
     function  GetMetafile: TMetafile;
protected
     FXSS         : TXLSSpreadSheet;
     FPrinterDC   : HDC;
     FLayers      : TXPaintLayers;
     FSkin        : TXLSBookSkin;

     FVertAdj     : double;

     FStdFontWidth: integer;
     FDevicePPIX  : integer;
     FDevicePPIY  : integer;
     FPPIY        : integer;

     FPaperWidth  : double;
     FPaperHeight : double;
     // The actually print area after subtracting margins, titles etc.
     FPrintWidth  : double;
     FPrintHeight : double;
     FPrintXMarg  : double;
     FPrintYMarg  : double;

     FPages       : TCellAreas;

     FMinCol      : integer;
     FMaxCol      : integer;
     FMinRow      : integer;
     FMaxRow      : integer;

     procedure CalcStdFontWidth;
     function  AreaIsEmpty(const C1,R1,C2,R2: integer): boolean;
     procedure GetPrinter;

     // Paper coords (centimeters) to pixels.
     function  MapX(const APixelWidth: integer; const ACmX: double): integer;
     function  MapY(const APixelHeight: integer; const ACmY: double): integer;
public
     constructor Create;
     destructor Destroy; override;

     function  Prepare: boolean;
     function  CalcPageCount: integer;

     function  ColWidthCm(const ACol: integer): double;
     function  RowHeightCm(const ARow: integer): double;
     function  HdrColHeightCm: double;
     function  HdrRowWidthCm: double;

     function  GetNextColBreak(var AWidth: double; var ACol: integer): boolean;
     function  GetNextRowBreak(var AHeight: double; var ARow: integer): boolean;

     procedure Paginate(const APage: integer);
     procedure PaintPreviewPage(ACanvas: TCanvas; const AXOffs,AYOffs,AX,AY,AWidth,AHeight: integer);

     procedure PrintPixelsPerInch(out APPIX,APPIY: integer);

     property XSS: TXLSSpreadSheet read FXSS write SetXSS;
     property Skin: TXLSBookSkin read FSkin;
     property GDI: TAXWGDI read GetGDI;
     property Metafile: TMetafile read GetMetafile;

     // Excel is not using a standard DPI of 72. VertAdjustment is used to adjust
     // the row height when printing. The value 0.97 is found by testing.
     property VertAdjustment: double read FVertAdj write FVertAdj;

     property PaperWidth : double read FPaperWidth;
     property PaperHeight: double read FPaperHeight;
     property PrintWidth : double read FPrintWidth;
     property PrintHeight: double read FPrintHeight;
     property PrintXMarg : double read FPrintXMarg;
     property PrintYMarg : double read FPrintYMarg;

     property MinCol     : integer read FMinCol;
     property MaxCol     : integer read FMaxCol;
     property MinRow     : integer read FMinRow;
     property MaxRow     : integer read FMaxRow;

     property Pages      : TCellAreas read FPages;
     end;

type TXLSPDFData = class(TObject)
public
     Printer : TXLSBookPrinter;
     Stream  : TStream;
     Page1   : integer;
     Page2   : integer;
     XMarg   : integer;
     YMarg   : integer;
     end;

type TXLSBookPrint = class(TComponent)
private
     procedure SetXSS(const Value: TXLSSpreadSheet);
protected
     FPrinter        : TXLSBookPrinter;
     FXSS            : TXLSSpreadSheet;
     FFirstPage      : integer;
     FLastPage       : integer;
     FPrintPDFEvent  : TNotifyEvent;
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure Print;

     procedure ExportToPDF(const AFilename: AxUCString); overload;
     function  ExportToPDF(AStream: TStream): boolean; overload;

     property FirstPage: integer read FFirstPage write FFirstPage;
     property LastPage: integer read FLastPage write FLastPage;
     property XSS: TXLSSpreadSheet read FXSS write SetXSS;

     property OnPrintPDF: TNotifyEvent read FPrintPDFEvent write FPrintPDFEvent;
     end;

type TXLSBookPrintPreview = class(TCustomControl)
private
     procedure SetXSS(const Value: TXLSSpreadSheet);
     function  GetPageCount: integer;
     procedure SetVertAdjustment(const Value: double);
     function  GetVertAdjustment: double;
     function  GetCurrPage: integer;
     procedure SetCurrPage(const Value: integer);
     procedure SetShowMargins(const Value: boolean);
protected
     FPrinter        : TXLSBookPrinter;
     FXSS            : TXLSSpreadSheet;

     FColorBackground: longword;
     FColorPaper     : longword;

     FPaperPixMarg   : integer;
     FPaperPixLeft   : integer;
     FPaperPixTop    : integer;
     FPaperPixWidth  : integer;
     FPaperPixHeight : integer;

     FPreviewArea    : TXYRect;

     FCurrPage       : integer;

     FShowMargins    : boolean;

     procedure CalcPaperPixSz;
     // Paper cooords (centimeters) to pixels.
     function  MapX(const ACmX: double): integer;
     function  MapY(const ACmY: double): integer;
     procedure PaintMargins;
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;

     procedure Paint; override;

     function  Execute: boolean;

     property VerticalAdjustment: double read GetVertAdjustment write SetVertAdjustment;
     property PageCount: integer read GetPageCount;
     property CurrPage: integer read GetCurrPage write SetCurrPage;
     property ShowMargins: boolean read FShowMargins write SetShowMargins;
published
     property XSS: TXLSSpreadSheet read FXSS write SetXSS;

     property Align;
     property DoubleBuffered;
     end;

implementation

const PREVIEW_OVERSAMPLING = 4;

{ TXLSBookPrintPreview }

procedure TXLSBookPrintPreview.CalcPaperPixSz;
var
  PScale: double;
  WScale: double;
  M     : integer;
begin
  FPaperPixMarg := 5;

  M := FPaperPixMarg * 2;

  PScale := FPrinter.PaperHeight / FPrinter.PaperWidth;
  WScale := Height / Width;

  if PScale > WScale then begin
    FPaperPixWidth := Round((Height - M)/ PScale);
    FPaperPixHeight := (Height - M);
    FPaperPixLeft := FPaperPixMarg + (Width div 2) - (FPaperPixWidth div 2);
    FPaperPixTop := FPaperPixMarg;
  end
  else begin
    FPaperPixWidth := (Width - M);
    FPaperPixHeight := Round((Width - M) * PScale);
    FPaperPixLeft := FPaperPixMarg;
    FPaperPixTop := FPaperPixMarg + (Height div 2) - (FPaperPixHeight div 2);
  end;
end;

constructor TXLSBookPrintPreview.Create(AOwner: TComponent);
begin
  inherited;

  DoubleBuffered := True;

  FPrinter := TXLSBookPrinter.Create;

  Width := 400;
  Height := 400;

  FShowMargins := False;

  SetXSS(Nil);
end;

destructor TXLSBookPrintPreview.Destroy;
begin
  FPrinter.Free;

  inherited;
end;

function TXLSBookPrintPreview.Execute: boolean;
begin
  Result := FPrinter.Prepare;
  if Result then begin
    Result := FPrinter.CalcPageCount > 0;
    if Result then begin
      FPrinter.Paginate(0);
    end;

    Repaint;
  end;
end;

function TXLSBookPrintPreview.GetCurrPage: integer;
begin
  Result := FCurrPage;
end;

function TXLSBookPrintPreview.GetPageCount: integer;
begin
  Result := FPrinter.Pages.Count;
end;

function TXLSBookPrintPreview.GetVertAdjustment: double;
begin
  Result := FPrinter.VertAdjustment;
end;

function TXLSBookPrintPreview.MapX(const ACmX: double): integer;
begin
  Result := FPreviewArea.X1 + Round((FPreviewArea.X2 - FPreviewArea.X1) * (ACmX / FPrinter.PaperWidth));
end;

function TXLSBookPrintPreview.MapY(const ACmY: double): integer;
begin
  Result := FPreviewArea.Y1 + Round((FPreviewArea.Y2 - FPreviewArea.Y1) * (ACmY / FPrinter.PaperHeight));
end;

procedure TXLSBookPrintPreview.Paint;
var
  ShadowWidth: integer;
begin
  inherited;

  if (FPrinter.PaperWidth <= 0) or (FPrinter.PaperHeight <= 0) then
    Exit;

  CalcPaperPixSz;

  Canvas.Pen.Color := XSS_SYS_COLOR_BLACK;
  Canvas.Brush.Color := FColorBackground;

  Canvas.Rectangle(0,0,Width,Height);

  Canvas.Brush.Color := $00FEFEFE; // FColorPaper;

  Canvas.Rectangle(FPaperPixLeft,FPaperPixTop,FPaperPixLeft + FPaperPixWidth,FPaperPixTop + FPaperPixHeight);

  ShadowWidth := 2;
  Canvas.Brush.Color := XSS_SYS_COLOR_BLACK;
  Canvas.Rectangle(FPaperPixLeft + FPaperPixWidth,FPaperPixTop + ShadowWidth,FPaperPixLeft + FPaperPixWidth + ShadowWidth,FPaperPixTop + FPaperPixHeight + ShadowWidth);
  Canvas.Rectangle(FPaperPixLeft + ShadowWidth,FPaperPixTop + FPaperPixHeight,FPaperPixLeft + FPaperPixWidth + ShadowWidth,FPaperPixTop + FPaperPixHeight + ShadowWidth);

  if not (csDesigning in ComponentState) then begin
    FPreviewArea.X1 := FPaperPixLeft + 1;
    FPreviewArea.Y1 := FPaperPixTop + 1;
    FPreviewArea.X2 := FPaperPixLeft + FPaperPixWidth - 2;
    FPreviewArea.Y2 := FPaperPixTop + FPaperPixHeight - 2;

    FPrinter.PaintPreviewPage(Canvas,
                              MapX(FXSS.XLSSheet.PrintSettings.MarginLeftCm) - MapX(0),
                              MapY(FXSS.XLSSheet.PrintSettings.MarginTopCm) - MapY(0),
                              FPreviewArea.X1,
                              FPreviewArea.Y1,
                              FPreviewArea.X2 - FPreviewArea.X1 - (MapX(FXSS.XLSSheet.PrintSettings.MarginLeftCm + FXSS.XLSSheet.PrintSettings.MarginRightCm) - MapX(0)),
                              FPreviewArea.Y2 - FPreviewArea.Y1 - (MapY(FXSS.XLSSheet.PrintSettings.MarginTopCm + FXSS.XLSSheet.PrintSettings.MarginBottomCm) - MapY(0)));

    if FShowMargins then
      PaintMargins;
  end;
end;

procedure TXLSBookPrintPreview.PaintMargins;
begin
  Canvas.Pen.Color := clBlack;
  Canvas.Brush.Style := bsClear;
  Canvas.Pen.Style := psDot;

  Canvas.MoveTo(MapX(FXSS.XLSSheet.PrintSettings.MarginLeftCm),MapY(0));
  Canvas.LineTo(MapX(FXSS.XLSSheet.PrintSettings.MarginLeftCm),MapY(FPrinter.PaperHeight) + 1);

  Canvas.MoveTo(MapX(FXSS.XLSSheet.PrintSettings.MarginLeftCm + FPrinter.PrintWidth),MapY(0));
  Canvas.LineTo(MapX(FXSS.XLSSheet.PrintSettings.MarginLeftCm + FPrinter.PrintWidth),MapY(FPrinter.PaperHeight) + 1);

  Canvas.MoveTo(MapX(0),MapY(FXSS.XLSSheet.PrintSettings.MarginTopCm));
  Canvas.LineTo(MapX(FPrinter.PaperWidth),MapY(FXSS.XLSSheet.PrintSettings.MarginTopCm));

  Canvas.MoveTo(MapX(0),MapY(FXSS.XLSSheet.PrintSettings.MarginTopCm + FPrinter.PrintHeight));
  Canvas.LineTo(MapX(FPrinter.PaperWidth),MapY(FXSS.XLSSheet.PrintSettings.MarginTopCm + FPrinter.PrintHeight));

//  Canvas.MoveTo(MapX(0),MapY(FXSS.XLSSheet.PrintSettings.HeaderMarginCm));
//  Canvas.LineTo(MapX(FPrinter.PaperWidth),MapY(FXSS.XLSSheet.PrintSettings.HeaderMarginCm));
//
//  Canvas.MoveTo(MapX(0),MapY(FPrinter.PaperHeight - FXSS.XLSSheet.PrintSettings.FooterMarginCm));
//  Canvas.LineTo(MapX(FPrinter.PaperWidth),MapY(FPrinter.PaperHeight - FXSS.XLSSheet.PrintSettings.FooterMarginCm));

  Canvas.Pen.Style := psSolid;
end;

procedure TXLSBookPrintPreview.SetCurrPage(const Value: integer);
begin
  FCurrPage := Value;
  FPrinter.Paginate(FCurrPage);
  Paint;
end;

procedure TXLSBookPrintPreview.SetShowMargins(const Value: boolean);
begin
  FShowMargins := Value;
  Paint
end;

procedure TXLSBookPrintPreview.SetVertAdjustment(const Value: double);
begin
  FPrinter.VertAdjustment := Fork(Value,0.8,1.0);
end;

procedure TXLSBookPrintPreview.SetXSS(const Value: TXLSSpreadSheet);
begin
  FXSS := Value;

  if FXSS <> Nil then begin
    FPrinter.XSS := FXSS;
    FColorBackground := RevRGB(FPrinter.Skin.Colors.WindowBkg)
  end
  else begin
    FPrinter.XSS := Nil;
    FColorBackground := XSS_SYS_COLOR_GRAY;
  end;
  FColorPaper := XSS_SYS_COLOR_WHITE;
end;

{ TXLSBookPrint }

constructor TXLSBookPrint.Create(AOwner: TComponent);
begin
  inherited;

  FPrinter := TXLSBookPrinter.Create;

  FFirstPage := 0;
  FLastPage := MAXINT;
end;

destructor TXLSBookPrint.Destroy;
begin
  FPrinter.Free;

  inherited;
end;

procedure TXLSBookPrint.ExportToPDF(const AFilename: AxUCString);
var
  Ok: boolean;
  Stream: TFileStream;
begin
  Stream := TFileStream.Create(AFilename,fmCreate);
  try
    Ok := ExportToPDF(Stream);
  finally
    Stream.Free;
  end;
  if not Ok then
    SysUtils.DeleteFile(AFilename);
end;

function TXLSBookPrint.ExportToPDF(AStream: TStream): boolean;
var
  Data: TXLSPDFData;
begin
  Result := False;
  if Assigned(FPrintPDFEvent) then begin
    Data := TXLSPDFData.Create;
    try
      if FPrinter.Prepare then begin
        if FPrinter.CalcPageCount > 0 then begin
          Data.Printer := FPrinter;
          Data.Stream := AStream;
          Data.Page1 := Min(FFirstPage,FPrinter.Pages.Count - 1);
          Data.Page2 := Min(FLastPage,FPrinter.Pages.Count - 1);

          // Resonable accurate. The real problem is that PaperWidth is used instead of PrintWidth when createing the metafile.
          Data.XMarg := Round((70 / 1.78) * (FPrinter.PrintXMarg / 2));
          Data.YMarg := Round((70 / 1.78) * (FPrinter.PrintYMarg / 2));
//          Data.XMarg := Round((FPrinter.Metafile.Width / FPrinter.PrintWidth) * FPrinter.PrintXMarg);
//          Data.YMarg := Round((FPrinter.Metafile.Height / FPrinter.PrintHeight) * FPrinter.PrintYMarg);

          FPrintPDFEvent(Data);

          Result := True;
        end;
      end;
    finally
      Data.Free;
    end;
  end;
end;

procedure TXLSBookPrint.Print;
var
  i    : integer;
  P1,P2: integer;
  XM,YM: integer;
begin
  if FPrinter.Prepare then begin
    if FPrinter.CalcPageCount > 0 then begin
      if psoPortrait in FXSS.XLSSheet.PrintSettings.Options then
        Printer.Orientation := poPortrait
      else
        Printer.Orientation := poLandscape;

      P1 := Min(FFirstPage,FPrinter.Pages.Count - 1);
      P2 := Min(FLastPage,FPrinter.Pages.Count - 1);

      Printer.BeginDoc;
      for i := P1 to P2 do begin
        FPrinter.Paginate(i);

{$ifdef DELPHI_2006_OR_LATER}
        FPrinter.Metafile.SetSize(Round(FPrinter.PaperWidth * (FPrinter.FDevicePPIX / 2.54)),Round(FPrinter.PaperHeight * (FPrinter.FDevicePPIY / 2.54)));
{$else}
        FPrinter.Metafile.Width := Round(FPrinter.PrintWidth * (FPrinter.FDevicePPIX / 2.54));
        FPrinter.Metafile.Height := Round(FPrinter.PrintHeight * (FPrinter.FDevicePPIY / 2.54));
{$endif}
        XM := Round((FPrinter.Metafile.Width / FPrinter.PrintWidth) * FPrinter.PrintXMarg);
        YM := Round((FPrinter.Metafile.Height / FPrinter.PrintHeight) * FPrinter.PrintYMarg);

        Printer.Canvas.Draw(XM,YM,FPrinter.Metafile);

        if i < P2 then
          Printer.NewPage;
      end;
      Printer.EndDoc;
    end;
  end;
end;

procedure TXLSBookPrint.SetXSS(const Value: TXLSSpreadSheet);
begin
  FXSS := Value;
  FPrinter.XSS := Value;
end;

{ TXLSBookPrinter }

function TXLSBookPrinter.AreaIsEmpty(const C1, R1, C2, R2: integer): boolean;
var
  C,R: integer;
  i: integer;
  Cell: TXLSCellItem;
  XF: TXc12XF;
begin
  Result := False;
  for R := R1 to R2 do begin
    for C := C1 to C2 do begin
      Cell := FXSS.XLSSheet.MMUCells.FindCell(C,R);
      if Cell.Data <> Nil then begin
        if FXSS.XLSSheet.MMUCells.CellType(@Cell) = xctBlank then begin
          i := FXSS.XLSSheet.MMUCells.GetStyle(@Cell);
          XF := FXSS.XLS.Manager.StyleSheet.XFs[i];
          if (XF.Fill <> FXSS.XLS.Manager.StyleSheet.Fills.DefaultFill) or (XF.Border <> FXSS.XLS.Manager.StyleSheet.Borders.DefaultBorder) then
            Exit;
        end
        else
          Exit;
      end;
    end;
  end;
  Result := True;
end;

procedure TXLSBookPrinter.CalcStdFontWidth;
var
  i: integer;
  S: AxUCString;
  Sz: TSize;
begin
  FStdFontWidth := 0;
  FXSS.XLS.Font.DeleteHandle;
  Windows.SelectObject(FPrinterDC,FXSS.XLS.Font.GetHandle(Windows.GetDeviceCaps(FPrinterDC,LOGPIXELSY)));
  for i := 1 to 10 do begin
    S := StrDigits[i];
    Windows.GetTextExtentPoint32W(FPrinterDC,PWideChar(S),1,Sz);
    if Sz.cx > FStdFontWidth then
      FStdFontWidth := Sz.cx;
  end;
  FXSS.XLS.Font.DeleteHandle;
end;

function TXLSBookPrinter.ColWidthCm(const ACol: integer): double;
begin
  Result := ((FXSS.XLSSheet.Columns[ACol].CharWidth * FStdFontWidth) / FDevicePPIX) * 2.54 * 0.95;
end;

constructor TXLSBookPrinter.Create;
begin
  FPages := TCellAreas.Create;

  FVertAdj := 0.97;

  FLayers := TXPaintLayers.Create(Nil);
  FLayers.LayerMode := plmMetafile;
  FSkin := TXLSBookSkin.Create(0,FLayers.MetafileLayer);
end;

destructor TXLSBookPrinter.Destroy;
begin
  FLayers.Free;
  FSkin.Free;
  Windows.DeleteDC(FPrinterDC);
  FPages.Free;

  inherited;
end;

function TXLSBookPrinter.GetGDI: TAXWGDI;
begin
  Result := FLayers.MetafileLayer;
end;

function TXLSBookPrinter.GetMetafile: TMetafile;
begin
  Result := FLayers.MetafileLayer.Metafile;
end;

function TXLSBookPrinter.GetNextColBreak(var AWidth: double; var ACol: integer): boolean;
begin
  Result := False;
  while ACol <= FMaxCol do begin
    if FXSS.XLSSheet.PrintSettings.VertPagebreaks.Find(ACol) <> Nil then
      Exit;
    if (AWidth + ColWidthCm(ACol)) >= FPrintWidth then
      Exit;
    AWidth := AWidth + ColWidthCm(ACol);
    Inc(ACol);
    Result := True;
    if ACol > FMaxCol then
      Exit;
  end;
end;

function TXLSBookPrinter.GetNextRowBreak(var AHeight: double; var ARow: integer): boolean;
begin
  Result := False;
  while ARow <= FMaxRow do begin
    if (FXSS.XLSSheet.PrintSettings.HorizPagebreaks.Find(ARow + 1) <> Nil) or ((AHeight + RowHeightCm(ARow)) >= FPrintHeight) then begin
      Inc(ARow);
      Exit;
    end;
    AHeight := AHeight + RowHeightCm(ARow);
    Inc(ARow);
    Result := True;
    if ARow > FMaxRow then
      Exit;
  end;
end;

procedure TXLSBookPrinter.GetPrinter;
var
  DeviceMode: THandle;
  DevMode: PDeviceMode;
  Device, Driver, Port: array[0..MAX_PATH] of Char;
begin
  Printer.GetPrinter(Device, Driver, Port, DeviceMode);
  DevMode := PDevMode(Windows.GlobalLock(DeviceMode));
  try
    FPrinterDC := Windows.CreateDC(Driver,Device,Nil,DevMode);

    if psoPortrait in FXSS.XLSSheet.PrintSettings.Options then begin
      DevMode.dmPaperWidth := Round(FXSS.XLSSheet.PrintSettings.PaperSizeDim.X * 100);
      DevMode.dmPaperLength := Round(FXSS.XLSSheet.PrintSettings.PaperSizeDim.Y * 100);
    end
    else begin
      DevMode.dmPaperWidth := Round(FXSS.XLSSheet.PrintSettings.PaperSizeDim.Y * 100);
      DevMode.dmPaperLength := Round(FXSS.XLSSheet.PrintSettings.PaperSizeDim.X * 100);
    end;
    Printer.SetPrinter(Device, Driver, Port, DeviceMode);

    FPaperWidth := DevMode.dmPaperWidth / 100;
    FPaperHeight := DevMode.dmPaperLength / 100;
  finally
    Windows.GlobalUnlock(DeviceMode);
  end;
end;

function TXLSBookPrinter.HdrColHeightCm: double;
begin
  Result := RowHeightCm(MAXINT);
end;

function TXLSBookPrinter.HdrRowWidthCm: double;
begin
  Result := 2.0; // TODO
end;

function TXLSBookPrinter.MapX(const APixelWidth: integer; const ACmX: double): integer;
begin
  Result := Round(APixelWidth * (ACmX / FPaperWidth));
end;

function TXLSBookPrinter.MapY(const APixelHeight: integer; const ACmY: double): integer;
begin
  Result := Round(APixelHeight * (ACmY / FPaperHeight));
end;

function TXLSBookPrinter.CalcPageCount: integer;
var
  C1,R1,C2,R2: integer;
  W,H: double;
begin
  W := 0;
  H := 0;

  C1 := 0; //FMinCol;
  R1 := 0; //FMinRow;
  C2 := C1;
  R2 := R1;

  while GetNextRowBreak(H,R2) do begin
    while GetNextColBreak(W,C2) do begin
      if not AreaIsEmpty(C1,R1,C2,R2) then
        FPages.Add(C1,R1,C2 - 1,R2 - 1);
      W := 0;
      C1 := C2;
    end;
    C1 := 0; //FMinCol;
    C2 := C1;

    R1 := R2;
    H := 0;
  end;

  Result := FPages.Count;
end;

procedure TXLSBookPrinter.Paginate(const APage: integer);
var
  Win: TXSSWindow;
  XSheet: TXLSBookSheet;
  ClipHandle: longword;
begin
  FXSS.XLS.Manager.StyleSheet.Fonts.DeleteHandles;

  FXSS.XLS.Manager.StyleSheet.OnPixelsPerInch := PrintPixelsPerInch;

  FLayers.MetafileLayer.BeginMetafile(FPrinterDC,FPaperWidth,FPaperHeight);

  FSkin.AssignSystemFont(FXSS.XLS.Font.Name,Round(FXSS.XLS.Font.Size * FSkin.GDI.Zoom),Integer(FXSS.XLS.Font.Charset));
  Win := TXSSWindow.Create(FSkin);
  XSheet := TXLSBookSheet.Create(Nil,Win,FXSS.Options,FLayers,FXSS.XLS.Manager,FXSS.XLS,FXSS.XLSSheet);
  try
    XSheet.BeginPrint(FVertAdj,Round(FStdFontWidth * FLayers.MetafileLayer.YScale));

    XSheet.Skin.Colors.SheetGridline := XSS_SYS_COLOR_BLACK;
    XSheet.SetSize(0,0,Round(FPrintWidth * 1000),Round(FPrintHeight * 1000));

    ClipHandle := FSkin.GDI.CreateClipRect(0,0,Round(XSheet.Columns.PageWidth * 1.001),Round(XSheet.Rows.PageHeight * 1.001));
    try
      XSheet.PaintPrint(FXSS.XLSSheet,FPages[APage].Col1,FPages[APage].Row1,FPages[APage].Col2,FPages[APage].Row2);
    finally
      FSkin.GDI.DeleteClipRect(ClipHandle);
    end;

    XSheet.EndPrint;
  finally
    XSheet.Free;
    Win.Free;
  end;

  FLayers.MetafileLayer.EndMetafile;

  FXSS.XLS.Manager.StyleSheet.OnPixelsPerInch := Nil;
  FXSS.XLS.Manager.StyleSheet.Fonts.DeleteHandles;
end;

procedure TXLSBookPrinter.PaintPreviewPage(ACanvas: TCanvas; const AXOffs,AYOffs,AX,AY,AWidth,AHeight: integer);
var
  W,H: integer;
  BMP: TBitmap;
begin
  BMP := TBitmap.Create;
  try
    W := AWidth * PREVIEW_OVERSAMPLING;
    H := AHeight * PREVIEW_OVERSAMPLING;
{$ifdef DELPHI_2006_OR_LATER}
    BMP.SetSize(W,H);
{$else}
    BMP.Width := W;
    BMP.Height := H;
{$endif}
    FLayers.MetafileLayer.PlayMetafile(BMP.Canvas,0,0,W,H);

    Windows.SetStretchBltMode(ACanvas.Handle,STRETCH_HALFTONE);
    Windows.StretchBlt(ACanvas.Handle,AX + AXOffs,AY + AYOffs,AWidth,AHeight,
                       BMP.Canvas.Handle,0,0,W - (AXOffs * PREVIEW_OVERSAMPLING),H - (AYOffs * PREVIEW_OVERSAMPLING),SRCCOPY);
//    Windows.StretchBlt(ACanvas.Handle,AX + AXOffs,AY + AYOffs,AWidth - AXOffs,AHeight - AYOffs,
//                       BMP.Canvas.Handle,0,0,W - (AXOffs * PREVIEW_OVERSAMPLING),H - (AYOffs * PREVIEW_OVERSAMPLING),SRCCOPY);
  finally
    BMP.Free
  end;

//  Printer.BeginDoc;
//
//  FLayers.MetafileLayer.PlayMetafile(Printer.Canvas,Round(FPrintXMarg * (FDevicePPIX / 2.54)),Round(FPrintYMarg * (FDevicePPIY / 2.54)),Printer.PageWidth,Printer.PageHeight);
//
//  Printer.EndDoc;
end;

function TXLSBookPrinter.Prepare: boolean;
var
  i: integer;
  C1,R1,C2,R2: integer;
begin
  Result := Printer.PrinterIndex >= 0;
  if not Result then
    Exit;

  if FXSS = Nil then begin
    FPaperWidth := 21.0;
    FPaperHeight := 29.7;

    FMinCol := 0;
    FMaxCol := 0;
    FMinRow := 0;
    FMaxRow := 0;
  end
  else begin
    GetPrinter;

    FLayers.MetafileLayer.BeginMetafile(FPrinterDC,FPaperWidth,FPaperHeight);
    try
      FDevicePPIX := FLayers.MetafileLayer.DevicePixelsPerInchX;
      FDevicePPIY := FLayers.MetafileLayer.DevicePixelsPerInchY;
      FPPIY := FLayers.MetafileLayer.PixelsPerInchY;
      CalcStdFontWidth;
    finally
      FLayers.MetafileLayer.EndMetafile;
    end;

    FXSS.XLSSheet.CalcDimensions;

    if FXSS.XLSSheet.PrintSettings.GetPrintArea(C1,R1,C2,R2) then begin
      FMinCol := Max(FXSS.XLSSheet.FirstCol,C1);
      FMaxCol := Min(FXSS.XLSSheet.LastCol,C2);
      FMinRow := Max(FXSS.XLSSheet.FirstRow,R1);
      FMaxRow := Min(FXSS.XLSSheet.LastRow,R2);
    end
    else begin
      FMinCol := FXSS.XLSSheet.FirstCol;
      FMaxCol := FXSS.XLSSheet.LastCol;
      FMinRow := FXSS.XLSSheet.FirstRow;
      FMaxRow := FXSS.XLSSheet.LastRow;
    end;

    FPrintXMarg := FXSS.XLSSheet.PrintSettings.MarginLeftCm;
    FPrintYMarg := FXSS.XLSSheet.PrintSettings.MarginTopCm;
    FPrintWidth  := FPaperWidth - FXSS.XLSSheet.PrintSettings.MarginLeftCm - FXSS.XLSSheet.PrintSettings.MarginRightCm;
    FPrintHeight := FPaperHeight - FXSS.XLSSheet.PrintSettings.MarginTopCm - FXSS.XLSSheet.PrintSettings.MarginBottomCm;
//                                   FXSS.XLSSheet.PrintSettings.HeaderMarginCm - FXSS.XLSSheet.PrintSettings.FooterMarginCm;

    if FXSS.XLSSheet.PrintSettings.GetPrintTitlesCols(C1,C2) then begin
      for i := C1 to C2 do
        FPrintWidth := FPrintWidth - ColWidthCm(i);
    end;
    if FXSS.XLSSheet.PrintSettings.GetPrintTitlesRows(R1,R2) then begin
      for i := R1 to R2 do
        FPrintHeight := FPrintHeight - RowHeightCm(i);
    end;

    if psoRowColHeading in FXSS.XLSSheet.PrintSettings.Options then begin
      FPrintWidth := FPrintWidth - HdrRowWidthCm;
      FPrintHeight := FPrintHeight - HdrColHeightCm;
    end;
  end;
end;

procedure TXLSBookPrinter.PrintPixelsPerInch(out APPIX, APPIY: integer);
begin
  APPIX := Round(GetDeviceCaps(FPrinterDC, LOGPIXELSX) * FLayers.MetafileLayer.XScale);
  APPIY := Round(GetDeviceCaps(FPrinterDC, LOGPIXELSY) * FLayers.MetafileLayer.YScale);
end;

function TXLSBookPrinter.RowHeightCm(const ARow: integer): double;
begin
//  Result := ((FXSS.XLSSheet.Rows[ARow].PixelHeight / FPPIY) * 2.54) * 0.9;
//  Result := ((FXSS.XLSSheet.Rows[ARow].HeightPt / 72) * 2.54) * FVertAdj;
  Result := ((FXSS.XLSSheet.Rows[ARow].HeightPt / 72) * 2.54) * 0.99;
end;

procedure TXLSBookPrinter.SetXSS(const Value: TXLSSpreadSheet);
begin
  FXSS := Value;
end;

end.
