unit ExcelColorPicker;

interface

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

uses
  SysUtils, Classes, vcl.Controls, vcl.Graphics, Messages, Windows,
{$ifdef DELPHI_XE3_OR_LATER}
     System.UITypes,
{$endif}
  Xc12Utils5, XLSUtils5, Math, vcl.Dialogs;

type TColorSwatch = class(TObject)
public
     Col,Row: integer;
     Color: TXc12Color;
     X,Y: integer;
     end;

const SWATCH_BORDERCOLOR = $00C5C5C5;
const SWATCH_MOUSECOLOR1 = $003694F2;
const SWATCH_MOUSECOLOR2 = $0094E2FF;
const SWATCH_SELECTEDCOLOR = $001048EF;

type TColorSwatches = class(TObject)
private
     procedure SetCols(const Value: integer);
     procedure SetRows(const Value: integer);
     procedure SetSwatchSize(const Value: integer);
     procedure SetSwatchXSpacing(const Value: integer);
     procedure SetSwatchYSpacing(const Value: integer);
     procedure SetY1(const Value: integer);
     procedure SetX1(const Value: integer);
     function  GetSwatches(Col, Row: integer): TColorSwatch;
protected
     FX1,FY1,FX2,FY2: integer;
     FHorizMarg,FVertMarg: integer;
     FCols,FRows: integer;
     FSwatchSize: integer;
     FSwatchXSpacing: integer;
     FSwatchYSpacing: integer;
     FSwatches: array of array of TColorSwatch;
     FCurrSwatch: TColorSwatch;
     FSelectedSwatch: TColorSwatch;
     FTempSwatch: TColorSwatch;

     procedure Clear;
     procedure InitSwatches; virtual;
     procedure PaintSwatch(Canvas: TCanvas; Swatch: TColorSwatch);
     procedure PaintMouseSwatch(Canvas: TCanvas; Swatch: TColorSwatch);
     function  Hit(X,Y: integer): boolean;
public
     constructor Create(Cols,Rows: integer);
     destructor Destroy; override;

     procedure Paint(Canvas: TCanvas);
     procedure MouseMove(Canvas: TCanvas; X,Y: integer);
     procedure MouseUp(Canvas: TCanvas; X,Y: integer);
     procedure ClearMouse(Canvas: TCanvas);
     function  ClientHit(X,Y: integer): boolean;
     procedure ClearSelected(Canvas: TCanvas);
     function  SetSelected(Canvas: TCanvas; Color: TXc12Color): boolean;

     property X1: integer read FX1 write SetX1;
     property Y1: integer read FY1 write SetY1;
     property X2: integer read FX2;
     property Y2: integer read FY2;
     property Cols: integer read FCols write SetCols;
     property Rows: integer read FRows write SetRows;
     property SwatchSize: integer read FSwatchSize write SetSwatchSize;
     property SwatchXSpacing: integer read FSwatchXSpacing write SetSwatchXSpacing;
     property SwatchYSpacing: integer read FSwatchYSpacing write SetSwatchYSpacing;
     property Swatches[Col,Row: integer]: TColorSwatch read GetSwatches; default;
     end;

type TSwatchContainer = class(TObject)
private
     procedure SetX1(const Value: integer);
     procedure SetY1(const Value: integer);
protected
     FX1,FY1,FX2,FY2: integer;
     FSwatchList: array of TColorSwatches;
     FSelectedColor: TXc12Color;

     procedure Initiate; virtual;
     procedure _ClearSelected;
public
     constructor Create;
     destructor Destroy; override;

     procedure Paint(Canvas: TCanvas);
     procedure MouseMove(Canvas: TCanvas; X,Y: integer);
     procedure MouseUp(Canvas: TCanvas; X,Y: integer);
     procedure ClearMouse(Canvas: TCanvas);
     procedure ClearSelected(Canvas: TCanvas);
     function  ClientHit(X,Y: integer): boolean;
     function  SwatchHit: boolean;
     function  SetSelected(Canvas: TCanvas; Color: TXc12Color): boolean;

     property X1: integer read FX1 write SetX1;
     property Y1: integer read FY1 write SetY1;
     property X2: integer read FX2;
     property Y2: integer read FY2;
     end;

type TEx97Swatches = class(TSwatchContainer)
private
protected
     procedure Initiate; override;
public
     constructor Create;
     end;

type TEx2007ThemeSwatches = class(TSwatchContainer)
protected
     FAdjust: array[0..9,0..4] of double;

     procedure Initiate; override;
public
     constructor Create(Compact: boolean);
     end;

type TEx2007StandardSwatches = class(TSwatchContainer)
private
protected
     procedure Initiate; override;
public
     constructor Create;
     end;

type TExcelColorPickerMode = (ecpmNone,ecpmExcel2007ThemeCompact,ecpmExcel2007Theme,ecpmExcel2007Standard,ecpmExcel97);

type TExcelColorPicker = class(TGraphicControl)
private
     procedure SetColorMode(const Value: TExcelColorPickerMode);
     function  GetExcelColor: TXc12Color;
     procedure SetExcelColor(const Value: TXc12Color);
     procedure SetLinkedPicker(const Value: TExcelColorPicker);
protected
     FColorMode    : TExcelColorPickerMode;
     FColorSwatches: TSwatchContainer;
     FClickEvent   : TNotifyEvent;
     FLinkedPicker : TExcelColorPicker;
     FParentPicker : TExcelColorPicker;

     procedure MouseMove(Shift: TShiftState; X, Y: Integer); override;
     procedure MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer); override;
     procedure WMMouseLeave(var Message: TMessage); message WM_MOUSELEAVE;
public
     constructor Create(AOwner: TComponent); override;
     destructor Destroy; override;
     procedure Paint; override;
     procedure SetBounds(ALeft, ATop, AWidth, AHeight: Integer); override;

     function  FindAndSelect(AColor: TXc12Color): boolean;

     property ExcelColor: TXc12Color read GetExcelColor write SetExcelColor;
published
     property ColorMode: TExcelColorPickerMode read FColorMode write SetColorMode;
     property LinkedPicker: TExcelColorPicker read FLinkedPicker write SetLinkedPicker;

     property OnClick: TNotifyEvent read FClickEvent write FClickEvent;
     end;

implementation

const Ex12StandardColors: array[0..9] of longword =
($00C00000,$00FF0000,$00FFC000,$00FFFF00,$0092D050,$0000B050,$0000B0F0,$000070C0,$00002060,$007030A0);

function RevRGB(RGB: longword): longword; {$ifdef D2006PLUS} inline; {$endif}
begin
  Result := ((RGB and $FF0000) shr 16) + (RGB and $00FF00) + ((RGB and $0000FF) shl 16);
end;


{ TExcelColorPicker }

constructor TExcelColorPicker.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  Color := clWhite;

  SetColorMode(ecpmExcel2007Theme);

//  OnMouseLeave := MouseLeave;
end;

destructor TExcelColorPicker.Destroy;
begin
  FColorSwatches.Free;
  inherited;
end;

function TExcelColorPicker.FindAndSelect(AColor: TXc12Color): boolean;
var
  ECP: TExcelColorPicker;
begin
  ECP := Self;

  while ECP <> Nil do begin
    Result := FColorSwatches.SetSelected(Canvas,AColor);
    if Result then
      Exit;
    ECP := ECP.LinkedPicker;
  end;
  Result := False;
end;

function TExcelColorPicker.GetExcelColor: TXc12Color;
begin
  Result := FColorSwatches.FSelectedColor;
end;

procedure TExcelColorPicker.MouseDown(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;

//  SetFocus;
end;

procedure TExcelColorPicker.MouseMove(Shift: TShiftState; X, Y: Integer);
begin
  inherited;

   if FParentPicker <> Nil then
     FParentPicker.FColorSwatches.ClearMouse(FParentPicker.Canvas);

   if FColorSwatches.ClientHit(X,Y) then
     FColorSwatches.MouseMove(Canvas,X,Y)
   else
     FColorSwatches.ClearMouse(Canvas);
end;

procedure TExcelColorPicker.MouseUp(Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  if FColorSwatches.ClientHit(X,Y) and FColorSwatches.SwatchHit then begin
    FColorSwatches.MouseUp(Canvas,X,Y);
    if Assigned(FLinkedPicker) then begin
      FLinkedPicker.FColorSwatches._ClearSelected;
      FLinkedPicker.FColorSwatches.FSelectedColor := FColorSwatches.FSelectedColor;
      FLinkedPicker.Invalidate;
      if (FParentPicker <> NIl) and Assigned(FParentPicker.FClickEvent) then
        FParentPicker.FClickEvent(Self)
      else if Assigned(FClickEvent) then
        FClickEvent(Self);
    end
    else if Assigned(FClickEvent) then
      FClickEvent(Self);
  end;
end;

procedure TExcelColorPicker.Paint;
begin
  inherited;
  Canvas.Brush.Color := Color;
  Canvas.Pen.Color := Canvas.Brush.Color;
  Canvas.Rectangle(0,0,Width,Height);

  FColorSwatches.Paint(Canvas);
end;

procedure TExcelColorPicker.SetBounds(ALeft, ATop, AWidth, AHeight: Integer);
begin
  AWidth := FColorSwatches.X2 - FColorSwatches.X1;
  AHeight := FColorSwatches.Y2 - FColorSwatches.Y1;
  inherited SetBounds(ALeft, ATop, AWidth, AHeight);
end;

procedure TExcelColorPicker.SetColorMode(const Value: TExcelColorPickerMode);
begin
  if (Value <> ecpmNone) and (Value <> FColorMode) then begin
    FColorMode := Value;

    if FColorSwatches <> Nil then
      FColorSwatches.Free;

    case FColorMode of
      ecpmExcel2007Theme:
        FColorSwatches := TEx2007ThemeSwatches.Create(False);
      ecpmExcel2007ThemeCompact:
        FColorSwatches := TEx2007ThemeSwatches.Create(True);
      ecpmExcel2007Standard:
        FColorSwatches := TEx2007StandardSwatches.Create;
      ecpmExcel97:
        FColorSwatches := TEx97Swatches.Create;
    end;
    if Parent <> Nil then begin
      Parent.Invalidate;
      UpdateBoundsRect(Rect(Left,Top,Left + (FColorSwatches.X2 - FColorSwatches.X1),Top + (FColorSwatches.Y2 - FColorSwatches.Y1))); ;
      if Canvas.HandleAllocated then
        Paint;
    end;
  end;
end;

procedure TExcelColorPicker.SetExcelColor(const Value: TXc12Color);
begin
  case Value.ColorType of
    exctAuto      : begin
      SetColorMode(ecpmExcel2007Theme);
      FColorSwatches.ClearSelected(Canvas);
      if FLinkedPicker <> Nil then
        FLinkedPicker.FColorSwatches.ClearSelected(FLinkedPicker.Canvas);
    end;
    exctIndexed   : SetColorMode(ecpmExcel97);
    exctRgb       : SetColorMode(ecpmExcel2007Theme);
    exctTheme     : SetColorMode(ecpmExcel2007Theme);
    exctUnassigned: SetColorMode(ecpmExcel2007Theme);
  end;
  if not FColorSwatches.SetSelected(Canvas,Value) and (FLinkedPicker <> Nil) then
    FLinkedPicker.FColorSwatches.SetSelected(FLinkedPicker.Canvas,Value);
end;

procedure TExcelColorPicker.SetLinkedPicker(const Value: TExcelColorPicker);
begin
  FLinkedPicker := Value;
  if Value <> Nil then begin
    Value.FLinkedPicker := Self;
    Value.FParentPicker := Self;
  end;
end;

procedure TExcelColorPicker.WMMouseLeave(var Message: TMessage);
begin
  FColorSwatches.ClearMouse(Canvas);
end;

{ TColorSwatches }

procedure TColorSwatches.InitSwatches;
var
  Col,Row: integer;
begin
  Clear;
  SetLength(FSwatches,FCols,FRows);
  for Row := 0 to FRows - 1 do begin
    for Col := 0 to FCols - 1 do begin
      FSwatches[Col,Row] := TColorSwatch.Create;
      FSwatches[Col,Row].Col := Col;
      FSwatches[Col,Row].Row := Row;
      FSwatches[Col,Row].Color.ARGB := RevRGB(clGray);
      FSwatches[Col,Row].X := FX1 + FHorizMarg + Col * (FSwatchSize + FSwatchXSpacing);
      FSwatches[Col,Row].Y := FY1 + FVertMarg + Row * (FSwatchSize + FSwatchYSpacing);
    end;
  end;
  FX2 := FSwatches[FCols - 1,0].X + FSwatchSize + FHorizMarg;
  FY2 := FSwatches[0,FRows - 1].Y + FSwatchSize + FVertMarg;
end;

procedure TColorSwatches.Clear;
var
  Col,Row: integer;
begin
  for Col := 0 to High(FSwatches) do begin
    for Row := 0 to High(FSwatches[Col]) do
      FSwatches[Col,Row].Free;
  end;
end;

procedure TColorSwatches.ClearMouse(Canvas: TCanvas);
var
  Tmp: TColorSwatch;
begin
  FTempSwatch := Nil;
  if FCurrSwatch <> Nil then begin
    Tmp := FCurrSwatch;
    FCurrSwatch := Nil;
    PaintSwatch(Canvas,Tmp);
  end;
end;

procedure TColorSwatches.ClearSelected(Canvas: TCanvas);
var
  Tmp: TColorSwatch;
begin
  FTempSwatch := Nil;
  if FSelectedSwatch <> Nil then begin
    Tmp := FSelectedSwatch;
    FSelectedSwatch := Nil;
    PaintSwatch(Canvas,Tmp);
  end;
end;

function TColorSwatches.ClientHit(X, Y: integer): boolean;
begin
  Result := (X >= FX1) and (X <= FX2) and (Y >= FY1) and (Y <= FY2);
end;

constructor TColorSwatches.Create(Cols, Rows: integer);
begin
  FCols := Cols;
  FRows := Rows;
  FSwatchSize := 14;
  FSwatchXSpacing := 4;
  FSwatchYSpacing := 4;
  FHorizMarg := Max(FSwatchXSpacing div 2,FSwatchYSpacing div 2);
  FVertMarg := FHorizMarg;
  InitSwatches;
end;

destructor TColorSwatches.Destroy;
begin
  Clear;
  inherited;
end;

function TColorSwatches.GetSwatches(Col, Row: integer): TColorSwatch;
begin
  Result := FSwatches[Col,Row];
end;

function TColorSwatches.Hit(X, Y: integer): boolean;
var
  Col,Row: integer;
  WX,WY: integer;
begin
  SetLength(FSwatches,FCols * FRows);
  if SwatchXSpacing > 0 then
    WX := SwatchXSpacing div 2
  else
    WX := 0;
  if SwatchYSpacing > 0 then
    WY := SwatchYSpacing div 2
  else
    WY := 0;
  for Row := 0 to FRows - 1 do begin
    for Col := 0 to FCols - 1 do begin
      FTempSwatch := FSwatches[Col,Row];
      Result := (X >= (FTempSwatch.X - WX)) and (X <= (FTempSwatch.X + SwatchSize + WX)) and (Y >= (FTempSwatch.Y - WY)) and (Y <= (FTempSwatch.Y + SwatchSize + WY));
      if Result then
        Exit;
    end;
  end;
  FTempSwatch := Nil;
  Result := False;
end;

procedure TColorSwatches.MouseMove(Canvas: TCanvas; X, Y: integer);
begin
  if Hit(X,Y) then begin
    if FCurrSwatch <> FTempSwatch then begin
      if FCurrSwatch <> Nil then
        PaintSwatch(Canvas,FCurrSwatch);
      FCurrSwatch := FTempSwatch;
      PaintMouseSwatch(Canvas,FCurrSwatch);
    end;
  end;
end;

procedure TColorSwatches.MouseUp(Canvas: TCanvas; X, Y: integer);
var
  Tmp: TColorSwatch;
begin
  if Hit(X,Y) and (FTempSwatch <> FSelectedSwatch) then begin
    if FSelectedSwatch <> Nil then begin
      Tmp := FSelectedSwatch;
      FSelectedSwatch := Nil;
      PaintSwatch(Canvas,Tmp);
    end;
    FSelectedSwatch := FTempSwatch;
    PaintSwatch(Canvas,FSelectedSwatch);
  end;
end;

procedure TColorSwatches.Paint(Canvas: TCanvas);
var
  Col,Row: integer;
begin

  Canvas.Brush.Color := clWhite;
  Canvas.Pen.Color := clWhite;
  Canvas.Rectangle(FX1,FY1,FX2,FY2);
  Canvas.Brush.Style := bsSolid;

  for Row := 0 to FRows - 1 do begin
    for Col := 0 to FCols - 1 do
      PaintSwatch(Canvas,FSwatches[Col,Row]);
  end;
end;

procedure TColorSwatches.PaintMouseSwatch(Canvas: TCanvas; Swatch: TColorSwatch);
begin
  Canvas.Brush.Color := RevRGB(Swatch.Color.ARGB);
  Canvas.Pen.Color := SWATCH_MOUSECOLOR1;
  Canvas.Rectangle(Swatch.X,Swatch.Y,Swatch.X + SwatchSize,Swatch.Y + SwatchSize);
  Canvas.Pen.Color := SWATCH_MOUSECOLOR2;
  Canvas.Rectangle(Swatch.X + 1,Swatch.Y + 1,Swatch.X + SwatchSize - 1,Swatch.Y + SwatchSize - 1);
end;

procedure TColorSwatches.PaintSwatch(Canvas: TCanvas; Swatch: TColorSwatch);
begin
  if Swatch = FSelectedSwatch then
    Canvas.Pen.Color := SWATCH_SELECTEDCOLOR
  else
    Canvas.Pen.Color := SWATCH_BORDERCOLOR;
  Canvas.Brush.Color := RevRGB(Swatch.Color.ARGB);
  Canvas.Rectangle(Swatch.X,Swatch.Y,Swatch.X + SwatchSize,Swatch.Y + SwatchSize);
  if Swatch = FSelectedSwatch then begin
    Canvas.Pen.Color := SWATCH_MOUSECOLOR2;
    Canvas.Brush.Style := bsClear;
    Canvas.Rectangle(Swatch.X + 1,Swatch.Y + 1,Swatch.X + SwatchSize - 1,Swatch.Y + SwatchSize - 1);
    Canvas.Brush.Style := bsSolid;
  end
  else if SwatchYSpacing = 0 then begin
    Canvas.Pen.Color := Swatch.Color.ARGB;
    if Swatch.Row > 0 then begin
      Canvas.MoveTo(Swatch.X + 1,Swatch.Y);
      Canvas.LineTo(Swatch.X + SwatchSize - 1,Swatch.Y);
    end;
    if Swatch.Row < (FRows - 1) then begin
      Canvas.MoveTo(Swatch.X + 1,Swatch.Y + SwatchSize - 1);
      Canvas.LineTo(Swatch.X + SwatchSize - 1,Swatch.Y + SwatchSize - 1);
    end;
  end;
end;

procedure TColorSwatches.SetCols(const Value: integer);
begin
  FCols := Value;
  InitSwatches;
end;

procedure TColorSwatches.SetY1(const Value: integer);
begin
  FY1 := Value;
  InitSwatches;
end;

procedure TColorSwatches.SetRows(const Value: integer);
begin
  FRows := Value;
  InitSwatches;
end;

function TColorSwatches.SetSelected(Canvas: TCanvas; Color: TXc12Color): boolean;
var
  Col,Row: integer;
begin
  Result := False;
  for Row := 0 to FRows - 1 do begin
    for Col := 0 to FCols - 1 do begin

      if Xc12ColorEqual(FSwatches[Col,Row].Color,Color) then begin
        if FSwatches[Col,Row] <> FSelectedSwatch then begin
          ClearSelected(Canvas);
          FSelectedSwatch := FSwatches[Col,Row];
          PaintSwatch(Canvas,FSwatches[Col,Row]);
        end;
        Result := True;
        Exit;
      end;

    end;
  end;
end;

procedure TColorSwatches.SetSwatchSize(const Value: integer);
begin
  FSwatchSize := Value;
  InitSwatches;
end;

procedure TColorSwatches.SetSwatchXSpacing(const Value: integer);
begin
  FSwatchXSpacing := Value;
  InitSwatches;
end;

procedure TColorSwatches.SetSwatchYSpacing(const Value: integer);
begin
  FSwatchYSpacing := Value;
  InitSwatches;
end;

procedure TColorSwatches.SetX1(const Value: integer);
begin
  FX1 := Value;
  InitSwatches;
end;

{ TEx2007ThemeSwatches }

constructor TEx2007ThemeSwatches.Create(Compact: boolean);
var
  Col: integer;
begin
  FAdjust[0,0] := -0.05;
  FAdjust[0,1] := -0.15;
  FAdjust[0,2] := -0.25;
  FAdjust[0,3] := -0.35;
  FAdjust[0,4] := -0.50;

  FAdjust[1,0] :=  0.50;
  FAdjust[1,1] :=  0.35;
  FAdjust[1,2] :=  0.25;
  FAdjust[1,3] :=  0.15;
  FAdjust[1,4] :=  0.05;

  FAdjust[2,0] := -0.10;
  FAdjust[2,1] := -0.25;
  FAdjust[2,2] := -0.50;
  FAdjust[2,3] := -0.75;
  FAdjust[2,4] := -0.90;

  for Col := 3 to 9 do begin
    FAdjust[Col,0] :=  0.80;
    FAdjust[Col,1] :=  0.60;
    FAdjust[Col,2] :=  0.40;
    FAdjust[Col,3] := -0.25;
    FAdjust[Col,4] := -0.50;
  end;

  SetLength(FSwatchList,2);

  FSwatchList[0] := TColorSwatches.Create(10,1);
  FSwatchList[1] := TColorSwatches.Create(10,5);
  FSwatchList[1].FY1 := 25;
  if Compact then
    FSwatchList[1].FSwatchYSpacing := 0;

  inherited Create;

  Initiate;
end;

procedure TEx2007ThemeSwatches.Initiate;
const
  ColOrder: array[0..9] of integer = (0,1,3,2,4,5,6,7,8,9);
var
  Col,Row: integer;
begin
  for Col := 0 to 9 do begin
    FSwatchList[0][Col,0].Color.ColorType := exctTheme;
    FSwatchList[0][Col,0].Color.Theme := TXc12ClrSchemeColor(ColOrder[Col]);
    FSwatchList[0][Col,0].Color.Tint := 0;
    FSwatchList[0][Col,0].Color.ARGB := Xc12DefColorSchemeRGB[TXc12ClrSchemeColor(ColOrder[Col])];
  end;

  for Col := 0 to FSwatchList[1].Cols - 1 do begin
    for Row := 0 to FSwatchList[1].Rows - 1 do begin
      FSwatchList[1][Col,Row].Color.ColorType := exctTheme;
      FSwatchList[1][Col,Row].Color.Theme := TXc12ClrSchemeColor(ColOrder[Col]);
      FSwatchList[1][Col,Row].Color.Tint := FAdjust[Col,Row];

      FSwatchList[1][Col,Row].Color.ARGB := Xc12ColorToRGB(FSwatchList[1][Col,Row].Color);
    end;
  end;

  inherited Initiate;
end;

{ TEx97Swatches }

constructor TEx97Swatches.Create;
var
  i: integer;
  PaletteInitiated: boolean;
begin
  for i := 0 to High(Xc12IndexColorPaletteRGB) do begin
    PaletteInitiated := Xc12IndexColorPaletteRGB[i] <> 0;
    if PaletteInitiated then
      Break;
  end;
  if not PaletteInitiated then
    Move(TXc12DefaultIndexColorPalette[0],Xc12IndexColorPaletteRGB[0],SizeOf(integer) * Length(Xc12IndexColorPaletteRGB));

  SetLength(FSwatchList,1);

  FSwatchList[0] := TColorSwatches.Create(8,7);

  inherited Create;

  Initiate;
end;

procedure TEx97Swatches.Initiate;
var
  Col,Row: integer;
begin
  for Row := 0 to FSwatchList[0].Rows - 1 do begin
    for Col := 0 to FSwatchList[0].Cols - 1 do begin
      FSwatchList[0][Col,Row].Color.ColorType := exctIndexed;
      FSwatchList[0][Col,Row].Color.Indexed := TXc12IndexColor(Col + ((Row + 1) * 8));
      FSwatchList[0][Col,Row].Color.ARGB := Xc12IndexColorPaletteRGB[Col + ((Row + 1) * 8)]
    end;
  end;

  inherited Initiate;
end;

{ TSwatchContainer }

procedure TSwatchContainer.ClearMouse(Canvas: TCanvas);
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do
    FSwatchList[i].ClearMouse(Canvas);
end;

procedure TSwatchContainer.ClearSelected(Canvas: TCanvas);
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do
    FSwatchList[i].ClearSelected(Canvas);
  FSelectedColor.ColorType := exctUnassigned;
end;

function TSwatchContainer.ClientHit(X, Y: integer): boolean;
begin
  Result := (X >= FX1) and (X <= FX2) and (Y >= FY1) and (Y <= FY2);
end;

constructor TSwatchContainer.Create;
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do
    FSwatchList[i].InitSwatches;
end;

destructor TSwatchContainer.Destroy;
var
  i: integer;
begin
   for i := 0 to High(FSwatchList) do
     FSwatchList[i].Free;

  inherited;
end;

procedure TSwatchContainer.Initiate;
var
  i: integer;
begin
  FX1 := MAXINT;
  FY1 := MAXINT;
  FX2 := 0;
  FY2 := 0;
  for i := 0 to High(FSwatchList) do begin
    FX1 := Min(FX1,FSwatchList[i].X1);
    FY1 := Min(FY1,FSwatchList[i].Y1);
    FX2 := Max(FX2,FSwatchList[i].X2);
    FY2 := Max(FY2,FSwatchList[i].Y2);
  end;
end;

procedure TSwatchContainer.MouseMove(Canvas: TCanvas; X, Y: integer);
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do begin
    if FSwatchList[i].ClientHit(X,Y) then
      FSwatchList[i].MouseMove(Canvas,X,Y)
    else
      FSwatchList[i].ClearMouse(Canvas);
  end;
end;

procedure TSwatchContainer.MouseUp(Canvas: TCanvas; X, Y: integer);
var
  i,j: integer;
begin
  for i := 0 to High(FSwatchList) do begin
    if FSwatchList[i].ClientHit(X,Y) then begin
      FSwatchList[i].MouseUp(Canvas,X,Y);
      for j := 0 to High(FSwatchList) do begin
        if j <> i then
          FSwatchList[j].ClearSelected(Canvas);
      end;
      if FSwatchList[i].FSelectedSwatch <> Nil then
        FSelectedColor := FSwatchList[i].FSelectedSwatch.Color;
      Break;
    end;
  end;
end;

procedure TSwatchContainer.Paint(Canvas: TCanvas);
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do
    FSwatchList[i].Paint(Canvas);
end;

function TSwatchContainer.SetSelected(Canvas: TCanvas; Color: TXc12Color): boolean;
var
  i,j: integer;
begin
  for i := 0 to High(FSwatchList) do begin
    if FSwatchList[i].SetSelected(Canvas,Color) then begin
      for j := 0 to High(FSwatchList) do begin
        if j <> i then
          FSwatchList[j].ClearSelected(Canvas);
      end;
      FSelectedColor := Color;
      Result := True;
      Exit;
    end;
  end;
  Result := False;
end;

procedure TSwatchContainer.SetX1(const Value: integer);
begin
  FX1 := Value;
end;

procedure TSwatchContainer.SetY1(const Value: integer);
begin
  FY1 := Value;
end;

function TSwatchContainer.SwatchHit: boolean;
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do begin
    Result := FSwatchList[i].FTempSwatch <> Nil;
    if Result then
      Exit;
  end;
  Result := False;
end;

procedure TSwatchContainer._ClearSelected;
var
  i: integer;
begin
  for i := 0 to High(FSwatchList) do
    FSwatchList[i].FSelectedSwatch := Nil;
  FSelectedColor.ColorType := exctUnassigned;
end;

{ TEx2007StandardSwatches }

constructor TEx2007StandardSwatches.Create;
begin
  SetLength(FSwatchList,1);

  FSwatchList[0] := TColorSwatches.Create(10,1);

  inherited Create;

  Initiate;
end;

procedure TEx2007StandardSwatches.Initiate;
var
  Col: integer;
begin
  for Col := 0 to FSwatchList[0].Cols - 1 do begin
    FSwatchList[0][Col,0].Color.ColorType := exctRgb;
    FSwatchList[0][Col,0].Color.OrigRGB := Ex12StandardColors[Col];
    FSwatchList[0][Col,0].Color.ARGB := Ex12StandardColors[Col];
  end;

  inherited Initiate;
end;

end.
