unit XBookDrawChart2;

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

uses Classes, SysUtils,
     Xc12Utils5,
     XBook_System_2, XBookWindows2, XBookSkin2;

const XSSChartColorBackg = XSS_SYS_COLOR_WHITE;
const XSSChartColorLine  = XSS_SYS_COLOR_BLACK;

type TXSSDrwChartData = class(TObject)
protected
     FSkin: TXLSBookSkin;
public
     constructor Create(ASkin: TXLSBookSkin);

     function CvtX(const AX: double): integer; virtual; abstract;
     function CvtY(const AY: double): integer; virtual; abstract;
     function CvtW(const AW: double): integer; virtual; abstract;
     function CvtH(const AH: double): integer; virtual; abstract;
     end;

type TXSSDrawObject = class(TObject)
protected
     FData: TXSSDrwChartData;
public
     constructor Create(AData: TXSSDrwChartData);
     end;

type TXSSDrwChartLayout = class(TXSSDrawObject)
private
     function GetColor: PXc12Color;
protected
     FX: double;
     FY: double;
     FW: double;
     FH: double;
     FColor: TXc12Color;
public
     constructor Create(AData: TXSSDrwChartData);

     procedure Paint;

     property X: double read FX write FX;
     property Y: double read FY write FY;
     property W: double read FW write FW;
     property H: double read FH write FH;
     property Color: PXc12Color read GetColor;
     end;

type TXSSDrwChartPlotArea = class(TXSSDrawObject)
protected
     FLayout: TXSSDrwChartLayout;
public
     constructor Create(AData: TXSSDrwChartData);
     destructor Destroy; override;

     procedure Paint;
     end;

type TXSSDrwChart = class(TXSSDrwChartData)
protected
     FX1: integer;
     FY1: integer;
     FX2: integer;
     FY2: integer;
     FPlotArea: TXSSDrwChartPlotArea;
public
     constructor Create(const AX1,AY1,AX2,AY2: integer);
     destructor Destroy; override;

     procedure Paint;

     function CvtX(const AX: double): integer; override;
     function CvtY(const AY: double): integer; override;
     function CvtW(const AW: double): integer; override;
     function CvtH(const AH: double): integer; override;
     end;

type TXSSDrwChartWin = class(TXSSClientWindow)
private
     function GetColor: PXc12Color;
     function GetLineColor: PXc12Color;
protected
     FColor: TXc12Color;
     FLineColor: TXc12Color;

     FChart: TXSSDrwChart;
public
     constructor Create(AParent: TXSSWindow);
     destructor Destroy; override;

     procedure Paint; override;

     property Color: PXc12Color read GetColor;
     property LineColor: PXc12Color read GetLineColor;
     end;

implementation

{ TXSSDrwChartLayout }

constructor TXSSDrwChartLayout.Create(AData: TXSSDrwChartData);
begin
  inherited Create(AData);

  FColor := RGBColorToXc12(XSSChartColorBackg);
end;

function TXSSDrwChartLayout.GetColor: PXc12Color;
begin
  Result := @FColor;
end;

procedure TXSSDrwChartLayout.Paint;
begin

end;

{ TXSSDrwChartPlotArea }

constructor TXSSDrwChartPlotArea.Create(AData: TXSSDrwChartData);
begin
  inherited Create(AData);

  FLayout := TXSSDrwChartLayout.Create(FData);
end;

destructor TXSSDrwChartPlotArea.Destroy;
begin
  FLayout.Free;
  inherited;
end;

procedure TXSSDrwChartPlotArea.Paint;
begin
  FLayout.Paint;
end;

{ TXSSDrwChart }

constructor TXSSDrwChartWin.Create(AParent: TXSSWindow);
begin
  inherited Create(AParent);
  FColor := RGBColorToXc12(XSSChartColorBackg);
  FLineColor := RGBColorToXc12(XSSChartColorLine);

  FChart := TXSSDrwChart.Create(FX1,FY1,FX2,FY2);
end;

destructor TXSSDrwChartWin.Destroy;
begin
  FChart.Free;
  inherited;
end;

function TXSSDrwChartWin.GetColor: PXc12Color;
begin
  Result := @FColor;
end;

function TXSSDrwChartWin.GetLineColor: PXc12Color;
begin
  Result := @FLineColor;
end;

procedure TXSSDrwChartWin.Paint;
begin
  FSkin.GDI.PenColor := Xc12ColorToRGB(FLineColor);
  FSkin.GDI.BrushColor := Xc12ColorToRGB(FColor);

  FSkin.GDI.Rectangle(X1,Y1,X2 - 1,Y2 - 1);

  FChart.Paint;
end;

{ TXSSDrwChart }

constructor TXSSDrwChart.Create(const AX1,AY1,AX2,AY2: integer);
begin
  FX1 := AX1;
  FY1 := AY1;
  FX2 := AX2;
  FY2 := AY2;
  FPlotArea := TXSSDrwChartPlotArea.Create(Self);
end;

function TXSSDrwChart.CvtH(const AH: double): integer;
begin
  Result := 0;
end;

function TXSSDrwChart.CvtW(const AW: double): integer;
begin
  Result := 0;
end;

function TXSSDrwChart.CvtX(const AX: double): integer;
begin
  Result := 0;
end;

function TXSSDrwChart.CvtY(const AY: double): integer;
begin
  Result := 0;
end;

destructor TXSSDrwChart.Destroy;
begin
  FPlotArea.Free;
  inherited;
end;

procedure TXSSDrwChart.Paint;
begin
  FPlotArea.Paint;
end;

{ TXSSDrwChartData }

constructor TXSSDrwChartData.Create(ASkin: TXLSBookSkin);
begin
  FSkin := ASkin;
end;

{ TXSSCharObject }

constructor TXSSDrawObject.Create(AData: TXSSDrwChartData);
begin
  FData := AData;
end;

end.
