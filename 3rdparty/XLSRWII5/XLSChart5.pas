unit XLSChart5;

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

uses Classes, SysUtils, Math, Contnrs,
     xpgParseDrawing,
     Xc12Utils5,
     XLSUtils5;

const XLSCHART_DEF_COLOR_FILL = $FFFFFF;
const XLSCHART_DEF_COLOR_LINE = $000000;

type TXLSChartType = (xctArea,xctArea3d,xctPie,xctPie3d,xctDoughnut,xctOfPie,xctSurface,xctSurface3d,xctLine,xctLine3d,xctStock,xctRadar,xctScatter,xctBar,xctBar3d,xctBubble);

type TXLSChartLayout = class(TObject)
private
     function  GetX2: double;
     function  GetY2: double;
     procedure SetX2(const Value: double);
     procedure SetY2(const Value: double);
protected
     FX: double;
     FY: double;
     FW: double;
     FH: double;
     FColor: longword;
     FAssigned: boolean;
public
     property X: double read FX write FX;
     property Y: double read FY write FY;
     property W: double read FW write FW;
     property H: double read FH write FH;

     property X1: double read FX write FX;
     property Y1: double read FY write FY;
     property X2: double read GetX2 write SetX2;
     property Y2: double read GetY2 write SetY2;

     property Color: longword read FColor write FColor;

     property Assigned: boolean read FAssigned write FAssigned;
     end;

type TXLSChartSerieValue = class(TObject)
protected
     FValue: double;
     FColor: longword;
public
     property Value: double read FValue write FValue;
     property Color: longword read FColor write FColor;
     end;

type TXLSChartCatAxisItem = class(TObject)
protected
     FValues: TStringList;
public
     constructor Create;
     destructor Destroy; override;

     property Values: TStringList read FValues;
     end;

type TXLSChartCatAxis = class(TObjectList)
private
     function  GetItems(Index: integer): TXLSChartCatAxisItem;
     function  GetPSourceArea: PXLS3dCellArea;
protected
     FValid: boolean;
     FSourceArea: TXLS3dCellArea;
public
     constructor Create;

     function Add: TXLSChartCatAxisItem;

     property Valid: boolean read FValid write FValid;
     property SourceArea: TXLS3dCellArea read FSourceArea write FSourceArea;
     property PSourceArea: PXLS3dCellArea read GetPSourceArea;
     property Items[Index: integer]: TXLSChartCatAxisItem read GetItems; default;
     end;

type TXLSChartLegendItem = class(TObject)
protected
     FText: AxUCString;
     FColor: longword;
public
     property Text: AxUCString read FText write FText;
     property Color: longword read FColor write FColor;
     end;

type TXLSChartLegend = class(TObjectList)
private
     function GetItems(Index: integer): TXLSChartLegendItem;
protected
     FLayout: TXLSChartLayout;
public
     constructor Create;
     destructor Destroy; override;

     function  Add(const AText: AxUCString): TXLSChartLegendItem;

     property Layout: TXLSChartLayout read FLayout;
     property Items[Index: integer]: TXLSChartLegendItem read GetItems; default;
     end;

type TXLSChartTitle = class(TObject)
protected
     FLayout  : TXLSChartLayout;
     FText    : AxUCString;
     FFontSize: double;
public
     constructor Create;
     destructor Destroy; override;

     property Layout: TXLSChartLayout read FLayout;
     property Text: AxUCString read FText write FText;
     property FontSize: double read FFontSize write FFontSize;
     end;

type TXLSChart = class;

     TXLSChartSerie = class(TObjectList)
private
     function  GetItems(Index: integer): TXLSChartSerieValue;
protected
     FParent : TXLSChart;
     FColor  : longword;
     FLable  : AxUCString;
     FMinVal : double;
     FMaxVal : double;
public
     constructor Create(AParent: TXLSChart);
     destructor Destroy; override;

     function  Add(const AValue: double): TXLSChartSerieValue;

     property Color: longword read FColor write FColor;
     property Lable: AxUCString read FLable write FLable;
     property MinVal: double read FMinVal write FMinVal;
     property MaxVal: double read FMaxVal write FMaxVal;
     property Items[Index: integer]: TXLSChartSerieValue read GetItems; default;
     end;

     TXLSChartCamera3D = class(TObject)
protected
     // Radians
     FRotateX: double;
     FRotateY: double;
     FLens   : double;
public
     constructor Create;

     property RotateX: double read FRotateX write FRotateX;
     property RotateY: double read FRotateY write FRotateY;
     property Lens: double read FLens write FLens;
     end;

     TXLSChart = class(TObjectList)
private
     function  GetItems(Index: integer): TXLSChartSerie;
protected
     FAnchor    : TCT_TwoCellAnchor;

     FLayout    : TXLSChartLayout;
     FColor     : longword;

     FCamera    : TXLSChartCamera3D;

     FTitle     : TXLSChartTitle;
     FLegend    : TXLSChartLegend;

     FMinVal    : double;
     FMaxVal    : double;
     FChartType : TXLSChartType;
     FHasValAxis: boolean;
     FHasCatAxis: boolean;
     FCatAxis   : TXLSChartCatAxis;
     FGapWidth  : double;
public
     constructor Create;
     destructor Destroy; override;

     function Add: TXLSChartSerie;

     function  SerieMaxCount: integer;

     property Anchor: TCT_TwoCellAnchor read FAnchor write FAnchor;
     property MinVal: double read FMinVal write FMinVal;
     property MaxVal: double read FMaxVal write FMaxVal;
     property ChartType: TXLSChartType read FChartType write FChartType;
     property HasValAxis: boolean read FHasValAxis write FHasValAxis;
     property HasCatAxis: boolean read FHasCatAxis write FHasCatAxis;
     property Camera: TXLSChartCamera3D read FCamera;
     property Title: TXLSChartTitle read FTitle;
     property Legend: TXLSChartLegend read FLegend;
     property CatAxis: TXLSChartCatAxis read FCatAxis;
     property Layout: TXLSChartLayout read FLayout;
     property GapWidth: double read FGapWidth write FGapWidth;
     property Color: longword read FColor write FColor;
     property Items[Index: integer]: TXLSChartSerie read GetItems; default;
     end;

type TXLSCharts = class(TObjectList)
private
     function  GetItems(Index: integer): TXLSChart;
protected
public
     constructor Create;
     destructor Destroy; override;

     function  Add: TXLSChart;

     property Items[Index: integer]: TXLSChart read GetItems; default;
     end;

implementation

{ TXLSChartSerie }

function TXLSChartSerie.Add(const AValue: double): TXLSChartSerieValue;
begin
  Result := TXLSChartSerieValue.Create;
  Result.Value := AValue;
  Result.Color := FColor;
  FMinVal := Min(FMinVal,AValue);
  FMaxVal := Max(FMaxVal,AValue);
  FParent.MinVal := Min(FParent.MinVal,AValue);
  FParent.MaxVal := Max(FParent.MaxVal,AValue);
  inherited Add(Result);
end;

constructor TXLSChartSerie.Create(AParent: TXLSChart);
begin
  inherited Create;
  FParent := AParent;
  FMinVal := MAXDOUBLE;
  FMaxVal := MINDOUBLE;
end;

destructor TXLSChartSerie.Destroy;
begin
  inherited;
end;

function TXLSChartSerie.GetItems(Index: integer): TXLSChartSerieValue;
begin
  Result := TXLSChartSerieValue(inherited Items[Index]);
end;

{ TXLSChart }

function TXLSChart.Add: TXLSChartSerie;
begin
  Result := TXLSChartSerie.Create(Self);
  inherited Add(Result);
end;

constructor TXLSChart.Create;
begin
  inherited Create;
  FMinVal := MAXDOUBLE;
  FMaxVal := MINDOUBLE;
  FHasValAxis := True;
  FHasCatAxis := True;
  FCatAxis := TXLSChartCatAxis.Create;
  FGapWidth := 1.50;

  FCamera := TXLSChartCamera3D.Create;

  FLayout := TXLSChartLayout.Create;
  FTitle := TXLSChartTitle.Create;
  FLegend := TXLSChartLegend.Create;
end;

destructor TXLSChart.Destroy;
begin
  FLayout.Free;
  FCatAxis.Free;
  FLegend.Free;
  FTitle.Free;
  FCamera.Free;
  inherited;
end;

function TXLSChart.GetItems(Index: integer): TXLSChartSerie;
begin
  Result := TXLSChartSerie(inherited Items[Index]);
end;

function TXLSChart.SerieMaxCount: integer;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    Result := Max(Result,Items[i].Count);
end;

{ TXLSCharts }

function TXLSCharts.Add: TXLSChart;
begin
  Result := TXLSChart.Create;
  inherited Add(Result);
end;

constructor TXLSCharts.Create;
begin
  inherited Create;
end;

destructor TXLSCharts.Destroy;
begin

  inherited;
end;

function TXLSCharts.GetItems(Index: integer): TXLSChart;
begin
  Result := TXLSChart(inherited Items[Index]);
end;

{ TXLSChartLayout }

function TXLSChartLayout.GetX2: double;
begin
  Result := FX + FW;
end;

function TXLSChartLayout.GetY2: double;
begin
  Result := FY + FH;
end;

procedure TXLSChartLayout.SetX2(const Value: double);
begin
  FW := Value - FX;
end;

procedure TXLSChartLayout.SetY2(const Value: double);
begin
  FH := Value - FY;
end;

{ TXLSChartCatAxisItem }

constructor TXLSChartCatAxisItem.Create;
begin
  FValues := TStringList.Create;
end;

destructor TXLSChartCatAxisItem.Destroy;
begin
  FValues.Free;
  inherited;
end;

{ TXLSChartCatAxis }

function TXLSChartCatAxis.Add: TXLSChartCatAxisItem;
begin
  Result := TXLSChartCatAxisItem.Create;
  inherited Add(Result);
end;

constructor TXLSChartCatAxis.Create;
begin
  inherited Create;

  FValid := True;
end;

function TXLSChartCatAxis.GetItems(Index: integer): TXLSChartCatAxisItem;
begin
  Result := TXLSChartCatAxisItem(inherited Items[Index]);
end;

function TXLSChartCatAxis.GetPSourceArea: PXLS3dCellArea;
begin
  Result := @FSourceArea;
end;

{ TXLSChartLegend }

function TXLSChartLegend.Add(const AText: AxUCString): TXLSChartLegendItem;
begin
  Result := TXLSChartLegendItem.Create;
  Result.Text := AText;
  inherited Add(Result);
end;

constructor TXLSChartLegend.Create;
begin
  inherited Create;

  FLayout := TXLSChartLayout.Create;
end;

destructor TXLSChartLegend.Destroy;
begin
  FLayout.Free;
  inherited;
end;

function TXLSChartLegend.GetItems(Index: integer): TXLSChartLegendItem;
begin
  Result := TXLSChartLegendItem(inherited Items[Index]);
end;

{ TXLSChartTitle }

constructor TXLSChartTitle.Create;
begin
  FLayout := TXLSChartLayout.Create;
end;

destructor TXLSChartTitle.Destroy;
begin
  FLayout.Free;
  inherited;
end;

{ TXLSChartCamera3D }

constructor TXLSChartCamera3D.Create;
begin
  FLens := 50;
end;

end.
