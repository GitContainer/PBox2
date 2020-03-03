unit XBookGeometryChart2;

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

uses Classes, SysUtils, Contnrs, Math,
     xpgParseDrawingCommon, xpgParseChart,
     Xc12Utils5, Xc12DataStyleSheet5,
     XLSUtils5, XLSDrawing5, XLSRelCells5,
     XBookPaintGDI2, XBookGeometry2;

const XSS_CHART_MAXDATAPOINTS      = $100000;
const XSS_CHART_DEF_COLOR          = $FFFFFF;
const XSS_CHART_DEF_LINECOLOR      = $000000;
const XSS_CHART_DEF_LINEWIDTH      = 0.75;
const XSS_CHART_DEF_LINECHARTWIDTH = 2.25;
const XSS_CHART_DEF_MARKER_SIZE    = 2.00;

const XSS_CHART_DEF_FILLCOLORS: array[0..5] of longword = (
      $003C6494,$00963D3B,$00799244,$00634D7E,$0039869B,$00C27535);

type TXSSChartData = class;

     TXSSChartCharRun = class(TObject)
protected
     FText  : AxUCString;
     FFont  : TXc12Font;
     FWidth : double;
     FHeight: double;
public
     destructor Destroy; override;

     property Text  : AxUCString read FText write FText;
     property Font  : TXc12Font read FFont write FFont;
     property Width : double read FWidth write FWidth;
     property Height: double read FHeight write FHeight;
     end;

     TXSSChartCharRuns = class(TObjectList)
private
     function GetItems(Index: integer): TXSSChartCharRun;
protected
public
     constructor Create;

     function  Add: TXSSChartCharRun;

     property Items[Index: integer]: TXSSChartCharRun read GetItems; default;
     end;

     TXSSChartParagraph = class(TObject)
private
     function  GetPlainText: AxUCString;
     procedure SetPlainText(const Value: AxUCString);
protected
     FRuns   : TXSSChartCharRuns;
     FData   : TXSSChartData;
     FDefFont: TXc12Font;

     procedure DoCharRun(ACharRun: TCT_RegularTextRun);
     procedure DoPara(APara: TCT_TextParagraph);
     procedure DoParas(AParas: TCT_TextParagraphXpgList);
public
     constructor Create(AData: TXSSChartData; AText: AxUCString; AFont: TXc12Font); overload;
     constructor Create(AData: TXSSChartData; AParas: TCT_TextParagraphXpgList; ADefFont: TXc12Font); overload;
     destructor Destroy; override;

     procedure UpdateFont(AFont: TXc12Font);

     property Runs     : TXSSChartCharRuns read FRuns;
     property PlainText: AxUCString read GetPlainText write SetPlainText;
     end;

     TXSSChartItemDataPoint = class(TObject)
protected
     FIndex: integer;
     FSpPr : TXLSDrwShapeProperies;
     FParas: TXSSChartParagraph;
     FData : TXSSChartData;
public
     constructor Create(AData: TXSSChartData); overload;
     constructor Create(AData: TXSSChartData; ADp: TCT_DPt); overload;
     destructor Destroy; override;

     procedure AddParas(AText: AxUCstring; AFont: TXc12Font); overload;
     procedure AddParas(AParas: TCT_TextParagraphXpgList); overload;

     function  FillRGB(ADefault: longword): longword;

     property Index: integer read FIndex write FIndex;
     property SpPr : TXLSDrwShapeProperies read FSpPr;
     property Paras: TXSSChartParagraph read FParas;
     end;

     TXSSChartItemDataPoints = class(TObjectList)
protected
     FData: TXSSChartData;

     function  GetItems(Index: integer): TXSSChartItemDataPoint;
     function AddDp(AIndex: integer): TXSSChartItemDataPoint;
public
     constructor Create(AData: TXSSChartData);

     function  Add: TXSSChartItemDataPoint; overload;
     function  Add(ADp: TCT_DPt): TXSSChartItemDataPoint; overload;
     procedure Add(ADps: TCT_DPtXpgList); overload;
     procedure Add(AdLbls: TCT_DLbls; AVals: TDynSingleArray); overload;

     function  Find(AIndex: integer): TXSSChartItemDataPoint;

     property Items[Index: integer]: TXSSChartItemDataPoint read GetItems;
     end;

     TXSSChartData = class(TObject)
private
     procedure SetFont(const Value: TXc12Font);
protected
     FPtWidth    : double;
     FPtHeight   : double;
     FDefMarg    : double;
     FDefPadding : double;
     FFont       : TXc12Font;
     FTitleFont  : TXc12Font;
     FColors     : array[0..(Length(XSS_CHART_DEF_FILLCOLORS) * 4) - 1] of longword;
     FCurrColor  : integer;
     FDataPoints : TXSSChartItemDataPoints;
     FShowValAxis: boolean;
     FGarbage    : TObjectList;
public
     constructor Create;
     destructor Destroy; override;

     function  MapL(APtSz: double): double;
     function  MapT(APtSz: double): double;
     function  MapR(APtSz: double): double;
     function  MapB(APtSz: double): double;

     function  GetChartFont(ATxt: TCT_TextBody): TXc12Font; overload;
     function  GetChartFont(ARPr: TCT_TextCharacterProperties; ADefFont: TXc12Font): TXc12Font; overload;
     procedure ResetColor;
     function  GetSerieColor: longword;

     procedure AddParas(AGeometry: TXSSGeometry; AX,AY: double; AParas: TXSSChartParagraph; AVertAlign: TXc12VertAlignment = cvaBottom);
     function ApplyDataPoint(AIndex: integer; AGeometry: TXSSGeometry; AX,AY: double; AAboveOrigo: boolean): TXSSChartItemDataPoint;

     property PtWidth    : double read FPtWidth write FPtWidth;
     property PtHeight   : double read FPtHeight write FPtHeight;
     property DefMarg    : double read FDefMarg write FDefMarg;
     property DefPadding : double read FDefPadding write FDefPadding;
     property Font       : TXc12Font read FFont write SetFont;
     property TitleFont  : TXc12Font read FTitleFont;
     property DataPoints : TXSSChartItemDataPoints read FDataPoints;
     property ShowValAxis: boolean read FShowValAxis write FShowValAxis;
     end;

type TXSSChartItem = class(TObject)
protected
     FGDI : TAXWGDI;
     FData: TXSSChartData;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);

     property Data: TXSSChartData read FData;
     end;

type TXSSChartItemTitle = class(TXSSChartItem)
private
     function  GetHeight: double;
     function  GetWidth: double;
protected
     FPara   : TXSSChartParagraph;
     FWidth  : double;
     FHeight : double;
     FDefFont: TXc12Font;

     function GetMargs: double;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);
     destructor Destroy; override;

     procedure PreBuild(ATitle: TCT_Title);

     procedure Build(ARect: TXSSGeometryRect; ATitle: TCT_Title);

     property Width : double read GetWidth;
     property Height: double read GetHeight;
     property Margs : double read GetMargs;
     end;

type TXSSChartItemChart = class(TXSSChartItem)
protected
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);
     destructor Destroy; override;
     end;

type TXSSChartItemLegendItem = class(TObject)
protected
     FColor : longword;
     FText  : AxUCString;
     FWidth : double;
     FHeight: double;
     FFont  : TXc12Font;
public
     destructor Destroy; override;

     property Color : longword read FColor write FColor;
     property Text  : AxUCString read FText write FText;
     property Width : double read FWidth write FWidth;
     property Height: double read FHeight write FHeight;
     property Font  : TXc12Font read FFont write FFont;
     end;

type TXSSChartItemLegend = class(TXSSChartItem)
private
     function GetItems(Index: integer): TXSSChartItemLegendItem;
     function GetPtWidth: double;
     function GetPtHeight: double;
protected
     FCTLegend : TCT_Legend;
     FItems    : TObjectList;
     FTxtWidth : double;
     FTxtHeight: double;
     FMargs    : double;
     FPadding  : double;
     FSpPr     : TXLSDrwShapeProperies;
     FFont     : TXc12Font;

     function  SerieMarkerSz: double; 
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; ALegendData: TCT_Legend);
     destructor Destroy; override;

     procedure Clear;

     procedure Build(ARect: TXSSGeometryRect; ALegend: TCT_Legend);

     procedure Add(AIndex: integer; AColor: longword; AText: AxUCString);

     function  Count: integer;

     property PtWidth : double read GetPtWidth;
     property PtHeight: double read GetPtHeight;
     property Margs   : double read FMargs write FMargs;
     property Padding : double read FPadding;
     property Items[Index: integer]: TXSSChartItemLegendItem read GetItems; default;
     end;

type TXSSChartItemAxis = class(TXSSChartItem)
protected
     FTextWidth : double;
     FTextHeight: double;
     FVisible   : boolean;
     FNumFormat : AxUCString;
     FFont      : TXc12Font;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);

     property Visible   : boolean read FVisible;
     property TextWidth : double read FTextWidth;
     property TextHeight: double read FTextHeight;
     property NumFormat : AxUCString read FNumFormat;
     property Font      : TXc12Font read FFont;
     end;


type TXSSChartItemValAxis = class(TXSSChartItemAxis)
protected
     FScaleMin  : double;
     FScaleMax  : double;
     FStepSize  : double;
     FAxisTicks : integer;
     FScaleOrigo: double;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);

     procedure CalcScale(AMinVal,AMaxVal: double; ATicks: integer);

     procedure Build(ARect: TXSSGeometryRect; AValAxis: TCT_ValAx);

     function  ScaleRange: double;

     property ScaleMin  : double read FScaleMin;
     property ScaleMax  : double read FScaleMax;
     property StepSize  : double read FStepSize;
     property AxisTicks : integer read FAxisTicks;
     property ScaleOrigo: double read FScaleOrigo;
     end;

type TXSSChartItemValAxises = class(TObjectList)
protected
     function GetItems(Index: integer): TXSSChartItemValAxis;
public
     constructor Create;

     function Add(AGDI: TAXWGDI; AData: TXSSChartData): TXSSChartItemValAxis;

     function First: TXSSChartItemValAxis;

     property Items[Index: integer]: TXSSChartItemValAxis read GetItems; default;
     end;

type TXSSChartItemCatAxis = class(TXSSChartItemAxis)
protected
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);

     procedure Build(ARect: TXSSGeometryRect; ACount: integer; ACats: TDynStringArray; ATxPr: TCT_TextBody; ATextBetweenTicks: boolean);
     end;

type TXSSChartItemChartSeries = class;

     TXSSChartItemChartSerie = class(TXSSChartItem)
private
     function  GetVal(Index: integer): double;
protected
     FOwner   : TXSSChartItemChartSeries;
     FVals    : TDynSingleArray;
     FMinVal  : double;
     FMaxVal  : double;
     FDefColor: longword;
     FColor   : longword;
     FName    : AxUCString;

     procedure GetLineStyle(ASpPr: TCT_ShapeProperties; var ALineColor: longword; var ALineWidth: double);
     procedure SetupDataPoints(ADPt: TCT_DPtXpgList; ADLbls: TCT_DLbls);
public
     constructor Create(AOwner: TXSSChartItemChartSeries; AGDI: TAXWGDI; AData: TXSSChartData);
     destructor Destroy; override;

     function  ValsSum: double;

     procedure PreBuild(AIndex: integer; AVal: TXLSRelCells; ACat: TCT_AxDataSource; AShared: TCT_SerieShared; AUseFill: boolean);
     procedure Build(ARect: TXSSGeometryRect; ABarSer: TCT_BarSer; AIndex: integer); overload;
     procedure Build(ARect: TXSSGeometryRect; ALineSer: TCT_LineSer; AHasMarkers: boolean); overload;
     procedure Build(ARect: TXSSGeometryRect; AAreaSer: TCT_AreaSer); overload;
     procedure Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; AScatterSer: TCT_ScatterSer); overload;
     procedure Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; ABubbleSer: TCT_BubbleSer); overload;
     procedure Build(ARect: TXSSGeometryRect; APieSer: TCT_PieSer; AStartAngle: double; ALegend: TXSSChartItemLegend); overload;
     procedure Build(ARect: TXSSGeometryRect; ARadarSer: TCT_RadarSer; ARadarStyle: TCT_RadarStyle); overload;

     function  Count: integer;
     function  Vals: TDynSingleArray;

     property MinVal: double read FMinVal write FMinVal;
     property MaxVal: double read FMaxVal write FMaxVal;
     property Color : longword read FColor write FColor;
     property Name  : AxUCString read FName;
     property Val[Index: integer]: double read GetVal;
     end;

     TXSSChartItemChartSeries = class(TObjectList)
private
     function  GetItems(Index: integer): TXSSChartItemChartSerie;
protected
     FGDI       : TAXWGDI;
     FData      : TXSSChartData;
     FMinVal    : double;
     FMaxVal    : double;
     FValCount  : integer;
     FValAxis   : TXSSChartItemValAxis;
     FCatAxis   : TXSSChartItemCatAxis;
     FCats      : TDynStringArray;
     FGapWidth  : double;

     procedure CalcMinMax(AGrouping: TST_BarGrouping);
     procedure AddGridlines(ARect: TXSSGeometryRect);
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
     destructor Destroy; override;

     function  Add: TXSSChartItemChartSerie;

     procedure PreBuild(AGrouping: TST_BarGrouping; ABarSeries: TCT_BarSeries); overload;
     procedure PreBuild(ALineSeries: TCT_LineSeries); overload;
     procedure PreBuild(AAreaSeries: TCT_AreaSeries); overload;
     procedure PreBuild(AScatterSeries: TCT_ScatterSeries); overload;
     procedure PreBuild(ABubbleSeries: TCT_BubbleSeries); overload;
     procedure PreBuild(ARadarSeries: TCT_RadarSeries); overload;
     procedure Build(ARect: TXSSGeometryRect; ABarSeries: TCT_BarSeries); overload;
     procedure Build(ARect: TXSSGeometryRect; ALineSeries: TCT_LineSeries; AHasMarkers: boolean); overload;
     procedure Build(ARect: TXSSGeometryRect; AAreaSeries: TCT_AreaSeries); overload;
     procedure Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; AScatterSeries: TCT_ScatterSeries); overload;
     procedure Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; ABubbleSeries: TCT_BubbleSeries); overload;
     procedure Build(ARect: TXSSGeometryRect; ARadarSeries: TCT_RadarSeries; ARadarStyle: TCT_RadarStyle; AValAxFont,ACatAxFont: TCT_TextBody); overload;
     procedure BuildStacked(ARect: TXSSGeometryRect; ABarSeries: TCT_BarSeries; APercent: boolean);

     function  MapY(AHeight,AVal: double): double;

     property MinVal    : double read FMinVal;
     property MaxVal    : double read FMaxVal;
     property ValCount  : integer read FValCount;
     property GapWidth  : double read FGapWidth write FGapWidth;
     property ValAxis   : TXSSChartItemValAxis read FValAxis;
     property CatAxis   : TXSSChartItemCatAxis read FCatAxis;
     property Items[Index: integer]: TXSSChartItemChartSerie read GetItems; default;
     end;

type TXSSChartItemSeriesChart = class(TXSSChartItemChart)
protected
     FSeries : TXSSChartItemChartSeries;
     FValAxis: TXSSChartItemValAxis;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
     destructor Destroy; override;

     procedure Clear;

     property Series : TXSSChartItemChartSeries read FSeries;
     end;

type TXSSChartItemBarChart = class(TXSSChartItemSeriesChart)
protected
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);

     procedure PreBuild(ALegend: TXSSChartItemLegend; ABarChart: TCT_BarChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ABarChart: TCT_BarChart);
     end;

type TXSSChartItemLineChart = class(TXSSChartItemSeriesChart)
protected
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);

     procedure PreBuild(ALegend: TXSSChartItemLegend; ALineChart: TCT_LineChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ALineChart: TCT_LineChart);
     end;

type TXSSChartItemAreaChart = class(TXSSChartItemSeriesChart)
protected
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);

     procedure PreBuild(ALegend: TXSSChartItemLegend; AAreaChart: TCT_AreaChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; AAreaChart: TCT_AreaChart);
     end;

type TXSSChartItemXYChart = class(TXSSChartItemSeriesChart)
protected
     FXAxis  : TXSSChartItemValAxis;
     FXSeries: TXSSChartItemChartSeries;

     procedure DoCatAxis(ARect: TXSSGeometryRect);
     procedure DoLegend(ALegend: TXSSChartItemLegend);
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxises: TXSSChartItemValAxises; ACatAxis: TXSSChartItemCatAxis);
     destructor Destroy; override;

     end;

type TXSSChartItemScatterChart = class(TXSSChartItemXYChart)
protected
public
     procedure PreBuild(ALegend: TXSSChartItemLegend; AScatterChart: TCT_ScatterChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; AScatterChart: TCT_ScatterChart);
     end;

type TXSSChartItemBubbleChart = class(TXSSChartItemXYChart)
protected
public
     procedure PreBuild(ALegend: TXSSChartItemLegend; ABubbleChart: TCT_BubbleChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ABubbleChart: TCT_BubbleChart);
     end;

type TXSSChartItemPieChart = class(TXSSChartItem)
protected
     FSerie: TXSSChartItemChartSerie;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);
     destructor Destroy; override;

     procedure PreBuild(ALegend: TXSSChartItemLegend; APieChart: TCT_PieChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; APieChart: TCT_PieChart);
     end;

type TXSSChartItemRadarChart = class(TXSSChartItemSeriesChart)
protected
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);

     procedure PreBuild(ALegend: TXSSChartItemLegend; ARadarChart: TCT_RadarChart);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ARadarChart: TCT_RadarChart; AValAxisFont,ACatAxFont: TCT_TextBody);
     end;

type TXSSChartItemPlotArea = class(TXSSChartItem)
protected
     FCharts   : TObjectList;
     FValAxises: TXSSChartItemValAxises;
     FCatAxis  : TXSSChartItemCatAxis;
     FMargs    : double;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);
     destructor Destroy; override;

     procedure Clear;

     procedure PreBuild(ALegend: TXSSChartItemLegend; AChartSpace: TCT_PlotArea);
     procedure Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; AChartSpace: TCT_PlotArea);

     property Margs: double read FMargs write FMargs;
     end;

type TXSSChartItemChartSpace = class(TXSSChartItem)
protected
     FPlotArea: TXSSChartItemPlotArea;
     FLegend  : TXSSChartItemLegend;
public
     constructor Create(AGDI: TAXWGDI; AData: TXSSChartData);
     destructor Destroy; override;

     procedure Clear;

     procedure Build(AGeometry: TXSSGeometry; AChartSpace: TCT_ChartSpace);
     end;

implementation

function GetSpPrLineColor(ASpPr: TCT_ShapeProperties; ADefault: longword): longword;
var
  SpPr: TXLSDrwShapeProperies;
begin
  Result := ADefault;

  if ASpPr <> Nil then begin
    SpPr := TXLSDrwShapeProperies.Create(ASpPr);
    try
      if SpPr.Line <> Nil then
        Result := SpPr.Line.RGB_(ADefault);
    finally
      SpPr.Free;
    end;
  end;
end;

function GetSpPrFillColor(ASpPr: TCT_ShapeProperties; ADefault: longword): longword;
var
  SpPr: TXLSDrwShapeProperies;
begin
  Result := ADefault;

  if ASpPr <> Nil then begin
    SpPr := TXLSDrwShapeProperies.Create(ASpPr);
    try
      Result := SpPr.FillRGB(ADefault);
    finally
      SpPr.Free;
    end;
  end;
end;

procedure ApplySpPr(ASpPr: TXLSDrwShapeProperies; AStyle: TXSSGeometryStyle); overload;
begin
  if ASpPr <> Nil then begin
    AStyle.Color := ASpPr.FillRGB(AStyle.Color);
    if ASpPr.Line <> Nil then begin
      AStyle.LineColor := ASpPr.Line.RGB_(AStyle.LineColor);
      AStyle.LineWidth := ASpPr.Line.Width;
    end;
  end;
end;

procedure ApplySpPr(ASpPr: TCT_ShapeProperties; AStyle: TXSSGeometryStyle); overload;
var
  SpPr: TXLSDrwShapeProperies;
begin
  if ASpPr <> Nil then begin
    SpPr := TXLSDrwShapeProperies.Create(ASpPr);
    try
      ApplySpPr(SpPr,AStyle);
    finally
      SpPr.Free;
    end;
  end;
end;

procedure ApplyManualLayout(ARect: TXSSGeometryRect; ALayout: TCT_ManualLayout);
begin
  if ALayout.X <> Nil then
    ARect.Pt1.X := ALayout.X.Val;
  if ALayout.Y <> Nil then
    ARect.Pt1.Y := ALayout.Y.Val;

  if ALayout.W <> Nil then
    ARect.Pt2.X := ARect.Pt1.X + ALayout.W.Val;
  if ALayout.H <> Nil then
    ARect.Pt2.Y := ARect.Pt1.Y + ALayout.H.Val;
end;

procedure ParaSize(AGDI: TAXWGDI; APara: TXSSChartParagraph; out AWidth, AHeight: double);
var
  i: integer;
begin
  AWidth := 0;
  AHeight := 0;

  for i := 0 to APara.Runs.Count - 1 do begin
    AGDI.SetFont(APara.Runs[i].Font);

    APara.Runs[i].Width := AGDI.TextWidthF(APara.Runs[i].Text);
    APara.Runs[i].Height := AGDI.TextHeightF(APara.Runs[i].Text);

    AWidth := AWidth + APara.Runs[i].Width;
    AHeight := Max(AHeight,APara.Runs[i].Height);
  end;

  AWidth := AGDI.PixToPtX(Round(AWidth));
  AHeight := AGDI.PixToPtY(Round(AHeight));
end;

function SetupMarker(AGeometry: TXSSGeometry; AMarker: TCT_Marker; AX,AY: double; ADefColor: longword; AForceMarker: boolean = False): TXSSGeometryMarker;
begin
  Result := Nil;

  if (AMarker = Nil) and not AForceMarker then
    Exit;

  if (AMarker <> Nil) and (AMarker.Symbol <> Nil) and (AMarker.Symbol.Val = stmsNone) then
    Exit;

  Result := AGeometry.AddMarker;;
  Result.Size := XSS_CHART_DEF_MARKER_SIZE;
  Result.MStyle := xgmsDiamond;

  if AForceMarker then
    Result.MStyle := xgmsSquare
  else if AMarker.Symbol <> Nil then begin
    case AMarker.Symbol.Val of
      stmsNone    : Result.MStyle := xgmsDiamond;
      stmsCircle  : Result.MStyle := xgmsCircle;
      stmsDash    : Result.MStyle := xgmsDash;
      stmsDiamond : Result.MStyle := xgmsDiamond;
      stmsDot     : Result.MStyle := xgmsDot;
      stmsPlus    : Result.MStyle := xgmsPlus;
      stmsSquare  : Result.MStyle := xgmsSquare;
      stmsStar    : Result.MStyle := xgmsStar;
      stmsTriangle: Result.MStyle := xgmsTriangle;
      stmsX       : Result.MStyle := xgmsX;
    end;

    if AMarker.Size <> Nil then
      Result.Size := AMarker.Size.Val / 4;
  end;


  if Result <> Nil then begin
    Result.Pt.X := AX;
    Result.Pt.Y := AY;
    Result.Style.Color := ADefColor;

    if (AMarker <> Nil) and (AMarker.SpPr <> Nil) then
      ApplySpPr(AMarker.SpPr,Result.Style);
  end;
end;

{ TXSSChartItemPlotArea }

procedure TXSSChartItemPlotArea.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; AChartSpace: TCT_PlotArea);
var
  i      : integer;
  Chart  : TXSSChartItem;
  Ser    : TXSSChartItemChartSeries;
  CatFont: TCT_TextBody;
  ValFont: TCT_TextBody;
begin
  ARect.Style.Color := XSS_CHART_DEF_COLOR;

  Chart := Nil;

  ValFont := Nil;
  CatFont := Nil;

  if AChartSpace.SpPr <> Nil then
    ApplySpPr(AChartSpace.SpPr,ARect.Style);

  if AChartSpace.ValAxis <> Nil then begin
    ValFont := AChartSpace.ValAxis[0].Shared.TxPr;
    if FData.ShowValAxis then begin
      // There is always one value axis, used for the grid lines.
      FValAxises[0].Build(ARect,AChartSpace.ValAxis[0]);
      for i := 1 to AChartSpace.ValAxis.Count - 1 do
        FValAxises.Add(FGDI,FData).Build(ARect,AChartSpace.ValAxis[i]);
    end;
  end;

  Ser := Nil;

  for i := 0 to FCharts.Count - 1 do begin
    Chart := TXSSChartItem(FCharts[i]);

    if Chart is TXSSChartItemBarChart then
      Ser := TXSSChartItemBarChart(Chart).Series
    else if Chart is TXSSChartItemLineChart then
      Ser := TXSSChartItemLineChart(Chart).Series
    else if Chart is TXSSChartItemAreaChart then
      Ser := TXSSChartItemAreaChart(Chart).Series
    else if Chart is TXSSChartItemScatterChart then
      Ser := TXSSChartItemScatterChart(Chart).Series
    else if Chart is TXSSChartItemBubbleChart then
      Ser := TXSSChartItemBubbleChart(Chart).Series
    else if Chart is TXSSChartItemRadarChart then
      Ser := TXSSChartItemRadarChart(Chart).Series;
  end;

  if (Ser <> Nil) and (AChartSpace.CatAxis <> Nil) then begin
    CatFont := AChartSpace.CatAxis[0].Shared.TxPr;
    if not (Chart is TXSSChartItemRadarChart) then
      FCatAxis.Build(ARect,Ser.ValCount,Ser.FCats,AChartSpace.CatAxis[0].Shared.TxPr,False);
  end;

  for i := 0 to FCharts.Count - 1 do begin
    Chart := TXSSChartItem(FCharts[i]);

    if Chart is TXSSChartItemBarChart then
      TXSSChartItemBarChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.BarChart)
    else if Chart is TXSSChartItemLineChart then
      TXSSChartItemLineChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.LineChart)
    else if Chart is TXSSChartItemAreaChart then
      TXSSChartItemAreaChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.AreaChart)
    else if Chart is TXSSChartItemPieChart then
      TXSSChartItemPieChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.PieChart)
    else if Chart is TXSSChartItemScatterChart then
      TXSSChartItemScatterChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.ScatterChart)
    else if Chart is TXSSChartItemBubbleChart then
      TXSSChartItemBubbleChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.BubbleChart)
    else if Chart is TXSSChartItemRadarChart then
      TXSSChartItemRadarChart(Chart).Build(ARect,AGeometry,ALegend,AChartSpace.RadarChart,ValFont,CatFont);
  end;
end;

procedure TXSSChartItemPlotArea.Clear;
begin
  FCharts.Clear;
end;

constructor TXSSChartItemPlotArea.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);

  FMargs := FData.DefMarg;

  FValAxises := TXSSChartItemValAxises.Create;
  FCatAxis := TXSSChartItemCatAxis.Create(FGDI,FData);

  FCharts := TObjectList.Create;
end;

destructor TXSSChartItemPlotArea.Destroy;
begin
  Clear;


  FValAxises.Free;
  FCatAxis.Free;

  FCharts.Free;

  inherited;
end;

procedure TXSSChartItemPlotArea.PreBuild(ALegend: TXSSChartItemLegend; AChartSpace: TCT_PlotArea);
begin
  FCharts.Clear;

  FValAxises.Add(FGDI,FData);

  if AChartSpace.BarChart <> Nil then begin
    FCharts.Add(TXSSChartItemBarChart.Create(FGDI,FData,FValAxises.First,FCatAxis));

    TXSSChartItemBarChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.BarChart);
  end;
  if AChartSpace.LineChart <> Nil then begin
    FCharts.Add(TXSSChartItemLineChart.Create(FGDI,FData,FValAxises.First,FCatAxis));

    TXSSChartItemLineChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.LineChart);
  end;
  if AChartSpace.AreaChart <> Nil then begin
    FCharts.Add(TXSSChartItemAreaChart.Create(FGDI,FData,FValAxises.First,FCatAxis));

    TXSSChartItemAreaChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.AreaChart);
  end;
  if AChartSpace.PieChart <> Nil then begin
    FCharts.Add(TXSSChartItemPieChart.Create(FGDI,FData));

    TXSSChartItemPieChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.PieChart);
  end;
  if AChartSpace.ScatterChart <> Nil then begin
    FCharts.Add(TXSSChartItemScatterChart.Create(FGDI,FData,FValAxises,FCatAxis));

    TXSSChartItemScatterChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.ScatterChart);
  end;
  if AChartSpace.BubbleChart <> Nil then begin
    FCharts.Add(TXSSChartItemBubbleChart.Create(FGDI,FData,FValAxises,FCatAxis));

    TXSSChartItemBubbleChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.BubbleChart);
  end;
  if AChartSpace.RadarChart <> Nil then begin
    FCharts.Add(TXSSChartItemRadarChart.Create(FGDI,FData,FValAxises.First,FCatAxis));

    TXSSChartItemRadarChart(FCharts[FCharts.Count - 1]).PreBuild(ALegend,AChartSpace.RadarChart);
  end;
end;

{ TXSSChartItemChartSpace }

procedure TXSSChartItemChartSpace.Build(AGeometry: TXSSGeometry; AChartSpace: TCT_ChartSpace);
var
  CSRect : TXSSGeometryRect;
  Rect   : TXSSGeometryRect;
  Rect2  : TXSSGeometryRect;
  L,T,R,B: double;
  TT     : double;
  Title  : TXSSChartItemTitle;
begin
  Clear;

  Rect2 := Nil;

  FData.ResetColor;

  CSRect := AGeometry.AddRect;
  CSRect.SetSize(0,0,1,1);
  CSRect.Style.LineColor := $000000;
  CSRect.Style.Color := $FFFFFF;

  if AChartSpace.SpPr <> Nil then
    ApplySpPr(AChartSpace.SpPr,CSRect.Style);

  L := FPlotArea.Margs;
  T := FPlotArea.Margs;
  R := FPlotArea.Margs;
  B := FPlotArea.Margs;

  TT := 0;

  if AChartSpace.Chart.Title <> Nil then begin
    Title := TXSSChartItemTitle.Create(FGDI,FData);
    try
      Title.PreBuild(AChartSpace.Chart.Title);

      Rect := CSRect.AddRect;

      Rect.SetSize(FData.MapL(Title.Margs),
                   FData.MapT(Title.Margs),
                   FData.MapL(Title.Margs + Title.Width),
                   FData.MapT(Title.Margs + Title.Height));

      Title.Build(Rect,AChartSpace.Chart.Title);

      TT := T + Title.Height;
      T := TT + Title.Margs;
    finally
      Title.Free;
    end;
  end;

  if AChartSpace.Chart.Legend <> Nil then begin
    FLegend := TXSSChartItemLegend.Create(FGDI,FData,AChartSpace.Chart.Legend);

    FPlotArea.PreBuild(FLegend,AChartSpace.Chart.PlotArea);

    Rect2 := CSRect.AddRect;

    Rect2.SetSize(0,0,FData.MapL(FLegend.PtWidth),FData.MapT(FLegend.PtHeight));

    if (AChartSpace.Chart.Legend.Layout <> Nil) and (AChartSpace.Chart.Legend.Layout.ManualLayout <> Nil) then
      ApplyManualLayout(Rect2,AChartSpace.Chart.Legend.Layout.ManualLayout);

    case AChartSpace.Chart.Legend.LegendPos.Val of
      stlpB : B := B + FLegend.PtHeight + FLegend.Margs;
      stlpTr: begin
        R := R + FLegend.PtWidth + FLegend.Margs;
        T := T + FLegend.PtHeight + FLegend.Margs;
      end;
      stlpL : L := L + FLegend.PtWidth + FLegend.Margs;
      stlpR : R := R + FLegend.PtWidth + FLegend.Margs;
      stlpT : T := T + FLegend.PtHeight + FLegend.Margs;
    end;

    case AChartSpace.Chart.Legend.LegendPos.Val of
      stlpB : begin
        Rect2.CenterHoriz(1);
        Rect2.Move(0,1 - Rect2.Height - FData.MapT(FLegend.Margs));
      end;
      stlpTr: begin
        Rect2.Move(1 - Rect2.Width - FData.MapL(FLegend.Margs),FData.MapT(TT + FLegend.Margs));
      end;
      stlpL : begin
        Rect2.CenterVert(1);
        Rect2.Move(FData.MapL(FLegend.Margs),0);
      end;
      stlpR : begin
        Rect2.CenterVert(1);
        Rect2.Move(1 - Rect2.Width - FData.MapL(FLegend.Margs),0);
      end;
      stlpT : begin
        Rect2.CenterHoriz(1);
        Rect2.Move(0,FData.MapT(TT + FLegend.Margs));
      end;
    end;
  end
  else
    FPlotArea.PreBuild(Nil,AChartSpace.Chart.PlotArea);

  Rect := CSRect.AddRect;
  Rect.SetSize(FData.MapL(L),FData.MapT(T),FData.MapR(R),FData.MapB(B));

  if (AChartSpace.Chart.PlotArea.Layout <> Nil) and (AChartSpace.Chart.PlotArea.Layout.ManualLayout <> Nil) then
    ApplyManualLayout(Rect,AChartSpace.Chart.PlotArea.Layout.ManualLayout);

  FPlotArea.Build(Rect,AGeometry,FLegend,AChartSpace.Chart.PlotArea);

  if FLegend <> Nil then
    FLegend.Build(Rect2,AChartSpace.Chart.Legend);
end;

procedure TXSSChartItemChartSpace.Clear;
begin
  if FLegend <> Nil then begin
    FLegend.Free;
    FLegend := Nil;
  end;

  FPlotArea.Clear;
end;

constructor TXSSChartItemChartSpace.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);

  FPlotArea := TXSSChartItemPlotArea.Create(FGDI,FData);
end;

destructor TXSSChartItemChartSpace.Destroy;
begin
  Clear;

  FPlotArea.Free;

  inherited;
end;

{ TXSSChartItem }

constructor TXSSChartItem.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  FGDI := AGDI;
  FData := AData;
end;

{ TXSSChartItemChart }

constructor TXSSChartItemChart.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);
end;

destructor TXSSChartItemChart.Destroy;
begin

  inherited;
end;

{ TXSSChartItemBarChart }

procedure TXSSChartItemBarChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ABarChart: TCT_BarChart);
var
  Grp: TST_BarGrouping;
begin
  if ABarChart.Shared.Grouping <> Nil then
    Grp := ABarChart.Shared.Grouping.Val
  else
    Grp := stbgStandard;
  case Grp of
    stbgPercentStacked: FSeries.BuildStacked(ARect,ABarChart.Shared.Series,True);
    stbgStacked       : FSeries.BuildStacked(ARect,ABarChart.Shared.Series,False);
    stbgStandard,
    stbgClustered     : FSeries.Build(ARect,ABarChart.Shared.Series);
  end;
end;

constructor TXSSChartItemBarChart.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create(AGDI,AData,AValAxis,ACatAxis);
end;

procedure TXSSChartItemBarChart.PreBuild(ALegend: TXSSChartItemLegend; ABarChart: TCT_BarChart);
var
  i   : integer;
  Grp: TST_BarGrouping;
begin
  if ABarChart.Shared.Grouping <> Nil then
    Grp := ABarChart.Shared.Grouping.Val
  else
    Grp := stbgStandard;
  FSeries.PreBuild(Grp,ABarChart.Shared.Series);

  if ALegend <> Nil then begin
    for i := 0 to FSeries.Count - 1 do
      ALegend.Add(i,FSeries[i].Color,FSeries[i].Name);
  end;
end;

{ TXSSChartItemBarChartSerie }

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; ABarSer: TCT_BarSer; AIndex: integer);
var
  i   : integer;
  Rect: TXSSGeometryRect;
  W   : double;
  H   : double;
  BarW: double;
  GapW: double;
  X   : double;
begin
  if Length(FVals) < 1 then
    Exit;

  SetupDataPoints(ABarSer.DPtXpgList,ABarSer.DLbls);

  H := ARect.Pt2.Y - ARect.Pt1.Y;

  W := (ARect.Pt2.X - ARect.Pt1.X) / Length(FVals);
  BarW := (W / (FOwner.GapWidth + FOwner.Count));
  GapW := BarW * FOwner.GapWidth;
  X := ARect.Pt1.X + (GapW / 2) + (BarW * AIndex);

  for i := 0 to High(FVals) do begin
    Rect := ARect.AddRect;
    Rect.Pt1.X := X;
    Rect.Pt2.X := X + BarW;
    X := X + W;

    Rect.Pt1.Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i]);
    Rect.Pt2.Y := ARect.Pt2.Y - FOwner.MapY(H,FOwner.ValAxis.ScaleOrigo);

    Rect.Style.LineColor := XLS_COLOR_NONE;

    Rect.Style.Color := FColor;

    FData.ApplyDataPoint(i,Rect,Rect.Pt1.X + BarW / 2,Rect.Pt1.Y,FVals[i] >= FOwner.ValAxis.ScaleOrigo);
  end;

  SetLength(FVals,0);
end;

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; ALineSer: TCT_LineSer; AHasMarkers: boolean);
var
  i     : integer;
  PLine : TXSSGeometryPolyline;
  LineCl: longword;
  LineW : double;
  W     : double;
  H     : double;
  X     : double;
  Y     : double;
begin
  if Length(FVals) < 1 then
    Exit;

  LineCl := FColor;
  LineW := XSS_CHART_DEF_LINECHARTWIDTH;

  GetLineStyle(ALineSer.Shared.SpPr,LineCl,LineW);

  SetupDataPoints(ALineSer.DPtXpgList,ALineSer.DLbls);

  H := ARect.Pt2.Y - ARect.Pt1.Y;

  W := (ARect.Pt2.X - ARect.Pt1.X) / Length(FVals);
  X := ARect.Pt1.X + W / 2;

  PLine := ARect.AddPolyline;
  PLine.Style.LineColor := LineCl;
  PLine.Style.LineWidth := LineW;

  for i := 0 to High(FVals) do begin
    Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i]);

    PLine.Add(X,Y);

    if AHasMarkers and (ALineSer.Marker <> Nil) then
      SetupMarker(PLine,ALineSer.Marker,X,Y,LineCl);

    FData.ApplyDataPoint(i,PLine,X,PLine[i].Y,FVals[i] >= FOwner.ValAxis.ScaleOrigo);

    X := X + W;
  end;

  SetLength(FVals,0);
end;

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; APieSer: TCT_PieSer; AStartAngle: double; ALegend: TXSSChartItemLegend);
var
  i    : integer;
  Sum  : double;
  V,V2 : double;
  VTxt : double;
  X,Y  : double;
  r    : double;
  S    : AxUCString;
  Slice: TXSSGeometryPizzaSlice;
  Dp   : TXSSChartItemDataPoint;
begin
  Sum := ValsSum;

  if Sum <= 0 then
    Exit;

  SetupDataPoints(APieSer.DPtXpgList,APieSer.DLbls);

  FData.ResetColor;

  V := (Pi * 2 - Pi / 2) + DegToRad(AStartAngle);
  r := (Min(FData.PtWidth,FData.PtHeight) / 2) * 0.65;

  for i := 0 to High(FVals) do begin
    Slice := ARect.AddPizzaSlice;
    Slice.Assign(ARect);
    Slice.Shrink(FData.MapL(FData.DefMarg),FData.MapT(FData.DefMarg));
    Slice.EndAngle := V;
    V2 := ((FVals[i] / Sum)) * Pi * 2;

    VTxt := V + (V2 / 2);
    V := V + V2;

    Slice.StartAngle := V;

    Slice.Style.Color := FData.GetSerieColor;
    Slice.Style.LineColor := XLS_COLOR_NONE;

    if (ALegend <> Nil) and (i < ALegend.Count) then
      ALegend[i].Color := Slice.Style.Color;

    if APieSer.DLbls <> Nil then begin
      S := '';
      if (APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowPercent <> Nil) and APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowPercent.Val then
        S := Format('%.0f%%',[(FVals[i] / Sum) * 100])
      else if (APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowVal <> Nil) and APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowVal.Val then
        S := Format('%.0f',[FVals[i]]);
    end;

    X := Cos(VTxt) * FData.MapL(r) + Slice.MidX;
    Y := Sin(VTxt) * FData.MapT(r) + Slice.MidY;

    Dp := FData.ApplyDataPoint(i,Slice,X,Y,False);

    if (APieSer.DLbls <> Nil) and (Dp = Nil) then begin
      S := '';
      if (APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowPercent <> Nil) and APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowPercent.Val then
        S := Format('%.0f%%',[(FVals[i] / Sum) * 100])
      else if (APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowPercent <> Nil) and APieSer.DLbls.Group_DLbls.EG_DLblSShared.ShowPercent.Val then
        S := Format('%.0f',[FVals[i]]);
      if S <> '' then
        ARect.AddText(FData.GetChartFont(APieSer.DLbls.Group_DLbls.EG_DLblSShared.TxPr),cvaCenter,chaCenter,X,Y,S);
    end;
  end;
end;

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; ABubbleSer: TCT_BubbleSer);
var
  i      : integer;
  W,H    : double;
  BSizes : TDynSingleArray;
  MaxArea: double;
  Circle : TXSSGeometryCircle;
begin
  if ABubbleSer.BubbleSize = Nil then
    Exit;

  ABubbleSer.BubbleSize.NumRef.RCells.GetFloatArray(BSizes,XSS_CHART_MAXDATAPOINTS);

  MaxArea := 0;
  for i := 0 to High(BSizes) do
    MaxArea := Max(MaxArea,BSizes[i]);

  MaxArea := Sqrt(MaxArea);

  H := ARect.Pt2.Y - ARect.Pt1.Y;
  W := ARect.Width;

  SetupDataPoints(ABubbleSer.DPtXpgList,ABubbleSer.DLbls);

  for i := 0 to Min(High(FVals),High(BSizes)) do begin
    if BSizes[i] > 0 then begin
      Circle := ARect.AddCircle;
      Circle.Pt.X := ARect.Pt1.X + (AXSerie.Val[i] / AXScaleRange) * W;
      Circle.Pt.Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i]);
      Circle.RefRect := ARect;
      Circle.MMScale := (Sqrt(BSizes[i]) / MaxArea) * 0.15;
      Circle.Style.Color := FDefColor;

      FData.ApplyDataPoint(i,Circle,Circle.Pt.X,Circle.Pt.Y,FVals[i] >= FOwner.ValAxis.ScaleOrigo);
    end;
  end;
end;

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; AScatterSer: TCT_ScatterSer);
var
  i     : integer;
  W,H   : double;
  Marker: TXSSGeometryMarker;
  Line  : TXSSGeometryLine;
  LineCl: longword;
  LineW : double;
begin
  H := ARect.Pt2.Y - ARect.Pt1.Y;
  W := ARect.Width;

  LineW := 0;
  LineCl := XLS_COLOR_NONE;
  GetLineStyle(AScatterSer.Shared.SpPr,LineCl,LineW);

  if (LineCl <> XLS_COLOR_NONE) and (LineW > 0) then begin
    for i := 0 to High(FVals) - 1 do begin
      Line := ARect.AddLine;
      Line.Pt1.X := ARect.Pt1.X + (AXSerie.Val[i] / AXScaleRange) * W;
      Line.Pt1.Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i]);
      Line.Pt2.X := ARect.Pt1.X + (AXSerie.Val[i + 1] / AXScaleRange) * W;
      Line.Pt2.Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i + 1]);
      Line.Style.LineColor := LineCl;
      Line.Style.LineWidth := LineW;
    end;
  end;

  SetupDataPoints(AScatterSer.DPtXpgList,AScatterSer.DLbls);

  for i := 0 to High(FVals) do begin
    Marker := ARect.AddMarker;
    Marker.Pt.X := ARect.Pt1.X + (AXSerie.Val[i] / AXScaleRange) * W;
    Marker.Pt.Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i]);
    Marker.MStyle := xgmsDiamond;
    Marker.Size := 2;
    Marker.Style.Color := FDefColor;

    FData.ApplyDataPoint(i,Marker,Marker.Pt.X,Marker.Pt.Y,FVals[i] >= FOwner.ValAxis.ScaleOrigo);
  end;
end;

function TXSSChartItemChartSerie.Count: integer;
begin
  Result := Length(FVals);
end;

constructor TXSSChartItemChartSerie.Create(AOwner: TXSSChartItemChartSeries; AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);

  FOwner := AOwner;
end;

destructor TXSSChartItemChartSerie.Destroy;
begin
  inherited;
end;

procedure TXSSChartItemChartSerie.GetLineStyle(ASpPr: TCT_ShapeProperties; var ALineColor: longword; var ALineWidth: double);
var
  SpPr  : TXLSDrwShapeProperies;
begin
  if ASpPr <> Nil then begin
    SpPr := TXLSDrwShapeProperies.Create(ASpPr);
    try
      if SpPr.Line <> Nil then begin
        ALineColor := SpPr.Line.RGB_(ALineColor);
        if SpPr.Line.Width < 1000 then
          ALineWidth := SpPr.Line.Width;
      end;
    finally
      SpPr.Free;
    end;
  end;
end;

function TXSSChartItemChartSerie.GetVal(Index: integer): double;
begin
  if Index < Count then
    Result := FVals[Index]
  else
    Result := 0;
end;

procedure TXSSChartItemChartSerie.PreBuild(AIndex: integer; AVal: TXLSRelCells; ACat: TCT_AxDataSource; AShared: TCT_SerieShared; AUseFill: boolean);
var
  i: integer;
begin
  if AVal <> Nil then
    AVal.GetFloatArray(FVals,XSS_CHART_MAXDATAPOINTS);

  if (ACat <> Nil) and (ACat.StrRef <> Nil) and (ACat.StrRef.RCells <> Nil) then
    ACat.StrRef.RCells.GetStrArray(FOwner.FCats,XSS_CHART_MAXDATAPOINTS);

  FMinVal := MAXDOUBLE;
  FMaxVal := MINDOUBLE;

  FDefColor := FData.GetSerieColor;
  if AUseFill then
    FColor := GetSpPrFillColor(AShared.SpPr,FDefColor)
  else
    FColor := GetSpPrLineColor(AShared.SpPr,FDefColor);

  if AShared.Tx <> Nil then
    FName := AShared.Tx.Text
  else
    FName := 'Serie ' + IntToStr(AIndex + 1);

  for i := 0 to High(FVals) do begin
    FMinVal := Min(FMinVal,FVals[i]);
    FMaxVal := Max(FMaxVal,FVals[i]);
  end;
end;

procedure TXSSChartItemChartSerie.SetupDataPoints(ADPt: TCT_DPtXpgList; ADLbls: TCT_DLbls);
begin
  FData.DataPoints.Clear;
  FData.DataPoints.Add(ADPt);
  if ADLbls <> Nil then
    FData.DataPoints.Add(ADLbls,FVals);
end;

function TXSSChartItemChartSerie.Vals: TDynSingleArray;
begin
  Result := FVals;
end;

function TXSSChartItemChartSerie.ValsSum: double;
var
  i: integer;
begin
  Result := 0;

  for i := 0 to High(FVals) do
    Result := Result + FVals[i];
end;

{ TXSSChartItemBarChartSeries }

function TXSSChartItemChartSeries.Add: TXSSChartItemChartSerie;
begin
  Result := TXSSChartItemChartSerie.Create(Self,FGDI,FData);

  inherited Add(Result);
end;

procedure TXSSChartItemChartSeries.AddGridlines(ARect: TXSSGeometryRect);
var
  H   : double;
  V   : double;
  D   : double;
  X,Y : double;
  Line: TXSSGeometryLine;
begin
  H := ARect.Pt2.Y - ARect.Pt1.Y;

  V := ValAxis.ScaleOrigo;
  D := Max(Abs(ValAxis.ScaleMax),Abs(ValAxis.ScaleMin)) / ValAxis.AxisTicks;

  if FValAxis.Visible then begin
    X := ARect.Pt1.X - FData.MapL(FData.DefPadding / 2);
    FGDI.SetFont(FData.Font);
  end
  else
    X := ARect.Pt1.X;

  while V <= ValAxis.ScaleMax do begin
    Line := ARect.AddLine;
    Y := ARect.Pt2.Y - MapY(H,V);

    Line.SetSize(X,Y,ARect.Pt2.X,Y);

    if FValAxis.Visible then
      ARect.AddText(FValAxis.Font,cvaCenter,chaRight,X - FData.MapL(FData.DefPadding),Y,Format('%.0f',[V]));

    V := V + D;
  end;

  V := ValAxis.ScaleOrigo - D;

  while V >= ValAxis.ScaleMin do begin
    Line := ARect.AddLine;
    Y := ARect.Pt2.Y - MapY(H,V);

    Line.SetSize(X,Y,ARect.Pt2.X,Y);

    if FValAxis.Visible then
      ARect.AddText(FValAxis.Font,cvaCenter,chaRight,X - FData.MapL(FData.DefPadding),Y,FloatToStr(V));

    V := V - D;
  end;
end;

procedure TXSSChartItemChartSeries.Build(ARect: TXSSGeometryRect; ABarSeries: TCT_BarSeries);
var
  i: integer;
begin
  AddGridlines(ARect);

  for i := 0 to ABarSeries.Count - 1 do begin
    Items[i].Build(ARect,ABarSeries[i],i);
  end;
end;

procedure TXSSChartItemChartSeries.Build(ARect: TXSSGeometryRect; ALineSeries: TCT_LineSeries; AHasMarkers: boolean);
var
  i: integer;
begin
  AddGridlines(ARect);

  for i := 0 to ALineSeries.Count - 1 do
    Items[i].Build(ARect,ALineSeries[i],AHasMarkers);
end;

procedure TXSSChartItemChartSeries.Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; AScatterSeries: TCT_ScatterSeries);
var
  i: integer;
begin
  AddGridlines(ARect);

  for i := 0 to AScatterSeries.Count - 1 do
    Items[i].Build(ARect,AXSerie,AXScaleRange,AScatterSeries[i]);
end;

procedure TXSSChartItemChartSeries.BuildStacked(ARect: TXSSGeometryRect; ABarSeries: TCT_BarSeries; APercent: boolean);
var
  i,j : integer;
  Sum : array of double;
  Rect: TXSSGeometryRect;
  W   : double;
  H   : double;
  BarW: double;
  GapW: double;
  X   : double;
  Y   : double;
begin
  AddGridlines(ARect);

  SetLength(Sum,FValCount);

  if APercent then begin
    for i := 0 to FValCount - 1 do begin
      Sum[i] := 0;
      for j := 0 to Count - 1 do
        Sum[i] := Sum[i] + Items[j].Val[i];
    end;
  end;

  H := ARect.Pt2.Y - ARect.Pt1.Y;

  W := (ARect.Pt2.X - ARect.Pt1.X) / FValCount;
  BarW := (W / (FGapWidth + 1));
  GapW := BarW * FGapWidth;

  X := ARect.Pt1.X + (GapW / 2);
  for i := 0 to FValCount - 1 do begin
    Y := ARect.Pt2.Y - MapY(H,ValAxis.ScaleOrigo);
    for j := 0 to Count - 1 do begin
      Rect := ARect.AddRect;
      Rect.Pt1.X := X;
      Rect.Pt2.X := X + BarW;

      Rect.Pt2.Y := Y;
      if APercent then begin
        if Sum[i] <> 0 then
          Rect.Pt1.Y := Rect.Pt2.Y - MapY(H,(Items[j].Val[i] / Sum[i]) * 100);
      end
      else
        Rect.Pt1.Y := Rect.Pt2.Y - MapY(H,Items[j].Val[i]);
      Y := Rect.Pt1.Y;

      Rect.Style.LineColor := XLS_COLOR_NONE;

      Rect.Style.Color := Items[j].Color;
    end;
    X := X + W;
  end;
end;

procedure TXSSChartItemChartSeries.CalcMinMax(AGrouping: TST_BarGrouping);
var
  i,j: integer;
  Sum: array of double;
begin
  FMinVal := MAXDOUBLE;
  FMaxVal := MINDOUBLE;
  FValCount := 0;

  for i := 0 to Count - 1 do
    FValCount := Max(FValCount,Items[i].Count);

  case AGrouping of
    stbgClustered,
    stbgStandard      : begin
      for i := 0 to Count - 1 do begin
        FMinVal := Min(FMinVal,Items[i].MinVal);
        FMaxVal := Max(FMaxVal,Items[i].MaxVal);
      end;
    end;
    stbgPercentStacked: begin
      FMinVal := 0;
      FMaxVal := 100;
    end;
    stbgStacked       : begin
      SetLength(Sum,FValCount);
      for i := 0 to FValCount - 1 do begin
        Sum[i] := 0;
        for j := 0 to Count - 1 do
          Sum[i] := Sum[i] + Items[j].Val[i];
      end;
      for i := 0 to FValCount - 1 do begin
        FMinVal := Min(FMinVal,Sum[i]);
        FMaxVal := Max(FMaxVal,Sum[i]);
      end;
    end;
  end;
end;

constructor TXSSChartItemChartSeries.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create;

  FGDI := AGDI;
  FData := AData;
  FValAxis := AValAxis;
  FCatAxis := ACatAxis;

  FGapWidth := 1.5;
end;

destructor TXSSChartItemChartSeries.Destroy;
begin

  inherited;
end;

function TXSSChartItemChartSeries.GetItems(Index: integer): TXSSChartItemChartSerie;
begin
  Result := TXSSChartItemChartSerie(inherited Items[Index]);
end;

function TXSSChartItemChartSeries.MapY(AHeight, AVal: double): double;
begin
  Result := AHeight * ((AVal - FValAxis.ScaleMin) / FValAxis.ScaleRange);
end;

procedure TXSSChartItemChartSeries.PreBuild(AAreaSeries: TCT_AreaSeries);
var
  i  : integer;
  Ser: TXSSChartItemChartSerie;
begin
  Clear;

  for i := 0 to AAreaSeries.Count - 1 do begin
    Ser := Add;

    Ser.PreBuild(i,AAreaSeries[i].Val.NumRef.RCells,AAreaSeries[i].Cat,AAreaSeries[i].Shared,False);
  end;

  CalcMinMax(stbgStandard);

  ValAxis.CalcScale(FMinVal,FMaxVal,6);
end;

procedure TXSSChartItemChartSeries.PreBuild(AScatterSeries: TCT_ScatterSeries);
var
  i  : integer;
  Ser: TXSSChartItemChartSerie;
begin
  Clear;

  for i := 0 to AScatterSeries.Count - 1 do begin
    Ser := Add;

    Ser.PreBuild(i,AScatterSeries[i].YVal.NumRef.RCells,Nil,AScatterSeries[i].Shared,False);
  end;

  CalcMinMax(stbgStandard);

  ValAxis.CalcScale(FMinVal,FMaxVal,6);
end;

procedure TXSSChartItemChartSeries.PreBuild(ALineSeries: TCT_LineSeries);
var
  i  : integer;
  Ser: TXSSChartItemChartSerie;
begin
  Clear;

  for i := 0 to ALineSeries.Count - 1 do begin
    Ser := Add;

    Ser.PreBuild(i,ALineSeries[i].Val.NumRef.RCells,ALineSeries[i].Cat,ALineSeries[i].Shared,False);
  end;

  CalcMinMax(stbgStandard);

  ValAxis.CalcScale(FMinVal,FMaxVal,6);
end;

procedure TXSSChartItemChartSeries.PreBuild(AGrouping: TST_BarGrouping; ABarSeries: TCT_BarSeries);
var
  i  : integer;
  Ser: TXSSChartItemChartSerie;
begin
  Clear;

  for i := 0 to ABarSeries.Count - 1 do begin
    Ser := Add;

    Ser.PreBuild(i,ABarSeries[i].Val.NumRef.RCells,ABarSeries[i].Cat,ABarSeries[i].Shared,True);
  end;

  CalcMinMax(AGrouping);

  ValAxis.CalcScale(FMinVal,FMaxVal,6);
end;

{ TXSSChartItemLegend }

procedure TXSSChartItemLegend.Add(AIndex: integer; AColor: longword; AText: AxUCString);
var
  Item: TXSSChartItemLegendItem;
  LE  : TCT_LegendEntry;
begin
  Item := TXSSChartItemLegendItem.Create;

  if FCTLegend.LegendEntryXpgList <> Nil then begin
    LE := FCTLegend.LegendEntryXpgList.FindIdx(FItems.Count);
    if LE <> Nil then
      Item.Font := FData.GetChartFont(LE.EG_LegendEntryData.TxPr);
  end;

  if Item.Font <> Nil then
    FGDI.SetFont(Item.Font)
  else
    FGDI.SetFont(FFont);

  Item.Color := AColor;
  Item.Text := AText;
  Item.Width := FPadding + SerieMarkerSz + FPadding + FGDI.PixToPtX(FGDI.TextWidth(Atext)) + FPadding + FPadding;
  Item.Height := FGDI.PixToPtX(FGDI.TextHeight(Atext));

  if FCTLegend.LegendPos.Val in [stlpT,stlpB] then
    FTxtWidth := FTxtWidth + Item.Width
  else
    FTxtWidth := Max(FTxtWidth,SerieMarkerSz + Item.Width);
  FTxtHeight := Max(FTxtHeight,Item.Height);

  FItems.Add(Item);
end;

procedure TXSSChartItemLegend.Build(ARect: TXSSGeometryRect; ALegend: TCT_Legend);
var
  i   : integer;
  Txt : TXSSGeometryText;
  Rect: TXSSGeometryRect;
  X,Y : double;
  D   : double;
begin
  ARect.Style.Color := XSS_CHART_DEF_COLOR;
  ARect.Style.LineColor := XLS_COLOR_NONE;
  ARect.Style.LineWidth := XSS_CHART_DEF_LINEWIDTH;

  ApplySpPr(FSpPr,ARect.Style);

  X := ARect.Pt1.X + FData.MapL(FPadding);
  Y := ARect.Pt1.Y + FData.MapT(FPadding);

  for i := 0 to FItems.Count - 1 do begin
    Rect := ARect.AddRect;
    D := (FTxtHeight - SerieMarkerSz) / 2;
    Rect.SetSize(X + FData.MapL(FPadding),Y + FData.MapT(D),X + FData.MapL(FPadding + SerieMarkerSz),Y + FData.MapT(D + SerieMarkerSz));
    Rect.Style.Color := Items[i].Color;

    Txt := ARect.AddText;
    Txt.Pt.X := FData.MapL(FPadding + SerieMarkerSz + FPadding) + X;
    if Items[i].Height < FTxtHeight then
      Txt.Pt.Y := Y + FData.MapT((FTxtHeight - Items[i].Height) / 2)
    else
      Txt.Pt.Y := Y;
    if Items[i].Font <> Nil then
      Txt.Font := Items[i].Font
    else
      Txt.Font := FFont;
    Txt.Text := Items[i].Text;

    if FCTLegend.LegendPos.Val in [stlpT,stlpB] then
      X := X + FData.MapL(Items[i].Width)
    else
      Y := Y + FData.MapT(FTxtHeight);
  end;
end;

procedure TXSSChartItemLegend.Clear;
begin
  FTxtWidth := 0;
  FTxtHeight := 0;

  FItems.Clear;

  FFont := Nil;

  if FSpPr <> Nil then begin
    FSpPr.Free;
    FSpPr := Nil;
  end;
end;

function TXSSChartItemLegend.Count: integer;
begin
  Result := FItems.Count;
end;

constructor TXSSChartItemLegend.Create(AGDI: TAXWGDI; AData: TXSSChartData; ALegendData: TCT_Legend);
begin
  inherited Create(AGDI,AData);

  FCTLegend := ALegendData;

  FItems := TObjectList.Create;

  FMargs := FData.DefMarg / 2;
  FPadding := FMargs;

  if FCTLegend.SpPr <> Nil then
    FSpPr := TXLSDrwShapeProperies.Create(FCTLegend.SpPr);

  FFont := FData.GetChartFont(FCTLegend.TxPr);
end;

destructor TXSSChartItemLegend.Destroy;
begin
  FItems.Free;

  Clear;

  inherited;
end;

function TXSSChartItemLegend.GetItems(Index: integer): TXSSChartItemLegendItem;
begin
  Result := TXSSChartItemLegendItem(FItems[Index]);
end;

function TXSSChartItemLegend.GetPtHeight: double;
begin
  if FCTLegend.LegendPos.Val in [stlpT,stlpB] then
    Result := FPadding + FTxtHeight + FPadding
  else
    Result := FPadding + FTxtHeight * FItems.Count + FPadding;
end;

function TXSSChartItemLegend.GetPtWidth: double;
begin
  Result := FPadding + FTxtWidth + FPadding;
end;

function TXSSChartItemLegend.SerieMarkerSz: double;
begin
  Result := FTxtHeight * 0.4;
end;

{ TXSSChartData }

procedure TXSSChartData.AddParas(AGeometry: TXSSGeometry; AX, AY: double; AParas: TXSSChartParagraph; AVertAlign: TXc12VertAlignment);
var
  i  : integer;
  Rich: TXSSGeometryRichText;
begin
  Rich := AGeometry.AddRichText;
  Rich.Pt.X := AX;
  Rich.Pt.Y := AY;
  Rich.VertAlign := AVertAlign;
  Rich.HorizAlign := chaCenter;

  for i := 0 to AParas.Runs.Count - 1 do
    Rich.Add(AParas.Runs[i].Text,AParas.Runs[i].Font);
end;

function TXSSChartData.ApplyDataPoint(AIndex: integer; AGeometry: TXSSGeometry; AX, AY: double; AAboveOrigo: boolean): TXSSChartItemDataPoint;
begin
  Result := FDataPoints.Find(AIndex);
  if Result <> Nil then begin
    if Result.SpPr <> Nil then
      ApplySpPr(Result.SpPr,AGeometry.Style);
    if Result.Paras <> Nil then begin
      if AAboveOrigo then
        AddParas(AGeometry,AX,AY,Result.Paras,cvaBottom)
      else
        AddParas(AGeometry,AX,AY,Result.Paras,cvaTop);
    end;
  end;
end;

constructor TXSSChartData.Create;
var
  i,j: integer;
begin
  FDataPoints := TXSSChartItemDataPoints.Create(Self);

  FShowValAxis := True;

  j := 0;

  for i := 0 to High(XSS_CHART_DEF_FILLCOLORS) do
    FColors[i] := XSS_CHART_DEF_FILLCOLORS[i];

  Inc(j,Length(XSS_CHART_DEF_FILLCOLORS));

  for i := 0 to High(XSS_CHART_DEF_FILLCOLORS) do
    FColors[i + j] := LightenColor(XSS_CHART_DEF_FILLCOLORS[i],0.2);

  Inc(j,Length(XSS_CHART_DEF_FILLCOLORS));

  for i := 0 to High(XSS_CHART_DEF_FILLCOLORS) do
    FColors[i + j] := LightenColor(XSS_CHART_DEF_FILLCOLORS[i],0.6);

  Inc(j,Length(XSS_CHART_DEF_FILLCOLORS));

  for i := 0 to High(XSS_CHART_DEF_FILLCOLORS) do
    FColors[i + j] := LightenColor(XSS_CHART_DEF_FILLCOLORS[i],0.8);

  FDefMarg := 10;
  FDefPadding := 5;

  FGarbage := TObjectList.Create;

  FTitleFont := TXc12Font.Create(Nil);
  FGarbage.Add(FTitleFont);
end;

destructor TXSSChartData.Destroy;
begin
  FDataPoints.Free;
  FGarbage.Free;

  inherited;
end;

function TXSSChartData.GetChartFont(ARPr: TCT_TextCharacterProperties; ADefFont: TXc12Font): TXc12Font;
var
  Fill: TXLSDrwColor;
begin
  if ARPr <> Nil then begin

    Result := TXc12Font.Create(Nil);

    FGarbage.Add(Result);

    Result.Assign(ADefFont);

    if ARPr.Sz < 100000 then
      Result.Size := ARPr.Sz / 100;
    if ARPr.B then
      Result.Style := Result.Style + [xfsBold];
    if ARPr.I then
      Result.Style := Result.Style + [xfsItalic];

    if (ARPr.Latin <> Nil) and (ARPr.Latin.Typeface <> '') then
      Result.Name := ARPr.Latin.Typeface;

     if ARPr.FillProperties.SolidFill <> Nil then begin
       Fill := TXLSDrwColor.Create(ARPr.FillProperties.SolidFill.ColorChoice);
       try
         Result.Color := RGBColorToXc12(RevRGB(Fill.AsRGB));
       finally
         Fill.Free;
       end;
     end;
  end
  else
    Result := ADefFont;
end;

function TXSSChartData.GetSerieColor: longword;
begin
  Result := FColors[FCurrColor];

  Inc(FCurrColor);
  if FCurrColor > High(FColors) then
    FCurrColor := 0;
end;

function TXSSChartData.GetChartFont(ATxt: TCT_TextBody): TXc12Font;
begin
  if(ATxt <> Nil) and (ATxt.Paras <> Nil) and (ATxt.Paras.Count > 0) and
    (ATxt.Paras[0].PPr <> Nil) and (ATxt.Paras[0].PPr.DefRPr <> Nil) then
    Result := GetChartFont(ATxt.Paras[0].PPr.DefRPr,FFont)
  else
    Result := FFont;
end;

function TXSSChartData.MapB(APtSz: double): double;
begin
  Result := 1 - (APtSz / FPtHeight);
end;

function TXSSChartData.MapL(APtSz: double): double;
begin
  Result := APtSz / FPtWidth;
end;

function TXSSChartData.MapR(APtSz: double): double;
begin
  Result := 1 - (APtSz / FPtWidth);
end;

function TXSSChartData.MapT(APtSz: double): double;
begin
  Result := APtSz / FPtHeight;
end;

procedure TXSSChartData.ResetColor;
begin
  FCurrColor := 0;
end;

procedure TXSSChartData.SetFont(const Value: TXc12Font);
begin
  FFont := Value;

  FTitleFont.Assign(FFont);
  FTitleFont.Size := 18;
  FTitleFont.Style := [xfsBold];
end;

{ TXSSChartItemLegendItem }

destructor TXSSChartItemLegendItem.Destroy;
begin

  inherited;
end;

{ TXSSChartItemDataPoint }

procedure TXSSChartItemDataPoint.AddParas(AParas: TCT_TextParagraphXpgList);
begin
  if FParas <> Nil then
    FParas.Free;

  FParas := TXSSChartParagraph.Create(FData,AParas,FData.Font);
end;

procedure TXSSChartItemDataPoint.AddParas(AText: AxUCstring; AFont: TXc12Font);
begin
  if FParas <> Nil then
    FParas.Free;

  FParas := TXSSChartParagraph.Create(FData,AText,AFont);
end;

constructor TXSSChartItemDataPoint.Create(AData: TXSSChartData; ADp: TCT_DPt);
begin
  FData := AData;

  if ADp.Idx <> Nil then
    FIndex := ADp.Idx.Val;

  if ADp.SpPr <> Nil then
    FSpPr := TXLSDrwShapeProperies.Create(ADp.SpPr);
end;

constructor TXSSChartItemDataPoint.Create(AData: TXSSChartData);
begin
  FData := AData;
end;

destructor TXSSChartItemDataPoint.Destroy;
begin
  if FSpPr <> Nil then
    FSpPr.Free;

  if FParas <> Nil then
    FParas.Free;

  inherited;
end;

function TXSSChartItemDataPoint.FillRGB(ADefault: longword): longword;
begin
  if FSpPr <> Nil then
    Result := FSpPr.FillRGB(ADefault)
  else
    Result := ADefault;
end;

{ TXSSChartItemDataPoints }

function TXSSChartItemDataPoints.Add(ADp: TCT_DPt): TXSSChartItemDataPoint;
begin
  Result := TXSSChartItemDataPoint.Create(FData,ADp);

  inherited Add(Result);
end;

procedure TXSSChartItemDataPoints.Add(ADps: TCT_DPtXpgList);
var
  i: integer;
begin
  if ADPs <> Nil then begin
    for i := 0 to ADPs.Count - 1 do
      Add(ADPs[i]);
  end;
end;

procedure TXSSChartItemDataPoints.Add(AdLbls: TCT_DLbls; AVals: TDynSingleArray);
var
  i   : integer;
  DP  : TXSSChartItemDataPoint;
  DLbl: TCT_DLbl;
  Fnt : TXc12Font;
begin
  if AdLbls <> Nil then begin
    Fnt := FData.Font;

    if AdLbls.Group_DLbls.EG_DLblSShared <> Nil then
      Fnt := FData.GetChartFont(AdLbls.Group_DLbls.EG_DLblSShared.TxPr);

    if (AdLbls.Group_DLbls.EG_DLblSShared <> Nil) and (AdLbls.Group_DLbls.EG_DLblSShared.ShowVal <> Nil) then begin
      for i := 0 to High(AVals) do begin
        DP := AddDp(i);
        DP.AddParas(FloatToStr(AVals[i]),Fnt);
      end;
    end;

    if AdLbls.DLblXpgList <> Nil then begin
      for i := 0 to AdLbls.DLblXpgList.Count - 1 do begin
        DLbl := AdLbls.DLblXpgList[i];
        if DLbl.Group_DLbl.Tx <> Nil then begin
          DP := AddDp(DLbl.Idx.Val);

//          DLbl.Group_DLbl.Tx.Create_Rich;
          if DLbl.Group_DLbl.Tx.Rich <> Nil then
            DP.AddParas(DLbl.Group_DLbl.Tx.Rich.Paras);
        end
        else if DLbl.Group_DLbl.EG_DLblShared.TxPr <> Nil then begin
          DP := Find(DLbl.Idx.Val);
          if (DP <> Nil) and (DP.Paras <> Nil) then
            DP.Paras.UpdateFont(FData.GetChartFont(DLbl.Group_DLbl.EG_DLblShared.TxPr));
        end
        else if (DLbl.Delete <> Nil) and DLbl.Delete.Val and (DLbl.Idx.Val < Count) then
          Delete(DLbl.Idx.Val);
      end;
    end;
  end;
end;

function TXSSChartItemDataPoints.AddDp(AIndex: integer): TXSSChartItemDataPoint;
begin
  Result := Find(AIndex);
  if Result = Nil then begin
    Result := Add;
    Result.Index := AIndex;
  end;
end;

function TXSSChartItemDataPoints.Add: TXSSChartItemDataPoint;
begin
  Result := TXSSChartItemDataPoint.Create(FData);

  inherited Add(Result);
end;

constructor TXSSChartItemDataPoints.Create(AData: TXSSChartData);
begin
  inherited Create;

  FData := AData;
end;

function TXSSChartItemDataPoints.Find(AIndex: integer): TXSSChartItemDataPoint;
var
  i: integer;
begin
  for i := 0 to Count - 1 do begin
    if Items[i].Index = AIndex then begin
      Result := Items[i];
      Exit;
    end;
  end;
  Result := Nil;
end;

function TXSSChartItemDataPoints.GetItems(Index: integer): TXSSChartItemDataPoint;
begin
  Result := TXSSChartItemDataPoint(inherited Items[Index]);
end;

{ TXSSChartCharRun }

destructor TXSSChartCharRun.Destroy;
begin

  inherited;
end;

{ TXSSChartCharRuns }

function TXSSChartCharRuns.Add: TXSSChartCharRun;
begin
  Result := TXSSChartCharRun.Create;

  inherited Add(Result);
end;

constructor TXSSChartCharRuns.Create;
begin
  inherited Create;
end;

function TXSSChartCharRuns.GetItems(Index: integer): TXSSChartCharRun;
begin
  Result := TXSSChartCharRun(inherited Items[Index]);
end;

{ TXSSChartParagraph }

constructor TXSSChartParagraph.Create(AData: TXSSChartData; AParas: TCT_TextParagraphXpgList; ADefFont: TXc12Font);
begin
  FData := AData;
  FDefFont := ADefFont;

  FRuns := TXSSChartCharRuns.Create;

  DoParas(AParas);
end;

constructor TXSSChartParagraph.Create(AData: TXSSChartData; AText: AxUCString; AFont: TXc12Font);
var
  Run: TXSSChartCharRun;
begin
  FData := AData;

  FRuns := TXSSChartCharRuns.Create;

  Run := FRuns.Add;
  Run.Text := AText;
  Run.Font := AFont;
end;

destructor TXSSChartParagraph.Destroy;
begin
  FRuns.Free;

  inherited;
end;

procedure TXSSChartParagraph.DoCharRun(ACharRun: TCT_RegularTextRun);
var
  Run: TXSSChartCharRun;
begin
  Run := FRuns.Add;
  Run.Text := ACharRun.T;

  if ACharRun.RPr <> Nil then
    Run.Font := FData.GetChartFont(ACharRun.Rpr,FDefFont)
  else
    Run.Font := FDefFont;
end;

procedure TXSSChartParagraph.DoPara(APara: TCT_TextParagraph);
var
  i: integer;
begin
  for i := 0 to APara.TextRuns.Count - 1 do begin
    if APara.TextRuns[i].Run <> Nil then
      DoCharRun(APara.TextRuns[i].Run);
  end;
end;

procedure TXSSChartParagraph.DoParas(AParas: TCT_TextParagraphXpgList);
var
  i: integer;
begin
  for i := 0 to AParas.Count - 1 do
    DoPara(AParas[i]);
end;

function TXSSChartParagraph.GetPlainText: AxUCString;
var
  i: integer;
begin
  Result := '';

  for i := 0 to FRuns.Count - 1 do
    Result := Result + FRuns[i].Text;
end;

procedure TXSSChartParagraph.SetPlainText(const Value: AxUCString);
var
  Run: TXSSChartCharRun;
begin
  while FRuns.Count > 1 do
    FRuns.Delete(FRuns.Count - 1);

  if FRuns.Count < 1 then
    Run := FRuns.Add
  else
    Run := FRuns[0];

  Run.Text := Value;
  if Run.Font = Nil then
    Run.Font := FData.Font;
end;

procedure TXSSChartParagraph.UpdateFont(AFont: TXc12Font);
begin
  if FRuns.Count > 0 then
    FRuns[0].Font := AFont;
end;

{ TXSSChartItemTitle }

procedure TXSSChartItemTitle.Build(ARect: TXSSGeometryRect; ATitle: TCT_Title);
var
  i   : integer;
  Rich: TXSSGeometryRichText;
begin
  if ATitle <> Nil then begin
    ARect.Style.Color := XSS_CHART_DEF_COLOR;
    ARect.Style.LineColor := XLS_COLOR_NONE;
    ARect.Style.LineWidth := XSS_CHART_DEF_LINEWIDTH;

    ApplySpPr(ATitle.SpPr,ARect.Style);

    Rich := ARect.AddRichText;
    Rich.Pt.X := FData.MapL(FData.DefPadding) + ARect.Pt1.X;
    Rich.Pt.Y := FData.MapL(FData.DefPadding) + ARect.Pt1.Y;
    Rich.HorizAlign := chaLeft;
    Rich.VertAlign := cvaTop;

    for i := 0 to FPara.Runs.Count - 1 do
      Rich.Add(FPara.Runs[i].Text,FPara.Runs[i].Font);

    ARect.CenterHoriz(1);
  end;
end;

constructor TXSSChartItemTitle.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);
end;

destructor TXSSChartItemTitle.Destroy;
begin
  if FPara <> Nil then
    FPara.Free;

  inherited;
end;

function TXSSChartItemTitle.GetHeight: double;
begin
  Result := FData.DefPadding + FHeight + FData.DefPadding;
end;

function TXSSChartItemTitle.GetMargs: double;
begin
  Result := FData.DefMarg;
end;

function TXSSChartItemTitle.GetWidth: double;
begin
  Result := FData.DefPadding + FWidth + FData.DefPadding;
end;

procedure TXSSChartItemTitle.PreBuild(ATitle: TCT_Title);
begin
  if (ATitle.Tx <> Nil) and (ATitle.Tx.Rich <> Nil) and (ATitle.Tx.Rich.Paras <> Nil) then
    FPara := TXSSChartParagraph.Create(FData,ATitle.Tx.Rich.Paras,FData.TitleFont)
  else
    FPara := TXSSChartParagraph.Create(FData,'Chart Title',FData.TitleFont);

  ParaSize(FGDI,FPara,FWidth,FHeight);
end;

{ TXSSChartItemValAxis }

procedure TXSSChartItemValAxis.Build(ARect: TXSSGeometryRect; AValAxis: TCT_ValAx);
var
  V: double;
  S: AxUCString;
begin
  FTextWidth := 0;

  FFont := FData.GetChartFont(AValAxis.Shared.TxPr);
  FGDI.SetFont(FFont);

  if FStepSize <> 0 then begin
    V := FScaleMin;
    while V <= FScaleMax do begin
      S := Format('%.0f',[V]);

      FTextWidth := Max(FTextWidth,FGDI.TextWidthF(S));
      FTextHeight := Max(FTextHeight,FGDI.TextHeightF(S));

      V := V + FStepSize;
    end;
  end;

  ARect.Pt1.X := ARect.Pt1.X + FData.MapL(FTextWidth) + FData.MapL(FData.DefPadding * 2);

  FVisible := True;
end;

procedure TXSSChartItemValAxis.CalcScale(AMinVal,AMaxVal: double; ATicks: integer);
var
  Range   : double;
  TempStep: double;
  Mag     : double;
  MagPow  : double;
  MagMsd  : double;
begin
  FAxisTicks := ATicks;
  Range := AMaxVal - AMinVal;
  if Range = 0 then
    Range := 100;
  if FAxisTicks < 2 then
    FAxisTicks := 2
  else if FAxisTicks > 2 then
    Dec(FAxisTicks,2);
  // Get raw step value
  TempStep := Range / FAxisTicks;
  // Calculate pretty step value
  Mag := Floor(Log10(TempStep));
  MagPow := Power(10,Mag);
  MagMsd := Round((TempStep / MagPow + 0.5));
  FStepSize := MagMsd * MagPow;

  FScaleMin := FStepSize * Floor(AMinVal / FStepSize);
  FScaleMax  := FStepSize * Ceil((AMaxVal / FStepSize));

  if FScaleMin > 0 then
    FScaleMin := 0;

  Inc(FAxisTicks,2);

  if (FScaleMin < 0) and (FScaleMax > 0) then
    FScaleOrigo := 0
  else
    FScaleOrigo := FScaleMin;
end;

constructor TXSSChartItemValAxis.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);

  FAxisTicks := 2;
end;

function TXSSChartItemValAxis.ScaleRange: double;
begin
  Result := FScaleMax - FScaleMin;
end;

{ TXSSChartItemAxis }

constructor TXSSChartItemAxis.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);

  FVisible := False;

  FFont := FData.Font;
end;

{ TXSSChartItemCatAxis }

procedure TXSSChartItemCatAxis.Build(ARect: TXSSGeometryRect; ACount: integer; ACats: TDynStringArray; ATxPr: TCT_TextBody; ATextBetweenTicks: boolean);
var
  i   : integer;
  W   : double;
  X,Y : double;
  Line: TXSSGeometryLine;
  Txt : TXSSGeometryText;
begin
  FFont := FData.GetChartFont(ATxPr);
  FGDI.SetFont(FFont);

  FTextHeight := FGDI.TextHeightF('8');

  ARect.Pt2.Y := ARect.Pt2.Y - FData.MapT(FTextHeight + FData.DefPadding);

  if ACount > 0 then begin
    if ATextBetweenTicks or (ACount = 1) then
      W := ARect.Width / ACount
    else begin
      W := ARect.Width / (ACount - 1);
      Dec(ACount);
    end;
    X := ARect.Pt1.X;

    for i := 0 to ACount do begin
      Y := ARect.Pt2.Y + FData.MapT(FData.DefPadding);

      Line := ARect.AddLine;
      Line.SetSize(X,ARect.Pt2.Y,X,Y);

      if not (ATextBetweenTicks) or (i < ACount) then begin
        Txt := ARect.AddText;
        Txt.Font := FFont;
        if ATextBetweenTicks then
          Txt.Pt.X := X + (W / 2)
        else
          Txt.Pt.X := X;;
        Txt.Pt.Y := Y;
        Txt.VertAlign := cvaTop;
        Txt.HorizAlign := chaCenter;
        if i <= High(ACats) then
          Txt.Text := ACats[i]
        else
          Txt.Text := IntToStr(i + 1);
      end;

      X := X + W;
    end;
  end;

  FVisible := True;
end;

constructor TXSSChartItemCatAxis.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);
end;

{ TXSSChartItemSeriesChart }

procedure TXSSChartItemSeriesChart.Clear;
begin
  FSeries.Clear;
end;

constructor TXSSChartItemSeriesChart.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create(AGDI,AData);

  FValAxis := AValAxis;

  FSeries := TXSSChartItemChartSeries.Create(FGDI,FData,FValAxis,ACatAxis);
end;

destructor TXSSChartItemSeriesChart.Destroy;
begin
  FSeries.Free;

  inherited;
end;

{ TXSSChartItemLineChart }

procedure TXSSChartItemLineChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ALineChart: TCT_LineChart);
begin
  FSeries.Build(ARect,ALineChart.Shared.Series,(ALineChart.Marker <> Nil) and ALineChart.Marker.Val);
end;

constructor TXSSChartItemLineChart.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create(AGDI,AData,AValAxis,ACatAxis);
end;

procedure TXSSChartItemLineChart.PreBuild(ALegend: TXSSChartItemLegend; ALineChart: TCT_LineChart);
var
  i: integer;
begin
  FSeries.PreBuild(ALineChart.Shared.Series);

  if ALegend <> Nil then begin
    for i := 0 to FSeries.Count - 1 do
      ALegend.Add(i,FSeries[i].Color,FSeries[i].Name);
  end;
end;

{ TXSSChartItemPieChart }

procedure TXSSChartItemPieChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; APieChart: TCT_PieChart);
var
  i   : integer;
  V   : double;
  Cats: TDynStringArray;
begin
  if APieChart.FirstSliceAng <> Nil then
    V := APieChart.FirstSliceAng.Val
  else
    V := 0;

  if (APieChart.Shared.Ser.Cat <> Nil) and (APieChart.Shared.Ser.Cat.StrRef <> Nil) and (APieChart.Shared.Ser.Cat.StrRef.RCells <> Nil) then begin
    APieChart.Shared.Ser.Cat.StrRef.RCells.GetStrArray(Cats,XSS_CHART_MAXDATAPOINTS);

    for i := 0 to Min(High(Cats),FSerie.Count - 1) do
      ALegend[i].Text := Cats[i];
  end;

  FSerie.Build(ARect,APieChart.Shared.Ser,V,ALegend);
end;

constructor TXSSChartItemPieChart.Create(AGDI: TAXWGDI; AData: TXSSChartData);
begin
  inherited Create(AGDI,AData);

  FSerie := TXSSChartItemChartSerie.Create(Nil,AGDI,FData);
end;

destructor TXSSChartItemPieChart.Destroy;
begin
  FSerie.Free;

  inherited;
end;

procedure TXSSChartItemPieChart.PreBuild(ALegend: TXSSChartItemLegend; APieChart: TCT_PieChart);
var
  i: integer;
begin
  FSerie.PreBuild(0,APieChart.Shared.Ser.Val.NumRef.RCells,Nil,APieChart.Shared.Ser.Shared,True);

  if ALegend <> Nil then begin
    for i := 0 to FSerie.Count - 1 do
      ALegend.Add(i,$FFFFFF,IntToStr(i + 1));
  end;
end;

{ TXSSChartItemValAxises }

function TXSSChartItemValAxises.Add(AGDI: TAXWGDI; AData: TXSSChartData): TXSSChartItemValAxis;
begin
  Result := TXSSChartItemValAxis.Create(AGDI,AData);

  inherited Add(Result);
end;

constructor TXSSChartItemValAxises.Create;
begin
  inherited Create;
end;

function TXSSChartItemValAxises.First: TXSSChartItemValAxis;
begin
  if Count > 0 then
    Result := Items[0]
  else
    Result := Nil;
end;

function TXSSChartItemValAxises.GetItems(Index: integer): TXSSChartItemValAxis;
begin
  Result := TXSSChartItemValAxis(inherited Items[Index]);
end;

{ TXSSChartItemScatterChart }

procedure TXSSChartItemScatterChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; AScatterChart: TCT_ScatterChart);
begin
  DoCatAxis(ARect);

  FSeries.Build(ARect,FXSeries[0],FXAxis.ScaleRange,AScatterChart.Series);
end;

procedure TXSSChartItemScatterChart.PreBuild(ALegend: TXSSChartItemLegend; AScatterChart: TCT_ScatterChart);
var
  Ser: TXSSChartItemChartSerie;
begin
  FSeries.PreBuild(AScatterChart.Series);

  FXSeries.Clear;

  Ser := FXSeries.Add;
  Ser.PreBuild(0,AScatterChart.Series[0].XVal.NumRef.RCells,Nil,AScatterChart.Series[0].Shared,False);

  FXSeries.CalcMinMax(stbgStandard);

  FXAxis.CalcScale(Ser.MinVal,Ser.MaxVal * 1.05,6);
  Ser.MaxVal := FXAxis.ScaleMax;

  DoLegend(ALegend);
end;

{ TXSSChartItemXYChart }

constructor TXSSChartItemXYChart.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxises: TXSSChartItemValAxises; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create(AGDI,AData,AValAxises[0],Nil);

  FXAxis := TXSSChartItemValAxis.Create(FGDI,FData);
  FXSeries := TXSSChartItemChartSeries.Create(FGDI,FData,FXAxis,ACatAxis);
end;

destructor TXSSChartItemXYChart.Destroy;
begin
  FXAxis.Free;
  FXSeries.Free;

  inherited;
end;

procedure TXSSChartItemXYChart.DoCatAxis(ARect: TXSSGeometryRect);
var
  i  : integer;
  v,d: double;
  Arr: TDynStringArray;
begin
  d := FXAxis.ScaleRange / FXSeries[0].Count;
  v := FXAxis.FScaleMin;
  SetLength(Arr,FXSeries[0].Count + 1);

  for i := 0 to FXSeries[0].Count do begin
    Arr[i] := Format('%.0f',[v]);
    v := v + d;
  end;

  FXSeries.CatAxis.Build(ARect,Length(Arr),Arr,Nil,False);
end;

procedure TXSSChartItemXYChart.DoLegend(ALegend: TXSSChartItemLegend);
var
  i: integer;
begin
  if ALegend <> Nil then begin
    for i := 0 to FSeries.Count - 1 do
      ALegend.Add(i,FSeries[i].Color,FSeries[i].Name);
  end;
end;

{ TXSSChartItemBubbleChart }

procedure TXSSChartItemBubbleChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ABubbleChart: TCT_BubbleChart);
begin
  DoCatAxis(ARect);

  FSeries.Build(ARect,FXSeries[0],FXAxis.ScaleRange,ABubbleChart.Series);
end;

procedure TXSSChartItemBubbleChart.PreBuild(ALegend: TXSSChartItemLegend; ABubbleChart: TCT_BubbleChart);
var
  Ser: TXSSChartItemChartSerie;
begin
  FSeries.PreBuild(ABubbleChart.Series);

  FXSeries.Clear;

  Ser := FXSeries.Add;
  Ser.PreBuild(0,ABubbleChart.Series[0].XVal.NumRef.RCells,Nil,ABubbleChart.Series[0].Shared,False);

  FXSeries.CalcMinMax(stbgStandard);

  FXAxis.CalcScale(Ser.MinVal,Ser.MaxVal * 1.05,6);
  Ser.MaxVal := FXAxis.ScaleMax;

  DoLegend(ALegend);
end;

procedure TXSSChartItemChartSeries.PreBuild(ABubbleSeries: TCT_BubbleSeries);
var
  i  : integer;
  Ser: TXSSChartItemChartSerie;
begin
  Clear;

  for i := 0 to ABubbleSeries.Count - 1 do begin
    Ser := Add;

    Ser.PreBuild(i,ABubbleSeries[i].YVal.NumRef.RCells,Nil,ABubbleSeries[i].Shared,False);
  end;

  CalcMinMax(stbgStandard);

  ValAxis.CalcScale(FMinVal,FMaxVal * 1.05,6);
end;

procedure TXSSChartItemChartSeries.Build(ARect: TXSSGeometryRect; AXSerie: TXSSChartItemChartSerie; AXScaleRange: double; ABubbleSeries: TCT_BubbleSeries);
var
  i: integer;
begin
  AddGridlines(ARect);

  for i := 0 to ABubbleSeries.Count - 1 do
    Items[i].Build(ARect,AXSerie,AXScaleRange,ABubbleSeries[i]);
end;

{ TXSSChartItemRadarChart }

procedure TXSSChartItemRadarChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; ARadarChart: TCT_RadarChart; AValAxisFont,ACatAxFont: TCT_TextBody);
begin
  FSeries.Build(ARect,ARadarChart.Series,ARadarChart.RadarStyle,AValAxisFont,ACatAxFont);
end;

constructor TXSSChartItemRadarChart.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create(AGDI,AData,AValAxis,ACatAxis);

  FData.ShowValAxis := False;
end;

procedure TXSSChartItemRadarChart.PreBuild(ALegend: TXSSChartItemLegend; ARadarChart: TCT_RadarChart);
var
  i: integer;
begin
  FSeries.PreBuild(ARadarChart.Series);

  if ALegend <> Nil then begin
    for i := 0 to FSeries.Count - 1 do
      ALegend.Add(i,FSeries[i].Color,FSeries[i].Name);
  end;
end;

procedure TXSSChartItemChartSeries.Build(ARect: TXSSGeometryRect; ARadarSeries: TCT_RadarSeries; ARadarStyle: TCT_RadarStyle; AValAxFont,ACatAxFont: TCT_TextBody);
var
  i    : integer;
  V    : double;
  Step : double;
  r    : double;
  X,Y  : double;
  S    : AxUCString;
  PLine: TXSSGeometryPolyline;
  Line : TXSSGeometryLine;
  Vals : TDynSingleArray;
  Font : TXc12Font;
begin
  if Count < 1 then
    Exit;

  for i := 0 to ARadarSeries.Count - 1 do
    Items[i].Build(ARect,ARadarSeries[i],ARadarStyle);

  Vals := Items[0].Vals;

  V := Pi * 2 - Pi / 2;
  r := (Min(FData.PtWidth,FData.PtHeight) / 2) * 0.65;

  Step := FValAxis.ScaleMin + FValAxis.StepSize;

  Font := FData.GetChartFont(AValAxFont);
  ARect.AddText(Font,cvaCenter,chaRight,ARect.MidX - FData.MapL(Font.Size),ARect.MidY,Format('%.0f',[FValAxis.ScaleMin]));

  while Step <= FValAxis.ScaleMax do begin
    PLine := ARect.AddPolyline;
    for i := 0 to High(Vals) do begin
      X := Cos(V) * FData.MapL(r * (Step / FValAxis.ScaleMax)) + ARect.MidX;
      Y := Sin(V) * FData.MapT(r * (Step / FValAxis.ScaleMax)) + ARect.MidY;

      PLine.Add(X,Y);

      if i = 0 then
        ARect.AddText(FData.GetChartFont(AValAxFont),cvaCenter,chaRight,X - FData.MapL(Font.Size),Y,Format('%.0f',[Step]));

      V := V + (Pi * 2) / Length(Vals);
    end;
    PLine.Close;

    Step := Step + FValAxis.StepSize;
  end;

  V := Pi * 2 - Pi / 2;
  for i := 0 to High(Vals) do begin
    X := Cos(V) * FData.MapL(r) + ARect.MidX;
    Y := Sin(V) * FData.MapT(r) + ARect.MidY;

    Line := ARect.AddLine;
    Line.Pt1.X := ARect.MidX;
    Line.Pt1.Y := ARect.MidY;
    Line.Pt2.X := X;
    Line.Pt2.Y := Y;

    if FCatAxis <> Nil then begin
      if Length(FCats) > 0 then begin
        if i <= High(FCats) then
          S := FCats[i];
      end
      else
        S := IntToStr(i + 1);

      Font := FData.GetChartFont(ACatAxFont);
      X := Cos(V) * FData.MapL(r + Font.Size) + ARect.MidX;
      Y := Sin(V) * FData.MapT(r + Font.Size) + ARect.MidY;

      ARect.AddText(FData.GetChartFont(ACatAxFont),cvaCenter,chaCenter,X,Y,S);
    end;

    V := V + (Pi * 2) / Length(Vals);
  end;
end;

procedure TXSSChartItemChartSeries.Build(ARect: TXSSGeometryRect; AAreaSeries: TCT_AreaSeries);
var
  i: integer;
begin
  AddGridlines(ARect);

  for i := 0 to AAreaSeries.Count - 1 do
    Items[i].Build(ARect,AAreaSeries[i]);
end;

procedure TXSSChartItemChartSeries.PreBuild(ARadarSeries: TCT_RadarSeries);
var
  i  : integer;
  Ser: TXSSChartItemChartSerie;
begin
  Clear;

  for i := 0 to ARadarSeries.Count - 1 do begin
    Ser := Add;

    Ser.PreBuild(i,ARadarSeries[i].Val.NumRef.RCells,ARadarSeries[i].Cat,ARadarSeries[i].Shared,False);
  end;

  CalcMinMax(stbgStandard);

  FValAxis.CalcScale(FMinVal,FMaxVal,6);
end;

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; ARadarSer: TCT_RadarSer; ARadarStyle: TCT_RadarStyle);
var
  i     : integer;
  V     : double;
  r     : double;
  X,Y   : double;
  PLine : TXSSGeometryPolyline;
  LineCl: longword;
  LineW : double;
begin
  SetupDataPoints(ARadarSer.DPtXpgList,ARadarSer.DLbls);

  FData.ResetColor;

  LineCl := FColor;
  LineW := XSS_CHART_DEF_LINECHARTWIDTH;

  V := Pi * 2 - Pi / 2;
  r := (Min(FData.PtWidth,FData.PtHeight) / 2) * 0.65;

  if (ARadarStyle <> Nil) and (ARadarStyle.Val = strsFilled) then
    PLine := ARect.AddPolygon
  else
    PLine := ARect.AddPolyline;

  GetLineStyle(ARadarSer.Shared.SpPr,LineCl,LineW);

  PLine.Style.Color := LineCl;
  PLine.Style.LineColor := LineCl;
  PLine.Style.LineWidth := LineW;

  ApplySpPr(ARadarSer.Shared.SpPr,PLine.Style);

  for i := 0 to High(Vals) do begin
    X := Cos(V) * FData.MapL(r * (Vals[i] / FOwner.ValAxis.ScaleMax)) + ARect.MidX;
    Y := Sin(V) * FData.MapT(r * (Vals[i] / FOwner.ValAxis.ScaleMax)) + ARect.MidY;

    PLine.Add(X,Y);

    FData.ApplyDataPoint(i,PLine,X,Y,False);

    SetupMarker(PLine,ARadarSer.Marker,X,Y,LineCl,(ARadarStyle <> Nil) and (ARadarStyle.Val = strsMarker));

    V := V + (Pi * 2) / Length(Vals);
  end;
  PLine.Close;
end;

procedure TXSSChartItemChartSerie.Build(ARect: TXSSGeometryRect; AAreaSer: TCT_AreaSer);
var
  i     : integer;
  PGon  : TXSSGeometryPolygon;
  LineCl: longword;
  LineW : double;
  W     : double;
  H     : double;
  X     : double;
  Y     : double;
begin
  if Length(FVals) < 1 then
    Exit;

  LineCl := FColor;
  LineW := XSS_CHART_DEF_LINECHARTWIDTH;

  GetLineStyle(AAreaSer.Shared.SpPr,LineCl,LineW);

  SetupDataPoints(AAreaSer.DPtXpgList,AAreaSer.DLbls);

  H := ARect.Pt2.Y - ARect.Pt1.Y;

  W := (ARect.Pt2.X - ARect.Pt1.X) / (Length(FVals) - 1);
  X := ARect.Pt1.X;

  PGon := ARect.AddPolygon;
  PGon.Style.Color := LineCl;
  PGon.Style.LineColor := LineCl;
  PGon.Style.LineWidth := LineW;

  for i := 0 to High(FVals) do begin
    Y := ARect.Pt2.Y - FOwner.MapY(H,FVals[i]);

    PGon.Add(X,Y);

    FData.ApplyDataPoint(i,PGon,X,PGon[i].Y,FVals[i] >= FOwner.ValAxis.ScaleOrigo);

    X := X + W;
  end;

  PGon.Add(ARect.Pt2.X,ARect.Pt2.Y);
  PGon.Add(ARect.Pt1.X,ARect.Pt2.Y);
  PGon.Close;

  SetLength(FVals,0);
end;

{ TXSSChartItemAreaChart }

procedure TXSSChartItemAreaChart.Build(ARect: TXSSGeometryRect; AGeometry: TXSSGeometry; ALegend: TXSSChartItemLegend; AAreaChart: TCT_AreaChart);
begin
  FSeries.Build(ARect,AAreaChart.Shared.Series);
end;

constructor TXSSChartItemAreaChart.Create(AGDI: TAXWGDI; AData: TXSSChartData; AValAxis: TXSSChartItemValAxis; ACatAxis: TXSSChartItemCatAxis);
begin
  inherited Create(AGDI,AData,AValAxis,ACatAxis);
end;

procedure TXSSChartItemAreaChart.PreBuild(ALegend: TXSSChartItemLegend; AAreaChart: TCT_AreaChart);
var
  i: integer;
begin
  FSeries.PreBuild(AAreaChart.Shared.Series);

  if ALegend <> Nil then begin
    for i := 0 to FSeries.Count - 1 do
      ALegend.Add(i,FSeries[i].Color,FSeries[i].Name);
  end;
end;

end.

